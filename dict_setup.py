"""
One-time ECDICT downloader and SQLite builder.
Run directly:  python dict_setup.py
Or imported by the app for background init.
"""
import csv
import io
import sqlite3
import sys
import threading
from pathlib import Path

import requests

ECDICT_URL = "https://raw.githubusercontent.com/skywind3000/ECDICT/master/ecdict.csv"
DB_PATH = Path(__file__).resolve().parent / "ecdict.db"


def is_ready() -> bool:
    return DB_PATH.exists() and DB_PATH.stat().st_size > 1_000_000


def build_db(progress_cb=None):
    """Download CSV and build SQLite.  progress_cb(pct, msg) optional."""

    def prog(pct, msg):
        if progress_cb:
            progress_cb(pct, msg)
        else:
            print(f"[{pct:3d}%] {msg}", flush=True)

    prog(0, "Connecting to ECDICT …")
    try:
        resp = requests.get(ECDICT_URL, stream=True, timeout=30)
        resp.raise_for_status()
    except Exception as e:
        prog(-1, f"Download failed: {e}")
        return False

    total = int(resp.headers.get("content-length", 0))
    downloaded = bytearray()
    received = 0
    prog(0, "Downloading …")
    for chunk in resp.iter_content(65536):
        downloaded.extend(chunk)
        received += len(chunk)
        if total:
            prog(int(received / total * 50), f"Downloading … {received // 1024} KB")

    prog(50, "Building local database …")
    tmp = DB_PATH.with_suffix(".tmp")
    con = sqlite3.connect(str(tmp))
    con.execute("PRAGMA journal_mode=WAL")
    con.execute("PRAGMA synchronous=NORMAL")
    con.execute("""
        CREATE TABLE IF NOT EXISTS dict (
            word      TEXT PRIMARY KEY,
            phonetic  TEXT,
            translation TEXT,
            definition  TEXT,
            tag       TEXT,
            collins   INTEGER
        )
    """)

    text = downloaded.decode("utf-8", errors="replace")
    reader = csv.DictReader(io.StringIO(text))
    batch = []
    row_count = 0
    for row in reader:
        word = (row.get("word") or "").strip().lower()
        if not word:
            continue
        batch.append((
            word,
            (row.get("phonetic") or "").strip(),
            (row.get("translation") or "").strip(),
            (row.get("definition") or "").strip(),
            (row.get("tag") or "").strip(),
            int(row.get("collins") or 0),
        ))
        row_count += 1
        if len(batch) >= 5000:
            con.executemany(
                "INSERT OR REPLACE INTO dict VALUES (?,?,?,?,?,?)", batch
            )
            batch.clear()
            pct = 50 + min(48, int(row_count / 3500))
            prog(pct, f"Importing … {row_count:,} words")

    if batch:
        con.executemany(
            "INSERT OR REPLACE INTO dict VALUES (?,?,?,?,?,?)", batch
        )
    con.execute("CREATE INDEX IF NOT EXISTS idx_word ON dict(word)")
    con.commit()
    con.close()

    tmp.replace(DB_PATH)
    prog(100, f"Done — {row_count:,} words imported to {DB_PATH.name}")
    return True


def lookup(word: str):
    """
    Return dict or None.
    Keys: word, phonetic, translation, definition, tag
    tag may contain: cet4 cet6 ielts gre toefl etc.
    """
    if not is_ready():
        return None
    word = word.strip().lower()
    if not word:
        return None
    try:
        con = sqlite3.connect(str(DB_PATH), check_same_thread=False)
        cur = con.execute(
            "SELECT word, phonetic, translation, definition, tag FROM dict WHERE word=?",
            (word,),
        )
        row = cur.fetchone()
        con.close()
        if not row:
            return None
        return {
            "word": row[0],
            "phonetic": row[1],
            "translation": row[2],
            "definition": row[3],
            "tag": row[4],
        }
    except Exception:
        return None


def init_async(progress_cb=None):
    """Kick off build_db in a daemon thread."""
    t = threading.Thread(target=build_db, args=(progress_cb,), daemon=True)
    t.start()
    return t


if __name__ == "__main__":
    success = build_db()
    sys.exit(0 if success else 1)
