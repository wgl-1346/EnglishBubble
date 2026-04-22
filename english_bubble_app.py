import json
import io
import os
import random
import re
import threading
import time
import tkinter as tk
from pathlib import Path
from tkinter import messagebox, ttk

import requests

try:
    import pyttsx3
except Exception:
    pyttsx3 = None

try:
    from deep_translator import GoogleTranslator
except Exception:
    GoogleTranslator = None

try:
    from PIL import Image, ImageDraw, ImageFilter, ImageFont, ImageTk
except Exception:
    Image = None
    ImageDraw = None
    ImageFilter = None
    ImageFont = None
    ImageTk = None

try:
    from pypinyin import pinyin, Style
except Exception:
    pinyin = None
    Style = None

try:
    import dict_setup as _dict_setup
except Exception:
    _dict_setup = None


APP_DIR = Path(__file__).resolve().parent
DATA_FILE = APP_DIR / "vocab_data.json"
CONFIG_FILE = APP_DIR / "app_config.json"


class VocabStore:
    def __init__(self, file_path: Path):
        self.file_path = file_path
        self.words = []
        self.load()

    def load(self):
        if self.file_path.exists():
            try:
                self.words = json.loads(self.file_path.read_text(encoding="utf-8"))
                if not isinstance(self.words, list):
                    self.words = []
            except Exception:
                self.words = []
        self._normalize_words()

    def save(self):
        self.file_path.write_text(
            json.dumps(self.words, ensure_ascii=False, indent=2), encoding="utf-8"
        )

    def _normalize_words(self):
        now = int(time.time())
        normalized = []
        for item in self.words:
            if not isinstance(item, dict):
                continue
            english = str(item.get("english", "")).strip()
            chinese = str(item.get("chinese", "")).strip()
            pronunciation = str(item.get("pronunciation", "")).strip()
            example = str(item.get("example", "")).strip()
            if not english or not chinese:
                continue
            normalized.append(
                {
                    "english": english,
                    "chinese": chinese,
                    "pronunciation": pronunciation,
                    "example": example,
                    "streak": int(item.get("streak", 0)),
                    "hard_count": int(item.get("hard_count", 0)),
                    "review_count": int(item.get("review_count", 0)),
                    "next_review_ts": int(item.get("next_review_ts", now)),
                }
            )
        self.words = normalized

    def add_word(self, english: str, chinese: str, pronunciation: str = "", example: str = ""):
        english = english.strip()
        chinese = chinese.strip()
        pronunciation = pronunciation.strip()
        example = example.strip()
        if not english or not chinese:
            return False

        for item in self.words:
            if item.get("english", "").strip().lower() == english.lower():
                item["chinese"] = chinese
                item["pronunciation"] = pronunciation
                item["example"] = example
                item["next_review_ts"] = int(time.time())
                self.save()
                return True

        self.words.append(
            {
                "english": english,
                "chinese": chinese,
                "pronunciation": pronunciation,
                "example": example,
                "streak": 0,
                "hard_count": 0,
                "review_count": 0,
                "next_review_ts": int(time.time()),
            }
        )
        self.save()
        return True

    def mark_known(self, idx: int):
        if idx < 0 or idx >= len(self.words):
            return
        item = self.words[idx]
        item["review_count"] = int(item.get("review_count", 0)) + 1
        item["streak"] = int(item.get("streak", 0)) + 1
        intervals = [1, 3, 7, 15, 30]
        days = intervals[min(item["streak"] - 1, len(intervals) - 1)]
        item["next_review_ts"] = int(time.time() + days * 24 * 3600)
        self.save()

    def mark_hard(self, idx: int):
        if idx < 0 or idx >= len(self.words):
            return
        item = self.words[idx]
        item["review_count"] = int(item.get("review_count", 0)) + 1
        item["hard_count"] = int(item.get("hard_count", 0)) + 1
        item["streak"] = 0
        item["next_review_ts"] = int(time.time())
        self.save()


class SpeechEngine:
    def __init__(self):
        self.engine = None
        self.lock = threading.Lock()
        if pyttsx3 is not None:
            try:
                self.engine = pyttsx3.init()
                self.engine.setProperty("rate", 160)
            except Exception:
                self.engine = None

    def available(self):
        return self.engine is not None

    def speak_async(self, text: str):
        if not self.engine or not text.strip():
            return

        def worker():
            with self.lock:
                self.engine.stop()
                self.engine.say(text)
                self.engine.runAndWait()

        threading.Thread(target=worker, daemon=True).start()


class TranslatorClient:
    def __init__(self):
        pass

    def query(self, text: str):
        text = text.strip()
        if not text:
            raise RuntimeError("输入不能为空")

        try:
            if self._contains_chinese(text):
                translated = self._translate_text(text, "zh-CN", "en")
                keyword = self._extract_english_keyword(translated)
                if self._is_single_english_word(translated):
                    phonetic, example = self._fetch_english_word_info(keyword)
                else:
                    phonetic, example = "N/A (phrase)", ""
                if not example:
                    example = self._build_default_example(translated)
                return (
                    f"English Translation: {translated}\n"
                    f"Pronunciation: {phonetic}\n"
                    f"Example: {example}"
                )

            # --- 优先查本地 ECDICT ---
            if self._is_single_english_word(text) and _dict_setup and _dict_setup.is_ready():
                local = _dict_setup.lookup(text)
                if local:
                    return self._format_local_result(local)

            translated = self._translate_text(text, "en", "zh-CN")
            # For English input, get Chinese pinyin and example instead of English phonetic
            pinyin_text, example = self._fetch_chinese_info(translated, text)
            if not example:
                example = self._build_chinese_example(translated)
            return (
                f"Chinese Meaning: {translated}\n"
                f"Example: {example}"
            )
        except Exception as e:
            raise RuntimeError(f"Translation failed: {str(e)}")

    def _contains_chinese(self, text: str):
        return re.search(r"[\u4e00-\u9fff]", text) is not None

    def _extract_english_keyword(self, text: str):
        tokens = re.findall(r"[A-Za-z][A-Za-z\-']*", text)
        if not tokens:
            return text.strip()
        return tokens[0]

    def _is_single_english_word(self, text: str):
        tokens = re.findall(r"[A-Za-z][A-Za-z\-']*", text.strip())
        return len(tokens) == 1 and len(tokens[0]) == len(text.strip())

    def _format_local_result(self, entry: dict) -> str:
        """Format an ECDICT entry into the standard output string."""
        word = entry.get("word", "")
        phonetic = entry.get("phonetic") or "N/A"
        tag = entry.get("tag") or ""

        # Build Chinese meaning from translation field
        raw_trans = entry.get("translation") or ""
        # ECDICT translation lines look like: "n. 苹果\nv. 苹果色的"
        # We take up to 3 lines and strip pos prefixes for display
        lines = [l.strip() for l in raw_trans.splitlines() if l.strip()][:3]
        chinese = "\n".join(lines) if lines else "N/A"

        # Build example sentence using definition field
        definition = entry.get("definition") or ""
        example_line = ""
        for dl in definition.splitlines():
            dl = dl.strip()
            if dl and len(dl) > 10:
                example_line = dl
                break
        if not example_line:
            example_line = self._build_chinese_example(chinese)

        # Level badge from tag
        level_map = {"cet4": "CET-4", "cet6": "CET-6", "ielts": "IELTS",
                     "gre": "GRE", "toefl": "TOEFL", "kaoyan": "考研"}
        badges = [level_map[t] for t in tag.lower().split() if t in level_map]
        level_str = "  [" + "/".join(badges) + "]" if badges else ""

        # Pinyin for top Chinese meaning
        first_cn = re.sub(r'^[a-z]+\.\s*', '', lines[0]) if lines else ""
        pron = self._get_chinese_pinyin(first_cn) if first_cn else "N/A"

        return (
            f"Chinese Meaning: {chinese}{level_str}\n"
            f"Pronunciation: {phonetic}\n"
            f"Example: {example_line}"
        )

    def _translate_text(self, text: str, source_lang: str, target_lang: str):
        """Translate text with MyMemory API and return original text on failure."""
        src = source_lang.split("-")[0]
        tgt = target_lang.split("-")[0]
        try:
            response = requests.get(
                "https://api.mymemory.translated.net/get",
                params={"q": text, "langpair": f"{src}|{tgt}"},
                timeout=10,
            )
            if response.status_code != 200:
                return text
            payload = response.json() or {}
            translated = (payload.get("responseData") or {}).get("translatedText", "")
            translated = str(translated).replace("&quot;", '"').strip()
            if not translated:
                return text
            return translated
        except Exception:
            return text

    def _fetch_english_word_info(self, word: str):
        word = word.strip()
        if not word:
            return "N/A", ""

        url = f"https://api.dictionaryapi.dev/api/v2/entries/en/{word}"
        try:
            response = requests.get(url, timeout=10)
            if response.status_code != 200:
                return "N/A", ""
            payload = response.json()
            if not isinstance(payload, list) or not payload:
                return "N/A", ""
            entry = payload[0]

            phonetic = self._pick_best_phonetic(entry)

            example = ""
            for meaning in entry.get("meanings", []):
                for definition in meaning.get("definitions", []):
                    sample = definition.get("example")
                    if sample:
                        example = sample
                        break
                if example:
                    break

            return phonetic, example
        except Exception:
            return "N/A", ""

    def _pick_best_phonetic(self, entry: dict):
        candidates = []
        if entry.get("phonetic"):
            candidates.append(str(entry.get("phonetic")))
        for ph in entry.get("phonetics", []):
            text = ph.get("text")
            if text:
                candidates.append(str(text))
        if not candidates:
            return "N/A"

        def score(value: str):
            s = 0
            if "/" in value or "[" in value:
                s += 2
            if "ˈ" in value or "ˌ" in value:
                s += 2
            if len(value) >= 4:
                s += 1
            return s

        best = sorted(candidates, key=lambda v: (score(v), len(v)), reverse=True)[0]
        return best

    def _build_default_example(self, word: str):
        if not word:
            return "I use this word in my daily English practice."
        return f"I often use the word '{word}' when I practice English."

    def _fetch_chinese_info(self, chinese_text: str, original_english: str):
        """Get pinyin for Chinese text parsed from English input"""
        pinyin_text = self._get_chinese_pinyin(chinese_text)
        # Don't fetch English example, let _build_chinese_example generate Chinese sentences
        return pinyin_text, ""

    def _get_chinese_pinyin(self, chinese_text: str):
        """Convert Chinese text to pinyin (e.g., 苹果 -> pingguo)"""
        if not pinyin or not Style:
            return "N/A"
        try:
            # Extract only Chinese characters
            chinese_chars = "".join(re.findall(r"[\u4e00-\u9fff]", chinese_text))
            if not chinese_chars:
                return "N/A"
            # Get pinyin with NORMAL style (without tone marks)
            result = pinyin(chinese_chars, style=Style.NORMAL, errors='default', heteronym=False)
            # Flatten and join: [['p'], ['i'], ['n'], ['g'], ['g', 'u', ['o']] -> pingguo
            pinyin_str = "".join([item[0] if isinstance(item, list) and item else str(item) for item in result])
            return pinyin_str if pinyin_str else "N/A"
        except Exception:
            return "N/A"

    def _build_chinese_example(self, word: str):
        """Build a Chinese example sentence for the given word"""
        if not word:
            return "这是学习英文时常见的词汇。"
        # Remove English words if any, keep only Chinese
        chinese_word = "".join(re.findall(r"[\u4e00-\u9fff]", word))
        if not chinese_word:
            chinese_word = word
        return f"他在学习中常常用到这个词'{chinese_word}'。"


class EnglishLearningApp:
    def __init__(self, root):
        self.root = root
        self.root.withdraw()
        self.store = VocabStore(DATA_FILE)
        self.translator = TranslatorClient()
        self.speech = SpeechEngine()

        self.main_window = None
        self.review_window = None
        self.query_window = None

        self.review_index = 0
        self.review_order = []
        self.review_mode = tk.StringVar(value="待复习")
        self.current_review_word_index = None
        self._drag_start_x = 0
        self._drag_start_y = 0
        self._did_drag = False
        self._query_drag_start_x = 0
        self._query_drag_start_y = 0
        self._query_restore_geometry = "253x187+220+120"
        self._query_is_maximized = False
        self._query_compact_size = (253, 187)
        self._query_result_size = (520, 420)
        self._result_hover = False
        self._result_need_x_scroll = False
        self._result_need_y_scroll = False
        self._result_hide_job = None
        self._review_drag_start_x = 0
        self._review_drag_start_y = 0
        self._review_restore_geometry = "680x430+260+140"
        self._review_is_maximized = False
        self._click_job = None
        self._clip_watch_active = False

        self._build_bubble()
        self._start_dict_init()

    def _start_dict_init(self):
        """If ECDICT not ready, start background download and show bubble tooltip."""
        if _dict_setup is None or _dict_setup.is_ready():
            return

        self._dict_status = "Downloading word database (first run)..."

        def progress(pct, msg):
            self._dict_status = f"[{pct}%] {msg}" if pct >= 0 else f"Dict error: {msg}"
            # Update status label if query window is open
            if hasattr(self, 'status_label') and self.query_window and self.query_window.winfo_exists():
                self.query_window.after(0, lambda m=self._dict_status: self.status_label.config(text=m))

        def on_done():
            self._dict_status = ""
            if _dict_setup.is_ready():
                if hasattr(self, 'status_label') and self.query_window and self.query_window.winfo_exists():
                    self.query_window.after(0, lambda: self.status_label.config(
                        text="Word database ready", fg="#7ec8a0"
                    ))
                    self.query_window.after(3000, lambda: self.status_label.config(text="", fg="#b9d3e6"))

        def worker():
            _dict_setup.build_db(progress_cb=progress)
            on_done()

        threading.Thread(target=worker, daemon=True).start()

    def _build_bubble(self):
        self.bubble = tk.Toplevel(self.root)
        self.bubble.overrideredirect(True)
        self.bubble.attributes("-topmost", True)
        self.bubble.attributes("-alpha", 0.95)
        bubble_size = 96
        self.bubble.geometry(f"{bubble_size}x{bubble_size}+1200+120")
        transparent_bg = "#002b45"
        self.bubble.configure(bg=transparent_bg)

        try:
            self.bubble.wm_attributes("-transparentcolor", transparent_bg)
        except tk.TclError:
            pass

        self.bubble_icon = self._create_bubble_icon(bubble_size)
        icon_label = tk.Label(self.bubble, image=self.bubble_icon, bg=transparent_bg, bd=0)
        icon_label.pack(fill="both", expand=True)

        icon_label.bind("<ButtonPress-1>", self._start_drag)
        icon_label.bind("<B1-Motion>", self._drag_bubble)
        icon_label.bind("<ButtonRelease-1>", self._on_left_release)
        icon_label.bind("<Double-Button-1>", self._on_double_click)
        icon_label.bind("<Button-3>", self._exit_app)

    def _create_bubble_icon(self, size: int):
        if not all([Image, ImageDraw, ImageFilter, ImageTk]):
            fallback = tk.PhotoImage(width=size, height=size)
            fallback.put("#00a8e8", to=(0, 0, size, size))
            return fallback

        scale = 8
        big = size * scale
        # Transparent background so LANCZOS anti-aliasing produces clean alpha edges
        img = Image.new("RGBA", (big, big), (0, 0, 0, 0))
        draw = ImageDraw.Draw(img)

        cx = cy = big // 2
        radius = int(big * 0.36)

        shadow = Image.new("RGBA", (big, big), (0, 0, 0, 0))
        shadow_draw = ImageDraw.Draw(shadow)
        shadow_box = (
            cx - radius + int(0.025 * big),
            cy - radius + int(0.05 * big),
            cx + radius + int(0.025 * big),
            cy + radius + int(0.05 * big),
        )
        shadow_draw.ellipse(shadow_box, fill=(0, 0, 0, 95))
        shadow = shadow.filter(ImageFilter.GaussianBlur(int(0.03 * big)))
        img.alpha_composite(shadow)

        for i in range(radius, 0, -1):
            t = i / radius
            r = int(0 + (36 * t))
            g = int(66 + (124 * t))
            b = int(138 + (117 * t))
            draw.ellipse((cx - i, cy - i, cx + i, cy + i), fill=(r, g, b, 255))

        border_w = int(0.018 * big)
        draw.ellipse(
            (cx - radius, cy - radius, cx + radius, cy + radius),
            outline=(205, 234, 255, 170),
            width=border_w,
        )
        draw.arc(
            (cx - radius, cy - radius, cx + radius, cy + radius),
            start=198,
            end=336,
            fill=(0, 70, 132, 145),
            width=border_w,
        )

        highlight = Image.new("RGBA", (big, big), (0, 0, 0, 0))
        hdraw = ImageDraw.Draw(highlight)
        hdraw.ellipse(
            (
                cx - int(radius * 0.88),
                cy - int(radius * 0.92),
                cx + int(radius * 0.38),
                cy - int(radius * 0.14),
            ),
            fill=(255, 255, 255, 145),
        )
        hdraw.ellipse(
            (
                cx - int(radius * 0.45),
                cy - int(radius * 0.72),
                cx - int(radius * 0.15),
                cy - int(radius * 0.43),
            ),
            fill=(255, 255, 255, 175),
        )
        highlight = highlight.filter(ImageFilter.GaussianBlur(int(0.016 * big)))
        img.alpha_composite(highlight)

        draw.text((cx, cy), "EN", anchor="mm", fill=(245, 252, 255, 250))

        badge_layer = Image.new("RGBA", (big, big), (0, 0, 0, 0))
        bdraw = ImageDraw.Draw(badge_layer)
        badge_r = int(radius * 0.28)
        left_cx = cx - int(radius * 0.88)
        right_cx = cx + int(radius * 0.88)
        badge_cy = cy
        badge_outline = max(2, int(0.012 * big))

        bdraw.ellipse(
            (left_cx - badge_r, badge_cy - badge_r, left_cx + badge_r, badge_cy + badge_r),
            fill=(239, 248, 255, 248),
            outline=(5, 94, 168, 255),
            width=badge_outline,
        )
        bdraw.ellipse(
            (right_cx - badge_r, badge_cy - badge_r, right_cx + badge_r, badge_cy + badge_r),
            fill=(239, 248, 255, 248),
            outline=(5, 94, 168, 255),
            width=badge_outline,
        )

        font = None
        if ImageFont is not None:
            try:
                font = ImageFont.truetype("arialbd.ttf", int(badge_r * 1.35))
            except Exception:
                try:
                    font = ImageFont.truetype("segoeuib.ttf", int(badge_r * 1.35))
                except Exception:
                    font = ImageFont.load_default()

        bdraw.text(
            (left_cx, badge_cy + int(0.008 * big)),
            "S",
            anchor="mm",
            fill=(8, 86, 150, 255),
            font=font,
            stroke_width=max(1, int(0.004 * big)),
            stroke_fill=(255, 255, 255, 220),
        )
        bdraw.text(
            (right_cx, badge_cy + int(0.008 * big)),
            "L",
            anchor="mm",
            fill=(8, 86, 150, 255),
            font=font,
            stroke_width=max(1, int(0.004 * big)),
            stroke_fill=(255, 255, 255, 220),
        )
        img.alpha_composite(badge_layer)

        # Downscale: LANCZOS produces smooth anti-aliased alpha channel
        icon = img.resize((size, size), Image.Resampling.LANCZOS)

        # Save as PNG (preserves full alpha) and reload as tk.PhotoImage
        # This gives true per-pixel transparency — no jagged edges from transparentcolor trick
        buf = io.BytesIO()
        icon.save(buf, format="PNG")
        buf.seek(0)
        return tk.PhotoImage(data=buf.getvalue())

    def _start_drag(self, event):
        self._drag_start_x = event.x
        self._drag_start_y = event.y
        self._did_drag = False

    def _on_left_release(self, _event):
        if self._did_drag:
            return
        # Use short delay so double-click can cancel single-click action
        bubble_w = self.bubble.winfo_width() or 96
        x = _event.x
        if self._click_job:
            self.root.after_cancel(self._click_job)
            self._click_job = None

        def do_single():
            self._click_job = None
            if x < bubble_w // 2:
                self._open_query_window()
            else:
                self._open_review_window()

        self._click_job = self.root.after(220, do_single)

    def _on_double_click(self, event):
        # Cancel pending single-click
        if self._click_job:
            self.root.after_cancel(self._click_job)
            self._click_job = None
        bubble_w = self.bubble.winfo_width() or 96
        if event.x < bubble_w // 2:
            self._start_screen_capture()

    def _start_screen_capture(self):
        """Double-click left: try OCR screen-region capture, fall back to clipboard watch."""
        try:
            import pytesseract
            from PIL import ImageGrab
            # Quick test if tesseract binary exists
            pytesseract.get_tesseract_version()
            self._show_region_overlay()
        except Exception:
            self._start_clipboard_watch()

    def _show_region_overlay(self):
        """Fullscreen transparent overlay for dragging a capture region."""
        overlay = tk.Toplevel(self.root)
        overlay.attributes("-fullscreen", True)
        overlay.attributes("-alpha", 0.25)
        overlay.attributes("-topmost", True)
        overlay.overrideredirect(True)
        overlay.configure(bg="#001a2e")

        canvas = tk.Canvas(overlay, cursor="crosshair", bg="#001a2e",
                           highlightthickness=0)
        canvas.pack(fill="both", expand=True)

        hint = tk.Label(overlay, text="Drag to select region  |  ESC = cancel",
                        bg="#163a54", fg="#eaf6ff",
                        font=("Segoe UI", 11), padx=12, pady=6)
        hint.place(relx=0.5, y=18, anchor="n")

        start = [0, 0]
        rect_id = [None]

        def on_press(e):
            start[0], start[1] = e.x_root, e.y_root

        def on_drag(e):
            if rect_id[0]:
                canvas.delete(rect_id[0])
            x1, y1 = min(start[0], e.x_root), min(start[1], e.y_root)
            x2, y2 = max(start[0], e.x_root), max(start[1], e.y_root)
            rect_id[0] = canvas.create_rectangle(
                x1, y1, x2, y2, outline="#00aaff", width=2,
                fill="#003a60", stipple="gray25"
            )

        def on_release(e):
            x1 = min(start[0], e.x_root)
            y1 = min(start[1], e.y_root)
            x2 = max(start[0], e.x_root)
            y2 = max(start[1], e.y_root)
            overlay.destroy()
            if x2 - x1 < 8 or y2 - y1 < 8:
                return
            self.root.after(120, lambda: self._ocr_region(x1, y1, x2, y2))

        canvas.bind("<ButtonPress-1>", on_press)
        canvas.bind("<B1-Motion>", on_drag)
        canvas.bind("<ButtonRelease-1>", on_release)
        canvas.bind("<Escape>", lambda e: overlay.destroy())
        canvas.focus_set()

    def _ocr_region(self, x1, y1, x2, y2):
        try:
            import pytesseract
            from PIL import ImageGrab
            img = ImageGrab.grab(bbox=(x1, y1, x2, y2))
            text = pytesseract.image_to_string(img, lang="eng").strip()
            text = " ".join(text.split())  # normalize whitespace
            if text:
                self._fill_query_and_search(text)
            else:
                messagebox.showwarning("Capture", "No text recognized. Try a larger region.")
        except Exception as exc:
            messagebox.showerror("OCR Error", str(exc))

    def _start_clipboard_watch(self):
        """Show a small indicator and watch clipboard for new text, then auto-search."""
        self._clip_watch_active = True

        # Small floating indicator near the bubble
        bx = self.bubble.winfo_x()
        by = self.bubble.winfo_y()
        ind = tk.Toplevel(self.root)
        ind.overrideredirect(True)
        ind.attributes("-topmost", True)
        ind.attributes("-alpha", 0.92)
        ind.geometry(f"260x36+{bx - 90}+{by - 48}")
        ind.configure(bg="#163a54")
        tk.Label(
            ind, text="📋  Select text & Ctrl+C  |  ESC cancel",
            bg="#163a54", fg="#eaf6ff", font=("Segoe UI", 9), padx=8
        ).pack(expand=True, fill="both")

        try:
            old_clip = self.root.clipboard_get()
        except Exception:
            old_clip = ""

        def cancel(e=None):
            self._clip_watch_active = False
            if ind.winfo_exists():
                ind.destroy()

        ind.bind("<Escape>", cancel)
        ind.bind("<Button-1>", cancel)
        ind.focus_set()

        def poll():
            if not self._clip_watch_active:
                return
            try:
                new_clip = self.root.clipboard_get()
            except Exception:
                new_clip = ""
            if new_clip.strip() and new_clip != old_clip:
                self._clip_watch_active = False
                if ind.winfo_exists():
                    ind.destroy()
                self._fill_query_and_search(new_clip.strip())
                return
            self.root.after(250, poll)

        self.root.after(250, poll)

    def _fill_query_and_search(self, text: str):
        """Open query window, fill text, and trigger search."""
        self._open_query_window()

        def fill():
            if hasattr(self, 'query_text') and self.query_window and self.query_window.winfo_exists():
                self.query_text.delete("1.0", "end")
                self.query_text.insert("1.0", text)
                self._start_query()

        self.root.after(150, fill)

    def _drag_bubble(self, event):
        x = self.bubble.winfo_x() + event.x - self._drag_start_x
        y = self.bubble.winfo_y() + event.y - self._drag_start_y
        self._did_drag = True
        w = self.bubble.winfo_width() or 88
        h = self.bubble.winfo_height() or 88
        self.bubble.geometry(f"{w}x{h}+{x}+{y}")

    def _exit_app(self, _event=None):
        self.root.quit()

    def _open_review_window(self):
        if self.review_window and self.review_window.winfo_exists():
            self.review_window.deiconify()
            self.review_window.lift()
            self.review_window.focus_force()
            return

        self.review_window = tk.Toplevel(self.root)
        self.review_window.geometry("680x430")
        self.review_window.attributes("-alpha", 0.9)
        self.review_window.overrideredirect(True)
        review_bg = "#0e2b40"
        panel_bg = "#163a54"
        card_bg = "#0f3047"
        font_fg = "#eaf6ff"
        muted_fg = "#b9d3e6"
        self.review_window.configure(bg=review_bg)

        shell = tk.Frame(self.review_window, bg=review_bg, bd=0, highlightthickness=1, highlightbackground="#35607e")
        shell.pack(fill="both", expand=True)

        title_bar = tk.Frame(shell, bg=panel_bg, height=34)
        title_bar.pack(fill="x")
        title_bar.pack_propagate(False)

        title_label = tk.Label(
            title_bar,
            text="单词背诵",
            bg=panel_bg,
            fg=font_fg,
            font=("Segoe UI", 10, "bold"),
        )
        title_label.pack(side="left", padx=12)

        tk.Button(
            title_bar,
            text="—",
            command=self._minimize_review_window,
            bg=panel_bg,
            fg=font_fg,
            activebackground="#2e5574",
            activeforeground="#ffffff",
            relief="flat",
            bd=0,
            width=4,
        ).pack(side="right")
        tk.Button(
            title_bar,
            text="□",
            command=self._toggle_review_window_maximize,
            bg=panel_bg,
            fg=font_fg,
            activebackground="#2e5574",
            activeforeground="#ffffff",
            relief="flat",
            bd=0,
            width=4,
        ).pack(side="right")
        tk.Button(
            title_bar,
            text="✕",
            command=self._close_review_window,
            bg=panel_bg,
            fg="#ffedf1",
            activebackground="#b3263f",
            activeforeground="#ffffff",
            relief="flat",
            bd=0,
            width=4,
        ).pack(side="right")

        title_bar.bind("<ButtonPress-1>", self._start_review_drag)
        title_bar.bind("<B1-Motion>", self._on_review_drag)
        title_label.bind("<ButtonPress-1>", self._start_review_drag)
        title_label.bind("<B1-Motion>", self._on_review_drag)

        self._rebuild_review_order(reset_index=True)

        outer = tk.Frame(shell, bg=review_bg, padx=14, pady=14)
        outer.pack(fill="both", expand=True)

        top_bar = tk.Frame(outer, bg=review_bg)
        top_bar.pack(fill="x", pady=(0, 8))
        tk.Label(top_bar, text="复习模式:", bg=review_bg, fg=font_fg, font=("Segoe UI", 10, "bold")).pack(side="left")
        mode_box = ttk.Combobox(
            top_bar,
            textvariable=self.review_mode,
            values=["待复习", "不熟", "全部"],
            width=8,
            state="readonly",
        )
        mode_box.pack(side="left", padx=6)
        mode_box.bind("<<ComboboxSelected>>", lambda _e: self._rebuild_review_order(reset_index=True))
        self.review_stats = tk.Label(top_bar, text="", bg=review_bg, fg=muted_fg, font=("Segoe UI", 10))
        self.review_stats.pack(side="right")

        self.card = tk.Frame(outer, bg=card_bg, padx=20, pady=20, highlightthickness=1, highlightbackground="#295372")
        self.card.pack(fill="both", expand=True)

        self.en_label = tk.Label(
            self.card,
            text="",
            bg=card_bg,
            fg=font_fg,
            font=("Segoe UI", 18, "bold"),
            wraplength=500,
        )
        self.en_label.pack(pady=(25, 10))

        self.cn_label = tk.Label(
            self.card,
            text="",
            bg=card_bg,
            fg=font_fg,
            font=("Segoe UI", 13),
            wraplength=620,
        )
        self.cn_label.pack(pady=(10, 20))

        self.meta_label = tk.Label(self.card, text="", bg=card_bg, fg=muted_fg, font=("Segoe UI", 10))
        self.meta_label.pack(pady=(0, 10))

        bar = tk.Frame(outer, bg=review_bg)
        bar.pack(fill="x", pady=(10, 0))

        btn_style = {
            "bg": "#2b7bbb",
            "fg": "#ffffff",
            "activebackground": "#3e8dca",
            "activeforeground": "#ffffff",
            "relief": "flat",
            "bd": 0,
            "padx": 10,
            "pady": 4,
        }
        tk.Button(bar, text="上一张", command=self._prev_card, **btn_style).pack(side="left", padx=4)
        tk.Button(bar, text="下一张", command=self._next_card, **btn_style).pack(side="left", padx=4)
        tk.Button(bar, text="打乱", command=self._shuffle_cards, **btn_style).pack(side="left", padx=4)
        tk.Button(bar, text="发音", command=self._speak_current_english, **btn_style).pack(side="left", padx=4)
        tk.Button(bar, text="记住了", command=self._mark_known_current, **btn_style).pack(side="right", padx=4)
        tk.Button(bar, text="不熟", command=self._mark_hard_current, **btn_style).pack(side="right", padx=4)
        tk.Button(bar, text="删除当前", command=self._delete_current_card, **btn_style).pack(side="right", padx=4)

        self._render_card()

    def _rebuild_review_order(self, reset_index=False):
        now = int(time.time())
        mode = self.review_mode.get()
        order = []
        for idx, item in enumerate(self.store.words):
            if mode == "不熟" and int(item.get("hard_count", 0)) <= 0:
                continue
            if mode == "待复习" and int(item.get("next_review_ts", now)) > now:
                continue
            order.append(idx)
        self.review_order = order
        if reset_index:
            self.review_index = 0

    def _render_card(self):
        if not self.store.words:
            self.en_label.config(text="暂无单词")
            self.cn_label.config(text="请去查询页面添加单词或句子")
            self.meta_label.config(text="")
            self.review_stats.config(text="总数 0")
            return

        if not self.review_order:
            self._rebuild_review_order(reset_index=True)
            if not self.review_order:
                self.en_label.config(text="当前模式下没有可复习卡片")
                self.cn_label.config(text="可切换到" + "全部" + "模式查看全部卡片")
                self.meta_label.config(text="")
                self.review_stats.config(text=f"总数 {len(self.store.words)} / 当前 0")
                self.current_review_word_index = None
                return

        self.review_index = max(0, min(self.review_index, len(self.review_order) - 1))
        word_idx = self.review_order[self.review_index]
        self.current_review_word_index = word_idx
        word = self.store.words[word_idx]
        now = int(time.time())
        due = int(word.get("next_review_ts", now))
        due_text = "已到期" if due <= now else f"{max(1, (due - now) // 3600)} 小时后"
        pronunciation = str(word.get("pronunciation", "")).strip()
        example = str(word.get("example", "")).strip()
        details = [word.get("chinese", "")]
        if pronunciation:
            details.append(f"Pronunciation: {pronunciation}")
        if example:
            details.append(f"Example: {example}")
        self.en_label.config(text=word.get("english", ""))
        self.cn_label.config(text="\n".join(details))
        self.meta_label.config(
            text=f"连续记住: {word.get('streak', 0)} | 不熟次数: {word.get('hard_count', 0)} | 下次复习: {due_text}"
        )
        self.review_stats.config(text=f"总数 {len(self.store.words)} / 当前 {len(self.review_order)}")

    def _prev_card(self):
        if not self.review_order:
            return
        self.review_index = (self.review_index - 1) % len(self.review_order)
        self._render_card()

    def _next_card(self):
        if not self.review_order:
            return
        self.review_index = (self.review_index + 1) % len(self.review_order)
        self._render_card()

    def _shuffle_cards(self):
        if not self.review_order:
            return
        random.shuffle(self.review_order)
        self.review_index = 0
        self._render_card()

    def _speak_current_english(self):
        if self.current_review_word_index is None:
            return
        if not self.speech.available():
            messagebox.showwarning("提示", "当前环境不可用发音引擎，请安装 pyttsx3")
            return
        text = self.store.words[self.current_review_word_index].get("english", "")
        self.speech.speak_async(text)

    def _mark_known_current(self):
        if self.current_review_word_index is None:
            return
        self.store.mark_known(self.current_review_word_index)
        self._rebuild_review_order(reset_index=False)
        self._render_card()

    def _mark_hard_current(self):
        if self.current_review_word_index is None:
            return
        self.store.mark_hard(self.current_review_word_index)
        self._rebuild_review_order(reset_index=False)
        self._render_card()

    def _delete_current_card(self):
        if self.current_review_word_index is None:
            return
        idx = self.current_review_word_index
        del self.store.words[idx]
        self.store.save()
        self._rebuild_review_order(reset_index=True)
        self._render_card()

    def _open_query_window(self):
        if self.query_window and self.query_window.winfo_exists():
            if not hasattr(self, "result_text"):
                self.query_window.destroy()
                self.query_window = None
            else:
                self.query_window.deiconify()
                self.query_window.lift()
                self.query_window.focus_force()
                return

        if self.query_window and self.query_window.winfo_exists():
            self.query_window.deiconify()
            self.query_window.lift()
            self.query_window.focus_force()
            return

        self.query_window = tk.Toplevel(self.root)
        self.query_window.geometry(f"{self._query_compact_size[0]}x{self._query_compact_size[1]}")
        self.query_window.minsize(self._query_compact_size[0], self._query_compact_size[1])
        self.query_window.attributes("-alpha", 0.9)
        self.query_window.overrideredirect(True)
        query_bg = "#0e2b40"
        panel_bg = "#163a54"
        field_bg = "#0f3047"
        font_fg = "#eaf6ff"
        muted_fg = "#b9d3e6"
        self.query_window.configure(bg=query_bg)

        shell = tk.Frame(self.query_window, bg=query_bg, bd=0, highlightthickness=1, highlightbackground="#35607e")
        shell.pack(fill="both", expand=True)

        title_bar = tk.Frame(shell, bg=panel_bg, height=34)
        title_bar.pack(fill="x")
        title_bar.pack_propagate(False)

        title_label = tk.Label(
            title_bar,
            text="Search",
            bg=panel_bg,
            fg=font_fg,
            font=("Segoe UI", 10, "bold"),
        )
        title_label.pack(side="left", padx=12)

        tk.Button(
            title_bar,
            text="—",
            command=self._minimize_query_window,
            bg=panel_bg,
            fg=font_fg,
            activebackground="#2e5574",
            activeforeground="#ffffff",
            relief="flat",
            bd=0,
            width=4,
        ).pack(side="right")
        tk.Button(
            title_bar,
            text="□",
            command=self._toggle_query_window_maximize,
            bg=panel_bg,
            fg=font_fg,
            activebackground="#2e5574",
            activeforeground="#ffffff",
            relief="flat",
            bd=0,
            width=4,
        ).pack(side="right")
        tk.Button(
            title_bar,
            text="✕",
            command=self._close_query_window,
            bg=panel_bg,
            fg="#ffedf1",
            activebackground="#b3263f",
            activeforeground="#ffffff",
            relief="flat",
            bd=0,
            width=4,
        ).pack(side="right")

        title_bar.bind("<ButtonPress-1>", self._start_query_drag)
        title_bar.bind("<B1-Motion>", self._on_query_drag)
        title_label.bind("<ButtonPress-1>", self._start_query_drag)
        title_label.bind("<B1-Motion>", self._on_query_drag)

        frame = tk.Frame(shell, bg=query_bg, padx=14, pady=14)
        frame.pack(fill="both", expand=True)

        self.query_text = tk.Text(
            frame,
            height=2,
            bg=field_bg,
            fg=font_fg,
            insertbackground=font_fg,
            relief="flat",
            bd=0,
            padx=10,
            pady=8,
        )
        self.query_text.pack(fill="x", pady=(4, 10))
        self.query_text.bind("<Return>", self._on_query_enter)

        query_row = tk.Frame(frame, bg=query_bg)
        query_row.pack(fill="x", pady=4)
        tk.Button(
            query_row,
            text="Search",
            command=self._start_query,
            bg="#2b7bbb",
            fg="#ffffff",
            activebackground="#3e8dca",
            activeforeground="#ffffff",
            relief="flat",
            padx=12,
            pady=4,
        ).pack(side="left")
        self.status_label = tk.Label(query_row, text="", bg=query_bg, fg=muted_fg, font=("Segoe UI", 10))
        self.status_label.pack(side="left", padx=10)
        # Show dict download progress if still running
        if _dict_setup and not _dict_setup.is_ready():
            init_msg = getattr(self, '_dict_status', 'Downloading word database...')
            self.status_label.config(text=init_msg)

        self.result_text = tk.Text(
            frame,
            height=8,
            bg="#081a28",
            fg="#ffffff",
            insertbackground=font_fg,
            relief="flat",
            bd=0,
            padx=10,
            pady=10,
            wrap="word",
            font=("Consolas", 10),
            selectbackground="#2d6a96",
            selectforeground="#ffffff",
        )
        self.result_wrap = tk.Frame(frame, bg=panel_bg, highlightthickness=1, highlightbackground="#295372")
        self.result_wrap.pack(fill="both", expand=True, pady=(4, 10))
        self.result_text.pack(in_=self.result_wrap, fill="both", expand=True)

        self.result_y_scroll = tk.Scrollbar(self.result_wrap, orient="vertical", command=self.result_text.yview)
        self.result_x_scroll = tk.Scrollbar(self.result_wrap, orient="horizontal", command=self.result_text.xview)
        self.result_text.configure(
            yscrollcommand=self._update_result_y_scroll,
            xscrollcommand=self._update_result_x_scroll,
        )
        self.result_text.tag_configure("result_content", foreground="#ffffff")

        for widget in (self.result_wrap, self.result_text, self.result_y_scroll, self.result_x_scroll):
            widget.bind("<Enter>", self._on_result_area_enter)
            widget.bind("<Leave>", self._on_result_area_leave)

        add_row = tk.Frame(frame, bg=query_bg)
        add_row.pack(fill="x")
        tk.Button(
            add_row,
            text="Add to Review",
            command=self._add_result_to_review,
            bg="#2b7bbb",
            fg="#ffffff",
            activebackground="#3e8dca",
            activeforeground="#ffffff",
            relief="flat",
            padx=14,
            pady=5,
        ).pack(side="right")

    def _start_review_drag(self, event):
        self._review_drag_start_x = event.x
        self._review_drag_start_y = event.y

    def _on_review_drag(self, event):
        if not (self.review_window and self.review_window.winfo_exists()):
            return
        if self._review_is_maximized:
            return
        x = self.review_window.winfo_x() + event.x - self._review_drag_start_x
        y = self.review_window.winfo_y() + event.y - self._review_drag_start_y
        w = self.review_window.winfo_width() or 680
        h = self.review_window.winfo_height() or 430
        self.review_window.geometry(f"{w}x{h}+{x}+{y}")

    def _toggle_review_window_maximize(self):
        if not (self.review_window and self.review_window.winfo_exists()):
            return
        if self._review_is_maximized:
            self.review_window.geometry(self._review_restore_geometry)
            self._review_is_maximized = False
        else:
            self._review_restore_geometry = self.review_window.geometry()
            sw = self.review_window.winfo_screenwidth()
            sh = self.review_window.winfo_screenheight()
            self.review_window.geometry(f"{sw}x{sh}+0+0")
            self._review_is_maximized = True

    def _minimize_review_window(self):
        if self.review_window and self.review_window.winfo_exists():
            self.review_window.iconify()

    def _close_review_window(self):
        if self.review_window and self.review_window.winfo_exists():
            self.review_window.destroy()
            self.review_window = None

    def _start_query_drag(self, event):
        self._query_drag_start_x = event.x
        self._query_drag_start_y = event.y

    def _update_result_y_scroll(self, first, last):
        self.result_y_scroll.set(first, last)
        self._result_need_y_scroll = not (float(first) <= 0.0 and float(last) >= 1.0)
        self._sync_result_scrollbars_visibility()

    def _update_result_x_scroll(self, first, last):
        self.result_x_scroll.set(first, last)
        self._result_need_x_scroll = not (float(first) <= 0.0 and float(last) >= 1.0)
        self._sync_result_scrollbars_visibility()

    def _on_result_area_enter(self, _event=None):
        self._result_hover = True
        if self._result_hide_job is not None:
            self.query_window.after_cancel(self._result_hide_job)
            self._result_hide_job = None
        self._sync_result_scrollbars_visibility()

    def _on_result_area_leave(self, _event=None):
        self._result_hover = False
        if self._result_hide_job is not None:
            self.query_window.after_cancel(self._result_hide_job)
        self._result_hide_job = self.query_window.after(120, self._sync_result_scrollbars_visibility)

    def _sync_result_scrollbars_visibility(self):
        if not (self.query_window and self.query_window.winfo_exists()):
            return
        show = self._result_hover
        if show and self._result_need_y_scroll:
            self.result_y_scroll.place(relx=1.0, rely=0.0, relheight=1.0, x=-2, anchor="ne")
        else:
            self.result_y_scroll.place_forget()

        if show and self._result_need_x_scroll:
            self.result_x_scroll.place(relx=0.0, rely=1.0, relwidth=1.0, y=-2, anchor="sw")
        else:
            self.result_x_scroll.place_forget()

    def _on_query_drag(self, event):
        if not (self.query_window and self.query_window.winfo_exists()):
            return
        if self._query_is_maximized:
            return
        x = self.query_window.winfo_x() + event.x - self._query_drag_start_x
        y = self.query_window.winfo_y() + event.y - self._query_drag_start_y
        w = self.query_window.winfo_width() or 760
        h = self.query_window.winfo_height() or 560
        self.query_window.geometry(f"{w}x{h}+{x}+{y}")

    def _toggle_query_window_maximize(self):
        if not (self.query_window and self.query_window.winfo_exists()):
            return
        if self._query_is_maximized:
            self.query_window.geometry(self._query_restore_geometry)
            self._query_is_maximized = False
        else:
            self._query_restore_geometry = self.query_window.geometry()
            sw = self.query_window.winfo_screenwidth()
            sh = self.query_window.winfo_screenheight()
            self.query_window.geometry(f"{sw}x{sh}+0+0")
            self._query_is_maximized = True

    def _minimize_query_window(self):
        if self.query_window and self.query_window.winfo_exists():
            self.query_window.withdraw()

    def _close_query_window(self):
        if self.query_window and self.query_window.winfo_exists():
            self.query_window.destroy()
            self.query_window = None

    def _start_query(self):
        text = self.query_text.get("1.0", "end").strip()
        if not text:
            messagebox.showwarning("Note", "Please enter a word or sentence to search")
            return

        self._ensure_query_result_visible()

        self.status_label.config(text="Searching...")
        self.result_text.delete("1.0", "end")

        def worker():
            try:
                result = self.translator.query(text)
                self.query_window.after(0, lambda: self._show_query_result(result))
            except Exception as exc:
                error_msg = str(exc)
                self.query_window.after(0, lambda: self._show_query_error(error_msg))

        threading.Thread(target=worker, daemon=True).start()

    def _on_query_enter(self, event):
        # Shift+Enter keeps newline input; Enter starts search.
        if event.state & 0x1:
            return None
        self._start_query()
        return "break"

    def _show_query_result(self, result: str):
        self.status_label.config(text="完成")
        self._ensure_query_result_visible()
        self.result_text.delete("1.0", "end")
        self.result_text.insert("1.0", result, "result_content")
        self.result_wrap.lift()
        self.result_text.lift()
        self.result_text.focus_set()
        self.result_text.yview_moveto(0.0)
        self.result_text.xview_moveto(0.0)
        self.status_label.config(fg="#b9d3e6", text=f"Done {len(result)} chars")
        self._sync_result_scrollbars_visibility()

    def _show_query_error(self, error: str):
        self.status_label.config(text="Failed")
        self._ensure_query_result_visible()
        self.result_text.delete("1.0", "end")
        self.result_text.insert("1.0", error, "result_content")
        self.result_wrap.lift()
        self.result_text.lift()
        self.result_text.focus_set()
        self.result_text.yview_moveto(0.0)
        self.result_text.xview_moveto(0.0)
        self.status_label.config(fg="#ffb8c6", text="Failed")
        self._sync_result_scrollbars_visibility()

    def _ensure_query_result_visible(self):
        if not (self.query_window and self.query_window.winfo_exists()):
            return
        if self._query_is_maximized:
            return

        self.query_window.update_idletasks()
        current_w = self.query_window.winfo_width()
        current_h = self.query_window.winfo_height()
        target_w = max(current_w, self._query_result_size[0])
        target_h = max(current_h, self._query_result_size[1])
        if current_w < target_w or current_h < target_h:
            x = self.query_window.winfo_x()
            y = self.query_window.winfo_y()
            self.query_window.geometry(f"{target_w}x{target_h}+{x}+{y}")
            self._query_restore_geometry = self.query_window.geometry()


    def _speak_query_input(self):
        text = self.query_text.get("1.0", "end").strip()
        if not text:
            return
        if not self.speech.available():
            messagebox.showwarning("提示", "当前环境不可用发音引擎，请安装 pyttsx3")
            return
        self.speech.speak_async(text)

    def _add_result_to_review(self):
        source_text = self.query_text.get("1.0", "end").strip()
        result = self.result_text.get("1.0", "end").strip()
        if not source_text:
            messagebox.showwarning("Note", "Please enter a word or sentence first")
            return
        if not result:
            messagebox.showwarning("Note", "Please search first, then add to review")
            return

        english = ""
        chinese = ""
        pronunciation = ""
        example = ""
        for raw_line in result.splitlines():
            line = raw_line.strip()
            if line.startswith("English Translation:"):
                english = line.split(":", 1)[1].strip()
                if self.translator._contains_chinese(source_text):
                    chinese = source_text
            elif line.startswith("Chinese Meaning:") or line.startswith("Chinese Translation:"):
                chinese = line.split(":", 1)[1].strip()
                if not self.translator._contains_chinese(source_text):
                    english = source_text
            elif line.startswith("Pronunciation:"):
                pronunciation = line.split(":", 1)[1].strip()
            elif line.startswith("Example:"):
                example = line.split(":", 1)[1].strip()

        if not english or not chinese:
            messagebox.showerror("Error", "Cannot parse result. Please run search again.")
            return

        ok = self.store.add_word(english, chinese, pronunciation, example)
        if ok:
            messagebox.showinfo("Info", "Added to review cards")
            if self.review_window and self.review_window.winfo_exists():
                self._rebuild_review_order(reset_index=True)
                self._render_card()
        else:
            messagebox.showerror("Error", "Failed to add")


def main():
    root = tk.Tk()
    style = ttk.Style()
    if "vista" in style.theme_names():
        style.theme_use("vista")
    EnglishLearningApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
