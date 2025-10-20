import tkinter as tk
from tkinter import ttk, filedialog
import tkinter.font as tkfont
import time
import threading
import random
import pyautogui # <-- The external application typing library
import json
import os
import re

# Configure PyAutoGUI settings
pyautogui.PAUSE = 0 # No pause between PyAutoGUI calls by default (we manage the delay ourselves)
pyautogui.FAILSAFE = True # Move the mouse to the top-left corner to stop the program


def py_typewrite(text, interval=0):
    """Compatibility wrapper that uses pyautogui.write if available, otherwise falls back to pyautogui.typewrite.

    Raises a RuntimeError with module file info if neither is available to help diagnose shadowing.
    """
    if hasattr(pyautogui, 'write'):
        return pyautogui.write(text, interval=interval)
    if hasattr(pyautogui, 'typewrite'):
        return pyautogui.typewrite(text, interval=interval)

    # Diagnostic output to help figure out why the functions are missing.
    try:
        module_file = getattr(pyautogui, '__file__', None)
        print("[diagnostic] pyautogui repr:", repr(pyautogui))
        print("[diagnostic] pyautogui type:", type(pyautogui))
        print("[diagnostic] pyautogui __file__:", module_file)
        print("[diagnostic] pyautogui dir() (first 50):", dir(pyautogui)[:50])
    except Exception as e:
        print("[diagnostic] error while introspecting pyautogui:", e)

    raise RuntimeError(
        "pyautogui has no 'write' or 'typewrite'. This usually means PyAutoGUI is not installed or a local file is shadowing it. "
        f"Module __file__: {module_file!r}.\nCheck your PYTHONPATH and ensure there's no local file named 'pyautogui.py' or a folder named 'pyautogui'."
    )


def _build_qwerty_neighbors():
    """Build a mapping of characters to nearby keys on a standard QWERTY keyboard.

    The mapping includes horizontal neighbors and approximate vertical neighbors by
    looking at the same column index in adjacent rows when possible.
    """
    rows = [
        "`1234567890-=",
        "qwertyuiop[]\\",
        "asdfghjkl;'",
        "zxcvbnm,./",
    ]
    neighbors = {}
    for r_idx, row in enumerate(rows):
        for i, ch in enumerate(row):
            s = set()
            # horizontal neighbors
            if i - 1 >= 0:
                s.add(row[i - 1])
            if i + 1 < len(row):
                s.add(row[i + 1])

            # vertical/diagonal neighbors from adjacent rows (approximate)
            for adj in (r_idx - 1, r_idx + 1):
                if 0 <= adj < len(rows):
                    adj_row = rows[adj]
                    # same column
                    if i < len(adj_row):
                        s.add(adj_row[i])
                    # left diagonal
                    if i - 1 >= 0 and i - 1 < len(adj_row):
                        s.add(adj_row[i - 1])
                    # right diagonal
                    if i + 1 < len(adj_row) and i + 1 < len(adj_row):
                        s.add(adj_row[i + 1])

            neighbors[ch] = ''.join(sorted(s))

    # add letters upper/lowercase variants and space
    full = {}
    for k, v in neighbors.items():
        full[k] = v
        if k.isalpha():
            full[k.upper()] = ''.join(x.upper() if x.isalpha() else x for x in v)

    # add a few common mappings for space and newline
    full[' '] = 'bvn '  # keys near the space bar on many layouts
    full['\n'] = full.get('\n', '\n')
    return full


_QW_NEIGHBORS = _build_qwerty_neighbors()


def get_nearby_char(target: str) -> str:
    """Return a nearby (adjacent-key) character for the given target char.

    If no neighbor is known, fall back to a random lowercase letter.
    """
    if not target:
        return ''
    # preserve case: mapping contains both lower and upper where applicable
    if target in _QW_NEIGHBORS and _QW_NEIGHBORS[target]:
        choices = _QW_NEIGHBORS[target]
        return random.choice(choices)

    # If char not in our map (e.g., emoji), fall back to an adjacent letter
    return random.choice('abcdefghijklmnopqrstuvwxyz')


# Abbreviations to ignore when deciding if a period ends a sentence
_ABBREVIATIONS = {"e.g", "i.e", "mr", "mrs", "dr", "jr", "sr", "vs", "etc", "prof", "rev", "st", "rd", "ave"}


def is_sentence_terminator(text: str, idx: int) -> bool:
    """Return True if the character at text[idx] is a real sentence terminator.

    Avoid treating common abbreviations like 'e.g.' or 'Mr.' as sentence ends.
    """
    if idx < 0 or idx >= len(text):
        return False
    ch = text[idx]
    if ch not in '.!?':
        return False

    # look backwards to find the word before the punctuation
    j = idx - 1
    while j >= 0 and text[j].isspace():
        j -= 1
    # collect letters/dots from the previous token up to some reasonable length
    start = j
    while start >= 0 and (text[start].isalpha() or text[start] == '.'):
        start -= 1
    word = text[start+1:j+1].lower()
    # strip trailing dot parts
    word = word.rstrip('.')
    if not word:
        return True
    if word in _ABBREVIATIONS:
        return False
    return True


def split_into_sentences(text: str):
    """Regex-based sentence splitter that returns (start,end,sentence_text).

    This uses punctuation-based splitting but tries to avoid common abbreviation splits and keeps
    dialog/list lines together when possible.
    """
    sentences = []
    # Normalize line endings
    t = text.replace('\r\n', '\n').replace('\r', '\n')

    # regex to find sentence-ending punctuation followed by space/newline and uppercase or quote
    pattern = re.compile(r"(?P<sent>.+?(?:[\.\?!]+|\n|$))", re.DOTALL)
    idx = 0
    for m in pattern.finditer(t):
        sent = m.group('sent')
        if not sent or sent.strip() == '':
            idx = m.end()
            continue
        start = idx
        end = m.end()
        s = sent.strip()
        # avoid splitting on single-letter initials and common abbreviations
        tail = s[-4:].lower()
        if any(s.lower().endswith(ab + '.') for ab in _ABBREVIATIONS):
            # keep going until next match
            idx = end
            continue

        sentences.append((start, end, s))
        idx = end

    # merge short lines that are bullets or dialog into previous sentence if appropriate
    merged = []
    for start, end, s in sentences:
        if merged:
            prev_start, prev_end, prev_s = merged[-1]
            # if current sentence is a single bullet/numbered line, merge
            if re.match(r'^[\s]*([-*•]|\d+\.)\s+', s):
                merged[-1] = (prev_start, end, (prev_s + '\n' + s).strip())
                continue
            # if current looks like dialog (starts with em-dash or name:), merge
            if re.match(r'^[\s]*(—|-|[A-Z][a-z]+:)', s):
                merged[-1] = (prev_start, end, (prev_s + ' ' + s).strip())
                continue
        merged.append((start, end, s))
    return merged


def classify_sentence(sent_text: str, full_text: str, start_idx: int, end_idx: int):
    """Return a richer set of tags for the sentence.

    Tags: quote, analysis, context, dialog, list, long
    """
    tags = set()
    s = sent_text.strip()
    lower = s.lower()

    # quote detection: starts/ends with quotes or enclosed
    if s.startswith(('"', '“', "'", '‘')) or s.endswith(('"', '”', "'", '’')):
        tags.add('quote')
    else:
        before = full_text[:start_idx]
        if before.count('"') % 2 == 1 or before.count("'") % 2 == 1:
            tags.add('quote')

    # dialog detection: em-dash, leading speaker tag or lines with quotes
    if re.match(r'^[\-—]\s*', s) or re.match(r'^[A-Z][a-z]+:\s', s) or lower.startswith('said '):
        tags.add('dialog')

    # list detection: leading bullet or numbered markers
    if re.match(r'^([-*•]|\d+\.)\s+', s):
        tags.add('list')

    # analysis detection: parentheses, 'note:', 'however', i.e./e.g.
    if '(' in s or any(k in lower for k in ('note:', 'however', 'i.e.', 'e.g.', 'viz', 'namely')) or ':' in s[:20]:
        tags.add('analysis')

    # context detection: common lead-ins
    if lower.startswith(('in conclusion', 'overall', 'context:', 'importantly', 'moreover', 'in summary', 'to conclude', 'therefore', 'consequently')):
        tags.add('context')

    # long sentence
    if len(s) > 200:
        tags.add('long')

    return tags


class ToolTip:
    """Simple tooltip for tkinter widgets."""
    def __init__(self, widget, text):
        self.widget = widget
        self.text = text
        self.tip = None
        widget.bind('<Enter>', self.show)
        widget.bind('<Leave>', self.hide)

    def show(self, _=None):
        if self.tip:
            return
        x = self.widget.winfo_rootx() + 20
        y = self.widget.winfo_rooty() + self.widget.winfo_height() + 10
        self.tip = tk.Toplevel(self.widget)
        self.tip.wm_overrideredirect(True)
        self.tip.wm_geometry(f"+{x}+{y}")
        lbl = ttk.Label(self.tip, text=self.text, relief='solid', padding=(6,4))
        lbl.pack()

    def hide(self, _=None):
        if self.tip:
            try:
                self.tip.destroy()
            except Exception:
                pass
            self.tip = None

class ExternalTypingSimulatorApp:
    def __init__(self, root):
        self.root = root
        root.title("HumanTyper — External Typing Simulator")
        root.geometry("840x560")
        root.minsize(760, 480)

        # Load config or defaults
        self.config_path = os.path.join(os.path.dirname(__file__), 'config.json')
        self.config = self.load_config()

        # --- Variables ---
        self.text_to_type = tk.StringVar(value=self.config.get('text_to_type', "This text will be typed into the active window."))
        self.typing_speed_wpm = tk.DoubleVar(value=self.config.get('typing_speed_wpm', 40))  # Default WPM
        # Thinking pause controls
        self.enable_thinking = tk.BooleanVar(value=self.config.get('enable_thinking', True))
        self.mid_sentence_pause_chance = tk.DoubleVar(value=self.config.get('mid_sentence_pause_chance', 0.05))
        self.mid_sentence_pause_seconds = tk.DoubleVar(value=self.config.get('mid_sentence_pause_seconds', 0.8))
        self.sentence_pause_seconds = tk.DoubleVar(value=self.config.get('sentence_pause_seconds', 1.6))
        # Major pause between paragraphs (preserve paragraph breaks)
        self.paragraph_pause_seconds = tk.DoubleVar(value=self.config.get('paragraph_pause_seconds', 20.0))
        # Sentence-structure pause multipliers
        self.quote_sentence_multiplier = tk.DoubleVar(value=self.config.get('quote_sentence_multiplier', 1.5))
        self.analysis_sentence_multiplier = tk.DoubleVar(value=self.config.get('analysis_sentence_multiplier', 1.8))
        self.context_sentence_multiplier = tk.DoubleVar(value=self.config.get('context_sentence_multiplier', 1.3))
            # UI advanced toggle
            self.show_advanced = tk.BooleanVar(value=self.config.get('show_advanced', True))
            self.is_typing = False

        # --- GUI Setup ---
        self.setup_ui()

    def setup_ui(self):
        # Apply a minimal dark theme
        style = ttk.Style(self.root)
        try:
            style.theme_use('clam')
        except Exception:
            pass
        # Basic dark color configuration
        bg = '#1e1f22'
        fg = '#dbe6f1'
        accent = '#3a7ca5'
        entry_bg = '#222326'
        scale_trough = '#2b2d30'
        self.root.configure(bg=bg)
        style.configure('.', background=bg, foreground=fg)
        style.configure('TLabel', background=bg, foreground=fg)
        style.configure('TButton', background=accent, foreground=fg)
        style.map('TButton', background=[('active', '#2f6b8d')])
        # Entry and scale appearances
        try:
            style.configure('TEntry', fieldbackground=entry_bg, foreground=fg)
            style.configure('TScale', troughcolor=scale_trough, background=bg)
            style.configure('Horizontal.TScale', troughcolor=scale_trough, background=bg)
        except Exception:
            # Some themes/platforms may not support these options; ignore failures
            pass

        # Set a readable UI font
        default_font = tkfont.nametofont("TkDefaultFont")
        default_font.configure(size=10)
        heading_font = default_font.copy()
        heading_font.configure(size=11, weight='bold')

        main_frame = ttk.Frame(self.root, padding="12")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        main_frame.columnconfigure(0, weight=1)

        # 1. Input Text Area (multiline to preserve indentation on paste)
        ttk.Label(main_frame, text="Text to Type (Ensure Target Window is Focused!):").grid(row=0, column=0, sticky=tk.W, pady=5)
        self.input_text = tk.Text(main_frame, width=80, height=6, font=default_font, wrap='none', undo=True)
        self.input_text.grid(row=1, column=0, sticky=(tk.W, tk.E), pady=5)
        # Put initial content into the Text widget
        self.input_text.insert('1.0', self.text_to_type.get())
        ToolTip(self.input_text, "Text that will be typed into the active window. Pasting preserves indentation.")
        # track modifications and sync to the StringVar for backward compatibility
        def _on_text_modified(event=None):
            try:
                text = self.input_text.get('1.0', 'end-1c')
                # Update the StringVar without causing recursion
                self.text_to_type.set(text)
                self.update_eta_display()
            except Exception:
                pass
            # clear the modified flag
            try:
                self.input_text.edit_modified(False)
            except Exception:
                pass

        # Bind modifications (covers paste, typing, and programmatic changes)
        self.input_text.bind('<<Modified>>', _on_text_modified)
        # Also update ETA on explicit key release (helpful for some paste scenarios)
        self.input_text.bind('<KeyRelease>', lambda e: self.update_eta_display())

        # 2. WPM Slider
        ttk.Label(main_frame, text="Typing Rate (WPM):").grid(row=2, column=0, sticky=tk.W, pady=5)
        speed_frame = ttk.Frame(main_frame)
        speed_frame.grid(row=3, column=0, sticky=(tk.W, tk.E))
        speed_frame.columnconfigure(0, weight=1)

        self.speed_slider = ttk.Scale(speed_frame, from_=10, to=100, orient=tk.HORIZONTAL,
                                      variable=self.typing_speed_wpm, command=self.update_slider_label)
        self.speed_slider.grid(row=0, column=0, sticky=(tk.W, tk.E), padx=5)

        self.speed_label = ttk.Label(speed_frame, text=f"{self.typing_speed_wpm.get():.0f} WPM")
        self.speed_label.grid(row=0, column=1, padx=5)

        # numeric entry for WPM
        self.wpm_entry = ttk.Entry(speed_frame, width=6, textvariable=self.typing_speed_wpm)
        self.wpm_entry.grid(row=0, column=2, padx=6)

        # Advanced toggle - hides advanced controls when unchecked
        self.advanced_check = ttk.Checkbutton(speed_frame, text='Advanced settings', variable=self.show_advanced, command=self.toggle_advanced)
        self.advanced_check.grid(row=0, column=3, padx=6)

        # 3. Simulate Button
        self.simulate_button = ttk.Button(main_frame, text="Start Typing (Switch to Target App in 3s)", command=self.start_typing_thread)
        self.simulate_button.grid(row=4, column=0, pady=15)
        ToolTip(self.simulate_button, "Starts typing after a short countdown. Move mouse to a corner to abort (PyAutoGUI failsafe).")

        # ETA label
        self.eta_label = ttk.Label(main_frame, text="ETA: --:--")
        self.eta_label.grid(row=4, column=1, sticky=tk.W, padx=6)

        # Progress bar showing typing progress
        self.progress = ttk.Progressbar(main_frame, orient='horizontal', mode='determinate')
        self.progress.grid(row=4, column=2, sticky=(tk.W, tk.E), padx=6)
        main_frame.columnconfigure(2, weight=0)
        ToolTip(self.progress, "Shows progress of the current typing run.")

        # 4. Instructions/Status
        self.status_label = ttk.Label(main_frame, text="Status: Ready. Click the button and quickly switch windows.")
        self.status_label.grid(row=5, column=0, pady=5)

        # 5. Thinking pause controls
        thinking_frame = ttk.Frame(main_frame, padding=(0,8,0,0))
        thinking_frame.grid(row=6, column=0, sticky=(tk.W, tk.E))
        thinking_frame.columnconfigure(1, weight=1)

        ttk.Checkbutton(thinking_frame, text="Enable thinking pauses", variable=self.enable_thinking).grid(row=0, column=0, sticky=tk.W)
        ttk.Label(thinking_frame, text="Mid-word pause chance:").grid(row=1, column=0, sticky=tk.W, pady=2)
        self.mid_pause_scale = ttk.Scale(thinking_frame, from_=0.0, to=0.3, orient=tk.HORIZONTAL, variable=self.mid_sentence_pause_chance, command=lambda *_: self.save_config())
        self.mid_pause_scale.grid(row=1, column=1, sticky=(tk.W, tk.E), padx=6)
        self.mid_pause_entry = ttk.Entry(thinking_frame, width=6, textvariable=self.mid_sentence_pause_chance)
        self.mid_pause_entry.grid(row=1, column=2, padx=6)

        ttk.Label(thinking_frame, text="Mid-word pause seconds:").grid(row=2, column=0, sticky=tk.W, pady=2)
        self.mid_seconds_scale = ttk.Scale(thinking_frame, from_=0.0, to=3.0, orient=tk.HORIZONTAL, variable=self.mid_sentence_pause_seconds, command=lambda *_: self.save_config())
        self.mid_seconds_scale.grid(row=2, column=1, sticky=(tk.W, tk.E), padx=6)
        self.mid_seconds_entry = ttk.Entry(thinking_frame, width=6, textvariable=self.mid_sentence_pause_seconds)
        self.mid_seconds_entry.grid(row=2, column=2, padx=6)

        ttk.Label(thinking_frame, text="Sentence pause seconds:").grid(row=3, column=0, sticky=tk.W, pady=2)
        self.sentence_scale = ttk.Scale(thinking_frame, from_=0.0, to=6.0, orient=tk.HORIZONTAL, variable=self.sentence_pause_seconds, command=lambda *_: self.save_config())
        self.sentence_scale.grid(row=3, column=1, sticky=(tk.W, tk.E), padx=6)
        self.sentence_entry = ttk.Entry(thinking_frame, width=6, textvariable=self.sentence_pause_seconds)
        self.sentence_entry.grid(row=3, column=2, padx=6)

        # Paragraph pause (major pause between paragraphs)
        ttk.Label(thinking_frame, text="Paragraph pause seconds:").grid(row=4, column=0, sticky=tk.W, pady=2)
        self.paragraph_scale = ttk.Scale(thinking_frame, from_=0.0, to=300.0, orient=tk.HORIZONTAL, variable=self.paragraph_pause_seconds, command=lambda *_: self.save_config())
        self.paragraph_scale.grid(row=4, column=1, sticky=(tk.W, tk.E), padx=6)
        self.paragraph_entry = ttk.Entry(thinking_frame, width=6, textvariable=self.paragraph_pause_seconds)
        self.paragraph_entry.grid(row=4, column=2, padx=6)

        # Sentence-structure multipliers
        ttk.Label(thinking_frame, text="Quote sentence multiplier:").grid(row=5, column=0, sticky=tk.W, pady=2)
        self.quote_scale = ttk.Scale(thinking_frame, from_=0.5, to=3.0, orient=tk.HORIZONTAL, variable=self.quote_sentence_multiplier, command=lambda *_: self.save_config())
        self.quote_scale.grid(row=5, column=1, sticky=(tk.W, tk.E), padx=6)
        self.quote_entry = ttk.Entry(thinking_frame, width=6, textvariable=self.quote_sentence_multiplier)
        self.quote_entry.grid(row=5, column=2, padx=6)

        ttk.Label(thinking_frame, text="Analysis sentence multiplier:").grid(row=6, column=0, sticky=tk.W, pady=2)
        self.analysis_scale = ttk.Scale(thinking_frame, from_=0.5, to=4.0, orient=tk.HORIZONTAL, variable=self.analysis_sentence_multiplier, command=lambda *_: self.save_config())
        self.analysis_scale.grid(row=6, column=1, sticky=(tk.W, tk.E), padx=6)
        self.analysis_entry = ttk.Entry(thinking_frame, width=6, textvariable=self.analysis_sentence_multiplier)
        self.analysis_entry.grid(row=6, column=2, padx=6)

        ttk.Label(thinking_frame, text="Context sentence multiplier:").grid(row=7, column=0, sticky=tk.W, pady=2)
        self.context_scale = ttk.Scale(thinking_frame, from_=0.5, to=3.0, orient=tk.HORIZONTAL, variable=self.context_sentence_multiplier, command=lambda *_: self.save_config())
        self.context_scale.grid(row=7, column=1, sticky=(tk.W, tk.E), padx=6)
        self.context_entry = ttk.Entry(thinking_frame, width=6, textvariable=self.context_sentence_multiplier)
        self.context_entry.grid(row=7, column=2, padx=6)

        # initially apply advanced visibility
        try:
            self.toggle_advanced()
        except Exception:
            pass

        # Presets (single dropdown)
        preset_frame = ttk.Frame(main_frame)
        preset_frame.grid(row=7, column=0, sticky=(tk.W, tk.E), pady=(8,0))
        ttk.Label(preset_frame, text="Preset:").grid(row=0, column=0, sticky=tk.W)
        self.preset_var = tk.StringVar(value='Normal')
        preset_values = ["Conservative", "Normal", "Deep Thinker", "Student"]
        self.preset_combo = ttk.Combobox(preset_frame, textvariable=self.preset_var, values=preset_values, state='readonly', width=20)
        self.preset_combo.grid(row=0, column=1, padx=6)
        # Map selection to our preset keys and apply
        self.preset_combo.bind('<<ComboboxSelected>>', lambda e: self.apply_preset('deep' if self.preset_var.get() == 'Deep Thinker' else self.preset_var.get().lower()))

        cfg_btn_frame = ttk.Frame(main_frame)
        cfg_btn_frame.grid(row=8, column=0, sticky=(tk.W, tk.E), pady=(6,0))
        ttk.Button(cfg_btn_frame, text="Save Config As...", command=self.save_config_as).grid(row=0, column=0, padx=6)
        ttk.Button(cfg_btn_frame, text="Load Config...", command=self.load_config_from_dialog).grid(row=0, column=1, padx=6)

        # Bind changes to save config
        self.typing_speed_wpm.trace_add('write', lambda *_: self.save_config())
        # Also update the WPM label when the variable changes (e.g., presets)
        self.typing_speed_wpm.trace_add('write', lambda *_: self.update_slider_label())
        # update ETA when settings change
        self.typing_speed_wpm.trace_add('write', lambda *_: self.update_eta_display())
        self.mid_sentence_pause_chance.trace_add('write', lambda *_: self.update_eta_display())
        self.mid_sentence_pause_seconds.trace_add('write', lambda *_: self.update_eta_display())
        self.sentence_pause_seconds.trace_add('write', lambda *_: self.update_eta_display())
        self.paragraph_pause_seconds.trace_add('write', lambda *_: self.update_eta_display())
        self.quote_sentence_multiplier.trace_add('write', lambda *_: self.update_eta_display())
        self.analysis_sentence_multiplier.trace_add('write', lambda *_: self.update_eta_display())
        self.context_sentence_multiplier.trace_add('write', lambda *_: self.update_eta_display())
        self.enable_thinking.trace_add('write', lambda *_: self.save_config())
        self.mid_sentence_pause_chance.trace_add('write', lambda *_: self.save_config())
        self.mid_sentence_pause_seconds.trace_add('write', lambda *_: self.save_config())
        self.sentence_pause_seconds.trace_add('write', lambda *_: self.save_config())
        self.paragraph_pause_seconds.trace_add('write', lambda *_: self.save_config())

        # Update ETA when the input text changes (covers paste and programmatic changes)
        self.text_to_type.trace_add('write', lambda *_: self.update_eta_display())



    def update_slider_label(self, *args):
        """Updates the WPM label next to the slider."""
        self.speed_label.config(text=f"{self.typing_speed_wpm.get():.0f} WPM")

    def get_delay_per_char(self):
        """Calculates the delay (in seconds) between characters based on WPM."""
        wpm = self.typing_speed_wpm.get()
        # 5 characters per word
        characters_per_minute = wpm * 5
        characters_per_second = characters_per_minute / 60
        
        # Delay is the reciprocal
        if characters_per_second > 0:
            return 1 / characters_per_second
        return 0.1 # Default safe minimum delay

    def estimate_remaining_seconds(self, text: str, idx: int) -> float:
        """Estimate remaining time in seconds to type the rest of `text` starting at index `idx`.

        Uses current WPM, a fixed mistake overhead, and enabled thinking pause settings.
        This is a best-effort estimate and updates during typing.
        """
        remaining = text[idx:]
        # base delay per char
        base = self.get_delay_per_char()
        chars = max(1, len(remaining))

        # approximate extra time for mistakes: assume mistake rate equals the mid-word pause chance
        mistake_rate = float(self.mid_sentence_pause_chance.get())
        # assume each mistake costs ~1.5 characters worth of time (insert + backspace + pause)
        mistake_overhead = 1.5 * base

        # thinking pauses: approximate how many pauses will occur
        thinking_overhead = 0.0
        if self.enable_thinking.get():
            # rough heuristic: one mid-word pause per (1/mid_chance) characters on average
            mid_chance = float(self.mid_sentence_pause_chance.get())
            if mid_chance > 0:
                expected_mid_pauses = chars * mid_chance
                thinking_overhead += expected_mid_pauses * float(self.mid_sentence_pause_seconds.get())

            # sentence-level pauses: split into sentences and classify
            sentences = split_into_sentences(remaining)
            for start_i, end_i, sent in sentences:
                pause = float(self.sentence_pause_seconds.get())
                tags = classify_sentence(sent, remaining, start_i, end_i)
                # apply multipliers
                if 'quote' in tags:
                    pause *= float(self.quote_sentence_multiplier.get())
                if 'analysis' in tags or 'long' in tags:
                    pause *= float(self.analysis_sentence_multiplier.get())
                if 'context' in tags:
                    pause *= float(self.context_sentence_multiplier.get())
                # dialog/list have smaller pauses
                if 'dialog' in tags or 'list' in tags:
                    pause *= 0.7
                thinking_overhead += pause

            # paragraph pauses: count double-newline occurrences as paragraph breaks
            paragraph_count = remaining.count('\n\n')
            thinking_overhead += paragraph_count * float(self.paragraph_pause_seconds.get())

        # total base typing time
        typing_time = chars * base
        mistakes_time = chars * mistake_rate * mistake_overhead

        total = typing_time + mistakes_time + thinking_overhead
        return total

    def update_eta_display(self, text=None, idx=0):
        try:
            if text is None:
                try:
                    text = self.input_text.get('1.0', 'end-1c')
                except Exception:
                    text = self.text_to_type.get()
            secs = self.estimate_remaining_seconds(text, idx)
            mins = int(secs) // 60
            sec = int(secs) % 60
            self.eta_label.config(text=f"ETA: {mins:02d}:{sec:02d}")
        except Exception:
            self.eta_label.config(text="ETA: --:--")

    def toggle_advanced(self):
        """Show or hide advanced controls besides WPM and preset."""
        show = self.show_advanced.get()
        # thinking_frame and preset_frame are created in setup_ui; hide their children
        try:
            # we know thinking_frame occupies several rows; hide/show by grid_remove/grid
            for widget in self.root.grid_slaves():
                # leave the main frame intact; hide advanced children by checking widget master
                pass
        except Exception:
            pass
        # A simpler approach: enable/disable the advanced widgets by toggling their state
        try:
            state = 'normal' if show else 'hidden'
            # hide/show by adjusting visibility of scales/entries in the thinking area
            for w in (self.mid_pause_scale, self.mid_pause_entry, self.mid_seconds_scale, self.mid_seconds_entry,
                      self.sentence_scale, self.sentence_entry, self.paragraph_scale, self.paragraph_entry,
                      self.quote_scale, self.quote_entry, self.analysis_scale, self.analysis_entry,
                      self.context_scale, self.context_entry):
                if show:
                    try:
                        w.grid()
                    except Exception:
                        pass
                else:
                    try:
                        w.grid_remove()
                    except Exception:
                        pass
        except Exception:
            pass

    def load_config(self):
        try:
            if os.path.exists(self.config_path):
                with open(self.config_path, 'r', encoding='utf-8') as f:
                    return json.load(f)
        except Exception:
            pass
        return {}

    def save_config(self):
        cfg = {
            'text_to_type': self.input_text.get('1.0', 'end-1c'),
            'typing_speed_wpm': float(self.typing_speed_wpm.get()),
            'enable_thinking': bool(self.enable_thinking.get()),
            'mid_sentence_pause_chance': float(self.mid_sentence_pause_chance.get()),
            'mid_sentence_pause_seconds': float(self.mid_sentence_pause_seconds.get()),
            'sentence_pause_seconds': float(self.sentence_pause_seconds.get()),
            'paragraph_pause_seconds': float(self.paragraph_pause_seconds.get()),
            'quote_sentence_multiplier': float(self.quote_sentence_multiplier.get()),
            'analysis_sentence_multiplier': float(self.analysis_sentence_multiplier.get()),
            'context_sentence_multiplier': float(self.context_sentence_multiplier.get()),
            'show_advanced': bool(self.show_advanced.get()),
        }
        try:
            with open(self.config_path, 'w', encoding='utf-8') as f:
                json.dump(cfg, f, indent=2)
        except Exception:
            pass

    def apply_preset(self, name: str):
        presets = {
            'conservative': {
                'typing_speed_wpm': 30,
                'enable_thinking': True,
                'mid_sentence_pause_chance': 0.02,
                'mid_sentence_pause_seconds': 0.5,
                'sentence_pause_seconds': 1.0,
                'paragraph_pause_seconds': 45.0,
            },
            'normal': {
                'typing_speed_wpm': 45,
                'enable_thinking': True,
                'mid_sentence_pause_chance': 0.05,
                'mid_sentence_pause_seconds': 0.8,
                'sentence_pause_seconds': 1.6,
                'paragraph_pause_seconds': 20.0,
                'quote_sentence_multiplier': 1.4,
                'analysis_sentence_multiplier': 1.6,
                'context_sentence_multiplier': 1.2,
            },
            'deep': {
                'typing_speed_wpm': 35,
                'enable_thinking': True,
                'mid_sentence_pause_chance': 0.12,
                'mid_sentence_pause_seconds': 1.6,
                'sentence_pause_seconds': 3.0,
                'paragraph_pause_seconds': 60.0,
                'quote_sentence_multiplier': 1.6,
                'analysis_sentence_multiplier': 2.0,
                'context_sentence_multiplier': 1.4,
            }
            ,
            'student': {
                # Average student typist: moderate speed, moderate mistakes
                'typing_speed_wpm': 38,
                'enable_thinking': True,
                'mid_sentence_pause_chance': 0.07,
                'mid_sentence_pause_seconds': 0.9,
                'sentence_pause_seconds': 1.8,
                'paragraph_pause_seconds': 15.0,
                'quote_sentence_multiplier': 1.3,
                'analysis_sentence_multiplier': 1.5,
                'context_sentence_multiplier': 1.2,
            }
        }
        p = presets.get(name)
        if not p:
            return
        self.typing_speed_wpm.set(p['typing_speed_wpm'])
        self.enable_thinking.set(p['enable_thinking'])
        self.mid_sentence_pause_chance.set(p['mid_sentence_pause_chance'])
        self.mid_sentence_pause_seconds.set(p['mid_sentence_pause_seconds'])
        self.sentence_pause_seconds.set(p['sentence_pause_seconds'])
        # paragraph pause may be absent in older presets; guard default
        self.paragraph_pause_seconds.set(p.get('paragraph_pause_seconds', self.paragraph_pause_seconds.get()))
        self.quote_sentence_multiplier.set(p.get('quote_sentence_multiplier', self.quote_sentence_multiplier.get()))
        self.analysis_sentence_multiplier.set(p.get('analysis_sentence_multiplier', self.analysis_sentence_multiplier.get()))
        self.context_sentence_multiplier.set(p.get('context_sentence_multiplier', self.context_sentence_multiplier.get()))
        self.save_config()

    def save_config_as(self):
        fpath = filedialog.asksaveasfilename(defaultextension='.json', filetypes=[('JSON files','*.json')])
        if not fpath:
            return
        cfg = {
            'text_to_type': self.input_text.get('1.0', 'end-1c'),
            'typing_speed_wpm': float(self.typing_speed_wpm.get()),
            'enable_thinking': bool(self.enable_thinking.get()),
            'mid_sentence_pause_chance': float(self.mid_sentence_pause_chance.get()),
            'mid_sentence_pause_seconds': float(self.mid_sentence_pause_seconds.get()),
            'sentence_pause_seconds': float(self.sentence_pause_seconds.get()),
        }
        try:
            with open(fpath, 'w', encoding='utf-8') as f:
                json.dump(cfg, f, indent=2)
        except Exception:
            pass

    def load_config_from_dialog(self):
        fpath = filedialog.askopenfilename(filetypes=[('JSON files','*.json')])
        if not fpath:
            return
        try:
            with open(fpath, 'r', encoding='utf-8') as f:
                cfg = json.load(f)
            # apply - write into the Text widget so indentation is preserved
            text_val = cfg.get('text_to_type', self.input_text.get('1.0', 'end-1c'))
            try:
                self.input_text.delete('1.0', 'end')
                self.input_text.insert('1.0', text_val)
                # ensure StringVar stays in sync
                try:
                    self.text_to_type.set(text_val)
                except Exception:
                    pass
            except Exception:
                self.text_to_type.set(text_val)
            self.typing_speed_wpm.set(cfg.get('typing_speed_wpm', self.typing_speed_wpm.get()))
            self.enable_thinking.set(cfg.get('enable_thinking', self.enable_thinking.get()))
            self.mid_sentence_pause_chance.set(cfg.get('mid_sentence_pause_chance', self.mid_sentence_pause_chance.get()))
            self.mid_sentence_pause_seconds.set(cfg.get('mid_sentence_pause_seconds', self.mid_sentence_pause_seconds.get()))
            self.sentence_pause_seconds.set(cfg.get('sentence_pause_seconds', self.sentence_pause_seconds.get()))
            self.save_config()
        except Exception:
            pass

    def simulate_typing(self):
        """The core logic that sends keystrokes with delays and mistakes via pyautogui."""
        if self.is_typing:
            return
        
        self.is_typing = True
        self.simulate_button.config(text="Typing...", state=tk.DISABLED)
        self.status_label.config(text="Status: Typing in the focused external application.")

        # Give the user a grace period to switch to the target application (e.g., Notepad)
        self.status_label.config(text="Switch to your target application NOW (3 seconds)...")
        self.root.update()
        time.sleep(3) 

        text = self.text_to_type.get()
        # Prefer reading directly from the Text widget to preserve indentation
        try:
            text = self.input_text.get('1.0', 'end-1c')
        except Exception:
            text = self.text_to_type.get()

        # Typing Simulation Loop
        try:
            # Initialize progress bar
            total_chars = max(1, len(text))
            self.progress['maximum'] = total_chars
            self.progress['value'] = 0
            self.root.update()
            # Pre-split sentences to know pause multipliers per sentence
            sentence_ranges = split_into_sentences(text)
            # Build a map of end_index -> multiplier
            end_to_multiplier = {}
            for s_start, s_end, s_text in sentence_ranges:
                tags = classify_sentence(s_text, text, s_start, s_end)
                multiplier = 1.0
                if 'quote' in tags:
                    multiplier *= float(self.quote_sentence_multiplier.get())
                if 'analysis' in tags or 'long' in tags:
                    multiplier *= float(self.analysis_sentence_multiplier.get())
                if 'context' in tags:
                    multiplier *= float(self.context_sentence_multiplier.get())
                if 'dialog' in tags or 'list' in tags:
                    multiplier *= 0.7
                end_to_multiplier[s_end - 1] = multiplier

            for i, char in enumerate(text):
                # Calculate base delay, then add a slight human-like random variation
                base_delay = self.get_delay_per_char()
                delay = base_delay * (random.uniform(0.8, 1.2))
                
                # --- Mistake Simulation (e.g., 5% chance of a typo) ---
                if random.random() < 0.05 and char:
                    # pick a nearby key based on QWERTY adjacency
                    wrong_char = get_nearby_char(char)
                    # fallback to a random letter if mapping missing
                    if not wrong_char:
                        wrong_char = random.choice('abcdefghijklmnopqrstuvwxyz')
                        if char.isupper():
                            wrong_char = wrong_char.upper()

                    py_typewrite(wrong_char, interval=0) # Type instantly
                    time.sleep(base_delay * 2) # Longer pause for the mistake

                    # Simulate the backspace to correct
                    pyautogui.press('backspace')
                    time.sleep(base_delay * 0.5) # Short delay for backspace press

                # --- Thinking pauses ---
                if self.enable_thinking.get():
                    # mid-word/word pause
                    if random.random() < self.mid_sentence_pause_chance.get() and not char.isspace():
                        time.sleep(self.mid_sentence_pause_seconds.get())
                    # longer pause on sentence end
                        if char in '.!?':
                            # base sentence pause
                            pause = float(self.sentence_pause_seconds.get())
                            # apply multiplier if this index corresponds to a sentence end
                            mult = end_to_multiplier.get(i, 1.0)
                            time.sleep(pause * mult)

                # --- Type the correct character ---
                py_typewrite(char, interval=0) # Type instantly, delay is managed by time.sleep
                time.sleep(delay)
                # update progress
                try:
                    self.progress['value'] = i + 1
                    self.root.update()
                except Exception:
                    pass
        except pyautogui.FailSafeException:
            # User moved mouse to a corner to abort
            self.status_label.config(text="Status: Aborted by PyAutoGUI failsafe (mouse moved to corner).")
        except Exception as e:
            self.status_label.config(text=f"Status: Error during simulation: {e}")
        finally:
            # Ensure state is reset
            try:
                self.progress['value'] = 0
            except Exception:
                pass
            self.simulate_button.config(text="Start Typing (Switch to Target App in 3s)", state=tk.NORMAL)
            self.is_typing = False


    def start_typing_thread(self):
        """Starts the typing simulation in a separate thread to keep the GUI responsive."""
        if not self.is_typing:
            typing_thread = threading.Thread(target=self.simulate_typing)
            typing_thread.start()

if __name__ == "__main__":
    root = tk.Tk()
    app = ExternalTypingSimulatorApp(root)
    root.mainloop()