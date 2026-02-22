"""
Pico HID Autoclicker – Windows GUI
Tab 1  Autoclicker – configure keys, min/max delay; Start/Stop via buttons or F12.
Tab 2  Conditional Clicker – OCR-based HP reading and template-matched status effects.
Tab 3  Status Effects Library – capture, preview, and manage status effect templates.
Requires: pyserial, pynput, mss, Pillow, opencv-python, numpy, easyocr.
Run: pip install -r requirements.txt
"""

import tkinter as tk
from tkinter import ttk, messagebox
import serial
import serial.tools.list_ports
import threading
import time
import random
import json
import os

try:
    from PIL import Image, ImageTk
    from screen_reader import (
        capture_region, get_unique_colors, color_present, count_color_pixels,
        read_hp_percentage, match_template,
    )
    from region_selector import RegionSelector
    from template_store import list_templates, save_template, load_template, delete_template

    CONDITIONAL_AVAILABLE = True
except ImportError:
    CONDITIONAL_AVAILABLE = False

# Must match key names in pico/code.py KEY_MAP
KEY_OPTIONS = (
    [c for c in "abcdefghijklmnopqrstuvwxyz"]
    + [str(i) for i in range(10)]
    + [f"f{i}" for i in range(1, 13)]
    + [
        "space", "enter", "tab", "escape", "backspace",
        "minus", "equals", "left_bracket", "right_bracket", "backslash",
        "semicolon", "quote", "grave", "comma", "period", "slash",
        "insert", "delete", "home", "end", "page_up", "page_down",
        "up", "down", "left", "right",
    ]
)

# Raspberry Pi Pico (0x2E8A) and Adafruit/CircuitPython (0x239A) USB VIDs
PICO_VIDS = (0x2E8A, 0x239A)
BAUD = 115200
CONFIG_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "config.json")


def find_pico_ports():
    """Return (list of (device, display_str), True if any Pico VID was found)."""
    pico_ports = []
    all_ports = []
    for p in serial.tools.list_ports.comports():
        desc = p.description or p.device
        display = f"{p.device} ({desc})"
        entry = (p.device, display)
        all_ports.append(entry)
        if p.vid is not None and p.vid in PICO_VIDS:
            pico_ports.append(entry)
    chosen = pico_ports if pico_ports else all_ports
    return (chosen, bool(pico_ports))


def _ask_template_name(parent):
    """Pop a simple dialog asking the user for a template name. Returns str or None."""
    import tkinter.simpledialog
    return tkinter.simpledialog.askstring("Template Name", "Enter a name for this template:", parent=parent)


class AutoclickerApp:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("Pico HID Autoclicker")
        self.root.minsize(520, 560)
        self.serial_port = None
        self.running = False
        self._lock = threading.Lock()

        # Conditional-clicker state
        self.region = None
        self.monitoring = False
        self.monitor_thread = None
        self.preview_active = False
        self._preview_image_ref = None

        self._build_ui()
        self._refresh_ports()
        self._load_settings()
        self._start_f12_listener()

    # ------------------------------------------------------------------ #
    #  UI construction                                                     #
    # ------------------------------------------------------------------ #

    def _build_ui(self):
        main = ttk.Frame(self.root, padding=10)
        main.pack(fill=tk.BOTH, expand=True)

        # --- Shared COM-port section (above tabs) ---
        port_frame = ttk.LabelFrame(main, text="Connection", padding=6)
        port_frame.pack(fill=tk.X, pady=(0, 8))
        port_row = ttk.Frame(port_frame)
        port_row.pack(fill=tk.X)
        ttk.Label(port_row, text="COM port:").pack(side=tk.LEFT, padx=(0, 6))
        self.port_var = tk.StringVar()
        self.port_combo = ttk.Combobox(
            port_row, textvariable=self.port_var, width=24, state="readonly"
        )
        self.port_combo.pack(side=tk.LEFT, padx=(0, 6))
        ttk.Button(port_row, text="Refresh", command=self._refresh_ports).pack(
            side=tk.LEFT
        )
        self.conn_status_var = tk.StringVar(value="")
        ttk.Label(port_frame, textvariable=self.conn_status_var, foreground="gray").pack(
            anchor=tk.W, pady=(4, 0)
        )

        # --- Tabbed notebook ---
        self.notebook = ttk.Notebook(main)
        self.notebook.pack(fill=tk.BOTH, expand=True)

        self._build_autoclicker_tab()
        self._build_conditional_tab()
        self._build_templates_tab()

    # ---- Tab 1: Autoclicker -----------------------------------------

    def _build_autoclicker_tab(self):
        tab = ttk.Frame(self.notebook, padding=10)
        self.notebook.add(tab, text="Autoclicker")

        # Keys
        key_frame = ttk.LabelFrame(tab, text="Keys to press (in order)", padding=6)
        key_frame.pack(fill=tk.X, pady=(0, 8))
        key_row = ttk.Frame(key_frame)
        key_row.pack(fill=tk.X)
        self.key_var = tk.StringVar()
        self.key_combo = ttk.Combobox(
            key_row, textvariable=self.key_var, values=KEY_OPTIONS, width=14
        )
        self.key_combo.pack(side=tk.LEFT, padx=(0, 6))
        ttk.Button(key_row, text="Add", command=self._add_key).pack(
            side=tk.LEFT, padx=(0, 6)
        )
        self.key_listbox = tk.Listbox(key_frame, height=4, selectmode=tk.SINGLE)
        self.key_listbox.pack(fill=tk.X, pady=(6, 0))
        ttk.Button(key_frame, text="Remove selected", command=self._remove_key).pack(
            anchor=tk.W, pady=(4, 0)
        )

        # Delay
        delay_frame = ttk.Frame(tab)
        delay_frame.pack(fill=tk.X, pady=(0, 8))
        ttk.Label(delay_frame, text="Min delay (ms):").pack(side=tk.LEFT, padx=(0, 6))
        self.min_delay_var = tk.StringVar(value="1000")
        ttk.Entry(delay_frame, textvariable=self.min_delay_var, width=10).pack(
            side=tk.LEFT, padx=(0, 16)
        )
        ttk.Label(delay_frame, text="Max delay (ms):").pack(side=tk.LEFT, padx=(0, 6))
        self.max_delay_var = tk.StringVar(value="1500")
        ttk.Entry(delay_frame, textvariable=self.max_delay_var, width=10).pack(
            side=tk.LEFT
        )

        # Start / Stop
        btn_frame = ttk.Frame(tab)
        btn_frame.pack(fill=tk.X, pady=(8, 0))
        self.start_btn = ttk.Button(btn_frame, text="Start", command=self._on_start)
        self.start_btn.pack(side=tk.LEFT, padx=(0, 8))
        self.stop_btn = ttk.Button(
            btn_frame, text="Stop", command=self._on_stop, state=tk.DISABLED
        )
        self.stop_btn.pack(side=tk.LEFT)
        ttk.Label(btn_frame, text="  F12 = Start/Stop", foreground="gray").pack(
            side=tk.LEFT, padx=(16, 0)
        )

        # Status
        self.status_var = tk.StringVar(value="Add keys and set delay.")
        ttk.Label(tab, textvariable=self.status_var, foreground="gray").pack(
            anchor=tk.W, pady=(8, 0)
        )

    # ---- Tab 2: Conditional Clicker ----------------------------------

    def _build_conditional_tab(self):
        tab = ttk.Frame(self.notebook, padding=10)
        self.notebook.add(tab, text="Conditional Clicker")

        if not CONDITIONAL_AVAILABLE:
            ttk.Label(
                tab,
                text=(
                    "Conditional Clicker requires mss and Pillow.\n"
                    "Run:  pip install mss Pillow"
                ),
                foreground="red",
            ).pack(pady=20)
            return

        # --- Region selection ---
        region_frame = ttk.LabelFrame(tab, text="Screen Region", padding=6)
        region_frame.pack(fill=tk.X, pady=(0, 6))

        btn_row = ttk.Frame(region_frame)
        btn_row.pack(fill=tk.X)
        ttk.Button(btn_row, text="Select Region", command=self._select_region).pack(
            side=tk.LEFT, padx=(0, 8)
        )
        self.region_status_var = tk.StringVar(value="No region selected")
        ttk.Label(btn_row, textvariable=self.region_status_var, foreground="gray").pack(
            side=tk.LEFT
        )

        coord_row = ttk.Frame(region_frame)
        coord_row.pack(fill=tk.X, pady=(4, 0))
        self.region_x_var = tk.StringVar(value="1106")
        self.region_y_var = tk.StringVar(value="18")
        self.region_w_var = tk.StringVar(value="67")
        self.region_h_var = tk.StringVar(value="17")
        for label, var in [
            ("X:", self.region_x_var),
            ("Y:", self.region_y_var),
            ("W:", self.region_w_var),
            ("H:", self.region_h_var),
        ]:
            ttk.Label(coord_row, text=label).pack(side=tk.LEFT, padx=(0, 2))
            ttk.Entry(coord_row, textvariable=var, width=6).pack(
                side=tk.LEFT, padx=(0, 8)
            )
        ttk.Button(coord_row, text="Apply", command=self._apply_region_coords).pack(
            side=tk.LEFT
        )

        # --- Live Preview + Detected Colors ---
        mid_frame = ttk.Frame(tab)
        mid_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 6))

        preview_frame = ttk.LabelFrame(mid_frame, text="Live Preview", padding=6)
        preview_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0, 4))
        self.preview_canvas = tk.Canvas(
            preview_frame, width=200, height=100, bg="black"
        )
        self.preview_canvas.pack(fill=tk.BOTH, expand=True)
        self.preview_btn = ttk.Button(
            preview_frame, text="Start Preview", command=self._toggle_preview
        )
        self.preview_btn.pack(pady=(4, 0))

        colors_frame = ttk.LabelFrame(mid_frame, text="Detected Colors", padding=6)
        colors_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(4, 0))
        colors_scroll = ttk.Frame(colors_frame)
        colors_scroll.pack(fill=tk.BOTH, expand=True)
        self.colors_canvas = tk.Canvas(colors_scroll, width=140, height=100)
        scrollbar = ttk.Scrollbar(
            colors_scroll, orient=tk.VERTICAL, command=self.colors_canvas.yview
        )
        self.colors_inner = ttk.Frame(self.colors_canvas)
        self.colors_inner.bind(
            "<Configure>",
            lambda _e: self.colors_canvas.configure(
                scrollregion=self.colors_canvas.bbox("all")
            ),
        )
        self.colors_canvas.create_window(
            (0, 0), window=self.colors_inner, anchor=tk.NW
        )
        self.colors_canvas.configure(yscrollcommand=scrollbar.set)
        self.colors_canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        # --- Trigger configuration ---
        trigger_frame = ttk.LabelFrame(tab, text="Trigger Configuration", padding=6)
        trigger_frame.pack(fill=tk.X, pady=(0, 6))

        trig_row1 = ttk.Frame(trigger_frame)
        trig_row1.pack(fill=tk.X, pady=(0, 4))
        ttk.Label(trig_row1, text="Stuck timeout (s):").pack(
            side=tk.LEFT, padx=(0, 4)
        )
        self.stuck_timeout_var = tk.StringVar(value="20")
        ttk.Entry(trig_row1, textvariable=self.stuck_timeout_var, width=6).pack(
            side=tk.LEFT, padx=(0, 12)
        )
        self.hp_live_var = tk.StringVar(value="HP: —")
        ttk.Label(trig_row1, textvariable=self.hp_live_var, foreground="blue").pack(
            side=tk.LEFT, padx=(12, 0)
        )

        # Unstuck movement sequence (triggered after 2 consecutive stucks)
        unstuck_row1 = ttk.Frame(trigger_frame)
        unstuck_row1.pack(fill=tk.X, pady=(0, 4))
        ttk.Label(unstuck_row1, text="Unstuck key 1:").pack(
            side=tk.LEFT, padx=(0, 4)
        )
        self.unstuck_key1_var = tk.StringVar(value="left")
        ttk.Combobox(
            unstuck_row1,
            textvariable=self.unstuck_key1_var,
            values=KEY_OPTIONS,
            width=10,
            state="readonly",
        ).pack(side=tk.LEFT, padx=(0, 12))
        ttk.Label(unstuck_row1, text="Hold (ms):").pack(side=tk.LEFT, padx=(0, 4))
        self.unstuck_dur1_var = tk.StringVar(value="1000")
        ttk.Entry(unstuck_row1, textvariable=self.unstuck_dur1_var, width=6).pack(
            side=tk.LEFT, padx=(0, 12)
        )
        ttk.Label(unstuck_row1, text="Key 2:").pack(side=tk.LEFT, padx=(0, 4))
        self.unstuck_key2_var = tk.StringVar(value="up")
        ttk.Combobox(
            unstuck_row1,
            textvariable=self.unstuck_key2_var,
            values=KEY_OPTIONS,
            width=10,
            state="readonly",
        ).pack(side=tk.LEFT, padx=(0, 12))
        ttk.Label(unstuck_row1, text="Hold (ms):").pack(side=tk.LEFT, padx=(0, 4))
        self.unstuck_dur2_var = tk.StringVar(value="1000")
        ttk.Entry(unstuck_row1, textvariable=self.unstuck_dur2_var, width=6).pack(
            side=tk.LEFT
        )

        # Targeting key (pressed when no HP / stuck)
        target_row = ttk.Frame(trigger_frame)
        target_row.pack(fill=tk.X, pady=(0, 4))
        ttk.Label(target_row, text="Targeting key:").pack(
            side=tk.LEFT, padx=(0, 4)
        )
        self.target_key_var = tk.StringVar(value="f4")
        ttk.Combobox(
            target_row,
            textvariable=self.target_key_var,
            values=KEY_OPTIONS,
            width=10,
            state="readonly",
        ).pack(side=tk.LEFT, padx=(0, 12))
        ttk.Label(target_row, text="Min (ms):").pack(side=tk.LEFT, padx=(0, 4))
        self.target_min_var = tk.StringVar(value="200")
        ttk.Entry(target_row, textvariable=self.target_min_var, width=6).pack(
            side=tk.LEFT, padx=(0, 12)
        )
        ttk.Label(target_row, text="Max (ms):").pack(side=tk.LEFT, padx=(0, 4))
        self.target_max_var = tk.StringVar(value="400")
        ttk.Entry(target_row, textvariable=self.target_max_var, width=6).pack(
            side=tk.LEFT
        )

        # No-target timeout
        notarget_row = ttk.Frame(trigger_frame)
        notarget_row.pack(fill=tk.X, pady=(0, 4))
        ttk.Label(notarget_row, text="No-target timeout (s):").pack(
            side=tk.LEFT, padx=(0, 4)
        )
        self.no_target_timeout_var = tk.StringVar(value="120")
        ttk.Entry(notarget_row, textvariable=self.no_target_timeout_var, width=6).pack(
            side=tk.LEFT, padx=(0, 12)
        )
        ttk.Label(
            notarget_row, text="  (stop bot if no target found; 0 = disabled)",
            foreground="gray",
        ).pack(side=tk.LEFT, padx=(4, 0))

        # Attack start delay (wait after targeting before attacks begin)
        engage_row = ttk.Frame(trigger_frame)
        engage_row.pack(fill=tk.X, pady=(0, 4))
        ttk.Label(engage_row, text="Attack start delay (ms):").pack(
            side=tk.LEFT, padx=(0, 4)
        )
        self.engage_delay_var = tk.StringVar(value="2500")
        ttk.Entry(engage_row, textvariable=self.engage_delay_var, width=8).pack(
            side=tk.LEFT
        )
        ttk.Label(
            engage_row, text="  (time to run to mob / cast status effects)",
            foreground="gray",
        ).pack(side=tk.LEFT, padx=(4, 0))

        # HP verification settings (to avoid false 0 / None OCR readings)
        hp_verify_frame = ttk.LabelFrame(trigger_frame, text="HP Reading Verification", padding=4)
        hp_verify_frame.pack(fill=tk.X, pady=(4, 4))

        hpv_row1 = ttk.Frame(hp_verify_frame)
        hpv_row1.pack(fill=tk.X, pady=(0, 2))
        ttk.Label(hpv_row1, text="Confirm zero/gone count:").pack(
            side=tk.LEFT, padx=(0, 4)
        )
        self.hp_confirm_count_var = tk.StringVar(value="3")
        ttk.Entry(hpv_row1, textvariable=self.hp_confirm_count_var, width=4).pack(
            side=tk.LEFT, padx=(0, 12)
        )
        ttk.Label(
            hpv_row1,
            text="(consecutive 0/None reads before acting)",
            foreground="gray",
        ).pack(side=tk.LEFT)

        hpv_row2 = ttk.Frame(hp_verify_frame)
        hpv_row2.pack(fill=tk.X)
        ttk.Label(hpv_row2, text="OCR threshold:").pack(
            side=tk.LEFT, padx=(0, 4)
        )
        self.ocr_threshold_var = tk.StringVar(value="180")
        ttk.Entry(hpv_row2, textvariable=self.ocr_threshold_var, width=5).pack(
            side=tk.LEFT, padx=(0, 12)
        )
        ttk.Label(hpv_row2, text="OCR scale:").pack(
            side=tk.LEFT, padx=(0, 4)
        )
        self.ocr_scale_var = tk.StringVar(value="3")
        ttk.Entry(hpv_row2, textvariable=self.ocr_scale_var, width=4).pack(
            side=tk.LEFT, padx=(0, 12)
        )
        ttk.Label(
            hpv_row2,
            text="(binary threshold 0-255; upscale factor)",
            foreground="gray",
        ).pack(side=tk.LEFT)

        # Attack keys (pressed while monster is targeted) – dynamic list
        attack_frame = ttk.LabelFrame(trigger_frame, text="Attack Keys", padding=4)
        attack_frame.pack(fill=tk.X, pady=(4, 0))

        self.attack_keys_container = ttk.Frame(attack_frame)
        self.attack_keys_container.pack(fill=tk.X)

        self.attack_key_rows = []
        self._add_attack_key_row("f1", "200", "1000")
        self._add_attack_key_row("f2", "200", "1000")

        ttk.Button(
            attack_frame, text="+ Add Attack Key", command=self._add_attack_key_row
        ).pack(anchor=tk.W, pady=(4, 0))

        # Status effect keys (pressed immediately on target, then re-applied)
        status_frame = ttk.LabelFrame(trigger_frame, text="Status Effect Keys", padding=4)
        status_frame.pack(fill=tk.X, pady=(4, 0))

        self.status_effect_keys_container = ttk.Frame(status_frame)
        self.status_effect_keys_container.pack(fill=tk.X)

        self.status_effect_key_rows = []

        ttk.Button(
            status_frame, text="+ Add Status Effect Key", command=self._add_status_effect_key_row
        ).pack(anchor=tk.W, pady=(4, 0))

        # Buffs (same template logic, checked on a timer only — not on new target)
        buff_frame = ttk.LabelFrame(trigger_frame, text="Buffs (check periodically)", padding=4)
        buff_frame.pack(fill=tk.X, pady=(4, 0))
        buff_interval_row = ttk.Frame(buff_frame)
        buff_interval_row.pack(fill=tk.X)
        ttk.Label(buff_interval_row, text="Check every (seconds):").pack(side=tk.LEFT, padx=(0, 4))
        self.buff_interval_var = tk.StringVar(value="10")
        ttk.Entry(buff_interval_row, textvariable=self.buff_interval_var, width=5).pack(side=tk.LEFT)
        self.buff_keys_container = ttk.Frame(buff_frame)
        self.buff_keys_container.pack(fill=tk.X)
        self.buff_key_rows = []
        ttk.Button(
            buff_frame, text="+ Add Buff Key", command=self._add_buff_key_row
        ).pack(anchor=tk.W, pady=(4, 0))

        # On-death key (pressed once when mob dies)
        death_frame = ttk.Frame(trigger_frame)
        death_frame.pack(fill=tk.X, pady=(6, 0))
        self.death_enabled_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(
            death_frame, text="On mob death, press:",
            variable=self.death_enabled_var,
        ).pack(side=tk.LEFT, padx=(0, 4))
        self.death_key_var = tk.StringVar(value="f3")
        ttk.Combobox(
            death_frame,
            textvariable=self.death_key_var,
            values=KEY_OPTIONS,
            width=10,
            state="readonly",
        ).pack(side=tk.LEFT)
        ttk.Label(death_frame, text="  Delay after (ms):").pack(
            side=tk.LEFT, padx=(8, 4)
        )
        self.death_delay_var = tk.StringVar(value="1000")
        ttk.Entry(death_frame, textvariable=self.death_delay_var, width=6).pack(
            side=tk.LEFT
        )

        # --- Monitor controls ---
        ctrl_frame = ttk.Frame(tab)
        ctrl_frame.pack(fill=tk.X, pady=(4, 0))
        self.monitor_start_btn = ttk.Button(
            ctrl_frame, text="Start Monitoring", command=self._start_monitoring
        )
        self.monitor_start_btn.pack(side=tk.LEFT, padx=(0, 8))
        self.monitor_stop_btn = ttk.Button(
            ctrl_frame,
            text="Stop Monitoring",
            command=self._stop_monitoring,
            state=tk.DISABLED,
        )
        self.monitor_stop_btn.pack(side=tk.LEFT)
        self.monitor_status_var = tk.StringVar(value="Idle")
        ttk.Label(
            ctrl_frame, textvariable=self.monitor_status_var, foreground="gray"
        ).pack(side=tk.LEFT, padx=(16, 0))

    # ---- Tab 3: Status Effects Library --------------------------------

    def _build_templates_tab(self):
        tab = ttk.Frame(self.notebook, padding=10)
        self.notebook.add(tab, text="Status Effects Library")

        if not CONDITIONAL_AVAILABLE:
            ttk.Label(
                tab,
                text="Requires mss, Pillow, opencv-python, easyocr.\nRun:  pip install -r requirements.txt",
                foreground="red",
            ).pack(pady=20)
            return

        paned = ttk.PanedWindow(tab, orient=tk.HORIZONTAL)
        paned.pack(fill=tk.BOTH, expand=True)

        # -- Left: template list --
        left = ttk.LabelFrame(paned, text="Templates", padding=6)
        paned.add(left, weight=1)

        self.tpl_listbox = tk.Listbox(left, height=12, exportselection=False)
        self.tpl_listbox.pack(fill=tk.BOTH, expand=True)
        self.tpl_listbox.bind("<<ListboxSelect>>", self._on_template_select)

        # -- Right: preview --
        right = ttk.LabelFrame(paned, text="Preview", padding=6)
        paned.add(right, weight=1)

        self.tpl_preview_canvas = tk.Canvas(right, width=150, height=150, bg="black")
        self.tpl_preview_canvas.pack(fill=tk.BOTH, expand=True)
        self.tpl_name_var = tk.StringVar(value="")
        ttk.Label(right, textvariable=self.tpl_name_var, font=("TkDefaultFont", 10, "bold")).pack(
            pady=(4, 0)
        )
        self._tpl_preview_ref = None

        # -- Bottom toolbar --
        toolbar = ttk.Frame(tab)
        toolbar.pack(fill=tk.X, pady=(6, 0))
        ttk.Button(toolbar, text="Capture New", command=self._capture_new_template).pack(
            side=tk.LEFT, padx=(0, 6)
        )
        ttk.Button(toolbar, text="Delete", command=self._delete_selected_template).pack(
            side=tk.LEFT, padx=(0, 6)
        )
        ttk.Button(toolbar, text="Test Match", command=self._test_match_template).pack(
            side=tk.LEFT
        )

        self._populate_template_list()

    def _populate_template_list(self):
        self.tpl_listbox.delete(0, tk.END)
        self._template_items = list_templates()
        for t in self._template_items:
            self.tpl_listbox.insert(tk.END, f"{t['name']}  ({t['slug']})")

    def _selected_template_slug(self):
        sel = self.tpl_listbox.curselection()
        if not sel:
            return None
        return self._template_items[sel[0]]["slug"]

    def _on_template_select(self, _event=None):
        slug = self._selected_template_slug()
        if not slug:
            return
        img = load_template(slug)
        if img is None:
            return
        cw = max(self.tpl_preview_canvas.winfo_width(), 100)
        ch = max(self.tpl_preview_canvas.winfo_height(), 100)
        iw, ih = img.size
        scale = min(cw / max(iw, 1), ch / max(ih, 1), 4.0)
        display = img.resize((max(1, int(iw * scale)), max(1, int(ih * scale))), Image.NEAREST)
        self._tpl_preview_ref = ImageTk.PhotoImage(display)
        self.tpl_preview_canvas.delete("all")
        self.tpl_preview_canvas.create_image(cw // 2, ch // 2, anchor=tk.CENTER, image=self._tpl_preview_ref)
        meta = [t for t in self._template_items if t["slug"] == slug]
        self.tpl_name_var.set(meta[0]["name"] if meta else slug)

    def _capture_new_template(self):
        def _on_captured(sx, sy, sw, sh):
            name = _ask_template_name(self.root)
            if not name:
                return
            img = capture_region(sx, sy, sw, sh)
            save_template(name, img, region=(sx, sy, sw, sh))
            self._populate_template_list()
            self._refresh_template_combos()

        RegionSelector(self.root, _on_captured)

    def _delete_selected_template(self):
        slug = self._selected_template_slug()
        if not slug:
            messagebox.showinfo("No selection", "Select a template to delete.")
            return
        delete_template(slug)
        self._populate_template_list()
        self._refresh_template_combos()
        self.tpl_preview_canvas.delete("all")
        self.tpl_name_var.set("")

    def _test_match_template(self):
        slug = self._selected_template_slug()
        if not slug:
            messagebox.showinfo("No selection", "Select a template to test.")
            return
        tpl_img = load_template(slug)
        if tpl_img is None:
            messagebox.showwarning("Missing", "Template image not found on disk.")
            return
        meta = [t for t in self._template_items if t["slug"] == slug]
        region = meta[0].get("region") if meta else None
        if not region or len(region) != 4:
            messagebox.showinfo("No region", "This template has no saved capture region.")
            return
        try:
            screen_img = capture_region(*region)
        except Exception as e:
            messagebox.showerror("Capture error", str(e))
            return
        is_match = match_template(screen_img, tpl_img, threshold=0.8)
        messagebox.showinfo(
            "Test Match",
            f"Template: {meta[0]['name']}\n"
            f"Region: {region}\n"
            f"Match: {'YES' if is_match else 'NO'}",
        )

    # ------------------------------------------------------------------ #
    #  Shared: serial / port helpers                                       #
    # ------------------------------------------------------------------ #

    def _refresh_ports(self):
        ports, is_pico = find_pico_ports()
        self.port_combo["values"] = [disp for _dev, disp in ports]
        if ports:
            self.port_combo.current(0)
            self.conn_status_var.set(
                "Pico port found."
                if is_pico
                else "No Pico auto-detected. Select your port above."
            )
        else:
            self.port_var.set("")
            self.conn_status_var.set(
                "No COM ports found. Connect device and Refresh."
            )

    def _get_port_device(self):
        val = self.port_var.get()
        if not val:
            return None
        return val.split(" ")[0].strip()

    def _serial_send(self, msg):
        with self._lock:
            if self.serial_port is None or not self.serial_port.is_open:
                return False
            try:
                self.serial_port.write((msg + "\n").encode("ascii"))
                self.serial_port.flush()
                return True
            except Exception:
                return False

    def _open_serial(self):
        with self._lock:
            if self.serial_port is not None and self.serial_port.is_open:
                return True
            self.serial_port = None
        dev = self._get_port_device()
        if not dev:
            messagebox.showwarning("No port", "Select a COM port first.")
            return False
        try:
            with self._lock:
                if self.serial_port is not None and self.serial_port.is_open:
                    return True
                self.serial_port = serial.Serial(
                    dev, BAUD, timeout=0.1, write_timeout=1
                )
                self.serial_port.dtr = False
                self.serial_port.rts = False
            return True
        except serial.SerialException as e:
            err = str(e).lower()
            if "already open" in err or "access" in err or "permission" in err:
                messagebox.showerror(
                    "Port in use",
                    "COM port is already in use.\n\n"
                    "Try: 1) Unplug the Pico, wait 2 s, plug back in.\n"
                    "     2) Close other apps using the port.\n"
                    "     3) Click Refresh and try again.",
                )
            else:
                messagebox.showerror("Serial error", str(e))
            return False
        except Exception as e:
            messagebox.showerror("Serial error", str(e))
            return False

    def _close_serial(self):
        with self._lock:
            if self.serial_port and self.serial_port.is_open:
                try:
                    self.serial_port.close()
                except Exception:
                    pass
                self.serial_port = None

    # ------------------------------------------------------------------ #
    #  Tab 1: Autoclicker logic                                            #
    # ------------------------------------------------------------------ #

    def _get_keys(self):
        keys = []
        for i in range(self.key_listbox.size()):
            keys.append(self.key_listbox.get(i))
        return keys

    def _add_key(self):
        k = self.key_var.get().strip().lower()
        if k and k in KEY_OPTIONS:
            self.key_listbox.insert(tk.END, k)
            self.status_var.set(f"Added key: {k}")
        elif k:
            messagebox.showwarning(
                "Unknown key", f"Key '{k}' not in list. Choose from dropdown."
            )

    def _remove_key(self):
        sel = self.key_listbox.curselection()
        if sel:
            self.key_listbox.delete(sel[0])

    def _on_start(self):
        keys = self._get_keys()
        if not keys:
            messagebox.showwarning("No keys", "Add at least one key.")
            return
        try:
            min_ms = int(self.min_delay_var.get())
            max_ms = int(self.max_delay_var.get())
        except ValueError:
            messagebox.showwarning(
                "Invalid delay", "Min and max delay must be numbers (ms)."
            )
            return
        if min_ms <= 0 or max_ms < min_ms:
            messagebox.showwarning("Invalid delay", "Min delay > 0 and max >= min.")
            return

        if self.serial_port is None or not self.serial_port.is_open:
            if not self._open_serial():
                return
            self.start_btn.config(state=tk.DISABLED)
            self.status_var.set("Starting...")
            self.root.after(
                500, lambda: self._send_start_after_delay(keys, min_ms, max_ms)
            )
            return
        self._do_send_start(keys, min_ms, max_ms)

    def _send_start_after_delay(self, keys, min_ms, max_ms):
        """Called after a short pause so the Pico is ready."""
        self._do_send_start(keys, min_ms, max_ms)

    def _do_send_start(self, keys, min_ms, max_ms):
        keys_str = ",".join(keys)
        cmd = f"START;{keys_str};{min_ms};{max_ms}"
        if not self._serial_send(cmd):
            self.start_btn.config(state=tk.NORMAL)
            messagebox.showerror(
                "Send failed", "Could not send START to Pico. Check connection."
            )
            return
        self.running = True
        self.start_btn.config(state=tk.DISABLED)
        self.stop_btn.config(state=tk.NORMAL)
        self.status_var.set("Running. Press F12 or Stop to stop.")

    def _on_stop(self):
        self._serial_send("STOP")
        self.running = False
        self.start_btn.config(state=tk.NORMAL)
        self.stop_btn.config(state=tk.DISABLED)
        self.status_var.set("Stopped.")

    def _on_f12(self):
        if self.running:
            self.root.after(0, self._on_stop)
        else:
            self.root.after(0, self._do_start_from_f12)

    def _do_start_from_f12(self):
        keys = self._get_keys()
        if not keys:
            self.status_var.set("Add at least one key before starting.")
            return
        try:
            min_ms = int(self.min_delay_var.get())
            max_ms = int(self.max_delay_var.get())
        except ValueError:
            self.status_var.set("Set valid min/max delay (numbers).")
            return
        if min_ms <= 0 or max_ms < min_ms:
            return
        if self.serial_port is None or not self.serial_port.is_open:
            if not self._open_serial():
                return
        cmd = f"START;{','.join(keys)};{min_ms};{max_ms}"
        if self._serial_send(cmd):
            self.running = True
            self.start_btn.config(state=tk.DISABLED)
            self.stop_btn.config(state=tk.NORMAL)
            self.status_var.set("Running (F12 to stop).")
        else:
            self.status_var.set("Failed to send START.")

    def _start_f12_listener(self):
        try:
            from pynput import keyboard as pynput_kb

            def on_press(key):
                try:
                    if key == pynput_kb.Key.f12:
                        self._on_f12()
                except AttributeError:
                    pass

            listener = pynput_kb.Listener(on_press=on_press)
            listener.daemon = True
            listener.start()
        except ImportError:
            self.conn_status_var.set(
                "pynput not installed; F12 hotkey disabled. pip install pynput"
            )

    # ------------------------------------------------------------------ #
    #  Tab 2: Conditional Clicker logic                                    #
    # ------------------------------------------------------------------ #

    def _select_region(self):
        RegionSelector(self.root, self._on_region_selected)

    def _on_region_selected(self, x, y, w, h):
        self.region = (x, y, w, h)
        self.region_x_var.set(str(x))
        self.region_y_var.set(str(y))
        self.region_w_var.set(str(w))
        self.region_h_var.set(str(h))
        self.region_status_var.set(f"Region: {x},{y} {w}x{h}")

    def _apply_region_coords(self):
        try:
            x = int(self.region_x_var.get())
            y = int(self.region_y_var.get())
            w = int(self.region_w_var.get())
            h = int(self.region_h_var.get())
        except ValueError:
            messagebox.showwarning("Invalid", "X, Y, W, H must be integers.")
            return
        if w <= 0 or h <= 0:
            messagebox.showwarning("Invalid", "Width and height must be > 0.")
            return
        self.region = (x, y, w, h)
        self.region_status_var.set(f"Region: {x},{y} {w}x{h}")

    # -- preview --

    def _toggle_preview(self):
        if self.preview_active:
            self.preview_active = False
            self.preview_btn.config(text="Start Preview")
        else:
            if not self.region:
                messagebox.showwarning("No region", "Select a screen region first.")
                return
            self.preview_active = True
            self.preview_btn.config(text="Stop Preview")
            self._update_preview()

    def _update_preview(self):
        if not self.preview_active or not self.region:
            return

        x, y, w, h = self.region
        try:
            image = capture_region(x, y, w, h)
        except Exception:
            self.root.after(500, self._update_preview)
            return

        cw = self.preview_canvas.winfo_width()
        ch = self.preview_canvas.winfo_height()
        if cw > 1 and ch > 1:
            scale = min(cw / w, ch / h)
            new_w = max(1, int(w * scale))
            new_h = max(1, int(h * scale))
            display = image.resize((new_w, new_h), Image.NEAREST)
        else:
            display = image

        self._preview_image_ref = ImageTk.PhotoImage(display)
        self.preview_canvas.delete("all")
        self.preview_canvas.create_image(
            cw // 2, ch // 2, anchor=tk.CENTER, image=self._preview_image_ref
        )

        colors = get_unique_colors(image, tolerance=30)
        self._display_colors(colors)

        self.root.after(500, self._update_preview)

    def _display_colors(self, colors):
        for widget in self.colors_inner.winfo_children():
            widget.destroy()
        for hex_color in colors[:50]:
            row = ttk.Frame(self.colors_inner)
            row.pack(fill=tk.X, pady=1)
            swatch = tk.Canvas(row, width=16, height=16, highlightthickness=0)
            swatch.pack(side=tk.LEFT, padx=(0, 4))
            try:
                swatch.configure(bg=hex_color)
            except tk.TclError:
                swatch.configure(bg="black")
            lbl = ttk.Label(row, text=hex_color, font=("Consolas", 9))
            lbl.pack(side=tk.LEFT)

    # -- attack key rows --

    def _add_attack_key_row(self, key="f1", min_ms="200", max_ms="1000"):
        row_frame = ttk.Frame(self.attack_keys_container)
        row_frame.pack(fill=tk.X, pady=1)

        key_var = tk.StringVar(value=key)
        min_var = tk.StringVar(value=min_ms)
        max_var = tk.StringVar(value=max_ms)

        ttk.Label(row_frame, text="Key:").pack(side=tk.LEFT, padx=(0, 2))
        ttk.Combobox(
            row_frame, textvariable=key_var, values=KEY_OPTIONS,
            width=8, state="readonly",
        ).pack(side=tk.LEFT, padx=(0, 8))
        ttk.Label(row_frame, text="Min (ms):").pack(side=tk.LEFT, padx=(0, 2))
        ttk.Entry(row_frame, textvariable=min_var, width=6).pack(
            side=tk.LEFT, padx=(0, 8)
        )
        ttk.Label(row_frame, text="Max (ms):").pack(side=tk.LEFT, padx=(0, 2))
        ttk.Entry(row_frame, textvariable=max_var, width=6).pack(
            side=tk.LEFT, padx=(0, 8)
        )

        entry = {"frame": row_frame, "key": key_var, "min": min_var, "max": max_var}
        self.attack_key_rows.append(entry)

        def _remove(e=entry):
            if len(self.attack_key_rows) <= 1:
                messagebox.showinfo("Cannot remove", "At least one attack key is required.")
                return
            e["frame"].destroy()
            self.attack_key_rows.remove(e)

        ttk.Button(row_frame, text="✕", width=3, command=_remove).pack(side=tk.LEFT)

    # -- status effect key rows --

    def _add_status_effect_key_row(
        self, key="f1", region_x="0", region_y="0", region_w="50", region_h="50",
        template_slug="", match_threshold="0.80", retry_min="1000", retry_max="2000",
    ):
        outer_frame = ttk.Frame(self.status_effect_keys_container)
        outer_frame.pack(fill=tk.X, pady=(2, 4))

        key_var = tk.StringVar(value=key)
        rx_var = tk.StringVar(value=region_x)
        ry_var = tk.StringVar(value=region_y)
        rw_var = tk.StringVar(value=region_w)
        rh_var = tk.StringVar(value=region_h)
        template_var = tk.StringVar(value=template_slug)
        threshold_var = tk.StringVar(value=match_threshold)
        retry_min_var = tk.StringVar(value=retry_min)
        retry_max_var = tk.StringVar(value=retry_max)

        row1 = ttk.Frame(outer_frame)
        row1.pack(fill=tk.X)
        ttk.Label(row1, text="Key:").pack(side=tk.LEFT, padx=(0, 2))
        ttk.Combobox(
            row1, textvariable=key_var, values=KEY_OPTIONS,
            width=8, state="readonly",
        ).pack(side=tk.LEFT, padx=(0, 8))
        for label, var in [
            ("X:", rx_var), ("Y:", ry_var), ("W:", rw_var), ("H:", rh_var),
        ]:
            ttk.Label(row1, text=label).pack(side=tk.LEFT, padx=(0, 2))
            ttk.Entry(row1, textvariable=var, width=5).pack(side=tk.LEFT, padx=(0, 4))

        def _select_se_region(rxv=rx_var, ryv=ry_var, rwv=rw_var, rhv=rh_var):
            def _cb(sx, sy, sw, sh):
                rxv.set(str(sx)); ryv.set(str(sy))
                rwv.set(str(sw)); rhv.set(str(sh))
            RegionSelector(self.root, _cb)

        ttk.Button(row1, text="Select Region", command=_select_se_region).pack(
            side=tk.LEFT, padx=(4, 0)
        )

        def _remove(e_ref=[None]):
            e_ref[0]["frame"].destroy()
            self.status_effect_key_rows.remove(e_ref[0])

        ttk.Button(row1, text="✕", width=3, command=_remove).pack(side=tk.RIGHT)

        row2 = ttk.Frame(outer_frame)
        row2.pack(fill=tk.X, pady=(2, 0))
        ttk.Label(row2, text="Template:").pack(side=tk.LEFT, padx=(0, 2))
        tpl_slugs = [t["slug"] for t in list_templates()]
        template_combo = ttk.Combobox(
            row2, textvariable=template_var, values=tpl_slugs, width=14,
        )
        template_combo.pack(side=tk.LEFT, padx=(0, 8))

        ttk.Label(row2, text="Threshold:").pack(side=tk.LEFT, padx=(0, 2))
        ttk.Entry(row2, textvariable=threshold_var, width=5).pack(side=tk.LEFT, padx=(0, 8))
        ttk.Label(row2, text="Retry Min (ms):").pack(side=tk.LEFT, padx=(0, 2))
        ttk.Entry(row2, textvariable=retry_min_var, width=6).pack(side=tk.LEFT, padx=(0, 4))
        ttk.Label(row2, text="Max:").pack(side=tk.LEFT, padx=(0, 2))
        ttk.Entry(row2, textvariable=retry_max_var, width=6).pack(side=tk.LEFT)

        status_lbl = ttk.Label(row2, text="", width=6)
        status_lbl.pack(side=tk.LEFT, padx=(8, 0))

        entry = {
            "frame": outer_frame, "key": key_var,
            "rx": rx_var, "ry": ry_var, "rw": rw_var, "rh": rh_var,
            "template": template_var, "threshold": threshold_var,
            "retry_min": retry_min_var, "retry_max": retry_max_var,
            "template_combo": template_combo,
            "status_label": status_lbl,
        }
        self.status_effect_key_rows.append(entry)
        _remove.__defaults__ = ([entry],)

    # -- buff key rows (same as status effects but checked on a timer, not on new target) --

    def _add_buff_key_row(
        self, key="f1", region_x="0", region_y="0", region_w="50", region_h="50",
        template_slug="", match_threshold="0.80",
    ):
        outer_frame = ttk.Frame(self.buff_keys_container)
        outer_frame.pack(fill=tk.X, pady=(2, 4))

        key_var = tk.StringVar(value=key)
        rx_var = tk.StringVar(value=region_x)
        ry_var = tk.StringVar(value=region_y)
        rw_var = tk.StringVar(value=region_w)
        rh_var = tk.StringVar(value=region_h)
        template_var = tk.StringVar(value=template_slug)
        threshold_var = tk.StringVar(value=match_threshold)

        row1 = ttk.Frame(outer_frame)
        row1.pack(fill=tk.X)
        ttk.Label(row1, text="Key:").pack(side=tk.LEFT, padx=(0, 2))
        ttk.Combobox(
            row1, textvariable=key_var, values=KEY_OPTIONS,
            width=8, state="readonly",
        ).pack(side=tk.LEFT, padx=(0, 8))
        for label, var in [
            ("X:", rx_var), ("Y:", ry_var), ("W:", rw_var), ("H:", rh_var),
        ]:
            ttk.Label(row1, text=label).pack(side=tk.LEFT, padx=(0, 2))
            ttk.Entry(row1, textvariable=var, width=5).pack(side=tk.LEFT, padx=(0, 4))

        def _select_buff_region(rxv=rx_var, ryv=ry_var, rwv=rw_var, rhv=rh_var):
            def _cb(sx, sy, sw, sh):
                rxv.set(str(sx)); ryv.set(str(sy))
                rwv.set(str(sw)); rhv.set(str(sh))
            RegionSelector(self.root, _cb)

        ttk.Button(row1, text="Select Region", command=_select_buff_region).pack(
            side=tk.LEFT, padx=(4, 0)
        )

        def _remove(e_ref=[None]):
            e_ref[0]["frame"].destroy()
            self.buff_key_rows.remove(e_ref[0])

        ttk.Button(row1, text="✕", width=3, command=_remove).pack(side=tk.RIGHT)

        row2 = ttk.Frame(outer_frame)
        row2.pack(fill=tk.X, pady=(2, 0))
        ttk.Label(row2, text="Template:").pack(side=tk.LEFT, padx=(0, 2))
        tpl_slugs = [t["slug"] for t in list_templates()]
        template_combo = ttk.Combobox(
            row2, textvariable=template_var, values=tpl_slugs, width=14,
        )
        template_combo.pack(side=tk.LEFT, padx=(0, 8))
        ttk.Label(row2, text="Threshold:").pack(side=tk.LEFT, padx=(0, 2))
        ttk.Entry(row2, textvariable=threshold_var, width=5).pack(side=tk.LEFT)

        entry = {
            "frame": outer_frame, "key": key_var,
            "rx": rx_var, "ry": ry_var, "rw": rw_var, "rh": rh_var,
            "template": template_var, "threshold": threshold_var,
            "template_combo": template_combo,
        }
        self.buff_key_rows.append(entry)
        _remove.__defaults__ = ([entry],)

    def _refresh_template_combos(self):
        """Refresh the values list on every status-effect and buff template combobox."""
        slugs = [t["slug"] for t in list_templates()]
        for row in self.status_effect_key_rows:
            combo = row.get("template_combo")
            if combo:
                combo["values"] = slugs
        for row in self.buff_key_rows:
            combo = row.get("template_combo")
            if combo:
                combo["values"] = slugs

    # -- monitoring --

    def _start_monitoring(self):
        if not self.region:
            messagebox.showwarning("No region", "Select a screen region first.")
            return
        try:
            stuck_s = float(self.stuck_timeout_var.get())
        except ValueError:
            messagebox.showwarning("Invalid", "Stuck timeout must be a number (s).")
            return
        if stuck_s <= 0:
            messagebox.showwarning("Invalid", "Stuck timeout must be > 0.")
            return
        unstuck_key1 = self.unstuck_key1_var.get()
        unstuck_key2 = self.unstuck_key2_var.get()
        if not unstuck_key1 or not unstuck_key2:
            messagebox.showwarning("No key", "Select both unstuck movement keys.")
            return
        try:
            unstuck_dur1 = int(self.unstuck_dur1_var.get())
            unstuck_dur2 = int(self.unstuck_dur2_var.get())
        except ValueError:
            messagebox.showwarning("Invalid", "Unstuck hold durations must be numbers (ms).")
            return
        if unstuck_dur1 <= 0 or unstuck_dur2 <= 0:
            messagebox.showwarning("Invalid", "Unstuck hold durations must be > 0.")
            return
        target_key = self.target_key_var.get()
        if not target_key:
            messagebox.showwarning("No key", "Select a targeting key.")
            return
        try:
            tgt_min = int(self.target_min_var.get())
            tgt_max = int(self.target_max_var.get())
        except ValueError:
            messagebox.showwarning("Invalid", "Targeting delay values must be numbers (ms).")
            return
        if tgt_min <= 0 or tgt_max < tgt_min:
            messagebox.showwarning("Invalid", "Targeting: min > 0 and max >= min.")
            return

        try:
            no_target_timeout_s = int(self.no_target_timeout_var.get())
        except ValueError:
            messagebox.showwarning("Invalid", "No-target timeout must be a number (seconds).")
            return
        if no_target_timeout_s < 0:
            messagebox.showwarning("Invalid", "No-target timeout must be >= 0.")
            return

        try:
            engage_delay_ms = int(self.engage_delay_var.get())
        except ValueError:
            messagebox.showwarning("Invalid", "Attack start delay must be a number (ms).")
            return
        if engage_delay_ms < 0:
            messagebox.showwarning("Invalid", "Attack start delay must be >= 0.")
            return

        try:
            hp_confirm_count = int(self.hp_confirm_count_var.get())
        except ValueError:
            messagebox.showwarning("Invalid", "HP confirm count must be an integer.")
            return
        if hp_confirm_count < 1:
            messagebox.showwarning("Invalid", "HP confirm count must be >= 1.")
            return

        try:
            ocr_threshold = int(self.ocr_threshold_var.get())
        except ValueError:
            messagebox.showwarning("Invalid", "OCR threshold must be an integer.")
            return
        if not 0 <= ocr_threshold <= 255:
            messagebox.showwarning("Invalid", "OCR threshold must be between 0 and 255.")
            return

        try:
            ocr_scale = int(self.ocr_scale_var.get())
        except ValueError:
            messagebox.showwarning("Invalid", "OCR scale must be an integer.")
            return
        if ocr_scale < 1:
            messagebox.showwarning("Invalid", "OCR scale must be >= 1.")
            return

        attack_keys = []
        for row in self.attack_key_rows:
            k = row["key"].get()
            if not k:
                messagebox.showwarning("No key", "All attack key slots must have a key selected.")
                return
            try:
                a_min = int(row["min"].get())
                a_max = int(row["max"].get())
            except ValueError:
                messagebox.showwarning("Invalid", f"Attack key '{k}': delay values must be numbers (ms).")
                return
            if a_min <= 0 or a_max < a_min:
                messagebox.showwarning("Invalid", f"Attack key '{k}': min > 0 and max >= min.")
                return
            attack_keys.append((k, a_min, a_max))

        if not attack_keys:
            messagebox.showwarning("No keys", "Add at least one attack key.")
            return

        status_effect_keys = []
        preloaded_templates = {}
        for row in self.status_effect_key_rows:
            k = row["key"].get()
            if not k:
                messagebox.showwarning("No key", "All status effect key slots must have a key selected.")
                return
            try:
                se_rx = int(row["rx"].get())
                se_ry = int(row["ry"].get())
                se_rw = int(row["rw"].get())
                se_rh = int(row["rh"].get())
            except ValueError:
                messagebox.showwarning("Invalid", f"Status effect '{k}': region X/Y/W/H must be integers.")
                return
            if se_rw <= 0 or se_rh <= 0:
                messagebox.showwarning("Invalid", f"Status effect '{k}': region W and H must be > 0.")
                return
            tpl_slug = row["template"].get().strip()
            if not tpl_slug:
                messagebox.showwarning("Invalid", f"Status effect '{k}': select a template.")
                return
            tpl_img = load_template(tpl_slug)
            if tpl_img is None:
                messagebox.showwarning("Missing", f"Status effect '{k}': template '{tpl_slug}' not found on disk.")
                return
            preloaded_templates[tpl_slug] = tpl_img
            try:
                se_threshold = float(row["threshold"].get())
            except ValueError:
                messagebox.showwarning("Invalid", f"Status effect '{k}': threshold must be a number.")
                return
            if not 0 <= se_threshold <= 1:
                messagebox.showwarning("Invalid", f"Status effect '{k}': threshold must be between 0 and 1.")
                return
            try:
                se_retry_min = int(row["retry_min"].get())
                se_retry_max = int(row["retry_max"].get())
            except ValueError:
                messagebox.showwarning("Invalid", f"Status effect '{k}': retry interval values must be numbers (ms).")
                return
            if se_retry_min <= 0 or se_retry_max < se_retry_min:
                messagebox.showwarning("Invalid", f"Status effect '{k}': retry min > 0 and max >= min.")
                return
            status_effect_keys.append({
                "key": k,
                "region": (se_rx, se_ry, se_rw, se_rh),
                "template_slug": tpl_slug,
                "threshold": se_threshold,
                "retry_min": se_retry_min,
                "retry_max": se_retry_max,
            })

        buff_keys = []
        preloaded_buff_templates = {}
        try:
            buff_interval_s = float(self.buff_interval_var.get())
        except ValueError:
            buff_interval_s = 10.0
        if buff_interval_s <= 0:
            buff_interval_s = 10.0
        for row in self.buff_key_rows:
            k = row["key"].get()
            if not k:
                messagebox.showwarning("No key", "All buff key slots must have a key selected (or remove the row).")
                return
            try:
                brx = int(row["rx"].get())
                bry = int(row["ry"].get())
                brw = int(row["rw"].get())
                brh = int(row["rh"].get())
            except ValueError:
                messagebox.showwarning("Invalid", f"Buff '{k}': region X/Y/W/H must be integers.")
                return
            if brw <= 0 or brh <= 0:
                messagebox.showwarning("Invalid", f"Buff '{k}': region W and H must be > 0.")
                return
            tpl_slug = row["template"].get().strip()
            if not tpl_slug:
                messagebox.showwarning("Invalid", f"Buff '{k}': select a template.")
                return
            tpl_img = load_template(tpl_slug)
            if tpl_img is None:
                messagebox.showwarning("Missing", f"Buff '{k}': template '{tpl_slug}' not found on disk.")
                return
            preloaded_buff_templates[tpl_slug] = tpl_img
            try:
                b_threshold = float(row["threshold"].get())
            except ValueError:
                messagebox.showwarning("Invalid", f"Buff '{k}': threshold must be a number.")
                return
            if not 0 <= b_threshold <= 1:
                messagebox.showwarning("Invalid", f"Buff '{k}': threshold must be between 0 and 1.")
                return
            buff_keys.append({
                "key": k,
                "region": (brx, bry, brw, brh),
                "template_slug": tpl_slug,
                "threshold": b_threshold,
            })

        death_key = None
        death_delay_ms = 0
        if self.death_enabled_var.get():
            death_key = self.death_key_var.get()
            if not death_key:
                messagebox.showwarning("No key", "Select an on-death key or disable the option.")
                return
            try:
                death_delay_ms = int(self.death_delay_var.get())
            except ValueError:
                messagebox.showwarning("Invalid", "Death delay must be a number (ms).")
                return
            if death_delay_ms < 0:
                messagebox.showwarning("Invalid", "Death delay must be >= 0.")
                return

        if not self._open_serial():
            return

        self.monitoring = True
        self.monitor_start_btn.config(state=tk.DISABLED)
        self.monitor_stop_btn.config(state=tk.NORMAL)
        self.monitor_status_var.set("Monitoring...")

        self.monitor_thread = threading.Thread(
            target=self._monitor_loop,
            args=(
                self.region, stuck_s,
                target_key, tgt_min, tgt_max,
                no_target_timeout_s,
                engage_delay_ms,
                attack_keys, status_effect_keys,
                preloaded_templates,
                buff_keys, preloaded_buff_templates, buff_interval_s,
                death_key, death_delay_ms,
                unstuck_key1, unstuck_dur1,
                unstuck_key2, unstuck_dur2,
                hp_confirm_count,
                ocr_threshold, ocr_scale,
            ),
            daemon=True,
        )
        self.monitor_thread.start()

    def _stop_monitoring(self):
        self.monitoring = False
        self._serial_send("STOP")
        self.monitor_start_btn.config(state=tk.NORMAL)
        self.monitor_stop_btn.config(state=tk.DISABLED)
        self.monitor_status_var.set("Idle")
        self.hp_live_var.set("HP: —")

    def _monitor_loop(
        self, region, stuck_s,
        target_key, tgt_min, tgt_max,
        no_target_timeout_s,
        engage_delay_ms,
        attack_keys, status_effect_keys,
        preloaded_templates,
        buff_keys, preloaded_buff_templates, buff_interval_s,
        death_key, death_delay_ms,
        unstuck_key1, unstuck_dur1,
        unstuck_key2, unstuck_dur2,
        hp_confirm_count,
        ocr_threshold, ocr_scale,
    ):
        """Background thread — OCR-based HP reading + template-based status effects + buffs.

        HP gone (None)  → press *target_key* (find next monster)
        HP visible (new)→ press status-effect keys immediately;
                          attacks delayed by engage_delay_ms
        HP visible      → press each attack key on its own independent timer
        Status effects  → template-match each effect's region; if not found,
                          re-apply at retry interval; once matched, stop retrying
        Buffs           → every buff_interval_s seconds, template-match each buff;
                          if not found, press key (never reset on new target)
        HP not dropping → if HP% hasn't decreased for stuck_s seconds,
                          press *target_key* (stuck, re-target)
        Stuck 2x in row → movement sequence (hold keys) then re-target
        HP was visible → now gone → press *death_key* once (if enabled)

        HP verification: a reading of 0 or None must occur *hp_confirm_count*
        consecutive times before it is accepted.  Until confirmed the last
        known-good HP value is used so the bot keeps attacking.
        """
        x, y, w, h = region
        engage_s = engage_delay_ms / 1000
        hp_since = None
        no_target_since = None
        prev_hp_visible = False
        prev_hp_pct = None
        stuck_count = 0

        # HP verification state — tracks consecutive "suspicious" reads
        hp_gone_streak = 0        # consecutive None reads
        hp_zero_streak = 0        # consecutive 0.0 reads
        last_good_hp = None       # last HP value that was > 0

        now = time.monotonic()
        next_press = [now] * len(attack_keys)

        se_applied = [False] * len(status_effect_keys)
        next_se_check = [now] * len(status_effect_keys)

        next_buff_check = now + buff_interval_s if buff_keys else float("inf")

        POLL_INTERVAL = 0.05

        def _set_status(msg):
            self.root.after(0, lambda m=msg: self.monitor_status_var.set(m))

        def _set_hp_live(txt):
            self.root.after(0, lambda t=txt: self.hp_live_var.set(t))

        def _set_se_label(idx, txt):
            if idx < len(self.status_effect_key_rows):
                row = self.status_effect_key_rows[idx]
                lbl = row.get("status_label")
                if lbl:
                    self.root.after(0, lambda t=txt, l=lbl: l.config(text=t))

        while self.monitoring:
            try:
                image = capture_region(x, y, w, h)
                raw_hp = read_hp_percentage(
                    image,
                    ocr_threshold=ocr_threshold,
                    ocr_scale=ocr_scale,
                )
            except Exception:
                time.sleep(0.5)
                continue

            # --- HP verification: require hp_confirm_count consecutive ---
            # --- suspicious reads before treating them as real.         ---
            if raw_hp is None:
                hp_gone_streak += 1
                hp_zero_streak = 0
            elif raw_hp == 0.0:
                hp_zero_streak += 1
                hp_gone_streak = 0
            else:
                hp_gone_streak = 0
                hp_zero_streak = 0
                last_good_hp = raw_hp

            gone_confirmed = hp_gone_streak >= hp_confirm_count
            zero_confirmed = hp_zero_streak >= hp_confirm_count

            if raw_hp is None and not gone_confirmed and last_good_hp is not None:
                hp_pct = last_good_hp
            elif raw_hp == 0.0 and not zero_confirmed and last_good_hp is not None:
                hp_pct = last_good_hp
            else:
                hp_pct = raw_hp

            hp_visible = hp_pct is not None

            if hp_pct is not None:
                verify_note = ""
                if raw_hp is None:
                    verify_note = f" [raw: None x{hp_gone_streak}]"
                elif raw_hp == 0.0 and not zero_confirmed:
                    verify_note = f" [raw: 0 x{hp_zero_streak}]"
                _set_hp_live(f"HP: {hp_pct:.1f}%{verify_note}")
            else:
                _set_hp_live("HP: —")

            now = time.monotonic()

            if hp_visible:
                no_target_since = None
                if hp_since is None:
                    hp_since = now
                    prev_hp_pct = hp_pct
                    next_press = [now + engage_s] * len(attack_keys)

                    se_applied = [False] * len(status_effect_keys)
                    for i, se in enumerate(status_effect_keys):
                        self._serial_send(f"PRESS;{se['key']}")
                        next_se_check[i] = now + random.uniform(
                            se["retry_min"] / 1000, se["retry_max"] / 1000
                        )
                        _set_se_label(i, "")

                if prev_hp_pct is not None and hp_pct < prev_hp_pct:
                    hp_since = now
                prev_hp_pct = hp_pct

                elapsed = now - hp_since
                if elapsed >= stuck_s:
                    stuck_count += 1
                    if stuck_count >= 2:
                        _set_status(f"Stuck x{stuck_count} \u2014 moving to unstuck")
                        self._serial_send(f"HOLD;{unstuck_key1};{unstuck_dur1}")
                        time.sleep(unstuck_dur1 / 1000)
                        self._serial_send(f"HOLD;{unstuck_key2};{unstuck_dur2}")
                        time.sleep(unstuck_dur2 / 1000)
                    else:
                        _set_status(f"Stuck ({int(elapsed)}s) \u2014 re-targeting")
                    self._serial_send(f"PRESS;{target_key}")
                    hp_since = None
                    prev_hp_pct = None
                    delay = random.uniform(tgt_min / 1000, tgt_max / 1000)
                    time.sleep(delay)
                else:
                    keys_desc = ", ".join(k for k, _, _ in attack_keys)
                    hp_info = f" [HP: {hp_pct:.1f}%]" if hp_pct is not None else ""
                    se_parts = []
                    for i, se in enumerate(status_effect_keys):
                        slug = se["template_slug"]
                        state = "ON" if se_applied[i] else "OFF"
                        se_parts.append(f"{slug}: {state}")
                    se_info = f" [{', '.join(se_parts)}]" if se_parts else ""
                    _set_status(f"Attacking [{keys_desc}] ({int(elapsed)}s){hp_info}{se_info}")

                    for i, (key, a_min, a_max) in enumerate(attack_keys):
                        if now >= next_press[i]:
                            self._serial_send(f"PRESS;{key}")
                            next_press[i] = now + random.uniform(a_min / 1000, a_max / 1000)

                    for i, se in enumerate(status_effect_keys):
                        if se_applied[i] or now < next_se_check[i]:
                            continue
                        try:
                            se_region = se["region"]
                            se_img = capture_region(*se_region)
                            tpl_img = preloaded_templates[se["template_slug"]]
                            if match_template(se_img, tpl_img, se["threshold"]):
                                se_applied[i] = True
                                _set_se_label(i, "ON")
                            else:
                                self._serial_send(f"PRESS;{se['key']}")
                                next_se_check[i] = now + random.uniform(
                                    se["retry_min"] / 1000, se["retry_max"] / 1000
                                )
                                _set_se_label(i, "OFF")
                        except Exception:
                            next_se_check[i] = now + 1.0

                    # Buffs: check every buff_interval_s seconds (never reset on new target)
                    if buff_keys and now >= next_buff_check:
                        for b in buff_keys:
                            try:
                                b_region = b["region"]
                                b_img = capture_region(*b_region)
                                tpl_img = preloaded_buff_templates[b["template_slug"]]
                                if not match_template(b_img, tpl_img, b["threshold"]):
                                    self._serial_send(f"PRESS;{b['key']}")
                            except Exception:
                                pass
                        next_buff_check = now + buff_interval_s

                    time.sleep(POLL_INTERVAL)
            else:
                if prev_hp_visible and death_key:
                    _set_status("Mob dead \u2014 pressing death key")
                    self._serial_send(f"PRESS;{death_key}")
                    time.sleep(death_delay_ms / 1000)

                if no_target_since is None:
                    no_target_since = now
                elif no_target_timeout_s > 0 and (now - no_target_since) >= no_target_timeout_s:
                    _set_status(f"No target for {no_target_timeout_s}s \u2014 stopping bot")
                    self.root.after(0, self._stop_monitoring)
                    break

                hp_since = None
                prev_hp_pct = None
                stuck_count = 0
                last_good_hp = None
                _set_status("No HP \u2014 targeting")
                self._serial_send(f"PRESS;{target_key}")
                delay = random.uniform(tgt_min / 1000, tgt_max / 1000)
                time.sleep(delay)

            prev_hp_visible = hp_visible

        _set_hp_live("HP: —")
        self.root.after(0, lambda: self.monitor_status_var.set("Idle"))

    # ------------------------------------------------------------------ #
    #  Settings persistence                                                #
    # ------------------------------------------------------------------ #

    def _save_settings(self):
        data = {
            "keys": list(self.key_listbox.get(0, tk.END)),
            "min_delay": self.min_delay_var.get(),
            "max_delay": self.max_delay_var.get(),
            "last_port": self.port_var.get(),
        }
        if CONDITIONAL_AVAILABLE and hasattr(self, "stuck_timeout_var"):
            data.update({
                "region": {
                    "x": self.region_x_var.get(),
                    "y": self.region_y_var.get(),
                    "w": self.region_w_var.get(),
                    "h": self.region_h_var.get(),
                },
                "stuck_timeout": self.stuck_timeout_var.get(),
                "unstuck_key1": self.unstuck_key1_var.get(),
                "unstuck_dur1": self.unstuck_dur1_var.get(),
                "unstuck_key2": self.unstuck_key2_var.get(),
                "unstuck_dur2": self.unstuck_dur2_var.get(),
                "target_key": self.target_key_var.get(),
                "target_min": self.target_min_var.get(),
                "target_max": self.target_max_var.get(),
                "engage_delay": self.engage_delay_var.get(),
                "hp_confirm_count": self.hp_confirm_count_var.get(),
                "ocr_threshold": self.ocr_threshold_var.get(),
                "ocr_scale": self.ocr_scale_var.get(),
                "no_target_timeout": self.no_target_timeout_var.get(),
                "attack_keys": [
                    {"key": r["key"].get(), "min": r["min"].get(), "max": r["max"].get()}
                    for r in self.attack_key_rows
                ],
                "status_effect_keys": [
                    {
                        "key": r["key"].get(),
                        "rx": r["rx"].get(), "ry": r["ry"].get(),
                        "rw": r["rw"].get(), "rh": r["rh"].get(),
                        "template_slug": r["template"].get(),
                        "match_threshold": r["threshold"].get(),
                        "retry_min": r["retry_min"].get(), "retry_max": r["retry_max"].get(),
                    }
                    for r in self.status_effect_key_rows
                ],
                "buff_interval": self.buff_interval_var.get(),
                "buff_keys": [
                    {
                        "key": r["key"].get(),
                        "rx": r["rx"].get(), "ry": r["ry"].get(),
                        "rw": r["rw"].get(), "rh": r["rh"].get(),
                        "template_slug": r["template"].get(),
                        "match_threshold": r["threshold"].get(),
                    }
                    for r in self.buff_key_rows
                ],
                "death_enabled": self.death_enabled_var.get(),
                "death_key": self.death_key_var.get(),
                "death_delay": self.death_delay_var.get(),
            })
        try:
            with open(CONFIG_PATH, "w", encoding="utf-8") as f:
                json.dump(data, f, indent=2)
        except Exception:
            pass

    def _load_settings(self):
        try:
            with open(CONFIG_PATH, "r", encoding="utf-8") as f:
                data = json.load(f)
        except (FileNotFoundError, json.JSONDecodeError, OSError):
            return

        # -- Autoclicker tab --
        for key in data.get("keys", []):
            self.key_listbox.insert(tk.END, key)
        if "min_delay" in data:
            self.min_delay_var.set(data["min_delay"])
        if "max_delay" in data:
            self.max_delay_var.set(data["max_delay"])

        # -- COM port (select if still available) --
        saved_port = data.get("last_port", "")
        if saved_port and saved_port in list(self.port_combo["values"]):
            self.port_var.set(saved_port)

        # -- Conditional Clicker tab --
        if not CONDITIONAL_AVAILABLE or not hasattr(self, "stuck_timeout_var"):
            return

        region = data.get("region", {})
        if "x" in region:
            self.region_x_var.set(region["x"])
        if "y" in region:
            self.region_y_var.set(region["y"])
        if "w" in region:
            self.region_w_var.set(region["w"])
        if "h" in region:
            self.region_h_var.set(region["h"])

        if "stuck_timeout" in data:
            self.stuck_timeout_var.set(data["stuck_timeout"])
        if "unstuck_key1" in data:
            self.unstuck_key1_var.set(data["unstuck_key1"])
        if "unstuck_dur1" in data:
            self.unstuck_dur1_var.set(data["unstuck_dur1"])
        if "unstuck_key2" in data:
            self.unstuck_key2_var.set(data["unstuck_key2"])
        if "unstuck_dur2" in data:
            self.unstuck_dur2_var.set(data["unstuck_dur2"])
        if "target_key" in data:
            self.target_key_var.set(data["target_key"])
        if "target_min" in data:
            self.target_min_var.set(data["target_min"])
        if "target_max" in data:
            self.target_max_var.set(data["target_max"])
        if "engage_delay" in data:
            self.engage_delay_var.set(data["engage_delay"])
        if "hp_confirm_count" in data:
            self.hp_confirm_count_var.set(data["hp_confirm_count"])
        if "ocr_threshold" in data:
            self.ocr_threshold_var.set(data["ocr_threshold"])
        if "ocr_scale" in data:
            self.ocr_scale_var.set(data["ocr_scale"])
        if "no_target_timeout" in data:
            self.no_target_timeout_var.set(data["no_target_timeout"])
        if "buff_interval" in data:
            self.buff_interval_var.set(data["buff_interval"])
        if "death_enabled" in data:
            self.death_enabled_var.set(data["death_enabled"])
        if "death_key" in data:
            self.death_key_var.set(data["death_key"])
        if "death_delay" in data:
            self.death_delay_var.set(data["death_delay"])

        saved_attacks = data.get("attack_keys", [])
        if saved_attacks:
            for row in list(self.attack_key_rows):
                row["frame"].destroy()
            self.attack_key_rows.clear()
            for ak in saved_attacks:
                self._add_attack_key_row(ak.get("key", "f1"), ak.get("min", "200"), ak.get("max", "1000"))

        saved_status_effects = data.get("status_effect_keys", [])
        if saved_status_effects:
            for row in list(self.status_effect_key_rows):
                row["frame"].destroy()
            self.status_effect_key_rows.clear()
            for se in saved_status_effects:
                self._add_status_effect_key_row(
                    key=se.get("key", "f1"),
                    region_x=se.get("rx", "0"), region_y=se.get("ry", "0"),
                    region_w=se.get("rw", "50"), region_h=se.get("rh", "50"),
                    template_slug=se.get("template_slug", ""),
                    match_threshold=se.get("match_threshold", "0.80"),
                    retry_min=se.get("retry_min", "1000"),
                    retry_max=se.get("retry_max", "2000"),
                )

        saved_buffs = data.get("buff_keys", [])
        if saved_buffs:
            for row in list(self.buff_key_rows):
                row["frame"].destroy()
            self.buff_key_rows.clear()
            for b in saved_buffs:
                self._add_buff_key_row(
                    key=b.get("key", "f1"),
                    region_x=b.get("rx", "0"), region_y=b.get("ry", "0"),
                    region_w=b.get("rw", "50"), region_h=b.get("rh", "50"),
                    template_slug=b.get("template_slug", ""),
                    match_threshold=b.get("match_threshold", "0.80"),
                )

    # ------------------------------------------------------------------ #
    #  App lifecycle                                                       #
    # ------------------------------------------------------------------ #

    def run(self):
        self.root.protocol("WM_DELETE_WINDOW", self._on_close)
        self.root.mainloop()

    def _on_close(self):
        self._save_settings()
        self.preview_active = False
        self.monitoring = False
        if self.running:
            self._on_stop()
        self._close_serial()
        self.root.destroy()


if __name__ == "__main__":
    app = AutoclickerApp()
    app.run()
