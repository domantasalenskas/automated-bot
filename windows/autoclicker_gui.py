"""
Pico HID Autoclicker – Windows GUI
Tab 1  Autoclicker – configure keys, min/max delay; Start/Stop via buttons or F12.
Tab 2  Conditional Clicker – monitor a screen region's colors and press a key
       when a trigger color disappears.
Requires: pyserial, pynput, mss, Pillow.  Run: pip install -r requirements.txt
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
    from screen_reader import capture_region, get_unique_colors, color_present
    from region_selector import RegionSelector

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
        ttk.Label(trig_row1, text="HP bar color:").pack(side=tk.LEFT, padx=(0, 4))
        self.hp_color_var = tk.StringVar(value="#892015")
        ttk.Entry(trig_row1, textvariable=self.hp_color_var, width=10).pack(
            side=tk.LEFT, padx=(0, 4)
        )
        self.hp_swatch = tk.Canvas(
            trig_row1, width=20, height=20, bg="#892015", highlightthickness=1
        )
        self.hp_swatch.pack(side=tk.LEFT)
        self.hp_color_var.trace_add("write", self._update_hp_swatch)

        trig_row2 = ttk.Frame(trigger_frame)
        trig_row2.pack(fill=tk.X, pady=(0, 4))
        ttk.Label(trig_row2, text="Tolerance:").pack(side=tk.LEFT, padx=(0, 4))
        self.tolerance_var = tk.StringVar(value="1")
        ttk.Entry(trig_row2, textvariable=self.tolerance_var, width=6).pack(
            side=tk.LEFT, padx=(0, 12)
        )
        ttk.Label(trig_row2, text="Stuck timeout (s):").pack(
            side=tk.LEFT, padx=(0, 4)
        )
        self.stuck_timeout_var = tk.StringVar(value="20")
        ttk.Entry(trig_row2, textvariable=self.stuck_timeout_var, width=6).pack(
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

        try:
            tolerance = int(self.tolerance_var.get())
        except ValueError:
            tolerance = 30
        colors = get_unique_colors(image, tolerance)
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
            lbl.bind(
                "<Button-1>",
                lambda _e, c=hex_color: self.hp_color_var.set(c),
            )
            swatch.bind(
                "<Button-1>",
                lambda _e, c=hex_color: self.hp_color_var.set(c),
            )

    def _update_hp_swatch(self, *_args):
        try:
            self.hp_swatch.configure(bg=self.hp_color_var.get().strip())
        except tk.TclError:
            pass

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

    def _add_status_effect_key_row(self, key="f1", min_ms="3000", max_ms="5000"):
        row_frame = ttk.Frame(self.status_effect_keys_container)
        row_frame.pack(fill=tk.X, pady=1)

        key_var = tk.StringVar(value=key)
        min_var = tk.StringVar(value=min_ms)
        max_var = tk.StringVar(value=max_ms)

        ttk.Label(row_frame, text="Key:").pack(side=tk.LEFT, padx=(0, 2))
        ttk.Combobox(
            row_frame, textvariable=key_var, values=KEY_OPTIONS,
            width=8, state="readonly",
        ).pack(side=tk.LEFT, padx=(0, 8))
        ttk.Label(row_frame, text="Re-apply Min (ms):").pack(side=tk.LEFT, padx=(0, 2))
        ttk.Entry(row_frame, textvariable=min_var, width=6).pack(
            side=tk.LEFT, padx=(0, 8)
        )
        ttk.Label(row_frame, text="Max (ms):").pack(side=tk.LEFT, padx=(0, 2))
        ttk.Entry(row_frame, textvariable=max_var, width=6).pack(
            side=tk.LEFT, padx=(0, 8)
        )

        entry = {"frame": row_frame, "key": key_var, "min": min_var, "max": max_var}
        self.status_effect_key_rows.append(entry)

        def _remove(e=entry):
            e["frame"].destroy()
            self.status_effect_key_rows.remove(e)

        ttk.Button(row_frame, text="✕", width=3, command=_remove).pack(side=tk.LEFT)

    # -- monitoring --

    def _start_monitoring(self):
        if not self.region:
            messagebox.showwarning("No region", "Select a screen region first.")
            return
        try:
            tolerance = int(self.tolerance_var.get())
        except ValueError:
            messagebox.showwarning("Invalid", "Tolerance must be a number.")
            return
        hp_color = self.hp_color_var.get().strip()
        if not hp_color.startswith("#") or len(hp_color) != 7:
            messagebox.showwarning(
                "Invalid", "HP bar color must be a hex code like #892015."
            )
            return
        try:
            stuck_s = float(self.stuck_timeout_var.get())
        except ValueError:
            messagebox.showwarning("Invalid", "Stuck timeout must be a number (s).")
            return
        if stuck_s <= 0:
            messagebox.showwarning("Invalid", "Stuck timeout must be > 0.")
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
            engage_delay_ms = int(self.engage_delay_var.get())
        except ValueError:
            messagebox.showwarning("Invalid", "Attack start delay must be a number (ms).")
            return
        if engage_delay_ms < 0:
            messagebox.showwarning("Invalid", "Attack start delay must be >= 0.")
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
        for row in self.status_effect_key_rows:
            k = row["key"].get()
            if not k:
                messagebox.showwarning("No key", "All status effect key slots must have a key selected.")
                return
            try:
                se_min = int(row["min"].get())
                se_max = int(row["max"].get())
            except ValueError:
                messagebox.showwarning("Invalid", f"Status effect key '{k}': delay values must be numbers (ms).")
                return
            if se_min <= 0 or se_max < se_min:
                messagebox.showwarning("Invalid", f"Status effect key '{k}': min > 0 and max >= min.")
                return
            status_effect_keys.append((k, se_min, se_max))

        death_key = None
        if self.death_enabled_var.get():
            death_key = self.death_key_var.get()
            if not death_key:
                messagebox.showwarning("No key", "Select an on-death key or disable the option.")
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
                self.region, hp_color, tolerance, stuck_s,
                target_key, tgt_min, tgt_max,
                engage_delay_ms,
                attack_keys, status_effect_keys, death_key,
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

    def _monitor_loop(
        self, region, hp_color, tolerance, stuck_s,
        target_key, tgt_min, tgt_max,
        engage_delay_ms,
        attack_keys, status_effect_keys, death_key,
    ):
        """Background thread:

        HP gone         → press *target_key* (find next monster)
        HP visible (new)→ press status-effect keys immediately;
                          attacks delayed by engage_delay_ms
        HP visible      → press each attack key on its own independent timer
        HP visible 20s+ → press *target_key* (stuck, re-target)
        HP was visible → now gone  → press *death_key* once (if enabled)
        """
        x, y, w, h = region
        engage_s = engage_delay_ms / 1000
        hp_since = None
        prev_hp_visible = False

        now = time.monotonic()
        next_press = [now] * len(attack_keys)
        next_se_press = [now] * len(status_effect_keys)

        POLL_INTERVAL = 0.05

        def _set_status(msg):
            self.root.after(0, lambda m=msg: self.monitor_status_var.set(m))

        while self.monitoring:
            try:
                image = capture_region(x, y, w, h)
                hp_visible = color_present(image, hp_color, tolerance)
            except Exception:
                time.sleep(0.5)
                continue

            now = time.monotonic()

            if hp_visible:
                if hp_since is None:
                    hp_since = now
                    next_press = [now + engage_s] * len(attack_keys)

                    for i, (key, se_min, se_max) in enumerate(status_effect_keys):
                        self._serial_send(f"PRESS;{key}")
                        next_se_press[i] = now + random.uniform(se_min / 1000, se_max / 1000)

                elapsed = now - hp_since
                if elapsed >= stuck_s:
                    _set_status(f"Stuck ({int(elapsed)}s) \u2014 re-targeting")
                    self._serial_send(f"PRESS;{target_key}")
                    hp_since = time.monotonic()
                    delay = random.uniform(tgt_min / 1000, tgt_max / 1000)
                    time.sleep(delay)
                else:
                    keys_desc = ", ".join(k for k, _, _ in attack_keys)
                    _set_status(f"Attacking [{keys_desc}] ({int(elapsed)}s)")
                    for i, (key, a_min, a_max) in enumerate(attack_keys):
                        if now >= next_press[i]:
                            self._serial_send(f"PRESS;{key}")
                            next_press[i] = now + random.uniform(a_min / 1000, a_max / 1000)
                    for i, (key, se_min, se_max) in enumerate(status_effect_keys):
                        if now >= next_se_press[i]:
                            self._serial_send(f"PRESS;{key}")
                            next_se_press[i] = now + random.uniform(se_min / 1000, se_max / 1000)
                    time.sleep(POLL_INTERVAL)
            else:
                if prev_hp_visible and death_key:
                    _set_status("Mob dead \u2014 pressing death key")
                    self._serial_send(f"PRESS;{death_key}")
                    time.sleep(0.1)

                hp_since = None
                _set_status("No HP \u2014 targeting")
                self._serial_send(f"PRESS;{target_key}")
                delay = random.uniform(tgt_min / 1000, tgt_max / 1000)
                time.sleep(delay)

            prev_hp_visible = hp_visible

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
        if CONDITIONAL_AVAILABLE and hasattr(self, "hp_color_var"):
            data.update({
                "region": {
                    "x": self.region_x_var.get(),
                    "y": self.region_y_var.get(),
                    "w": self.region_w_var.get(),
                    "h": self.region_h_var.get(),
                },
                "hp_color": self.hp_color_var.get(),
                "tolerance": self.tolerance_var.get(),
                "stuck_timeout": self.stuck_timeout_var.get(),
                "target_key": self.target_key_var.get(),
                "target_min": self.target_min_var.get(),
                "target_max": self.target_max_var.get(),
                "engage_delay": self.engage_delay_var.get(),
                "attack_keys": [
                    {"key": r["key"].get(), "min": r["min"].get(), "max": r["max"].get()}
                    for r in self.attack_key_rows
                ],
                "status_effect_keys": [
                    {"key": r["key"].get(), "min": r["min"].get(), "max": r["max"].get()}
                    for r in self.status_effect_key_rows
                ],
                "death_enabled": self.death_enabled_var.get(),
                "death_key": self.death_key_var.get(),
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
        if not CONDITIONAL_AVAILABLE or not hasattr(self, "hp_color_var"):
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

        if "hp_color" in data:
            self.hp_color_var.set(data["hp_color"])
        if "tolerance" in data:
            self.tolerance_var.set(data["tolerance"])
        if "stuck_timeout" in data:
            self.stuck_timeout_var.set(data["stuck_timeout"])
        if "target_key" in data:
            self.target_key_var.set(data["target_key"])
        if "target_min" in data:
            self.target_min_var.set(data["target_min"])
        if "target_max" in data:
            self.target_max_var.set(data["target_max"])
        if "engage_delay" in data:
            self.engage_delay_var.set(data["engage_delay"])
        if "death_enabled" in data:
            self.death_enabled_var.set(data["death_enabled"])
        if "death_key" in data:
            self.death_key_var.set(data["death_key"])

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
                self._add_status_effect_key_row(se.get("key", "f1"), se.get("min", "3000"), se.get("max", "5000"))

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
