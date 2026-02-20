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
        ttk.Label(trig_row2, text="Key to press:").pack(side=tk.LEFT, padx=(0, 4))
        self.cond_key_var = tk.StringVar(value="f4")
        ttk.Combobox(
            trig_row2,
            textvariable=self.cond_key_var,
            values=KEY_OPTIONS,
            width=10,
            state="readonly",
        ).pack(side=tk.LEFT)

        trig_row3 = ttk.Frame(trigger_frame)
        trig_row3.pack(fill=tk.X)
        ttk.Label(trig_row3, text="Min delay (ms):").pack(
            side=tk.LEFT, padx=(0, 4)
        )
        self.cond_min_delay_var = tk.StringVar(value="200")
        ttk.Entry(trig_row3, textvariable=self.cond_min_delay_var, width=8).pack(
            side=tk.LEFT, padx=(0, 12)
        )
        ttk.Label(trig_row3, text="Max delay (ms):").pack(
            side=tk.LEFT, padx=(0, 4)
        )
        self.cond_max_delay_var = tk.StringVar(value="400")
        ttk.Entry(trig_row3, textvariable=self.cond_max_delay_var, width=8).pack(
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
            min_ms = int(self.cond_min_delay_var.get())
            max_ms = int(self.cond_max_delay_var.get())
        except ValueError:
            messagebox.showwarning("Invalid", "Min/max delay must be numbers (ms).")
            return
        if min_ms <= 0 or max_ms < min_ms:
            messagebox.showwarning("Invalid", "Min delay > 0 and max >= min.")
            return
        key = self.cond_key_var.get()
        if not key:
            messagebox.showwarning("No key", "Select a key to press.")
            return

        if not self._open_serial():
            return

        self.monitoring = True
        self.monitor_start_btn.config(state=tk.DISABLED)
        self.monitor_stop_btn.config(state=tk.NORMAL)
        self.monitor_status_var.set("Monitoring...")

        self.monitor_thread = threading.Thread(
            target=self._monitor_loop,
            args=(self.region, hp_color, tolerance, key, min_ms, max_ms),
            daemon=True,
        )
        self.monitor_thread.start()

    def _stop_monitoring(self):
        self.monitoring = False
        self._serial_send("STOP")
        self.monitor_start_btn.config(state=tk.NORMAL)
        self.monitor_stop_btn.config(state=tk.DISABLED)
        self.monitor_status_var.set("Idle")

    def _monitor_loop(self, region, hp_color, tolerance, key, min_ms, max_ms):
        """Background thread: HP bar visible → idle; HP bar gone → press key."""
        x, y, w, h = region

        def _set_status(msg):
            self.root.after(0, lambda m=msg: self.monitor_status_var.set(m))

        while self.monitoring:
            try:
                image = capture_region(x, y, w, h)
                hp_visible = color_present(image, hp_color, tolerance)
            except Exception:
                time.sleep(0.5)
                continue

            if hp_visible:
                _set_status("Monster alive \u2014 waiting")
                time.sleep(0.2)
            else:
                _set_status("No HP \u2014 pressing key")
                self._serial_send(f"PRESS;{key}")
                delay = random.uniform(min_ms / 1000, max_ms / 1000)
                time.sleep(delay)

        self.root.after(0, lambda: self.monitor_status_var.set("Idle"))

    # ------------------------------------------------------------------ #
    #  App lifecycle                                                       #
    # ------------------------------------------------------------------ #

    def run(self):
        self.root.protocol("WM_DELETE_WINDOW", self._on_close)
        self.root.mainloop()

    def _on_close(self):
        self.preview_active = False
        self.monitoring = False
        if self.running:
            self._on_stop()
        self._close_serial()
        self.root.destroy()


if __name__ == "__main__":
    app = AutoclickerApp()
    app.run()
