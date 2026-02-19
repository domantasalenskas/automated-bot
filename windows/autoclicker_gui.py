"""
Pico HID Autoclicker - Windows GUI
Configure keys, min/max delay; Start/Stop via buttons or F12.
Requires: pyserial, pynput. Run: pip install -r requirements.txt
"""

import tkinter as tk
from tkinter import ttk, messagebox
import serial
import serial.tools.list_ports
import threading

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
        self.root.minsize(400, 380)
        self.serial_port = None
        self.running = False
        self._lock = threading.Lock()

        self._build_ui()
        self._refresh_ports()
        self._start_f12_listener()

    def _build_ui(self):
        main = ttk.Frame(self.root, padding=10)
        main.pack(fill=tk.BOTH, expand=True)

        # COM port
        port_frame = ttk.Frame(main)
        port_frame.pack(fill=tk.X, pady=(0, 8))
        ttk.Label(port_frame, text="COM port:").pack(side=tk.LEFT, padx=(0, 6))
        self.port_var = tk.StringVar()
        self.port_combo = ttk.Combobox(port_frame, textvariable=self.port_var, width=24, state="readonly")
        self.port_combo.pack(side=tk.LEFT, padx=(0, 6))
        ttk.Button(port_frame, text="Refresh", command=self._refresh_ports).pack(side=tk.LEFT)

        # Keys
        key_frame = ttk.LabelFrame(main, text="Keys to press (in order)", padding=6)
        key_frame.pack(fill=tk.X, pady=(0, 8))
        key_row = ttk.Frame(key_frame)
        key_row.pack(fill=tk.X)
        self.key_var = tk.StringVar()
        self.key_combo = ttk.Combobox(key_row, textvariable=self.key_var, values=KEY_OPTIONS, width=14)
        self.key_combo.pack(side=tk.LEFT, padx=(0, 6))
        ttk.Button(key_row, text="Add", command=self._add_key).pack(side=tk.LEFT, padx=(0, 6))
        self.key_listbox = tk.Listbox(key_frame, height=4, selectmode=tk.SINGLE)
        self.key_listbox.pack(fill=tk.X, pady=(6, 0))
        ttk.Button(key_frame, text="Remove selected", command=self._remove_key).pack(anchor=tk.W, pady=(4, 0))

        # Delay
        delay_frame = ttk.Frame(main)
        delay_frame.pack(fill=tk.X, pady=(0, 8))
        ttk.Label(delay_frame, text="Min delay (ms):").pack(side=tk.LEFT, padx=(0, 6))
        self.min_delay_var = tk.StringVar(value="1000")
        ttk.Entry(delay_frame, textvariable=self.min_delay_var, width=10).pack(side=tk.LEFT, padx=(0, 16))
        ttk.Label(delay_frame, text="Max delay (ms):").pack(side=tk.LEFT, padx=(0, 6))
        self.max_delay_var = tk.StringVar(value="1500")
        ttk.Entry(delay_frame, textvariable=self.max_delay_var, width=10).pack(side=tk.LEFT)

        # Start / Stop
        btn_frame = ttk.Frame(main)
        btn_frame.pack(fill=tk.X, pady=(8, 0))
        self.start_btn = ttk.Button(btn_frame, text="Start", command=self._on_start)
        self.start_btn.pack(side=tk.LEFT, padx=(0, 8))
        self.stop_btn = ttk.Button(btn_frame, text="Stop", command=self._on_stop, state=tk.DISABLED)
        self.stop_btn.pack(side=tk.LEFT)
        ttk.Label(btn_frame, text="  F12 = Start/Stop", foreground="gray").pack(side=tk.LEFT, padx=(16, 0))

        # Status
        self.status_var = tk.StringVar(value="Select COM port and add keys.")
        ttk.Label(main, textvariable=self.status_var, foreground="gray").pack(anchor=tk.W, pady=(8, 0))

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
            messagebox.showwarning("Unknown key", f"Key '{k}' not in list. Choose from dropdown.")

    def _remove_key(self):
        sel = self.key_listbox.curselection()
        if sel:
            self.key_listbox.delete(sel[0])

    def _refresh_ports(self):
        ports, is_pico = find_pico_ports()
        self.port_combo["values"] = [disp for _dev, disp in ports]
        if ports:
            self.port_combo.current(0)
            self.status_var.set(
                "Pico port found. Add keys and set delay."
                if is_pico
                else "No Pico auto-detected. Select your Thonny port above."
            )
        else:
            self.port_var.set("")
            self.status_var.set("No COM ports found. Connect the device and click Refresh.")

    def _get_port_device(self):
        val = self.port_var.get()
        if not val:
            return None
        # "COM3 (CircuitPython ...)" -> COM3
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
        # Already have an open port - reuse it
        with self._lock:
            if self.serial_port is not None and self.serial_port.is_open:
                return True
            # Clear stale closed port so we don't try to reuse it
            self.serial_port = None
        dev = self._get_port_device()
        if not dev:
            messagebox.showwarning("No port", "Select a COM port first.")
            return False
        try:
            with self._lock:
                if self.serial_port is not None and self.serial_port.is_open:
                    return True
                self.serial_port = serial.Serial(dev, BAUD, timeout=0.1, write_timeout=1)
                self.serial_port.dtr = False
                self.serial_port.rts = False
            return True
        except serial.SerialException as e:
            err = str(e).lower()
            if "already open" in err or "access" in err or "permission" in err:
                messagebox.showerror(
                    "Port in use",
                    "COM port is already in use.\n\n"
                    "Try: 1) Unplug the Pico, wait 2 seconds, plug it back in.\n"
                    "     2) Close any other app using the port (e.g. Thonny, serial monitor).\n"
                    "     3) Click Refresh and Start again.",
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

    def _on_start(self):
        keys = self._get_keys()
        if not keys:
            messagebox.showwarning("No keys", "Add at least one key.")
            return
        try:
            min_ms = int(self.min_delay_var.get())
            max_ms = int(self.max_delay_var.get())
        except ValueError:
            messagebox.showwarning("Invalid delay", "Min and max delay must be numbers (ms).")
            return
        if min_ms <= 0 or max_ms < min_ms:
            messagebox.showwarning("Invalid delay", "Min delay > 0 and max >= min.")
            return

        if self.serial_port is None or not self.serial_port.is_open:
            if not self._open_serial():
                return
            self.start_btn.config(state=tk.DISABLED)
            self.status_var.set("Starting...")
            # Short delay in case the board still reset on connect
            self.root.after(500, lambda: self._send_start_after_delay(keys, min_ms, max_ms))
            return
        self._do_send_start(keys, min_ms, max_ms)

    def _send_start_after_delay(self, keys, min_ms, max_ms):
        """Called 1.2s after opening port so Pico is ready."""
        self._do_send_start(keys, min_ms, max_ms)

    def _do_send_start(self, keys, min_ms, max_ms):
        keys_str = ",".join(keys)
        cmd = f"START;{keys_str};{min_ms};{max_ms}"
        if not self._serial_send(cmd):
            self.start_btn.config(state=tk.NORMAL)
            messagebox.showerror("Send failed", "Could not send START to Pico. Check connection.")
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
            self.status_var.set("pynput not installed; F12 hotkey disabled. pip install pynput")

    def run(self):
        self.root.protocol("WM_DELETE_WINDOW", self._on_close)
        self.root.mainloop()

    def _on_close(self):
        if self.running:
            self._on_stop()
        self._close_serial()
        self.root.destroy()


if __name__ == "__main__":
    app = AutoclickerApp()
    app.run()
