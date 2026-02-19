# Pico HID Autoclicker

Raspberry Pi Pico (CircuitPython v10) acts as a USB HID keyboard. A Windows GUI configures which keys to press and the delay range; the Pico sends key-down + key-up in sequence at a random interval between min and max delay. Start/Stop from the app or with **F12**.

## Parts

- **Pico** (`pico/code.py`): Runs on the board. Listens on USB serial for `START;keys;min_ms;max_ms` and `STOP`; sends the key sequence at a random delay (min_ms–max_ms) until STOP.
- **Windows** (`windows/`): GUI to pick COM port, add keys, set min/max delay (ms), and Start/Stop. F12 toggles run state globally.

## Pico setup (CircuitPython v10)

1. Install [CircuitPython 10](https://circuitpython.org/board/raspberry_pi_pico/) on the Pico (drag UF2 to the device).
2. Install the HID library:
   - Download [Adafruit_CircuitPython_HID](https://github.com/adafruit/Adafruit_CircuitPython_HID/releases) and copy the `adafruit_hid` folder onto the Pico’s USB drive (next to `code.py`).
3. Copy `pico/code.py` from this repo to the Pico’s drive as `code.py`. The board will reboot and run it.

The Pico will show up as a **keyboard** (HID) and as a **serial port** (COM port on Windows). You can use the default console serial; no `boot.py` change is required. If you enable a second serial with `usb_cdc.data` in `boot.py`, the code will use it automatically.

## Windows setup

1. Python 3.8+.
2. Install dependencies:
   ```bash
   cd windows
   pip install -r requirements.txt
   ```
3. Run the GUI:
   ```bash
   python autoclicker_gui.py
   ```

## Usage

1. Connect the Pico via USB. In the app, click **Refresh** if the COM port list is empty; select the Pico’s COM port.
2. Add keys: choose a key from the dropdown and click **Add**. Order is the order they are pressed (e.g. add `a` then `b` to press a, then b, each cycle).
3. Set **Min delay (ms)** and **Max delay (ms)**. Each cycle waits a random time in this range (e.g. 1000 and 1500 → 1–1.5 s between sequences).
4. Click **Start** to begin. The Pico will repeatedly wait a random delay, then press the keys in order (key down + key up for each).
5. Click **Stop** or press **F12** (anywhere) to stop.

## Protocol (serial)

- **Start:** `START;<key1>,<key2>,...;<min_ms>;<max_ms>\n`  
  Example: `START;a,b;1000;1500`
- **Stop:** `STOP\n`

Keys are lowercase names: `a`–`z`, `0`–`9`, `f1`–`f12`, `space`, `enter`, `tab`, `escape`, `backspace`, and others (see the GUI dropdown).

## License

MIT
