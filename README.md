# Pico HID Autoclicker

Raspberry Pi Pico (CircuitPython v10) acts as a USB HID keyboard. A Windows GUI configures which keys to press and the delay range; the Pico sends key-down + key-up in sequence at a random interval between min and max delay. **F12** toggles Start/Stop globally.

The GUI has five tabs:

| Tab | Purpose |
|-----|---------|
| **Autoclicker** | Simple key sequence with random delay |
| **Conditional Clicker** | OCR-based HP monitoring, template-matched status effects, auto-targeting, stuck detection |
| **Status Effects Library** | Capture, preview, and manage template images for pattern matching |
| **Mouse Clicker** | Click screen coordinates in sequence with active/pause cycling |
| **Profiles** | Save and load full configuration sets as JSON files |

## Project structure

```
pico/
  code.py                  CircuitPython firmware (USB HID keyboard + serial listener)

windows/
  autoclicker_gui.py       Main Tkinter GUI (all five tabs)
  screen_reader.py         Screen capture, color analysis, OCR (EasyOCR), template matching (OpenCV)
  region_selector.py       Fullscreen overlay for drag-to-select screen regions
  template_store.py        Save/load/delete template PNG images with JSON metadata
  templates/               Stored template images and templates.json
  configs/                 Profile JSON files
  requirements.txt         Python dependencies

send_test_command.py       CLI utility to send serial commands to the Pico without the GUI
```

## Pico setup (CircuitPython v10)

1. Install [CircuitPython 10](https://circuitpython.org/board/raspberry_pi_pico/) on the Pico (drag the UF2 to the device).
2. Install the HID library:
   - Download [Adafruit_CircuitPython_HID](https://github.com/adafruit/Adafruit_CircuitPython_HID/releases) and copy the `adafruit_hid` folder onto the Pico's USB drive (next to `code.py`).
3. Copy `pico/code.py` from this repo to the Pico's drive as `code.py`. The board will reboot and run it.

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

### Dependencies

| Package | Used for |
|---------|----------|
| `pyserial` | Serial communication with the Pico |
| `pynput` | Global F12 hotkey listener |
| `mss` | Fast screen capture |
| `Pillow` | Image processing |
| `opencv-python` | Template matching (normalised cross-correlation) |
| `numpy` | Image array operations |
| `easyocr` | Reading HP percentage text from screen |

The Conditional Clicker, Status Effects Library, and their dependencies (`mss`, `Pillow`, `opencv-python`, `numpy`, `easyocr`) are optional — the Autoclicker and Mouse Clicker tabs work with only `pyserial` and `pynput`.

## Usage

### Tab 1 — Autoclicker

1. Connect the Pico via USB. Click **Refresh** if the COM port list is empty; select the Pico's COM port.
2. Add keys: choose a key from the dropdown and click **Add**. Order is the order they are pressed each cycle.
3. Set **Min delay (ms)** and **Max delay (ms)** — each cycle waits a random time in this range.
4. Click **Start** to begin. Click **Stop** or press **F12** to stop.

### Tab 2 — Conditional Clicker

Monitors a screen region via OCR to read an HP percentage and uses template matching to detect status effects. Designed for game automation with these features:

- **Screen region** — select or type coordinates for the HP bar area. Live preview shows the captured region and detected colors.
- **Targeting key** — pressed when no HP is visible (no target). Configurable min/max delay.
- **Attack keys** — pressed in rotation while a target is alive. Each key has its own min/max delay.
- **Instant keys** — pressed once immediately when a new target is acquired.
- **On-death key** — optionally pressed when the target's HP reaches zero.
- **Stuck detection** — if HP doesn't change for a configurable timeout, the bot presses an unstuck movement sequence (two directional key holds).
- **No-target timeout** — stops the bot if no target is found within a configurable duration.
- **Attack start delay** — wait time after targeting before attacks begin (time to approach the target).
- **Status effect keys** — bound to template images; pressed when the template is not detected on screen.
- **Buff keys** — like status effects but checked on a periodic timer rather than every cycle.
- **HP reading verification** — configurable OCR threshold, upscale factor, dimmer fallback threshold, and a "gone timeout" to filter out transient OCR misreads.

### Tab 3 — Status Effects Library

Capture and manage template images used by the Conditional Clicker's status effect and buff detection:

- **Capture** — select a screen region to save as a named template PNG.
- **Preview** — view saved templates with their metadata.
- **Test Match** — check if a template is currently visible on screen.
- **Delete** — remove templates from the library.

Templates are stored in `windows/templates/` as PNG files with a `templates.json` metadata sidecar.

### Tab 4 — Mouse Clicker

Automates mouse clicks at specific screen coordinates (runs on the PC, no Pico needed):

- **Coordinate tracker** — shows live mouse X/Y position to help identify targets.
- **Click targets** — add multiple (x, y) pairs; they are clicked in order each cycle. Each target has a **Capture** button to grab the current mouse position.
- **Timing** — min/max delay between clicks and a start delay before clicking begins.
- **Active/Pause cycle** — click for N seconds, then pause for M seconds, and repeat. Set pause to 0 for continuous clicking.
- **F12** toggles start/stop when this tab is active.

### Tab 5 — Profiles

Save and load the entire application state (all tab settings) as named JSON profiles:

- **Save / Save As** — write current settings to a profile in `windows/configs/`.
- **Load** — restore all tabs from a saved profile.
- **Delete** — remove a profile.
- Settings are auto-saved to the active profile on exit.

Legacy `config.json` files are automatically migrated to the profiles directory on first run.

## Serial protocol

All commands are newline-terminated and sent at 115200 baud.

| Command | Format | Example |
|---------|--------|---------|
| **Start** | `START;<key1>,<key2>,...;<min_ms>;<max_ms>\n` | `START;a,b;1000;1500` |
| **Stop** | `STOP\n` | `STOP` |
| **Press** | `PRESS;<key>\n` | `PRESS;f1` |
| **Hold** | `HOLD;<key>;<duration_ms>\n` | `HOLD;left;1000` |

Supported key names: `a`–`z`, `0`–`9`, `f1`–`f12`, `space`, `enter`, `tab`, `escape`, `backspace`, `minus`, `equals`, `left_bracket`, `right_bracket`, `backslash`, `semicolon`, `quote`, `grave`, `comma`, `period`, `slash`, `insert`, `delete`, `home`, `end`, `page_up`, `page_down`, `up`, `down`, `left`, `right`.

## Testing without the GUI

```bash
python send_test_command.py COM3
python send_test_command.py COM3 '{"keys": ["a"], "interval_ms": 3000, "randomness_ms": 500, "running": true}'
```

## License

MIT
