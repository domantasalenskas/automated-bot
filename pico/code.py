# SPDX-FileCopyrightText: 2025
# SPDX-License-Identifier: MIT
"""
Pico HID Autoclicker - CircuitPython v10
Listens on USB serial for START/STOP; sends key sequence at random interval.
Copy this file to the Pico as code.py (with adafruit_hid on the board).
"""

import time
import random
import usb_cdc
import usb_hid
from adafruit_hid.keyboard import Keyboard
from adafruit_hid.keycode import Keycode

# Use console so the same COM port the PC opens (e.g. Thonny / this app) is what we read
serial = usb_cdc.console

keyboard = Keyboard(usb_hid.devices)

# Key name (from GUI) -> Keycode
KEY_MAP = {}
for c in "abcdefghijklmnopqrstuvwxyz":
    KEY_MAP[c] = getattr(Keycode, c.upper())
for i in range(10):
    name = str(i)
    KEY_MAP[name] = [Keycode.ZERO, Keycode.ONE, Keycode.TWO, Keycode.THREE, Keycode.FOUR,
                     Keycode.FIVE, Keycode.SIX, Keycode.SEVEN, Keycode.EIGHT, Keycode.NINE][i]
for i in range(1, 13):
    KEY_MAP["f" + str(i)] = getattr(Keycode, "F" + str(i))
KEY_MAP["space"] = Keycode.SPACE
KEY_MAP["enter"] = Keycode.ENTER
KEY_MAP["tab"] = Keycode.TAB
KEY_MAP["escape"] = Keycode.ESCAPE
KEY_MAP["backspace"] = Keycode.BACKSPACE
KEY_MAP["minus"] = Keycode.MINUS
KEY_MAP["equals"] = Keycode.EQUALS
KEY_MAP["left_bracket"] = Keycode.LEFT_BRACKET
KEY_MAP["right_bracket"] = Keycode.RIGHT_BRACKET
KEY_MAP["backslash"] = Keycode.BACKSLASH
KEY_MAP["semicolon"] = Keycode.SEMICOLON
KEY_MAP["quote"] = Keycode.QUOTE
KEY_MAP["grave"] = Keycode.GRAVE_ACCENT
KEY_MAP["comma"] = Keycode.COMMA
KEY_MAP["period"] = Keycode.PERIOD
KEY_MAP["slash"] = Keycode.FORWARD_SLASH
KEY_MAP["insert"] = Keycode.INSERT
KEY_MAP["delete"] = Keycode.DELETE
KEY_MAP["home"] = Keycode.HOME
KEY_MAP["end"] = Keycode.END
KEY_MAP["page_up"] = Keycode.PAGE_UP
KEY_MAP["page_down"] = Keycode.PAGE_DOWN
KEY_MAP["up"] = Keycode.UP_ARROW
KEY_MAP["down"] = Keycode.DOWN_ARROW
KEY_MAP["left"] = Keycode.LEFT_ARROW
KEY_MAP["right"] = Keycode.RIGHT_ARROW

BAUD = 115200
CMD_START = "START"
CMD_STOP = "STOP"
CMD_PRESS = "PRESS"
CMD_HOLD = "HOLD"


def read_serial_line():
    """Non-blocking read; return a full line (without newline) or None."""
    if serial.in_waiting == 0:
        return None
    raw = serial.read(serial.in_waiting)
    if not raw:
        return None
    return raw.decode("utf-8", "ignore").strip()


def parse_start(line):
    """Parse START;keys;min_ms;max_ms -> (key_codes_list, min_ms, max_ms) or None."""
    parts = line.split(";")
    if len(parts) != 4:
        return None
    cmd, keys_str, min_s, max_s = parts
    if cmd.strip().upper() != CMD_START:
        return None
    try:
        min_ms = int(min_s)
        max_ms = int(max_s)
    except ValueError:
        return None
    if min_ms < 0 or max_ms < min_ms:
        return None
    key_names = [k.strip().lower() for k in keys_str.split(",") if k.strip()]
    key_codes = []
    for name in key_names:
        if name not in KEY_MAP:
            return None
        key_codes.append(KEY_MAP[name])
    return (key_codes, min_ms, max_ms)


def run_loop(keys, min_ms, max_ms):
    """Send key sequence repeatedly with random delay; check serial for STOP."""
    running = True
    buf = ""
    while running:
        # Check for STOP (non-blocking)
        if serial.in_waiting:
            raw = serial.read(serial.in_waiting)
            buf += raw.decode("utf-8", "ignore")
            if CMD_STOP in buf or buf.strip().upper() == CMD_STOP:
                running = False
                break
            if "\n" in buf or "\r" in buf:
                lines = buf.replace("\r", "\n").split("\n")
                buf = lines[-1]
                for line in lines[:-1]:
                    if line.strip().upper() == CMD_STOP:
                        running = False
                        break
                if not running:
                    break

        if not running:
            break

        # Wait random delay (seconds)
        delay_s = random.uniform(min_ms / 1000.0, max_ms / 1000.0)
        time.sleep(delay_s)

        if not running:
            break

        # Send each key: down then up
        for kc in keys:
            keyboard.press(kc)
            keyboard.release(kc)
            time.sleep(0.02)

    # Drain any remaining serial
    if serial.in_waiting:
        serial.read(serial.in_waiting)


def main():
    line_buf = ""
    while True:
        if serial.in_waiting:
            raw = serial.read(serial.in_waiting)
            line_buf += raw.decode("utf-8", "ignore")
        else:
            time.sleep(0.01)
            continue

        while "\n" in line_buf or "\r" in line_buf:
            parts = line_buf.replace("\r\n", "\n").replace("\r", "\n").split("\n", 1)
            if len(parts) < 2:
                break
            line, line_buf = parts[0].strip(), parts[1]
            if not line:
                continue

            if line.upper() == CMD_STOP:
                continue

            if line.upper().startswith(CMD_PRESS + ";"):
                parts = line.split(";")
                if len(parts) == 2:
                    key_name = parts[1].strip().lower()
                    if key_name in KEY_MAP:
                        keyboard.press(KEY_MAP[key_name])
                        time.sleep(0.05)
                        keyboard.release_all()

            elif line.upper().startswith(CMD_HOLD + ";"):
                parts = line.split(";")
                if len(parts) == 3:
                    key_name = parts[1].strip().lower()
                    try:
                        duration_ms = int(parts[2].strip())
                    except ValueError:
                        duration_ms = 0
                    if key_name in KEY_MAP and duration_ms > 0:
                        keyboard.press(KEY_MAP[key_name])
                        time.sleep(duration_ms / 1000.0)
                        keyboard.release_all()

            elif line.upper().startswith(CMD_START + ";"):
                parsed = parse_start(line)
                if parsed:
                    keys, min_ms, max_ms = parsed
                    run_loop(keys, min_ms, max_ms)


main()
