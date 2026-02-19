"""
Send one JSON command to the Pico (for testing without the full app).
Close Thonny (and any other app using the port) first.

Requires: pip install pyserial

Usage:
  python send_test_command.py COM3
  python send_test_command.py COM3 "{\"keys\": [\"KEY_SPACE\"], \"interval_ms\": 3000, \"randomness_ms\": 500, \"running\": true}"
"""
import sys
try:
    import serial
except ImportError:
    print("Missing dependency. Run: pip install pyserial")
    sys.exit(1)

DEFAULT_JSON = '{"keys": ["a"], "interval_ms": 3000, "randomness_ms": 500, "running": true}'

def main():
    if len(sys.argv) < 2:
        print(__doc__.strip())
        print("\nExample: python send_test_command.py COM3")
        sys.exit(1)
    port = sys.argv[1]
    payload = sys.argv[2] if len(sys.argv) > 2 else DEFAULT_JSON
    line = payload.strip() + "\n"
    try:
        with serial.Serial(port, 115200, timeout=1) as ser:
            ser.write(line.encode("utf-8"))
        print("Sent:", line.strip())
    except serial.SerialException as e:
        print("Error:", e)
        sys.exit(1)

if __name__ == "__main__":
    main()
