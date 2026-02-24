"""
Screen capture and color/image analysis utilities for the Conditional Clicker.
Uses mss for fast screen grabs, Pillow for image processing, OpenCV for
template matching, and EasyOCR for reading HP percentage text.
"""

import mss
import numpy as np
import cv2
from PIL import Image

MAX_ANALYSIS_PIXELS = 10000


def capture_region(x, y, w, h):
    """Grab a screen region and return a PIL RGB Image."""
    with mss.mss() as sct:
        monitor = {"top": y, "left": x, "width": w, "height": h}
        screenshot = sct.grab(monitor)
        return Image.frombytes("RGB", screenshot.size, screenshot.rgb)


def get_unique_colors(image, tolerance=30):
    """Extract unique hex colors, grouping pixels within Euclidean RGB distance *tolerance*.

    Returns a list of hex strings sorted by frequency (most common first).
    Large images are downsampled before analysis to keep things fast.
    """
    w, h = image.size
    total = w * h
    if total > MAX_ANALYSIS_PIXELS:
        scale = (MAX_ANALYSIS_PIXELS / total) ** 0.5
        image = image.resize(
            (max(1, int(w * scale)), max(1, int(h * scale))), Image.NEAREST
        )

    tol_sq = tolerance * tolerance
    groups = []  # [(representative_rgb_tuple, count), ...]

    for pixel in image.getdata():
        r, g, b = pixel[:3]
        matched = False
        for i, (rep, count) in enumerate(groups):
            dr = r - rep[0]
            dg = g - rep[1]
            db = b - rep[2]
            if dr * dr + dg * dg + db * db <= tol_sq:
                groups[i] = (rep, count + 1)
                matched = True
                break
        if not matched:
            groups.append(((r, g, b), 1))

    groups.sort(key=lambda g: g[1], reverse=True)
    return [f"#{r:02X}{g:02X}{b:02X}" for (r, g, b), _ in groups]


def color_present(image, hex_color, tolerance=30):
    """Return True if any pixel matches *hex_color* within Euclidean RGB *tolerance*."""
    hex_color = hex_color.lstrip("#")
    tr = int(hex_color[0:2], 16)
    tg = int(hex_color[2:4], 16)
    tb = int(hex_color[4:6], 16)
    tol_sq = tolerance * tolerance

    for pixel in image.getdata():
        r, g, b = pixel[:3]
        dr = r - tr
        dg = g - tg
        db = b - tb
        if dr * dr + dg * dg + db * db <= tol_sq:
            return True
    return False


def count_color_pixels(image, hex_color, tolerance=30):
    """Return the number of pixels matching *hex_color* within Euclidean RGB *tolerance*."""
    hex_color = hex_color.lstrip("#")
    tr = int(hex_color[0:2], 16)
    tg = int(hex_color[2:4], 16)
    tb = int(hex_color[4:6], 16)
    tol_sq = tolerance * tolerance
    count = 0

    for pixel in image.getdata():
        r, g, b = pixel[:3]
        dr = r - tr
        dg = g - tg
        db = b - tb
        if dr * dr + dg * dg + db * db <= tol_sq:
            count += 1
    return count


# ---------------------------------------------------------------------------
#  EasyOCR lazy singleton – loaded once on first call to avoid re-loading
#  the ~200 MB model on every HP check.
# ---------------------------------------------------------------------------

_ocr_reader = None


def _get_ocr_reader():
    global _ocr_reader
    if _ocr_reader is None:
        import easyocr
        _ocr_reader = easyocr.Reader(["en"], gpu=False, verbose=False)
    return _ocr_reader


def _ocr_threshold_attempt(
    gray: np.ndarray,
    threshold: int,
    ocr_scale: int,
    reader,
) -> float | None:
    """Run OCR on *gray* image with a single binary threshold. Returns parsed value or None."""
    _, thresh = cv2.threshold(gray, threshold, 255, cv2.THRESH_BINARY)
    h, w = thresh.shape
    upscaled = cv2.resize(thresh, (w * ocr_scale, h * ocr_scale), interpolation=cv2.INTER_NEAREST)
    results = reader.readtext(upscaled, allowlist="0123456789.%", detail=0)
    text = "".join(results).strip().replace("%", "")
    if not text:
        return None
    try:
        value = float(text)
        if 0 <= value <= 100:
            return value
    except ValueError:
        pass
    return None


def read_hp_percentage(
    image: Image.Image,
    ocr_threshold: int = 125,
    ocr_scale: int = 5,
    dimmer_threshold: int = 50,
) -> float | None:
    """Read the HP percentage text from *image* (a tightly cropped HP bar region).

    *ocr_threshold*    — binary threshold applied to grayscale (0-255).
    *ocr_scale*        — integer upscale factor before OCR.
    *dimmer_threshold* — fallback binary threshold for dim text (e.g. "0.00%"
                         on an empty HP bar).  Set to 0 to disable the fallback.

    Returns the numeric value (e.g. ``99.09``) or ``None`` if the text
    cannot be parsed.
    """
    gray = cv2.cvtColor(np.array(image), cv2.COLOR_RGB2GRAY)
    reader = _get_ocr_reader()

    result = _ocr_threshold_attempt(gray, ocr_threshold, ocr_scale, reader)
    if result is not None:
        return result

    if dimmer_threshold > 0:
        return _ocr_threshold_attempt(gray, dimmer_threshold, ocr_scale, reader)
    return None


def match_template(screen_image: Image.Image, template_image: Image.Image,
                   threshold: float = 0.8) -> bool:
    """Return ``True`` if *template_image* is found inside *screen_image*.

    Uses OpenCV normalised cross-correlation (``TM_CCOEFF_NORMED``).
    """
    screen_bgr = cv2.cvtColor(np.array(screen_image), cv2.COLOR_RGB2BGR)
    template_bgr = cv2.cvtColor(np.array(template_image), cv2.COLOR_RGB2BGR)

    sh, sw = screen_bgr.shape[:2]
    th, tw = template_bgr.shape[:2]
    if th > sh or tw > sw:
        return False

    result = cv2.matchTemplate(screen_bgr, template_bgr, cv2.TM_CCOEFF_NORMED)
    _, max_val, _, _ = cv2.minMaxLoc(result)
    return max_val >= threshold
