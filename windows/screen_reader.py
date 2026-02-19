"""
Screen capture and color analysis utilities for the Conditional Clicker.
Uses mss for fast screen grabs and Pillow for image processing.
"""

import mss
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
