"""
Fullscreen screenshot overlay for drag-to-select screen region.

Approach (reliable on Windows):
1. Capture the entire screen with mss.
2. Display the frozen screenshot in a borderless Toplevel.
3. User clicks and drags to draw a selection rectangle.
4. On release, the absolute screen coordinates are passed to a callback.
"""

import tkinter as tk

import mss as mss_module
from PIL import Image, ImageTk


class RegionSelector:
    """Show a frozen-screenshot overlay; let the user drag-select a rectangle.

    Usage::

        RegionSelector(parent_window, callback)

    *callback(x, y, w, h)* is called with absolute screen coordinates
    when the user finishes the selection.  Pressing Escape cancels.
    """

    def __init__(self, parent, callback):
        self.callback = callback
        self.start_x = 0
        self.start_y = 0
        self.rect_id = None

        with mss_module.mss() as sct:
            monitor = sct.monitors[0]
            screenshot = sct.grab(monitor)
            self.screenshot_img = Image.frombytes("RGB", screenshot.size, screenshot.rgb)
            self.offset_x = monitor["left"]
            self.offset_y = monitor["top"]
            scr_w = monitor["width"]
            scr_h = monitor["height"]

        self.overlay = tk.Toplevel(parent)
        self.overlay.overrideredirect(True)
        self.overlay.geometry(f"{scr_w}x{scr_h}+{self.offset_x}+{self.offset_y}")
        self.overlay.attributes("-topmost", True)
        self.overlay.configure(cursor="crosshair")

        self.tk_image = ImageTk.PhotoImage(self.screenshot_img)
        self.canvas = tk.Canvas(self.overlay, highlightthickness=0)
        self.canvas.pack(fill=tk.BOTH, expand=True)
        self.canvas.create_image(0, 0, anchor=tk.NW, image=self.tk_image)

        self.canvas.bind("<ButtonPress-1>", self._on_press)
        self.canvas.bind("<B1-Motion>", self._on_drag)
        self.canvas.bind("<ButtonRelease-1>", self._on_release)
        self.overlay.bind("<Escape>", lambda _e: self.overlay.destroy())

    def _on_press(self, event):
        self.start_x = event.x
        self.start_y = event.y
        if self.rect_id:
            self.canvas.delete(self.rect_id)
        self.rect_id = self.canvas.create_rectangle(
            self.start_x, self.start_y, self.start_x, self.start_y,
            outline="red", width=2,
        )

    def _on_drag(self, event):
        if self.rect_id:
            self.canvas.coords(
                self.rect_id, self.start_x, self.start_y, event.x, event.y
            )

    def _on_release(self, event):
        x1 = min(self.start_x, event.x)
        y1 = min(self.start_y, event.y)
        x2 = max(self.start_x, event.x)
        y2 = max(self.start_y, event.y)
        w = x2 - x1
        h = y2 - y1
        self.overlay.destroy()
        if w > 2 and h > 2:
            self.callback(x1 + self.offset_x, y1 + self.offset_y, w, h)
