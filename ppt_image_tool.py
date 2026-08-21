from __future__ import annotations

import json
import os
import subprocess
import sys
import threading
import tkinter as tk
import tkinter.font as tkfont
from pathlib import Path
from tkinter import filedialog, messagebox, ttk
from typing import Callable, Iterable, Literal

from PIL import Image, ImageDraw, ImageOps, ImageTk
from tkinterdnd2 import TkinterDnD

from converter_core import (
    ImagesToPptOptions,
    PptExportOptions,
    SUPPORTED_IMAGES,
    SUPPORTED_PRESENTATIONS,
    detect_export_backends,
    export_presentations_to_images,
    images_to_presentation,
    natural_sort_key,
)
from native_drop import NativeFileDrop

APP_NAME = "PPT 图片互转工具"
APP_VERSION = "3.3.0"
CONFIG_PATH = Path.home() / ".ppt_image_converter_config.json"

RADIUS_CARD = 12
RADIUS_CONTROL = 8
RADIUS_SMALL = 6

# Ant Design-inspired semantic tokens. Keep all UI color decisions centralized.
COLORS = {
    "app": "#F5F5F5",
    "surface": "#FFFFFF",
    "surface_alt": "#FAFAFA",
    "sidebar": "#FFFFFF",
    "sidebar_hover": "#F5F5F5",
    "sidebar_active": "#E6F4FF",
    "primary": "#1677FF",
    "primary_hover": "#4096FF",
    "primary_active": "#0958D9",
    "primary_soft": "#E6F4FF",
    "primary_border": "#91CAFF",
    "text": "#1F1F1F",
    "muted": "#595959",
    "subtle": "#8C8C8C",
    "disabled": "#BFBFBF",
    "border": "#D9D9D9",
    "border_soft": "#F0F0F0",
    "success": "#52C41A",
    "success_soft": "#F6FFED",
    "warning": "#FAAD14",
    "warning_soft": "#FFFBE6",
    "danger": "#FF4D4F",
    "danger_soft": "#FFF2F0",
    "code": "#141414",
}

FONT = "Microsoft YaHei UI"
Mode = Literal["ppt", "images"]


def resource_path(name: str) -> Path:
    base = Path(getattr(sys, "_MEIPASS", Path(__file__).resolve().parent))
    return base / name


def human_size(size: int) -> str:
    value = float(max(0, size))
    units = ("B", "KB", "MB", "GB")
    for unit in units:
        if value < 1024 or unit == units[-1]:
            return f"{value:.0f} {unit}" if unit == "B" else f"{value:.1f} {unit}"
        value /= 1024
    return f"{value:.1f} GB"


def unique_existing(paths: Iterable[Path]) -> list[Path]:
    output: list[Path] = []
    seen: set[str] = set()
    for raw in paths:
        try:
            item = raw.expanduser().resolve()
        except Exception:
            item = raw
        key = os.path.normcase(str(item))
        if key not in seen and item.exists():
            output.append(item)
            seen.add(key)
    return output


def rounded_points(x1: float, y1: float, x2: float, y2: float, radius: float) -> list[float]:
    """Return smooth polygon points for a rounded rectangle."""
    radius = max(0.0, min(radius, (x2 - x1) / 2, (y2 - y1) / 2))
    return [
        x1 + radius, y1,
        x2 - radius, y1,
        x2, y1,
        x2, y1 + radius,
        x2, y2 - radius,
        x2, y2,
        x2 - radius, y2,
        x1 + radius, y2,
        x1, y2,
        x1, y2 - radius,
        x1, y1 + radius,
        x1, y1,
    ]


def draw_icon(canvas: tk.Canvas, name: str, cx: float, cy: float, color: str, size: float = 16, width: int = 2) -> None:
    """Draw a small vector icon without relying on platform-specific icon fonts."""
    s = size / 2
    kw = {"fill": color, "width": width, "capstyle": "round", "joinstyle": "round"}
    if name == "plus":
        canvas.create_line(cx - s * .55, cy, cx + s * .55, cy, **kw)
        canvas.create_line(cx, cy - s * .55, cx, cy + s * .55, **kw)
    elif name == "trash":
        canvas.create_rectangle(cx - s * .45, cy - s * .25, cx + s * .45, cy + s * .62, outline=color, width=width)
        canvas.create_line(cx - s * .62, cy - s * .48, cx + s * .62, cy - s * .48, **kw)
        canvas.create_line(cx - s * .26, cy - s * .68, cx + s * .26, cy - s * .68, **kw)
        canvas.create_line(cx - s * .18, cy - s * .08, cx - s * .18, cy + s * .4, **kw)
        canvas.create_line(cx + s * .18, cy - s * .08, cx + s * .18, cy + s * .4, **kw)
    elif name == "folder":
        pts = [cx - s * .72, cy - s * .38, cx - s * .18, cy - s * .38, cx, cy - s * .16, cx + s * .72, cy - s * .16,
               cx + s * .62, cy + s * .52, cx - s * .72, cy + s * .52]
        canvas.create_polygon(pts, fill="", outline=color, width=width, joinstyle="round")
    elif name == "sort":
        canvas.create_line(cx - s * .35, cy - s * .58, cx - s * .35, cy + s * .52, **kw)
        canvas.create_line(cx - s * .58, cy - s * .32, cx - s * .35, cy - s * .58, cx - s * .12, cy - s * .32, **kw)
        canvas.create_line(cx + s * .35, cy + s * .58, cx + s * .35, cy - s * .52, **kw)
        canvas.create_line(cx + s * .12, cy + s * .32, cx + s * .35, cy + s * .58, cx + s * .58, cy + s * .32, **kw)
    elif name == "play":
        canvas.create_polygon(cx - s * .34, cy - s * .55, cx + s * .58, cy, cx - s * .34, cy + s * .55, fill=color, outline="")
    elif name == "stop":
        canvas.create_rectangle(cx - s * .42, cy - s * .42, cx + s * .42, cy + s * .42, fill=color, outline="")
    elif name == "external":
        canvas.create_rectangle(cx - s * .62, cy - s * .35, cx + s * .28, cy + s * .55, outline=color, width=width)
        canvas.create_line(cx - s * .02, cy - s * .5, cx + s * .55, cy - s * .5, cx + s * .55, cy + s * .07, **kw)
        canvas.create_line(cx + s * .55, cy - s * .5, cx - s * .06, cy + s * .1, **kw)
    elif name == "history":
        canvas.create_arc(cx - s * .62, cy - s * .62, cx + s * .62, cy + s * .62, start=38, extent=285, style="arc", outline=color, width=width)
        canvas.create_line(cx - s * .58, cy - s * .3, cx - s * .65, cy + s * .08, cx - s * .28, cy + s * .02, **kw)
        canvas.create_line(cx, cy - s * .34, cx, cy + s * .05, cx + s * .3, cy + s * .24, **kw)
    elif name == "file":
        canvas.create_polygon(cx - s * .5, cy - s * .65, cx + s * .18, cy - s * .65, cx + s * .5, cy - s * .32,
                              cx + s * .5, cy + s * .65, cx - s * .5, cy + s * .65, fill="", outline=color, width=width)
        canvas.create_line(cx + s * .18, cy - s * .65, cx + s * .18, cy - s * .32, cx + s * .5, cy - s * .32, **kw)
        canvas.create_line(cx - s * .27, cy + s * .05, cx + s * .25, cy + s * .05, **kw)
        canvas.create_line(cx - s * .27, cy + s * .32, cx + s * .18, cy + s * .32, **kw)
    elif name == "image":
        canvas.create_rectangle(cx - s * .65, cy - s * .52, cx + s * .65, cy + s * .52, outline=color, width=width)
        canvas.create_oval(cx + s * .22, cy - s * .3, cx + s * .43, cy - s * .09, outline=color, width=width)
        canvas.create_line(cx - s * .5, cy + s * .32, cx - s * .12, cy - s * .03, cx + s * .08, cy + s * .15,
                           cx + s * .3, cy - s * .05, cx + s * .53, cy + s * .26, **kw)
    elif name == "settings":
        canvas.create_oval(cx - s * .22, cy - s * .22, cx + s * .22, cy + s * .22, outline=color, width=width)
        import math
        for angle in range(0, 360, 60):
            a = math.radians(angle)
            canvas.create_line(cx + math.cos(a) * s * .42, cy + math.sin(a) * s * .42,
                               cx + math.cos(a) * s * .67, cy + math.sin(a) * s * .67, **kw)
    elif name == "upload":
        canvas.create_line(cx, cy + s * .35, cx, cy - s * .55, **kw)
        canvas.create_line(cx - s * .32, cy - s * .22, cx, cy - s * .55, cx + s * .32, cy - s * .22, **kw)
        canvas.create_line(cx - s * .55, cy + s * .22, cx - s * .55, cy + s * .58, cx + s * .55, cy + s * .58, cx + s * .55, cy + s * .22, **kw)
    elif name == "info":
        canvas.create_oval(cx - s * .55, cy - s * .55, cx + s * .55, cy + s * .55, outline=color, width=width)
        canvas.create_line(cx, cy - s * .08, cx, cy + s * .32, **kw)
        canvas.create_oval(cx - 1, cy - s * .34 - 1, cx + 1, cy - s * .34 + 1, fill=color, outline=color)
    elif name == "warning":
        canvas.create_polygon(cx, cy - s * .62, cx + s * .6, cy + s * .5, cx - s * .6, cy + s * .5,
                              fill="", outline=color, width=width)
        canvas.create_line(cx, cy - s * .22, cx, cy + s * .12, **kw)
        canvas.create_oval(cx - 1, cy + s * .28 - 1, cx + 1, cy + s * .28 + 1, fill=color, outline=color)
    elif name == "check":
        canvas.create_line(cx - s * .55, cy, cx - s * .12, cy + s * .38, cx + s * .62, cy - s * .48, **kw)
    elif name == "zoom":
        canvas.create_oval(cx - s * .58, cy - s * .58, cx + s * .2, cy + s * .2, outline=color, width=width)
        canvas.create_line(cx + s * .08, cy + s * .08, cx + s * .62, cy + s * .62, **kw)
        canvas.create_line(cx - s * .33, cy - s * .18, cx - s * .05, cy - s * .18, **kw)
        canvas.create_line(cx - s * .19, cy - s * .32, cx - s * .19, cy - s * .04, **kw)
    elif name == "minus":
        canvas.create_line(cx - s * .55, cy, cx + s * .55, cy, **kw)
    elif name == "chevron_left":
        canvas.create_line(cx + s * .25, cy - s * .52, cx - s * .28, cy, cx + s * .25, cy + s * .52, **kw)
    elif name == "chevron_right":
        canvas.create_line(cx - s * .25, cy - s * .52, cx + s * .28, cy, cx - s * .25, cy + s * .52, **kw)
    elif name == "fit":
        canvas.create_line(cx - s * .58, cy - s * .2, cx - s * .58, cy - s * .58, cx - s * .2, cy - s * .58, **kw)
        canvas.create_line(cx + s * .2, cy - s * .58, cx + s * .58, cy - s * .58, cx + s * .58, cy - s * .2, **kw)
        canvas.create_line(cx - s * .58, cy + s * .2, cx - s * .58, cy + s * .58, cx - s * .2, cy + s * .58, **kw)
        canvas.create_line(cx + s * .2, cy + s * .58, cx + s * .58, cy + s * .58, cx + s * .58, cy + s * .2, **kw)
    elif name == "actual":
        canvas.create_rectangle(cx - s * .56, cy - s * .46, cx + s * .56, cy + s * .46, outline=color, width=width)
        canvas.create_line(cx - s * .32, cy - s * .2, cx + s * .32, cy - s * .2, **kw)
        canvas.create_line(cx - s * .32, cy + s * .08, cx + s * .18, cy + s * .08, **kw)
    elif name == "close":
        canvas.create_line(cx - s * .48, cy - s * .48, cx + s * .48, cy + s * .48, **kw)
        canvas.create_line(cx + s * .48, cy - s * .48, cx - s * .48, cy + s * .48, **kw)
    elif name == "ppt":
        canvas.create_rectangle(cx - s * .64, cy - s * .55, cx + s * .64, cy + s * .55, outline=color, width=width)
        canvas.create_line(cx - s * .2, cy - s * .55, cx - s * .2, cy + s * .55, **kw)
        canvas.create_text(cx - s * .42, cy, text="P", fill=color, font=("Segoe UI", max(7, int(size * .42)), "bold"))
        canvas.create_line(cx + s * .02, cy - s * .22, cx + s * .42, cy - s * .22, **kw)
        canvas.create_line(cx + s * .02, cy + s * .05, cx + s * .5, cy + s * .05, **kw)
    else:
        canvas.create_oval(cx - 1, cy - 1, cx + 1, cy + 1, fill=color, outline=color)


class AntButton(tk.Canvas):
    """Rounded button with vector icons and consistent Ant-like states."""

    def __init__(self, parent, text: str = "", command: Callable[[], None] | None = None, kind: str = "default",
                 icon: str | None = None, width: int = 92, height: int = 36, radius: int = RADIUS_CONTROL,
                 state: str = "normal", **kwargs) -> None:
        try:
            parent_bg = parent.cget("bg")
        except Exception:
            parent_bg = COLORS["surface"]
        super().__init__(parent, width=width, height=height, bg=parent_bg, highlightthickness=0, bd=0,
                         cursor="hand2", takefocus=1, **kwargs)
        self.text = text
        self.command = command
        self.kind = kind
        self.icon = icon
        self.radius = radius
        self.state_value = state
        self.hovered = False
        self.pressed = False
        self.focused = False
        self.bind("<Configure>", self._redraw, add="+")
        self.bind("<Enter>", self._enter, add="+")
        self.bind("<Leave>", self._leave, add="+")
        self.bind("<ButtonPress-1>", self._press, add="+")
        self.bind("<ButtonRelease-1>", self._release, add="+")
        self.bind("<FocusIn>", self._focus_in, add="+")
        self.bind("<FocusOut>", self._focus_out, add="+")
        self.bind("<Return>", self._keyboard_activate, add="+")
        self.bind("<space>", self._keyboard_activate, add="+")
        self.after_idle(self._redraw)

    def configure(self, cnf=None, **kwargs):
        options = dict(cnf or {})
        options.update(kwargs)
        for key in ("state", "text", "command", "kind", "icon"):
            if key in options:
                value = options.pop(key)
                if key == "state":
                    self.state_value = str(value)
                else:
                    setattr(self, key, value)
        if options:
            super().configure(**options)
        super().configure(cursor="arrow" if self.state_value == "disabled" else "hand2")
        self._redraw()

    config = configure

    def _enter(self, _event=None) -> None:
        if self.state_value != "disabled":
            self.hovered = True
            self._redraw()

    def _leave(self, _event=None) -> None:
        self.hovered = False
        self.pressed = False
        self._redraw()

    def _press(self, _event=None) -> None:
        if self.state_value != "disabled":
            self.focus_set()
            self.pressed = True
            self._redraw()

    def _release(self, event=None) -> None:
        if self.state_value == "disabled":
            return
        was_pressed = self.pressed
        self.pressed = False
        self._redraw()
        if was_pressed and event is not None and 0 <= event.x <= self.winfo_width() and 0 <= event.y <= self.winfo_height() and self.command:
            self.command()

    def _keyboard_activate(self, _event=None):
        if self.state_value != "disabled" and self.command:
            self.command()
        return "break"

    def _focus_in(self, _event=None) -> None:
        self.focused = True
        self._redraw()

    def _focus_out(self, _event=None) -> None:
        self.focused = False
        self._redraw()

    def _palette(self) -> tuple[str, str, str]:
        disabled = self.state_value == "disabled"
        if self.kind == "primary":
            return (("#BAE0FF" if disabled else (COLORS["primary_active"] if self.pressed else COLORS["primary_hover"] if self.hovered else COLORS["primary"])), "", "#FFFFFF")
        if self.kind == "danger":
            return ((COLORS["surface_alt"] if disabled else "#FFCCC7" if self.hovered else COLORS["danger_soft"]), "", COLORS["disabled"] if disabled else COLORS["danger"])
        if self.kind == "danger_text":
            return ((COLORS["danger_soft"] if self.hovered and not disabled else self.cget("bg")), "", COLORS["disabled"] if disabled else COLORS["danger"])
        if self.kind == "text":
            return ((COLORS["surface_alt"] if self.hovered and not disabled else self.cget("bg")), "", COLORS["disabled"] if disabled else COLORS["primary"] if self.hovered else COLORS["muted"])
        if self.kind == "overlay":
            return (("#000000" if self.pressed else "#262626" if self.hovered else "#434343"), "", "#FFFFFF")
        return ((COLORS["primary_soft"] if self.pressed else COLORS["surface_alt"] if self.hovered else COLORS["surface"]), COLORS["primary"] if (self.hovered or self.focused) and not disabled else COLORS["border"], COLORS["disabled"] if disabled else COLORS["primary"] if self.hovered else COLORS["text"])

    def _redraw(self, _event=None) -> None:
        if not self.winfo_exists():
            return
        self.delete("all")
        width = max(24, self.winfo_width())
        height = max(24, self.winfo_height())
        fill, outline, foreground = self._palette()
        if self.kind not in ("text", "danger_text"):
            self.create_polygon(rounded_points(1, 1, width - 2, height - 2, self.radius), smooth=True, splinesteps=24,
                                fill=fill, outline=outline, width=2 if self.focused else 1)
        elif fill != self.cget("bg"):
            self.create_polygon(rounded_points(1, 1, width - 2, height - 2, self.radius), smooth=True, splinesteps=24,
                                fill=fill, outline="")
        has_text = bool(self.text)
        font_weight = "bold" if self.kind == "primary" else "normal"
        text_font = tkfont.Font(family=FONT, size=9, weight=font_weight)
        display_text = self.text
        available = width - (30 if self.icon else 16)
        while display_text and text_font.measure(display_text) > available:
            display_text = display_text[:-1]
        if display_text != self.text and len(display_text) >= 2:
            display_text = display_text[:-1] + "…"
        text_width = text_font.measure(display_text) if display_text else 0
        group_width = text_width + (22 if self.icon and display_text else 14 if self.icon else 0)
        group_left = max(8, (width - group_width) / 2)
        icon_x = group_left + 7
        if self.icon:
            draw_icon(self, self.icon, icon_x, height / 2, foreground, size=14, width=2)
        if has_text and display_text:
            text_x = group_left + (18 if self.icon else 0)
            self.create_text(text_x, height / 2, text=display_text, fill=foreground,
                             font=(FONT, 9, font_weight), anchor="w")


class AntCheck(tk.Canvas):
    """Custom checkbox with a rounded box and vector checkmark."""

    def __init__(self, parent, text: str, variable: tk.BooleanVar, width: int = 265,
                 command: Callable[[], None] | None = None) -> None:
        super().__init__(parent, width=width, height=26, bg=parent.cget("bg"), highlightthickness=0, bd=0,
                         cursor="hand2", takefocus=1)
        self.text = text
        self.variable = variable
        self.command = command
        self.hovered = False
        self._trace = variable.trace_add("write", lambda *_: self._redraw())
        self.bind("<Configure>", self._redraw, add="+")
        self.bind("<Enter>", lambda _e: self._set_hover(True), add="+")
        self.bind("<Leave>", lambda _e: self._set_hover(False), add="+")
        self.bind("<Button-1>", self._toggle, add="+")
        self.bind("<space>", self._toggle, add="+")
        self.bind("<Destroy>", self._destroy, add="+")
        self.after_idle(self._redraw)

    def _destroy(self, _event=None) -> None:
        try:
            self.variable.trace_remove("write", self._trace)
        except Exception:
            pass

    def _set_hover(self, value: bool) -> None:
        self.hovered = value
        self._redraw()

    def _toggle(self, _event=None):
        self.focus_set()
        self.variable.set(not self.variable.get())
        if self.command:
            self.command()
        return "break"

    def _redraw(self, _event=None) -> None:
        self.delete("all")
        checked = bool(self.variable.get())
        outline = COLORS["primary"] if checked or self.hovered else COLORS["border"]
        fill = COLORS["primary"] if checked else COLORS["surface"]
        self.create_polygon(rounded_points(2, 5, 18, 21, 4), smooth=True, splinesteps=20, fill=fill,
                            outline=outline, width=1)
        if checked:
            draw_icon(self, "check", 10, 13, "#FFFFFF", size=9, width=2)
        self.create_text(27, 13, anchor="w", text=self.text, fill=COLORS["text"], font=(FONT, 9))


class AntProgress(tk.Canvas):
    """Rounded progress bar with a subtle active shimmer and percentage pill."""

    def __init__(self, parent, width: int = 300, height: int = 28) -> None:
        super().__init__(parent, width=width, height=height, bg=parent.cget("bg"), highlightthickness=0, bd=0)
        self.value = 0.0
        self.running = False
        self.phase = 0
        self._after_id: str | None = None
        self.bind("<Configure>", self._redraw, add="+")
        self.bind("<Destroy>", self._destroy, add="+")

    def _destroy(self, _event=None) -> None:
        if self._after_id:
            try:
                self.after_cancel(self._after_id)
            except Exception:
                pass
            self._after_id = None

    def configure(self, cnf=None, **kwargs):
        options = dict(cnf or {})
        options.update(kwargs)
        if "value" in options:
            try:
                self.value = max(0.0, min(100.0, float(options.pop("value"))))
            except Exception:
                self.value = 0.0
        if "running" in options:
            self.running = bool(options.pop("running"))
            if self.running:
                self._ensure_animation()
            elif self._after_id:
                try:
                    self.after_cancel(self._after_id)
                except Exception:
                    pass
                self._after_id = None
        if options:
            super().configure(**options)
        self._redraw()

    config = configure

    def _ensure_animation(self) -> None:
        if not self._after_id and self.winfo_exists():
            self._after_id = self.after(55, self._animate)

    def _animate(self) -> None:
        self._after_id = None
        if not self.running or not self.winfo_exists():
            return
        self.phase = (self.phase + 7) % 100
        self._redraw()
        self._ensure_animation()

    def _redraw(self, _event=None) -> None:
        if not self.winfo_exists():
            return
        self.delete("all")
        width = max(100, self.winfo_width())
        height = max(24, self.winfo_height())
        cy = height / 2
        pill_w = 42
        track_left, track_right = 2, width - pill_w - 8
        top, bottom = cy - 4, cy + 4
        self.create_polygon(rounded_points(track_left, top, track_right, bottom, 4), smooth=True, splinesteps=20,
                            fill="#EEF3FB", outline="")
        fill_width = (track_right - track_left) * self.value / 100
        if fill_width > 1:
            right = min(track_right, track_left + fill_width)
            color = COLORS["success"] if self.value >= 100 else COLORS["primary"]
            self.create_polygon(rounded_points(track_left, top, max(track_left + 8, right), bottom, 4), smooth=True,
                                splinesteps=20, fill=color, outline="")
            if self.running and right - track_left > 28:
                shimmer_width = 28
                usable = max(1, right - track_left + shimmer_width)
                shimmer_x = track_left - shimmer_width + usable * self.phase / 100
                shimmer_right = min(right, shimmer_x + shimmer_width)
                shimmer_left = max(track_left, shimmer_x)
                if shimmer_right > shimmer_left:
                    self.create_rectangle(shimmer_left, top + 1, shimmer_right, bottom - 1, fill="#69B1FF", outline="")
        pill_fill = COLORS["success_soft"] if self.value >= 100 else COLORS["primary_soft"]
        pill_fg = "#389E0D" if self.value >= 100 else COLORS["primary_active"]
        self.create_polygon(rounded_points(width - pill_w, cy - 10, width - 1, cy + 10, 10), smooth=True,
                            splinesteps=24, fill=pill_fill, outline="")
        self.create_text(width - pill_w / 2 - .5, cy, text=f"{int(round(self.value))}%", fill=pill_fg,
                         font=("Segoe UI", 8, "bold"))


class AntEntry(tk.Canvas):
    """Rounded text field used for output paths."""

    def __init__(self, parent, textvariable: tk.StringVar, width: int = 180) -> None:
        super().__init__(parent, width=width, height=36, bg=parent.cget("bg"), highlightthickness=0, bd=0,
                         cursor="xterm")
        self.variable = textvariable
        self.hovered = False
        self.focused = False
        self.entry = tk.Entry(self, textvariable=textvariable, bd=0, relief="flat", bg=COLORS["surface"],
                              fg=COLORS["text"], insertbackground=COLORS["primary"], font=(FONT, 9),
                              highlightthickness=0)
        self.window_id = self.create_window((11, 19), window=self.entry, anchor="w")
        self.bind("<Configure>", self._redraw, add="+")
        self.bind("<Enter>", lambda _e: self._hover(True), add="+")
        self.bind("<Leave>", lambda _e: self._hover(False), add="+")
        self.bind("<Button-1>", lambda _e: self.entry.focus_set(), add="+")
        self.entry.bind("<FocusIn>", lambda _e: self._focus(True), add="+")
        self.entry.bind("<FocusOut>", lambda _e: self._focus(False), add="+")
        self.after_idle(self._redraw)

    def _hover(self, value: bool) -> None:
        self.hovered = value
        self._redraw()

    def _focus(self, value: bool) -> None:
        self.focused = value
        self._redraw()

    def _redraw(self, _event=None) -> None:
        self.delete("field")
        width = max(80, self.winfo_width())
        height = max(36, self.winfo_height())
        outline = COLORS["primary"] if self.focused else COLORS["primary_border"] if self.hovered else COLORS["border"]
        self.create_polygon(rounded_points(1, 1, width - 2, height - 2, RADIUS_CONTROL), smooth=True,
                            splinesteps=24, fill=COLORS["surface"], outline=outline,
                            width=2 if self.focused else 1, tags="field")
        self.tag_lower("field")
        self.coords(self.window_id, 11, height / 2)
        self.itemconfigure(self.window_id, width=max(30, width - 22), height=max(22, height - 12))


class AntSpinbox(tk.Canvas):
    """Compact rounded number input with integrated steppers."""

    def __init__(self, parent, textvariable: tk.IntVar, from_: int, to: int, increment: int = 1,
                 width: int = 115) -> None:
        super().__init__(parent, width=width, height=36, bg=parent.cget("bg"), highlightthickness=0, bd=0,
                         cursor="xterm")
        self.variable = textvariable
        self.minimum = from_
        self.maximum = to
        self.increment = increment
        self.hovered = False
        self.focused = False
        self.entry = tk.Entry(self, textvariable=textvariable, bd=0, relief="flat", bg=COLORS["surface"],
                              fg=COLORS["text"], insertbackground=COLORS["primary"], font=("Segoe UI", 9),
                              highlightthickness=0, justify="left")
        self.window_id = self.create_window((11, 19), window=self.entry, anchor="w")
        self.bind("<Configure>", self._redraw, add="+")
        self.bind("<Enter>", lambda _e: self._hover(True), add="+")
        self.bind("<Leave>", lambda _e: self._hover(False), add="+")
        self.bind("<Button-1>", self._click, add="+")
        self.entry.bind("<FocusIn>", lambda _e: self._focus(True), add="+")
        self.entry.bind("<FocusOut>", self._normalize, add="+")
        self.entry.bind("<Up>", lambda _e: self._step(1), add="+")
        self.entry.bind("<Down>", lambda _e: self._step(-1), add="+")
        self.after_idle(self._redraw)

    def _hover(self, value: bool) -> None:
        self.hovered = value
        self._redraw()

    def _focus(self, value: bool) -> None:
        self.focused = value
        self._redraw()

    def _normalize(self, _event=None) -> None:
        self.focused = False
        try:
            value = int(self.variable.get())
        except Exception:
            value = self.minimum
        self.variable.set(max(self.minimum, min(self.maximum, value)))
        self._redraw()

    def _step(self, direction: int):
        try:
            current = int(self.variable.get())
        except Exception:
            current = self.minimum
        self.variable.set(max(self.minimum, min(self.maximum, current + direction * self.increment)))
        self._redraw()
        return "break"

    def _click(self, event) -> None:
        if event.x >= self.winfo_width() - 28:
            self._step(1 if event.y < self.winfo_height() / 2 else -1)
        else:
            self.entry.focus_set()

    def _redraw(self, _event=None) -> None:
        self.delete("field")
        width = max(82, self.winfo_width())
        height = max(36, self.winfo_height())
        outline = COLORS["primary"] if self.focused else COLORS["primary_border"] if self.hovered else COLORS["border"]
        self.create_polygon(rounded_points(1, 1, width - 2, height - 2, RADIUS_CONTROL), smooth=True,
                            splinesteps=24, fill=COLORS["surface"], outline=outline,
                            width=2 if self.focused else 1, tags="field")
        divider_x = width - 28
        self.create_line(divider_x, 5, divider_x, height - 5, fill=COLORS["border_soft"], tags="field")
        arrow_color = COLORS["primary"] if self.hovered else COLORS["subtle"]
        self.create_line(width - 18, height / 2 - 5, width - 14, height / 2 - 9, width - 10,
                         height / 2 - 5, fill=arrow_color, width=1.5, joinstyle="round", tags="field")
        self.create_line(width - 18, height / 2 + 5, width - 14, height / 2 + 9, width - 10,
                         height / 2 + 5, fill=arrow_color, width=1.5, joinstyle="round", tags="field")
        self.tag_lower("field")
        self.coords(self.window_id, 11, height / 2)
        self.itemconfigure(self.window_id, width=max(25, width - 48), height=max(22, height - 12))


class StatusTag(tk.Canvas):
    """Rounded status pill used in the header."""

    def __init__(self, parent, text: str, command: Callable[[], None] | None = None, width: int = 166) -> None:
        super().__init__(parent, width=width, height=34, bg=parent.cget("bg"), highlightthickness=0, bd=0,
                         cursor="hand2")
        self.text = text
        self.state_name = "neutral"
        self.command = command
        self.bind("<Configure>", self._redraw, add="+")
        self.bind("<Button-1>", lambda _e: self.command() if self.command else None)
        self.after_idle(self._redraw)

    def set(self, text: str, state: str = "neutral") -> None:
        self.text = text
        self.state_name = state
        self._redraw()

    def _redraw(self, _event=None) -> None:
        self.delete("all")
        palettes = {
            "success": (COLORS["success_soft"], "#B7EB8F", COLORS["success"], "#389E0D"),
            "warning": (COLORS["warning_soft"], "#FFE58F", COLORS["warning"], "#D48806"),
            "neutral": (COLORS["surface_alt"], COLORS["border_soft"], COLORS["subtle"], COLORS["muted"]),
        }
        fill, outline, dot, fg = palettes.get(self.state_name, palettes["neutral"])
        width = max(90, self.winfo_width())
        height = max(30, self.winfo_height())
        self.create_polygon(rounded_points(1, 1, width - 2, height - 2, RADIUS_CONTROL), smooth=True,
                            splinesteps=24, fill=fill, outline=outline, width=1)
        self.create_oval(12, height / 2 - 3, 18, height / 2 + 3, fill=dot, outline="")
        self.create_text(25, height / 2, anchor="w", text=self.text, fill=fg, font=(FONT, 8))


class SidebarItem(tk.Canvas):
    """Rounded sidebar menu item with a vector icon."""

    def __init__(self, parent, icon: str, title: str, subtitle: str, command: Callable[[], None]) -> None:
        super().__init__(parent, height=58, bg=COLORS["sidebar"], highlightthickness=0, bd=0, cursor="hand2")
        self.icon_name = icon
        self.title_text = title
        self.subtitle_text = subtitle
        self.command = command
        self.selected = False
        self.hovered = False
        self.bind("<Configure>", self._redraw, add="+")
        self.bind("<Enter>", lambda _e: self._hover(True), add="+")
        self.bind("<Leave>", lambda _e: self._hover(False), add="+")
        self.bind("<Button-1>", lambda _e: self.command())
        self.after_idle(self._redraw)

    def _hover(self, value: bool) -> None:
        self.hovered = value
        self._redraw()

    def set_selected(self, selected: bool) -> None:
        self.selected = selected
        self._redraw()

    def _redraw(self, _event=None) -> None:
        self.delete("all")
        width = max(120, self.winfo_width())
        fill = COLORS["sidebar_active"] if self.selected else COLORS["sidebar_hover"] if self.hovered else COLORS["sidebar"]
        if fill != COLORS["sidebar"]:
            self.create_polygon(rounded_points(2, 2, width - 2, 56, RADIUS_CONTROL), smooth=True,
                                splinesteps=24, fill=fill, outline="")
        icon_color = COLORS["primary"] if self.selected else COLORS["muted"]
        draw_icon(self, self.icon_name, 25, 28, icon_color, size=17, width=2)
        self.create_text(48, 20, anchor="w", text=self.title_text,
                         fill=COLORS["primary_active"] if self.selected else COLORS["text"], font=(FONT, 9, "bold"))
        self.create_text(48, 39, anchor="w", text=self.subtitle_text,
                         fill=COLORS["primary"] if self.selected else COLORS["subtle"], font=(FONT, 7))


class Card(tk.Canvas):
    """Rounded Ant Card surface with a subtle shadow and inner content frame."""

    def __init__(self, parent, radius: int = RADIUS_CARD, shadow: bool = True, **kwargs) -> None:
        self._explicit_width = "width" in kwargs
        self._explicit_height = "height" in kwargs
        try:
            outer_bg = parent.cget("bg")
        except Exception:
            outer_bg = COLORS["app"]
        super().__init__(
            parent,
            bg=outer_bg,
            highlightthickness=0,
            bd=0,
            relief="flat",
            **kwargs,
        )
        self.radius = radius
        self.shadow = shadow
        self.inset = 6
        self.shadow_offset = 3 if shadow else 0
        self.content = tk.Frame(self, bg=COLORS["surface"], bd=0, highlightthickness=0)
        self._window_id = self.create_window((self.inset, self.inset), window=self.content, anchor="nw")
        self.bind("<Configure>", self._redraw, add="+")
        self.content.bind("<Configure>", self._sync_requested_size, add="+")
        self.after_idle(self._redraw)

    def _draw_round_rect(self, x1: float, y1: float, x2: float, y2: float, radius: int, **kwargs) -> int:
        return self.create_polygon(
            rounded_points(x1, y1, x2, y2, radius),
            smooth=True,
            splinesteps=32,
            **kwargs,
        )

    def _redraw(self, _event=None) -> None:
        width = max(24, self.winfo_width())
        height = max(24, self.winfo_height())
        self.delete("card-bg")
        bottom = height - self.shadow_offset - 1
        if self.shadow:
            self._draw_round_rect(
                2,
                3,
                width - 2,
                height - 1,
                self.radius,
                fill="#E8E8E8",
                outline="",
                tags="card-bg",
            )
        self._draw_round_rect(
            1,
            1,
            width - 2,
            bottom,
            self.radius,
            fill=COLORS["surface"],
            outline=COLORS["border_soft"],
            width=1,
            tags="card-bg",
        )
        content_width = max(1, width - self.inset * 2)
        content_height = max(1, bottom - self.inset * 2 + 1)
        self.coords(self._window_id, self.inset, self.inset)
        self.itemconfigure(self._window_id, width=content_width, height=content_height)
        self.tag_lower("card-bg")

    def _sync_requested_size(self, _event=None) -> None:
        if not self._explicit_height:
            requested = self.content.winfo_reqheight() + self.inset * 2 + self.shadow_offset + 1
            if requested > 1 and abs(int(float(self.cget("height"))) - requested) > 1:
                self.configure(height=requested)
        if not self._explicit_width:
            requested = self.content.winfo_reqwidth() + self.inset * 2 + 1
            current = int(float(self.cget("width")))
            if requested > current and current <= 2:
                self.configure(width=requested)


class Tooltip:
    def __init__(self, widget: tk.Widget, text: str) -> None:
        self.widget = widget
        self.text = text
        self.window: tk.Toplevel | None = None
        widget.bind("<Enter>", self._show, add="+")
        widget.bind("<Leave>", self._hide, add="+")
        widget.bind("<ButtonPress>", self._hide, add="+")

    def _show(self, _event=None) -> None:
        if self.window or not self.text:
            return
        x = self.widget.winfo_rootx() + 8
        y = self.widget.winfo_rooty() + self.widget.winfo_height() + 7
        self.window = tk.Toplevel(self.widget)
        self.window.overrideredirect(True)
        self.window.attributes("-topmost", True)
        tk.Label(
            self.window,
            text=self.text,
            bg="#262626",
            fg="#FFFFFF",
            padx=9,
            pady=6,
            font=(FONT, 8),
        ).pack()
        self.window.geometry(f"+{x}+{y}")

    def _hide(self, _event=None) -> None:
        if self.window:
            self.window.destroy()
            self.window = None


class StepBar(tk.Canvas):
    """Three-step indicator with labels separated from the connector line."""

    def __init__(self, parent, steps: tuple[str, ...]) -> None:
        super().__init__(parent, height=48, bg=COLORS["surface"], highlightthickness=0, bd=0)
        self.steps = steps
        self.current = 0
        self.finished = False
        self.error = False
        self.bind("<Configure>", self._redraw, add="+")
        self.after_idle(self._redraw)

    def set_current(self, current: int, finished: bool = False, error: bool = False) -> None:
        self.current = max(0, min(len(self.steps) - 1, current))
        self.finished = finished
        self.error = error
        self._redraw()

    def _redraw(self, _event=None) -> None:
        self.delete("all")
        if not self.steps:
            return
        width = max(330, self.winfo_width())
        node_y, label_y = 15, 37
        left, right = 42, width - 42
        positions = [left + (right - left) * i / max(1, len(self.steps) - 1) for i in range(len(self.steps))]
        for index in range(len(self.steps) - 1):
            color = COLORS["primary"] if index < self.current else COLORS["border_soft"]
            self.create_line(positions[index] + 12, node_y, positions[index + 1] - 12, node_y,
                             fill=color, width=2)
        for index, title in enumerate(self.steps):
            done = index < self.current or (self.finished and index == self.current)
            active = index == self.current and not self.finished
            node_color = COLORS["danger"] if self.error and active else COLORS["primary"]
            fill = node_color if (done or active) else COLORS["surface"]
            outline = node_color if (done or active) else COLORS["border"]
            self.create_oval(positions[index] - 10, node_y - 10, positions[index] + 10, node_y + 10,
                             fill=fill, outline=outline, width=1)
            if done:
                draw_icon(self, "check", positions[index], node_y, "#FFFFFF", size=9, width=2)
            else:
                self.create_text(positions[index], node_y, text=str(index + 1),
                                 fill="#FFFFFF" if active else COLORS["subtle"],
                                 font=("Segoe UI", 8, "bold"))
            label_color = node_color if active else COLORS["text"] if done else COLORS["subtle"]
            self.create_text(positions[index], label_y, text=title, fill=label_color,
                             font=(FONT, 8, "bold" if active or done else "normal"), anchor="center")


class PresetChoice(tk.Canvas):
    """Rounded selectable preset card inspired by Ant radio-card patterns."""

    def __init__(self, parent, title: str, subtitle: str, command: Callable[[], None]) -> None:
        super().__init__(
            parent,
            height=48,
            bg=COLORS["surface"],
            highlightthickness=0,
            bd=0,
            cursor="hand2",
        )
        self.title = title
        self.subtitle = subtitle
        self.command = command
        self.selected = False
        self.hovered = False
        self.bind("<Configure>", self._redraw, add="+")
        self.bind("<Button-1>", lambda _event: self.command())
        self.bind("<Enter>", self._hover, add="+")
        self.bind("<Leave>", self._leave, add="+")

    def _hover(self, _event=None) -> None:
        self.hovered = True
        self._redraw()

    def _leave(self, _event=None) -> None:
        self.hovered = False
        self._redraw()

    def set_selected(self, selected: bool) -> None:
        self.selected = selected
        self._redraw()

    def _redraw(self, _event=None) -> None:
        self.delete("all")
        width = max(60, self.winfo_width())
        height = max(44, self.winfo_height())
        fill = COLORS["primary_soft"] if self.selected else COLORS["surface"]
        outline = COLORS["primary"] if self.selected else (COLORS["primary_border"] if self.hovered else COLORS["border"])
        self.create_polygon(
            rounded_points(1, 1, width - 2, height - 2, RADIUS_CONTROL),
            smooth=True,
            splinesteps=28,
            fill=fill,
            outline=outline,
            width=1,
        )
        title_color = COLORS["primary_active"] if self.selected else COLORS["text"]
        self.create_text(11, 14, anchor="w", text=self.title, fill=title_color, font=(FONT, 9, "bold"))
        self.create_text(11, 32, anchor="w", text=self.subtitle, fill=COLORS["subtle"], font=(FONT, 7))
        if self.selected:
            self.create_oval(width - 25, 9, width - 11, 23, fill=COLORS["primary"], outline="")
            draw_icon(self, "check", width - 18, 16, "#FFFFFF", size=8, width=2)


class DropZone(tk.Canvas):
    """Upload.Dragger-inspired drop target."""

    def __init__(self, parent, title: str, subtitle: str, accept_text: str, command: Callable[[], None]) -> None:
        super().__init__(
            parent,
            height=80,
            bg=COLORS["surface"],
            highlightthickness=0,
            bd=0,
            cursor="hand2",
        )
        self.title = title
        self.subtitle = subtitle
        self.accept_text = accept_text
        self.command = command
        self.hovered = False
        self.bind("<Configure>", self._redraw)
        self.bind("<Button-1>", lambda _event: self.command())
        self.bind("<Enter>", self._enter)
        self.bind("<Leave>", self._leave)

    def _enter(self, _event=None) -> None:
        self.hovered = True
        self._redraw()

    def _leave(self, _event=None) -> None:
        self.hovered = False
        self._redraw()

    def pulse(self) -> None:
        self.hovered = True
        self._redraw()
        self.after(520, self._leave)

    def _redraw(self, _event=None) -> None:
        self.delete("all")
        width = max(10, self.winfo_width())
        height = max(10, self.winfo_height())
        fill = COLORS["primary_soft"] if self.hovered else COLORS["surface"]
        outline = COLORS["primary"] if self.hovered else COLORS["border"]
        shape = rounded_points(2, 2, width - 3, height - 3, RADIUS_CARD)
        self.create_polygon(shape, smooth=True, splinesteps=28, fill=fill, outline="")
        self.create_line(
            shape + shape[:2],
            smooth=True,
            splinesteps=28,
            fill=outline,
            width=2,
            dash=(6, 4),
            joinstyle="round",
        )

        cx = 43
        cy = height / 2
        self.create_oval(cx - 20, cy - 20, cx + 20, cy + 20, fill="#FFFFFF", outline=COLORS["primary_border"], width=1)
        draw_icon(self, "upload", cx, cy, COLORS["primary"], size=18, width=2)

        text_x = 77
        self.create_text(text_x, cy - 15, anchor="w", text=self.title, fill=COLORS["text"], font=(FONT, 10, "bold"))
        if width >= 660:
            self.create_text(text_x, cy + 6, anchor="w", text=self.subtitle, fill=COLORS["muted"], font=(FONT, 8))
            self.create_text(text_x, cy + 24, anchor="w", text=self.accept_text, fill=COLORS["subtle"], font=(FONT, 7))
        else:
            self.create_text(text_x, cy + 10, anchor="w", text=self.accept_text, fill=COLORS["subtle"], font=(FONT, 7))

        button_w = 92
        button_left = width - button_w - 18
        self.create_polygon(rounded_points(button_left, cy - 17, width - 18, cy + 17, RADIUS_CONTROL), smooth=True,
                            splinesteps=24, fill="#FFFFFF", outline=COLORS["primary_border"] if not self.hovered else COLORS["primary"])
        self.create_text(button_left + button_w / 2, cy, text="选择文件", fill=COLORS["primary"], font=(FONT, 8, "bold"))


class AntSelect(tk.Canvas):
    """Rounded, keyboard-friendly Select component with an Ant-style popup menu."""

    _opened: "AntSelect | None" = None

    def __init__(self, parent, textvariable: tk.StringVar, values: tuple[str, ...], width: int = 118) -> None:
        super().__init__(
            parent,
            height=38,
            width=width,
            bg=COLORS["surface"],
            highlightthickness=0,
            bd=0,
            cursor="hand2",
            takefocus=1,
        )
        self.variable = textvariable
        self.values = tuple(values)
        self.hovered = False
        self.focused = False
        self.popup: tk.Toplevel | None = None
        self._trace_id = self.variable.trace_add("write", lambda *_args: self._redraw())
        self.bind("<Configure>", self._redraw, add="+")
        self.bind("<Enter>", self._on_enter, add="+")
        self.bind("<Leave>", self._on_leave, add="+")
        self.bind("<FocusIn>", self._on_focus_in, add="+")
        self.bind("<FocusOut>", self._on_focus_out, add="+")
        self.bind("<Button-1>", self._toggle_popup)
        self.bind("<Return>", self._toggle_popup)
        self.bind("<space>", self._toggle_popup)
        self.bind("<Down>", self._open_from_keyboard)
        self.bind("<Escape>", lambda _event: self.close_popup())
        self.bind("<Destroy>", self._on_destroy, add="+")
        self.after_idle(self._redraw)

    def _on_destroy(self, _event=None) -> None:
        try:
            self.variable.trace_remove("write", self._trace_id)
        except Exception:
            pass
        self.close_popup()

    def _on_enter(self, _event=None) -> None:
        self.hovered = True
        self._redraw()

    def _on_leave(self, _event=None) -> None:
        self.hovered = False
        self._redraw()

    def _on_focus_in(self, _event=None) -> None:
        self.focused = True
        self._redraw()

    def _on_focus_out(self, _event=None) -> None:
        self.focused = False
        self._redraw()

    def _open_from_keyboard(self, _event=None):
        if not self.popup:
            self.open_popup()
        return "break"

    def _toggle_popup(self, _event=None):
        self.focus_set()
        if self.popup and self.popup.winfo_exists():
            self.close_popup()
        else:
            self.open_popup()
        return "break"

    def _redraw(self, _event=None) -> None:
        if not self.winfo_exists():
            return
        self.delete("all")
        width = max(80, self.winfo_width())
        height = max(36, self.winfo_height())
        opened = bool(self.popup and self.popup.winfo_exists())
        outline = COLORS["primary"] if (self.focused or opened) else (COLORS["primary_border"] if self.hovered else COLORS["border"])
        fill = COLORS["primary_soft"] if opened else COLORS["surface"]
        self.create_polygon(
            rounded_points(1, 1, width - 2, height - 2, 8),
            smooth=True,
            splinesteps=24,
            fill=fill,
            outline=outline,
            width=2 if (self.focused or opened) else 1,
        )
        value = self.variable.get() or (self.values[0] if self.values else "")
        self.create_text(12, height / 2, anchor="w", text=value, fill=COLORS["text"], font=(FONT, 9))
        cx = width - 17
        cy = height / 2
        arrow = COLORS["primary"] if (opened or self.hovered) else COLORS["subtle"]
        if opened:
            self.create_line(cx - 4, cy + 2, cx, cy - 2, cx + 4, cy + 2, fill=arrow, width=2, joinstyle="round")
        else:
            self.create_line(cx - 4, cy - 2, cx, cy + 2, cx + 4, cy - 2, fill=arrow, width=2, joinstyle="round")

    def open_popup(self) -> None:
        if AntSelect._opened and AntSelect._opened is not self:
            AntSelect._opened.close_popup()
        if not self.values:
            return
        self.update_idletasks()
        popup = tk.Toplevel(self)
        popup.overrideredirect(True)
        popup.transient(self.winfo_toplevel())
        popup.configure(bg="#D9D9D9")
        try:
            popup.attributes("-topmost", True)
            popup.attributes("-alpha", 0.99)
        except Exception:
            pass
        self.popup = popup
        AntSelect._opened = self

        width = max(self.winfo_width(), 150)
        item_height = 34
        popup_height = item_height * len(self.values) + 8
        x = self.winfo_rootx()
        y = self.winfo_rooty() + self.winfo_height() + 5
        if y + popup_height > self.winfo_screenheight() - 12:
            y = self.winfo_rooty() - popup_height - 5
        popup.geometry(f"{width}x{popup_height}+{x}+{y}")

        shell = tk.Frame(popup, bg="#FFFFFF", bd=0, highlightbackground=COLORS["border_soft"], highlightthickness=1)
        shell.pack(fill="both", expand=True)
        current = self.variable.get()
        for index, value in enumerate(self.values):
            selected = value == current
            bg = COLORS["primary_soft"] if selected else "#FFFFFF"
            fg = COLORS["primary_active"] if selected else COLORS["text"]
            row = tk.Frame(shell, bg=bg, cursor="hand2", height=item_height)
            row.pack(fill="x", padx=4, pady=(4 if index == 0 else 0, 0))
            row.pack_propagate(False)
            label = tk.Label(
                row,
                text=value,
                anchor="w",
                bg=bg,
                fg=fg,
                padx=10,
                font=(FONT, 9, "bold" if selected else "normal"),
                cursor="hand2",
            )
            label.pack(side="left", fill="both", expand=True)
            check = None
            if selected:
                check = tk.Label(row, text="✓", bg=bg, fg=COLORS["primary"], font=("Segoe UI", 9, "bold"))
                check.pack(side="right", padx=(0, 10))

            def enter(_event, frame=row, text_label=label, check_label=check, is_selected=selected):
                hover_bg = COLORS["primary_soft"] if is_selected else COLORS["surface_alt"]
                frame.configure(bg=hover_bg)
                text_label.configure(bg=hover_bg)
                if check_label:
                    check_label.configure(bg=hover_bg)

            def leave(_event, frame=row, text_label=label, check_label=check, is_selected=selected):
                normal_bg = COLORS["primary_soft"] if is_selected else "#FFFFFF"
                frame.configure(bg=normal_bg)
                text_label.configure(bg=normal_bg)
                if check_label:
                    check_label.configure(bg=normal_bg)

            def choose(_event, selected_value=value):
                self.variable.set(selected_value)
                self.event_generate("<<ComboboxSelected>>")
                self.close_popup()
                self.focus_set()

            widgets = (row, label) if check is None else (row, label, check)
            for widget in widgets:
                widget.bind("<Enter>", enter, add="+")
                widget.bind("<Leave>", leave, add="+")
                widget.bind("<Button-1>", choose)

        popup.bind("<Escape>", lambda _event: self.close_popup())
        popup.bind("<FocusOut>", self._popup_focus_out, add="+")
        popup.after(40, popup.focus_force)
        self._redraw()

    def _popup_focus_out(self, _event=None) -> None:
        if self.popup:
            self.popup.after(70, self._close_if_focus_left)

    def _close_if_focus_left(self) -> None:
        if not self.popup or not self.popup.winfo_exists():
            return
        focused = self.popup.focus_get()
        if focused is None or focused.winfo_toplevel() is not self.popup:
            self.close_popup()

    def close_popup(self) -> None:
        if self.popup and self.popup.winfo_exists():
            try:
                self.popup.destroy()
            except Exception:
                pass
        self.popup = None
        if AntSelect._opened is self:
            AntSelect._opened = None
        if self.winfo_exists():
            self._redraw()


class AutoHideScrollbar(tk.Canvas):
    """Auto-hiding rounded scrollbar with a slim Ant-style thumb."""

    def __init__(self, parent, orient: str = "vertical", command=None, **kwargs) -> None:
        self.orient = orient
        self.command = command
        try:
            bg = parent.cget("bg")
        except Exception:
            bg = COLORS["surface"]
        width = 10 if orient == "vertical" else kwargs.pop("width", 100)
        height = kwargs.pop("height", 10) if orient == "horizontal" else kwargs.pop("height", 100)
        super().__init__(parent, width=width, height=height, bg=bg, highlightthickness=0, bd=0, cursor="arrow", **kwargs)
        self.first = 0.0
        self.last = 1.0
        self._pack_options: dict[str, object] | None = None
        self._hidden = False
        self._hovered = False
        self._dragging = False
        self._drag_offset = 0.0
        self._thumb_start = 0.0
        self._thumb_end = 0.0
        self.bind("<Configure>", self._redraw, add="+")
        self.bind("<Enter>", lambda _e: self._set_hover(True), add="+")
        self.bind("<Leave>", lambda _e: self._set_hover(False), add="+")
        self.bind("<ButtonPress-1>", self._press, add="+")
        self.bind("<B1-Motion>", self._drag, add="+")
        self.bind("<ButtonRelease-1>", self._release, add="+")

    def pack(self, cnf=None, **kwargs):
        options = {}
        if cnf:
            options.update(cnf)
        options.update(kwargs)
        if options:
            self._pack_options = options.copy()
        self._hidden = False
        return super().pack(**options)

    def set(self, first, last) -> None:
        self.first, self.last = float(first), float(last)
        full = self.first <= 0.0 and self.last >= 1.0
        if full and not self._hidden and self.winfo_manager() == "pack":
            super().pack_forget()
            self._hidden = True
        elif not full and self._hidden and self._pack_options:
            super().pack(**self._pack_options)
            self._hidden = False
        self._redraw()

    def _set_hover(self, value: bool) -> None:
        self._hovered = value
        self._redraw()

    def _geometry(self) -> tuple[float, float, float]:
        length = self.winfo_height() if self.orient == "vertical" else self.winfo_width()
        track = max(1.0, length - 8)
        thumb_len = max(26.0, track * max(0.0, self.last - self.first))
        thumb_len = min(track, thumb_len)
        max_start = max(0.0, track - thumb_len)
        start = 4 + max_start * (self.first / max(1e-9, 1.0 - (self.last - self.first)))
        end = start + thumb_len
        self._thumb_start, self._thumb_end = start, end
        return start, end, track

    def _redraw(self, _event=None) -> None:
        if not self.winfo_exists() or self._hidden:
            return
        self.delete("all")
        start, end, _track = self._geometry()
        color = COLORS["subtle"] if self._dragging else COLORS["disabled"] if self._hovered else "#D9D9D9"
        if self.orient == "vertical":
            cx = self.winfo_width() / 2
            self.create_polygon(rounded_points(cx - 3, start, cx + 3, end, 3), smooth=True, splinesteps=20,
                                fill=color, outline="")
        else:
            cy = self.winfo_height() / 2
            self.create_polygon(rounded_points(start, cy - 3, end, cy + 3, 3), smooth=True, splinesteps=20,
                                fill=color, outline="")

    def _axis_value(self, event) -> float:
        return float(event.y if self.orient == "vertical" else event.x)

    def _press(self, event) -> None:
        pos = self._axis_value(event)
        if self._thumb_start <= pos <= self._thumb_end:
            self._dragging = True
            self._drag_offset = pos - self._thumb_start
            self.configure(cursor="sb_v_double_arrow" if self.orient == "vertical" else "sb_h_double_arrow")
        elif self.command:
            direction = -1 if pos < self._thumb_start else 1
            self.command("scroll", direction, "pages")
        self._redraw()

    def _drag(self, event) -> None:
        if not self._dragging or not self.command:
            return
        length = self.winfo_height() if self.orient == "vertical" else self.winfo_width()
        track = max(1.0, length - 8)
        thumb_len = max(1.0, self._thumb_end - self._thumb_start)
        max_start = max(1.0, track - thumb_len)
        position = max(0.0, min(max_start, self._axis_value(event) - 4 - self._drag_offset))
        self.command("moveto", position / max_start)

    def _release(self, _event=None) -> None:
        self._dragging = False
        self.configure(cursor="arrow")
        self._redraw()


class ScrollablePanel(tk.Frame):
    """Scrollable settings body with a sticky action area outside it."""

    def __init__(self, parent) -> None:
        super().__init__(parent, bg=COLORS["surface"])
        self.canvas = tk.Canvas(self, bg=COLORS["surface"], highlightthickness=0, bd=0)
        self.scrollbar = AutoHideScrollbar(self, orient="vertical", command=self.canvas.yview)
        self.canvas.configure(yscrollcommand=self.scrollbar.set)
        self.canvas.pack(side="left", fill="both", expand=True)
        self.scrollbar.pack(side="right", fill="y", padx=(4, 2), pady=6, before=self.canvas)

        self.viewport = tk.Frame(self.canvas, bg=COLORS["surface"])
        self.content = tk.Frame(self.viewport, bg=COLORS["surface"])
        self.content.pack(fill="both", expand=True, padx=14, pady=(9, 7))
        self.window_id = self.canvas.create_window((0, 0), window=self.viewport, anchor="nw")
        self.viewport.bind("<Configure>", self._update_scrollregion)
        self.canvas.bind("<Configure>", self._resize_content)
        self.canvas.bind("<Enter>", self._bind_wheel)
        self.canvas.bind("<Leave>", self._unbind_wheel)

    def _update_scrollregion(self, _event=None) -> None:
        self.canvas.configure(scrollregion=self.canvas.bbox("all"))

    def _resize_content(self, event) -> None:
        self.canvas.itemconfigure(self.window_id, width=max(1, event.width))

    def _bind_wheel(self, _event=None) -> None:
        self.canvas.bind_all("<MouseWheel>", self._on_mousewheel)

    def _unbind_wheel(self, _event=None) -> None:
        self.canvas.unbind_all("<MouseWheel>")

    def _on_mousewheel(self, event) -> None:
        delta = -1 if event.delta > 0 else 1
        self.canvas.yview_scroll(delta * 3, "units")


class ImagePreviewDialog:
    """Modal image viewer with navigation, zoom, fit and original-size controls."""

    def __init__(self, app: "PptImageConverterApp", paths: list[Path], index: int = 0) -> None:
        self.app = app
        self.paths = paths
        self.index = max(0, min(len(paths) - 1, index))
        self.image: Image.Image | None = None
        self.photo: ImageTk.PhotoImage | None = None
        self.scale = 1.0
        self.mode = "fit"
        self.offset_x = 0.0
        self.offset_y = 0.0
        self.drag_origin: tuple[int, int] | None = None

        self.window = tk.Toplevel(app.root)
        self.window.title("查看大图")
        self.window.minsize(720, 520)
        self.window.configure(bg=COLORS["surface"])
        self.window.transient(app.root)
        app.root.update_idletasks()
        root_w = max(900, app.root.winfo_width())
        root_h = max(650, app.root.winfo_height())
        width = min(1080, max(760, root_w - 90))
        height = min(760, max(540, root_h - 70))
        x = app.root.winfo_rootx() + max(0, (root_w - width) // 2)
        y = app.root.winfo_rooty() + max(0, (root_h - height) // 2)
        self.window.geometry(f"{width}x{height}+{x}+{y}")
        self.window.grab_set()
        self.window.protocol("WM_DELETE_WINDOW", self.close)

        header = tk.Frame(self.window, bg=COLORS["surface"], height=58,
                          highlightbackground=COLORS["border_soft"], highlightthickness=1)
        header.pack(fill="x")
        header.pack_propagate(False)
        title_wrap = tk.Frame(header, bg=COLORS["surface"])
        title_wrap.pack(side="left", fill="x", expand=True, padx=18, pady=10)
        self.title_var = tk.StringVar()
        self.meta_var = tk.StringVar()
        tk.Label(title_wrap, textvariable=self.title_var, bg=COLORS["surface"], fg=COLORS["text"],
                 font=(FONT, 11, "bold"), anchor="w").pack(fill="x")
        tk.Label(title_wrap, textvariable=self.meta_var, bg=COLORS["surface"], fg=COLORS["subtle"],
                 font=(FONT, 8), anchor="w").pack(fill="x", pady=(2, 0))
        AntButton(header, icon="close", kind="text", width=38, height=36, command=self.close).pack(side="right", padx=14)

        viewer = tk.Frame(self.window, bg="#141414")
        viewer.pack(fill="both", expand=True)
        self.canvas = tk.Canvas(viewer, bg="#141414", highlightthickness=0, bd=0, cursor="fleur")
        self.canvas.pack(fill="both", expand=True)
        self.prev_button = AntButton(viewer, icon="chevron_left", kind="overlay", width=42, height=48,
                                     radius=12, command=lambda: self._move(-1))
        self.next_button = AntButton(viewer, icon="chevron_right", kind="overlay", width=42, height=48,
                                     radius=12, command=lambda: self._move(1))
        self.prev_button.place(relx=0.0, rely=0.5, x=18, anchor="w")
        self.next_button.place(relx=1.0, rely=0.5, x=-18, anchor="e")

        toolbar = tk.Frame(self.window, bg=COLORS["surface"], height=58,
                           highlightbackground=COLORS["border_soft"], highlightthickness=1)
        toolbar.pack(fill="x")
        toolbar.pack_propagate(False)
        left = tk.Frame(toolbar, bg=COLORS["surface"])
        left.pack(side="left", padx=16, pady=10)
        AntButton(left, icon="chevron_left", kind="default", width=38, height=36,
                  command=lambda: self._move(-1)).pack(side="left")
        AntButton(left, icon="chevron_right", kind="default", width=38, height=36,
                  command=lambda: self._move(1)).pack(side="left", padx=(6, 0))

        center = tk.Frame(toolbar, bg=COLORS["surface"])
        center.pack(side="left", padx=(18, 0), pady=10)
        AntButton(center, icon="minus", kind="default", width=38, height=36,
                  command=lambda: self._zoom(1 / 1.2)).pack(side="left")
        self.zoom_var = tk.StringVar(value="100%")
        tk.Label(center, textvariable=self.zoom_var, bg=COLORS["surface_alt"], fg=COLORS["muted"],
                 font=("Segoe UI", 9, "bold"), width=7, padx=4, pady=8).pack(side="left", padx=6)
        AntButton(center, icon="plus", kind="default", width=38, height=36,
                  command=lambda: self._zoom(1.2)).pack(side="left")
        AntButton(center, text="适应窗口", icon="fit", kind="default", width=104, height=36,
                  command=self._fit).pack(side="left", padx=(10, 0))
        AntButton(center, text="原始尺寸", icon="actual", kind="default", width=104, height=36,
                  command=self._actual).pack(side="left", padx=(6, 0))

        AntButton(toolbar, text="打开所在目录", icon="folder", kind="default", width=122, height=36,
                  command=self._open_folder).pack(side="right", padx=16, pady=10)

        self.canvas.bind("<Configure>", lambda _e: self._render(), add="+")
        self.canvas.bind("<MouseWheel>", self._wheel, add="+")
        self.canvas.bind("<ButtonPress-1>", self._drag_start, add="+")
        self.canvas.bind("<B1-Motion>", self._drag_move, add="+")
        self.canvas.bind("<ButtonRelease-1>", self._drag_end, add="+")
        self.window.bind("<Left>", lambda _e: self._move(-1))
        self.window.bind("<Right>", lambda _e: self._move(1))
        self.window.bind("<Escape>", lambda _e: self.close())
        self.window.bind("<Control-0>", lambda _e: self._actual())
        self.window.bind("<Control-Key-0>", lambda _e: self._actual())
        self.window.after(60, self._load_current)

    def close(self) -> None:
        try:
            self.window.grab_release()
        except Exception:
            pass
        if self.window.winfo_exists():
            self.window.destroy()
        if getattr(self.app, "_preview_dialog", None) is self:
            self.app._preview_dialog = None

    def _load_current(self) -> None:
        if not self.paths:
            self.close()
            return
        path = self.paths[self.index]
        try:
            with Image.open(path) as opened:
                raw = ImageOps.exif_transpose(opened).convert("RGBA")
                white = Image.new("RGBA", raw.size, "white")
                white.alpha_composite(raw)
                self.image = white.convert("RGB")
        except Exception as exc:
            self.image = None
            self.canvas.delete("all")
            self.canvas.create_text(self.canvas.winfo_width() / 2, self.canvas.winfo_height() / 2,
                                    text=f"图片无法预览\n{exc}", fill="#FFFFFF", font=(FONT, 11), justify="center")
            return
        self.title_var.set(path.name)
        try:
            size = human_size(path.stat().st_size)
        except OSError:
            size = "—"
        self.meta_var.set(f"第 {self.index + 1} / {len(self.paths)} 张 · {self.image.width}×{self.image.height} · {size}")
        self.mode = "fit"
        self.offset_x = self.offset_y = 0.0
        self._render()
        state_prev = "normal" if self.index > 0 else "disabled"
        state_next = "normal" if self.index < len(self.paths) - 1 else "disabled"
        self.prev_button.configure(state=state_prev)
        self.next_button.configure(state=state_next)

    def _fit_scale(self) -> float:
        if not self.image:
            return 1.0
        cw = max(100, self.canvas.winfo_width() - 110)
        ch = max(100, self.canvas.winfo_height() - 80)
        return max(0.02, min(cw / self.image.width, ch / self.image.height))

    def _render(self) -> None:
        if not self.image or not self.canvas.winfo_exists():
            return
        scale = self._fit_scale() if self.mode == "fit" else 1.0 if self.mode == "actual" else self.scale
        scale = max(0.02, min(8.0, scale))
        self.scale = scale
        width = max(1, int(self.image.width * scale))
        height = max(1, int(self.image.height * scale))
        # Guard against accidental huge allocations while still allowing practical zooming.
        if width * height > 36_000_000:
            guard = (36_000_000 / (width * height)) ** 0.5
            width = max(1, int(width * guard))
            height = max(1, int(height * guard))
            self.scale *= guard
        resized = self.image.resize((width, height), Image.Resampling.LANCZOS)
        self.photo = ImageTk.PhotoImage(resized)
        self.canvas.delete("all")
        x = self.canvas.winfo_width() / 2 + self.offset_x
        y = self.canvas.winfo_height() / 2 + self.offset_y
        self.canvas.create_image(x, y, image=self.photo, anchor="center")
        self.zoom_var.set(f"{int(round(self.scale * 100))}%")

    def _zoom(self, factor: float) -> None:
        if not self.image:
            return
        current = self._fit_scale() if self.mode == "fit" else 1.0 if self.mode == "actual" else self.scale
        self.scale = max(0.05, min(8.0, current * factor))
        self.mode = "custom"
        self._render()

    def _fit(self) -> None:
        self.mode = "fit"
        self.offset_x = self.offset_y = 0.0
        self._render()

    def _actual(self) -> None:
        self.mode = "actual"
        self.offset_x = self.offset_y = 0.0
        self._render()

    def _move(self, direction: int) -> None:
        target = self.index + direction
        if 0 <= target < len(self.paths):
            self.index = target
            self._load_current()

    def _wheel(self, event) -> str:
        self._zoom(1.18 if event.delta > 0 else 1 / 1.18)
        return "break"

    def _drag_start(self, event) -> None:
        self.drag_origin = (event.x, event.y)

    def _drag_move(self, event) -> None:
        if not self.drag_origin:
            return
        dx = event.x - self.drag_origin[0]
        dy = event.y - self.drag_origin[1]
        self.offset_x += dx
        self.offset_y += dy
        self.drag_origin = (event.x, event.y)
        self._render()

    def _drag_end(self, _event=None) -> None:
        self.drag_origin = None

    def _open_folder(self) -> None:
        if self.paths:
            self.app._open_path(str(self.paths[self.index].parent))


class PptImageConverterApp:
    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title(f"{APP_NAME} v{APP_VERSION}")
        self.root.geometry("1200x800")
        self.root.minsize(1000, 680)
        self.root.configure(bg=COLORS["app"])
        self._app_icon: ImageTk.PhotoImage | None = None
        try:
            icon_path = resource_path("app_icon.png")
            if icon_path.exists():
                self._app_icon = ImageTk.PhotoImage(file=str(icon_path))
                self.root.iconphoto(True, self._app_icon)
        except Exception:
            pass

        self.cancel_event = threading.Event()
        self.busy = False
        self.config = self._load_config()
        self.active_mode: Mode = "ppt"
        self._preview_photo: ImageTk.PhotoImage | None = None
        self._image_thumbnails: dict[str, ImageTk.PhotoImage] = {}
        self._image_drag_item: str | None = None
        self._image_drag_start_y = 0
        self._hover_image_item: str | None = None
        self._hover_after_id: str | None = None
        self._hover_preview_window: tk.Toplevel | None = None
        self._hover_preview_photo: ImageTk.PhotoImage | None = None
        self._preview_dialog: ImagePreviewDialog | None = None
        self._backend_status: dict[str, object] = {}
        self._toast: tk.Frame | None = None
        self._toast_after_id: str | None = None
        self._last_task_success: dict[Mode, bool] = {"ppt": False, "images": False}

        self._configure_style()
        self._build_ui()
        self._restore_config()
        self._bind_shortcuts()
        self.root.bind("<Configure>", self._on_root_resize, add="+")
        self._drop_adapter = NativeFileDrop(self.root, self._handle_native_drop)
        self.root.after(120, self._apply_initial_mode)
        self.root.after(260, self._detect_environment_async)
        self.root.protocol("WM_DELETE_WINDOW", self._on_close)

    # ---------- Ant Design-inspired visual system ----------
    def _configure_style(self) -> None:
        style = ttk.Style(self.root)
        if "clam" in style.theme_names():
            style.theme_use("clam")

        style.configure("App.TFrame", background=COLORS["app"])
        style.configure("Surface.TFrame", background=COLORS["surface"])
        style.configure("TLabel", background=COLORS["surface"], foreground=COLORS["text"], font=(FONT, 9))
        style.configure("TCheckbutton", background=COLORS["surface"], foreground=COLORS["text"], font=(FONT, 9))
        style.map("TCheckbutton", background=[("active", COLORS["surface"])])

        style.configure(
            "Primary.TButton",
            background=COLORS["primary"],
            foreground="#FFFFFF",
            borderwidth=0,
            focusthickness=2,
            focuscolor=COLORS["primary_border"],
            padding=(16, 10),
            font=(FONT, 10, "bold"),
        )
        style.map(
            "Primary.TButton",
            background=[("pressed", COLORS["primary_active"]), ("active", COLORS["primary_hover"]), ("disabled", "#BAE0FF")],
            foreground=[("disabled", "#FFFFFF")],
        )
        style.configure(
            "Default.TButton",
            background="#FFFFFF",
            foreground=COLORS["text"],
            bordercolor=COLORS["border"],
            lightcolor=COLORS["border"],
            darkcolor=COLORS["border"],
            borderwidth=1,
            padding=(11, 8),
            font=(FONT, 9),
        )
        style.map(
            "Default.TButton",
            background=[("active", COLORS["surface_alt"]), ("disabled", COLORS["surface_alt"])],
            foreground=[("active", COLORS["primary"]), ("disabled", COLORS["disabled"])],
            bordercolor=[("active", COLORS["primary"]), ("disabled", COLORS["border_soft"])],
        )
        style.configure(
            "Text.TButton",
            background=COLORS["surface"],
            foreground=COLORS["muted"],
            borderwidth=0,
            padding=(7, 6),
            font=(FONT, 9),
        )
        style.map("Text.TButton", background=[("active", COLORS["surface_alt"])], foreground=[("active", COLORS["primary"])])
        style.configure(
            "DangerText.TButton",
            background=COLORS["surface"],
            foreground=COLORS["danger"],
            borderwidth=0,
            padding=(7, 6),
            font=(FONT, 9),
        )
        style.map("DangerText.TButton", background=[("active", COLORS["danger_soft"])])
        style.configure(
            "Danger.TButton",
            background=COLORS["danger_soft"],
            foreground=COLORS["danger"],
            borderwidth=0,
            padding=(10, 8),
            font=(FONT, 9),
        )
        style.map("Danger.TButton", background=[("active", "#FFCCC7"), ("disabled", COLORS["surface_alt"])], foreground=[("disabled", COLORS["disabled"])])

        style.configure(
            "Treeview",
            background="#FFFFFF",
            fieldbackground="#FFFFFF",
            foreground=COLORS["text"],
            rowheight=38,
            borderwidth=0,
            relief="flat",
            font=(FONT, 9),
        )
        style.map(
            "Treeview",
            background=[("selected", COLORS["primary_soft"])],
            foreground=[("selected", COLORS["text"])],
        )
        style.configure(
            "Treeview.Heading",
            background=COLORS["surface_alt"],
            foreground=COLORS["muted"],
            borderwidth=0,
            relief="flat",
            padding=(8, 9),
            font=(FONT, 8, "bold"),
        )
        style.map("Treeview.Heading", background=[("active", COLORS["surface_alt"])])
        style.configure(
            "Image.Treeview",
            background="#FFFFFF",
            fieldbackground="#FFFFFF",
            foreground=COLORS["text"],
            rowheight=66,
            borderwidth=0,
            relief="flat",
            font=(FONT, 9),
        )
        style.map("Image.Treeview", background=[("selected", COLORS["primary_soft"])], foreground=[("selected", COLORS["text"])])
        style.configure(
            "Image.Treeview.Heading",
            background=COLORS["surface_alt"],
            foreground=COLORS["muted"],
            borderwidth=0,
            relief="flat",
            padding=(8, 9),
            font=(FONT, 8, "bold"),
        )
        style.map("Image.Treeview.Heading", background=[("active", COLORS["surface_alt"])])
        style.configure(
            "TEntry",
            padding=8,
            fieldbackground="#FFFFFF",
            bordercolor=COLORS["border"],
            lightcolor=COLORS["border"],
            darkcolor=COLORS["border"],
            insertcolor=COLORS["primary"],
        )
        style.map("TEntry", bordercolor=[("focus", COLORS["primary"])], lightcolor=[("focus", COLORS["primary"])], darkcolor=[("focus", COLORS["primary"])])
        style.configure("TSpinbox", padding=7, fieldbackground="#FFFFFF", bordercolor=COLORS["border"], arrowsize=14)
        style.map("TSpinbox", bordercolor=[("focus", COLORS["primary"])], lightcolor=[("focus", COLORS["primary"])], darkcolor=[("focus", COLORS["primary"])])
        style.configure("Horizontal.TProgressbar", troughcolor=COLORS["border_soft"], background=COLORS["primary"], borderwidth=0, thickness=6)

        # Slim Ant-style scrollbars: no arrow buttons and clearer hover feedback.
        style.layout(
            "Vertical.Ant.TScrollbar",
            [("Vertical.Scrollbar.trough", {"sticky": "ns", "children": [("Vertical.Scrollbar.thumb", {"expand": "1", "sticky": "nswe"})]})],
        )
        style.configure(
            "Vertical.Ant.TScrollbar",
            troughcolor=COLORS["surface"],
            background="#D9D9D9",
            bordercolor=COLORS["surface"],
            lightcolor="#D9D9D9",
            darkcolor="#D9D9D9",
            borderwidth=0,
            relief="flat",
            width=8,
        )
        style.map(
            "Vertical.Ant.TScrollbar",
            background=[("pressed", COLORS["subtle"]), ("active", COLORS["disabled"])],
            lightcolor=[("pressed", COLORS["subtle"]), ("active", COLORS["disabled"])],
            darkcolor=[("pressed", COLORS["subtle"]), ("active", COLORS["disabled"])],
        )
        style.layout(
            "Horizontal.Ant.TScrollbar",
            [("Horizontal.Scrollbar.trough", {"sticky": "ew", "children": [("Horizontal.Scrollbar.thumb", {"expand": "1", "sticky": "nswe"})]})],
        )
        style.configure(
            "Horizontal.Ant.TScrollbar",
            troughcolor=COLORS["surface"],
            background="#D9D9D9",
            borderwidth=0,
            relief="flat",
            width=8,
        )

    def _build_ui(self) -> None:
        shell = tk.Frame(self.root, bg=COLORS["app"])
        shell.pack(fill="both", expand=True)

        self.sidebar = tk.Frame(shell, bg=COLORS["sidebar"], width=194, highlightbackground=COLORS["border_soft"], highlightthickness=1)
        self.sidebar.pack(side="left", fill="y")
        self.sidebar.pack_propagate(False)
        self._build_sidebar()

        body = tk.Frame(shell, bg=COLORS["app"])
        body.pack(side="left", fill="both", expand=True)
        self._build_header(body)
        self._build_status_bar(body)

        self.page_host = tk.Frame(body, bg=COLORS["app"])
        self.page_host.pack(fill="both", expand=True, padx=18, pady=(14, 12))
        self.ppt_page = tk.Frame(self.page_host, bg=COLORS["app"])
        self.image_page = tk.Frame(self.page_host, bg=COLORS["app"])
        self._build_ppt_page()
        self._build_image_page()

    def _build_sidebar(self) -> None:
        brand = tk.Frame(self.sidebar, bg=COLORS["sidebar"])
        brand.pack(fill="x", padx=16, pady=(20, 22))
        logo = tk.Canvas(brand, width=42, height=42, bg=COLORS["sidebar"], highlightthickness=0, bd=0)
        logo.pack(side="left")
        logo.create_polygon(rounded_points(1, 1, 41, 41, 10), smooth=True, splinesteps=24, fill=COLORS["primary"], outline="")
        logo.create_text(21, 21, text="P", fill="#FFFFFF", font=("Segoe UI", 15, "bold"))
        brand_text = tk.Frame(brand, bg=COLORS["sidebar"])
        brand_text.pack(side="left", padx=(10, 0))
        tk.Label(brand_text, text="PPT 图片互转", bg=COLORS["sidebar"], fg=COLORS["text"], font=(FONT, 11, "bold")).pack(anchor="w")
        tk.Label(brand_text, text=f"Desktop · v{APP_VERSION}", bg=COLORS["sidebar"], fg=COLORS["subtle"], font=("Segoe UI", 8)).pack(anchor="w", pady=(2, 0))

        tk.Label(self.sidebar, text="转换工作区", bg=COLORS["sidebar"], fg=COLORS["subtle"], font=(FONT, 8, "bold")).pack(anchor="w", padx=17, pady=(0, 7))
        self.nav_ppt = self._sidebar_button("ppt", "PPT 导出图片", "批量导出 PNG / JPG", lambda: self._switch_mode("ppt"))
        self.nav_images = self._sidebar_button("image", "图片合成 PPT", "一张图片生成一页", lambda: self._switch_mode("images"))

        help_section = tk.Frame(self.sidebar, bg=COLORS["sidebar"])
        help_section.pack(side="bottom", fill="x", padx=14, pady=14)
        tips = Card(help_section, radius=RADIUS_CONTROL, shadow=False)
        tips.pack(fill="x")
        tips_body = tips.content
        icon = tk.Canvas(tips_body, width=26, height=26, bg=COLORS["surface"], highlightthickness=0, bd=0)
        icon.pack(anchor="w", padx=10, pady=(9, 2))
        icon.create_oval(2, 2, 24, 24, fill=COLORS["primary_soft"], outline="")
        draw_icon(icon, "upload", 13, 13, COLORS["primary"], size=12, width=2)
        tk.Label(tips_body, text="拖入即可开始", bg=COLORS["surface"], fg=COLORS["primary_active"], font=(FONT, 8, "bold")).pack(anchor="w", padx=10)
        tk.Label(tips_body, text="Ctrl + Enter 快速转换", bg=COLORS["surface"], fg=COLORS["subtle"], font=(FONT, 7)).pack(anchor="w", padx=10, pady=(2, 8))
        help_btn = AntButton(help_section, text="使用说明", icon="info", kind="text", height=32, command=self._open_help)
        help_btn.pack(fill="x", pady=(6, 0))

    def _sidebar_button(self, icon: str, title: str, subtitle: str, command: Callable[[], None]) -> SidebarItem:
        holder = SidebarItem(self.sidebar, icon, title, subtitle, command)
        holder.pack(fill="x", padx=8, pady=2)
        return holder

    def _nav_hover(self, holder: SidebarItem, active: bool) -> None:
        holder._hover(active)

    def _build_header(self, parent: tk.Frame) -> None:
        header = tk.Frame(parent, bg=COLORS["surface"], height=58, highlightbackground=COLORS["border_soft"], highlightthickness=1)
        header.pack(fill="x")
        header.pack_propagate(False)
        left = tk.Frame(header, bg=COLORS["surface"])
        left.pack(side="left", padx=20, pady=9)
        self.header_breadcrumb = tk.Label(left, text="转换工具 / PPT 导出图片", bg=COLORS["surface"], fg=COLORS["subtle"], font=(FONT, 8))
        self.header_breadcrumb.pack(anchor="w")
        self.header_title = tk.Label(left, text="PPT 导出图片", bg=COLORS["surface"], fg=COLORS["text"], font=(FONT, 12, "bold"))
        self.header_title.pack(anchor="w", pady=(2, 0))
        self.header_subtitle = tk.Label(left, text="批量将演示文稿导出为高清图片", bg=COLORS["surface"], fg=COLORS["muted"], font=(FONT, 8))
        # Keep subtitle for compatibility but hide it in compact header.

        right = tk.Frame(header, bg=COLORS["surface"])
        right.pack(side="right", padx=20)
        self.env_tag = StatusTag(right, "正在检测导出环境", command=self._show_backend_status, width=170)
        self.env_tag.pack(side="right")
        Tooltip(self.env_tag, "点击查看 PowerPoint、LibreOffice 与拖拽环境")

    def _build_status_bar(self, parent: tk.Frame) -> None:
        footer = tk.Frame(parent, bg=COLORS["surface"], height=44, highlightbackground=COLORS["border_soft"], highlightthickness=1)
        footer.pack(fill="x", side="bottom")
        footer.pack_propagate(False)
        status_wrap = tk.Frame(footer, bg=COLORS["surface"])
        status_wrap.pack(fill="both", expand=True, padx=18)
        self.status_dot = tk.Label(status_wrap, text="●", bg=COLORS["surface"], fg=COLORS["success"], font=("Segoe UI", 8))
        self.status_dot.pack(side="left")
        self.status_var = tk.StringVar(value="就绪，可拖入文件开始")
        tk.Label(status_wrap, textvariable=self.status_var, bg=COLORS["surface"], fg=COLORS["muted"], font=(FONT, 8)).pack(side="left", padx=(6, 12))
        self.progress = AntProgress(status_wrap, width=275, height=26)
        self.progress.pack(side="right")
        self.log_button = AntButton(status_wrap, text="处理记录", icon="history", kind="text", width=92, height=32,
                                    command=self._show_log_dialog)
        self.log_button.pack(side="right", padx=(0, 10))

    # ---------- Pages ----------
    def _page_heading(self, parent: tk.Frame, title: str, description: str, count_var: tk.StringVar, mode: Mode) -> StepBar:
        # The window header already carries the page title. Keep this area compact and task-focused.
        card = Card(parent, shadow=False, height=66)
        card.pack(fill="x", pady=(0, 12))
        surface = card.content
        icon = tk.Canvas(surface, width=34, height=34, bg=COLORS["surface"], highlightthickness=0, bd=0)
        icon.pack(side="left", padx=(14, 5), pady=15)
        icon.create_oval(2, 2, 32, 32, fill=COLORS["primary_soft"], outline="")
        draw_icon(icon, "ppt" if mode == "ppt" else "image", 17, 17, COLORS["primary"], size=15, width=2)
        steps = StepBar(surface, ("导入文件", "调整参数", "开始转换"))
        steps.pack(side="left", fill="x", expand=True, padx=(0, 8), pady=8)
        badge = tk.Canvas(surface, width=148, height=32, bg=COLORS["surface"], highlightthickness=0, bd=0)
        badge.pack(side="right", padx=(6, 14), pady=17)

        def redraw_badge(*_args) -> None:
            badge.delete("all")
            badge.create_polygon(rounded_points(1, 1, 147, 31, RADIUS_CONTROL), smooth=True, splinesteps=24,
                                 fill=COLORS["primary_soft"], outline="")
            draw_icon(badge, "image" if mode == "images" else "file", 15, 16, COLORS["primary"], size=11, width=2)
            badge.create_text(27, 16, anchor="w", text=count_var.get(), fill=COLORS["primary_active"], font=(FONT, 8, "bold"))

        count_var.trace_add("write", redraw_badge)
        redraw_badge()
        return steps

    def _build_ppt_page(self) -> None:
        self.ppt_count_var = tk.StringVar(value="0 个文件")
        self.ppt_steps = self._page_heading(self.ppt_page, "将 PPT 导出为图片", "支持批量处理并自动选择可用导出引擎", self.ppt_count_var, "ppt")

        self.ppt_dropzone = DropZone(
            self.ppt_page,
            "拖入 PPT 或包含演示文稿的文件夹",
            "自动识别文件类型、去重并加入任务",
            "PPT / PPTX / PPTM / 文件夹",
            self._add_ppts,
        )
        self.ppt_dropzone.pack(fill="x", pady=(0, 12))

        self.ppt_log_card = self._build_log_card(self.ppt_page, "ppt")

        main = tk.Frame(self.ppt_page, bg=COLORS["app"])
        main.pack(fill="both", expand=True)
        main.grid_columnconfigure(0, weight=1, minsize=430)
        main.grid_columnconfigure(1, minsize=344)
        main.grid_rowconfigure(0, weight=1)

        files_card = Card(main)
        files_card.grid(row=0, column=0, sticky="nsew", padx=(0, 12))
        self._build_ppt_file_card(files_card.content)

        settings_card = Card(main, width=344)
        settings_card.grid(row=0, column=1, sticky="nsew")
        settings_card.grid_propagate(False)
        self._build_ppt_settings(settings_card.content)

    def _build_ppt_file_card(self, parent: tk.Frame) -> None:
        head = tk.Frame(parent, bg=COLORS["surface"])
        head.pack(fill="x", padx=14, pady=(11, 9))
        title = tk.Frame(head, bg=COLORS["surface"])
        title.pack(side="left")
        title_row = tk.Frame(title, bg=COLORS["surface"])
        title_row.pack(anchor="w")
        title_icon = tk.Canvas(title_row, width=24, height=24, bg=COLORS["surface"], highlightthickness=0, bd=0)
        title_icon.pack(side="left", padx=(0, 7))
        title_icon.create_oval(2, 2, 22, 22, fill=COLORS["primary_soft"], outline="")
        draw_icon(title_icon, "file", 12, 12, COLORS["primary"], size=11, width=2)
        tk.Label(title_row, text="待处理文件", bg=COLORS["surface"], fg=COLORS["text"], font=(FONT, 10, "bold")).pack(side="left")
        tk.Label(title, text="支持多选，双击打开所在目录", bg=COLORS["surface"], fg=COLORS["subtle"], font=(FONT, 7)).pack(anchor="w", padx=(31, 0), pady=(1, 0))
        actions = tk.Frame(head, bg=COLORS["surface"])
        actions.pack(side="right")
        add_btn = AntButton(actions, text="添加", icon="plus", kind="default", width=76, height=34, command=self._add_ppts)
        add_btn.pack(side="left")
        delete_btn = AntButton(actions, icon="trash", kind="text", width=36, height=34, command=self._remove_ppts)
        delete_btn.pack(side="left", padx=(3, 0))
        clear_btn = AntButton(actions, icon="trash", kind="danger_text", width=36, height=34, command=self._clear_ppts)
        clear_btn.pack(side="left", padx=(1, 0))
        Tooltip(add_btn, "也可以将 PPT 或文件夹直接拖入窗口")
        Tooltip(delete_btn, "删除选中文件")
        Tooltip(clear_btn, "清空全部文件")

        divider = tk.Frame(parent, bg=COLORS["border_soft"], height=1)
        divider.pack(fill="x")
        area = tk.Frame(parent, bg=COLORS["surface"])
        area.pack(fill="both", expand=True, padx=1, pady=(0, 1))
        columns = ("name", "type", "size", "folder", "path")
        self.ppt_tree = ttk.Treeview(area, columns=columns, displaycolumns=("name", "type", "size", "folder"), show="headings", selectmode="extended")
        headings = {"name": "文件名", "type": "格式", "size": "大小", "folder": "所在位置"}
        for column, title_text in headings.items():
            self.ppt_tree.heading(column, text=title_text)
        self.ppt_tree.column("name", width=205, anchor="w")
        self.ppt_tree.column("type", width=62, anchor="center", stretch=False)
        self.ppt_tree.column("size", width=72, anchor="e", stretch=False)
        self.ppt_tree.column("folder", width=138, anchor="w")
        self.ppt_tree.column("path", width=0, stretch=False)
        scroll = AutoHideScrollbar(area, orient="vertical", command=self.ppt_tree.yview)
        self.ppt_tree.configure(yscrollcommand=scroll.set)
        self.ppt_tree.pack(side="left", fill="both", expand=True)
        scroll.pack(side="right", fill="y", padx=(3, 2), pady=5, before=self.ppt_tree)
        self.ppt_tree.bind("<<TreeviewSelect>>", self._preview_selected_ppt)
        self.ppt_tree.bind("<Double-1>", lambda _event: self._open_selected_source("ppt"))
        self.ppt_empty = self._empty_overlay(area, "暂无 PPT 文件", "拖入文件，或点击右上角“添加”")

    def _build_ppt_settings(self, parent: tk.Frame) -> None:
        action = tk.Frame(parent, bg=COLORS["surface"])
        action.pack(fill="x", side="bottom", padx=14, pady=(0, 14))
        self.ppt_start_button = AntButton(action, text="开始导出图片", icon="play", kind="primary", height=40,
                                          command=self._start_ppt_export, state="disabled")
        self.ppt_start_button.pack(fill="x")
        self.ppt_ready_var = tk.StringVar(value="请先导入 PPT 文件")
        row = tk.Frame(action, bg=COLORS["surface"])
        row.pack(fill="x", pady=(6, 0))
        self.ppt_cancel_button = AntButton(row, text="取消", icon="stop", kind="danger", height=32,
                                           command=self._cancel_task, state="disabled")
        self.ppt_cancel_button.pack(side="left", fill="x", expand=True)
        AntButton(row, text="打开目录", icon="external", kind="default", height=32,
                  command=lambda: self._open_path(self.ppt_output_var.get())).pack(side="left", fill="x", expand=True, padx=(7, 0))

        self.ppt_settings_scroll = ScrollablePanel(parent)
        self.ppt_settings_scroll.pack(fill="both", expand=True)
        body = self.ppt_settings_scroll.content
        self._settings_heading(body, "导出设置", "格式、清晰度与保存位置")

        presets = tk.Frame(body, bg=COLORS["surface"])
        presets.pack(fill="x", pady=(0, 5))
        for column in range(3):
            presets.grid_columnconfigure(column, weight=1)
        self.ppt_presets = {
            "web": PresetChoice(presets, "电商", "PNG · 200 DPI", lambda: self._apply_ppt_preset("web")),
            "print": PresetChoice(presets, "高清", "PNG · 300 DPI", lambda: self._apply_ppt_preset("print")),
            "light": PresetChoice(presets, "轻量", "JPG · 160 DPI", lambda: self._apply_ppt_preset("light")),
        }
        for column, key in enumerate(("web", "print", "light")):
            self.ppt_presets[key].grid(row=0, column=column, sticky="ew", padx=(0 if column == 0 else 4, 0 if column == 2 else 4))
        self.ppt_presets["web"].set_selected(True)

        self.ppt_output_var = tk.StringVar()
        self.ppt_format_var = tk.StringVar(value="PNG")
        self.ppt_dpi_var = tk.IntVar(value=200)
        self.jpg_quality_var = tk.IntVar(value=92)
        self.backend_var = tk.StringVar(value="自动选择")
        self.separate_folder_var = tk.BooleanVar(value=True)

        self._field_entry(body, "输出文件夹", self.ppt_output_var, self._choose_ppt_output, "选择")

        pair = tk.Frame(body, bg=COLORS["surface"])
        pair.pack(fill="x", pady=(5, 0))
        pair.grid_columnconfigure(0, weight=1)
        pair.grid_columnconfigure(1, weight=1)
        left = tk.Frame(pair, bg=COLORS["surface"])
        left.grid(row=0, column=0, sticky="ew", padx=(0, 5))
        right = tk.Frame(pair, bg=COLORS["surface"])
        right.grid(row=0, column=1, sticky="ew", padx=(5, 0))
        self._form_label(left, "图片格式")
        AntSelect(left, self.ppt_format_var, ("PNG", "JPG")).pack(fill="x")
        self._form_label(right, "清晰度 DPI")
        AntSpinbox(right, self.ppt_dpi_var, 72, 600, 10).pack(fill="x")

        pair2 = tk.Frame(body, bg=COLORS["surface"])
        pair2.pack(fill="x", pady=(5, 0))
        pair2.grid_columnconfigure(0, weight=1)
        pair2.grid_columnconfigure(1, weight=1)
        left2 = tk.Frame(pair2, bg=COLORS["surface"])
        left2.grid(row=0, column=0, sticky="ew", padx=(0, 5))
        right2 = tk.Frame(pair2, bg=COLORS["surface"])
        right2.grid(row=0, column=1, sticky="ew", padx=(5, 0))
        self._form_label(left2, "JPG 质量")
        AntSpinbox(left2, self.jpg_quality_var, 1, 100, 1).pack(fill="x")
        self._form_label(right2, "导出引擎")
        AntSelect(right2, self.backend_var, ("自动选择", "Microsoft PowerPoint", "LibreOffice")).pack(fill="x")

        AntCheck(body, "每个 PPT 使用独立文件夹", self.separate_folder_var).pack(anchor="w", pady=(5, 0))
        self._alert(body, "自动选择会优先使用 PowerPoint，失败后尝试 LibreOffice。", "info")
        self.ppt_output_var.trace_add("write", lambda *_args: self._refresh_action_state())

    def _build_image_page(self) -> None:
        self.image_count_var = tk.StringVar(value="0 张图片")
        self.image_steps = self._page_heading(self.image_page, "将图片合成为 PPT", "一张图片生成一页，默认保持比例不拉伸", self.image_count_var, "images")

        self.image_dropzone = DropZone(
            self.image_page,
            "拖入图片或图片文件夹",
            "自动识别、去重；导入后可拖动列表调整页序",
            "PNG / JPG / WEBP / TIFF / 文件夹",
            self._add_images,
        )
        self.image_dropzone.pack(fill="x", pady=(0, 12))

        self.image_log_card = self._build_log_card(self.image_page, "images")

        main = tk.Frame(self.image_page, bg=COLORS["app"])
        main.pack(fill="both", expand=True)
        main.grid_columnconfigure(0, weight=1, minsize=430)
        main.grid_columnconfigure(1, minsize=344)
        main.grid_rowconfigure(0, weight=1)

        files_card = Card(main)
        files_card.grid(row=0, column=0, sticky="nsew", padx=(0, 12))
        self._build_image_file_card(files_card.content)

        settings_card = Card(main, width=344)
        settings_card.grid(row=0, column=1, sticky="nsew")
        settings_card.grid_propagate(False)
        self._build_image_settings(settings_card.content)

    def _build_image_file_card(self, parent: tk.Frame) -> None:
        head = tk.Frame(parent, bg=COLORS["surface"])
        head.pack(fill="x", padx=14, pady=(11, 9))
        title = tk.Frame(head, bg=COLORS["surface"])
        title.pack(side="left")
        title_row = tk.Frame(title, bg=COLORS["surface"])
        title_row.pack(anchor="w")
        title_icon = tk.Canvas(title_row, width=24, height=24, bg=COLORS["surface"], highlightthickness=0, bd=0)
        title_icon.pack(side="left", padx=(0, 7))
        title_icon.create_oval(2, 2, 22, 22, fill=COLORS["primary_soft"], outline="")
        draw_icon(title_icon, "image", 12, 12, COLORS["primary"], size=11, width=2)
        tk.Label(title_row, text="幻灯片顺序", bg=COLORS["surface"], fg=COLORS["text"], font=(FONT, 10, "bold")).pack(side="left")
        tk.Label(title, text="缩略图可悬停放大，拖动列表调整页序", bg=COLORS["surface"], fg=COLORS["subtle"],
                 font=(FONT, 7)).pack(anchor="w", padx=(31, 0), pady=(1, 0))
        actions = tk.Frame(head, bg=COLORS["surface"])
        actions.pack(side="right")
        add_btn = AntButton(actions, text="添加", icon="plus", kind="default", width=76, height=34, command=self._add_images)
        add_btn.pack(side="left")
        folder_btn = AntButton(actions, icon="folder", kind="text", width=36, height=34, command=self._add_image_folder)
        folder_btn.pack(side="left", padx=(2, 0))
        sort_btn = AntButton(actions, icon="sort", kind="text", width=36, height=34, command=self._sort_images)
        sort_btn.pack(side="left")
        delete_btn = AntButton(actions, icon="trash", kind="text", width=36, height=34, command=self._remove_images)
        delete_btn.pack(side="left")
        clear_btn = AntButton(actions, icon="trash", kind="danger_text", width=36, height=34, command=self._clear_images)
        clear_btn.pack(side="left")
        Tooltip(add_btn, "支持多选；也可以把图片或文件夹直接拖入")
        Tooltip(folder_btn, "导入整个图片文件夹")
        Tooltip(sort_btn, "按文件名自然排序")
        Tooltip(delete_btn, "删除选中图片")
        Tooltip(clear_btn, "清空全部图片")

        divider = tk.Frame(parent, bg=COLORS["border_soft"], height=1)
        divider.pack(fill="x")
        tree_area = tk.Frame(parent, bg=COLORS["surface"])
        tree_area.pack(fill="both", expand=True, padx=1, pady=(0, 1))
        columns = ("index", "name", "dimensions", "size", "folder", "path")
        self.image_tree = ttk.Treeview(
            tree_area,
            columns=columns,
            displaycolumns=("index", "name", "dimensions", "size", "folder"),
            show="tree headings",
            selectmode="extended",
            style="Image.Treeview",
        )
        self.image_tree.heading("#0", text="预览")
        self.image_tree.column("#0", width=66, minwidth=66, anchor="center", stretch=False)
        headings = {"index": "页序", "name": "文件名", "dimensions": "像素尺寸", "size": "大小", "folder": "所在位置"}
        for column, title_text in headings.items():
            self.image_tree.heading(column, text=title_text)
        self.image_tree.column("index", width=44, anchor="center", stretch=False)
        self.image_tree.column("name", width=154, minwidth=112, anchor="w")
        self.image_tree.column("dimensions", width=84, anchor="center", stretch=False)
        self.image_tree.column("size", width=60, anchor="e", stretch=False)
        self.image_tree.column("folder", width=86, minwidth=64, anchor="w")
        self.image_tree.column("path", width=0, stretch=False)
        scroll = AutoHideScrollbar(tree_area, orient="vertical", command=self.image_tree.yview)
        self.image_tree.configure(yscrollcommand=scroll.set)
        self.image_tree.pack(side="left", fill="both", expand=True)
        scroll.pack(side="right", fill="y", padx=(3, 2), pady=5, before=self.image_tree)
        self.image_tree.bind("<<TreeviewSelect>>", self._preview_selected_image)
        self.image_tree.bind("<Double-1>", self._on_image_tree_double_click)
        self.image_tree.bind("<Motion>", self._image_tree_motion, add="+")
        self.image_tree.bind("<Leave>", self._image_tree_leave, add="+")
        self.image_tree.bind("<Button-3>", self._image_tree_context_menu, add="+")
        self.image_tree.bind("<ButtonPress-1>", self._image_drag_press, add="+")
        self.image_tree.bind("<B1-Motion>", self._image_drag_motion, add="+")
        self.image_tree.bind("<ButtonRelease-1>", self._image_drag_release, add="+")
        self.image_empty = self._empty_overlay(tree_area, "暂无图片", "拖入图片后，可预览并拖动调整页序")

        self.image_zoom_button = AntButton(self.image_tree, icon="zoom", kind="overlay", width=27, height=27,
                                           radius=RADIUS_SMALL, command=self._open_hovered_image_preview)
        self.image_zoom_button.place_forget()
        self.image_zoom_button.bind("<Enter>", lambda _e: self._cancel_hover_hide(), add="+")
        self.image_zoom_button.bind("<Leave>", lambda _e: self._schedule_hover_hide(), add="+")
        Tooltip(self.image_zoom_button, "查看大图")

    def _build_image_settings(self, parent: tk.Frame) -> None:
        self.pptx_output_var = tk.StringVar()
        self.slide_size_var = tk.StringVar(value="16:9")
        self.fit_mode_var = tk.StringVar(value="完整显示（不裁切）")
        self.background_var = tk.StringVar(value="白色")

        action = tk.Frame(parent, bg=COLORS["surface"])
        action.pack(fill="x", side="bottom", padx=14, pady=(0, 14))
        path_row = tk.Frame(action, bg=COLORS["surface"])
        path_row.pack(fill="x", pady=(0, 7))
        tk.Label(path_row, text="保存位置", bg=COLORS["surface"], fg=COLORS["muted"],
                 font=(FONT, 8)).pack(side="left", padx=(0, 7))
        AntEntry(path_row, self.pptx_output_var).pack(side="left", fill="x", expand=True)
        AntButton(path_row, text="另存为", icon="folder", kind="default", width=78, height=36,
                  command=self._choose_pptx_output).pack(side="left", padx=(6, 0))

        self.image_start_button = AntButton(action, text="生成 PPT 文件", icon="play", kind="primary", height=40,
                                            command=self._start_images_to_ppt, state="disabled")
        self.image_start_button.pack(fill="x")
        self.image_ready_var = tk.StringVar(value="请先导入图片")
        row = tk.Frame(action, bg=COLORS["surface"])
        row.pack(fill="x", pady=(6, 0))
        self.image_cancel_button = AntButton(row, text="取消", icon="stop", kind="danger", height=32,
                                             command=self._cancel_task, state="disabled")
        self.image_cancel_button.pack(side="left", fill="x", expand=True)
        AntButton(row, text="打开目录", icon="external", kind="default", height=32,
                  command=self._open_pptx_parent).pack(side="left", fill="x", expand=True, padx=(7, 0))

        self.image_settings_scroll = ScrollablePanel(parent)
        self.image_settings_scroll.pack(fill="both", expand=True)
        body = self.image_settings_scroll.content
        self._settings_heading(body, "PPT 设置", "页面比例、背景与图片排版")

        self._settings_section(body, "快速预设", "check")
        presets = tk.Frame(body, bg=COLORS["surface"])
        presets.pack(fill="x", pady=(0, 5))
        for column in range(3):
            presets.grid_columnconfigure(column, weight=1)
        self.image_presets = {
            "wide": PresetChoice(presets, "16:9", "完整显示", lambda: self._apply_image_preset("wide")),
            "source": PresetChoice(presets, "首图比例", "不裁切", lambda: self._apply_image_preset("source")),
            "cover": PresetChoice(presets, "铺满", "居中裁切", lambda: self._apply_image_preset("cover")),
        }
        for column, key in enumerate(("wide", "source", "cover")):
            self.image_presets[key].grid(row=0, column=column, sticky="ew",
                                         padx=(0 if column == 0 else 4, 0 if column == 2 else 4))
        self.image_presets["wide"].set_selected(True)

        self._settings_section(body, "页面设置", "ppt")
        pair = tk.Frame(body, bg=COLORS["surface"])
        pair.pack(fill="x", pady=(0, 5))
        pair.grid_columnconfigure(0, weight=1)
        pair.grid_columnconfigure(1, weight=1)
        left = tk.Frame(pair, bg=COLORS["surface"])
        left.grid(row=0, column=0, sticky="ew", padx=(0, 5))
        right = tk.Frame(pair, bg=COLORS["surface"])
        right.grid(row=0, column=1, sticky="ew", padx=(5, 0))
        self._form_label(left, "页面比例")
        AntSelect(left, self.slide_size_var, ("16:9", "4:3", "A4横向", "A4竖向", "按首图比例")).pack(fill="x")
        self._form_label(right, "页面背景")
        AntSelect(right, self.background_var, ("白色", "黑色")).pack(fill="x")

        self._settings_section(body, "图片布局", "image")
        AntSelect(body, self.fit_mode_var, ("完整显示（不裁切）", "铺满页面（居中裁切）")).pack(fill="x")
        self._alert(body, "完整显示不拉伸；铺满页面会从中心等比裁切。", "info")
        self.pptx_output_var.trace_add("write", lambda *_args: self._refresh_action_state())

    def _build_log_card(self, parent: tk.Frame, mode: Mode) -> tk.Frame:
        # Keep logs available without consuming permanent page height. The status-bar button opens a drawer window.
        holder = tk.Frame(parent, bg=COLORS["app"])
        log = tk.Text(holder, height=1, wrap="word", bd=0, bg=COLORS["code"], fg="#D9D9D9",
                      insertbackground="#FFFFFF", font=("Consolas", 9), padx=12, pady=9)
        log.configure(state="disabled")
        if mode == "ppt":
            self.ppt_log = log
        else:
            self.image_log = log
        return holder

    def _empty_overlay(self, parent: tk.Frame, title: str, subtitle: str) -> tk.Frame:
        overlay = tk.Frame(parent, bg="#FFFFFF")
        icon = tk.Canvas(overlay, width=50, height=50, bg="#FFFFFF", highlightthickness=0, bd=0)
        icon.pack(pady=(0, 2))
        icon.create_polygon(rounded_points(2, 2, 48, 48, 12), smooth=True, splinesteps=24,
                            fill=COLORS["surface_alt"], outline=COLORS["border_soft"])
        draw_icon(icon, "file", 25, 25, COLORS["disabled"], size=22, width=2)
        tk.Label(overlay, text=title, bg="#FFFFFF", fg=COLORS["text"], font=(FONT, 10, "bold")).pack()
        tk.Label(overlay, text=subtitle, bg="#FFFFFF", fg=COLORS["subtle"], font=(FONT, 8)).pack(pady=(5, 0))
        overlay.place(relx=0.5, rely=0.5, anchor="center")
        return overlay

    # ---------- Common components ----------
    def _settings_heading(self, parent: tk.Frame, title: str, subtitle: str) -> None:
        row = tk.Frame(parent, bg=COLORS["surface"])
        row.pack(fill="x", pady=(0, 6))
        icon = tk.Canvas(row, width=28, height=28, bg=COLORS["surface"], highlightthickness=0, bd=0)
        icon.pack(side="left", padx=(0, 8))
        icon.create_polygon(rounded_points(1, 1, 27, 27, RADIUS_CONTROL), smooth=True, splinesteps=24,
                            fill=COLORS["primary_soft"], outline="")
        draw_icon(icon, "settings", 14, 14, COLORS["primary"], size=13, width=2)
        text = tk.Frame(row, bg=COLORS["surface"])
        text.pack(side="left", fill="x", expand=True)
        tk.Label(text, text=title, bg=COLORS["surface"], fg=COLORS["text"], font=(FONT, 10, "bold")).pack(anchor="w")
        tk.Label(text, text=subtitle, bg=COLORS["surface"], fg=COLORS["subtle"], font=(FONT, 7)).pack(anchor="w", pady=(1, 0))

    def _settings_section(self, parent: tk.Frame, title: str, icon_name: str) -> None:
        row = tk.Frame(parent, bg=COLORS["surface"])
        row.pack(fill="x", pady=(5, 3))
        icon = tk.Canvas(row, width=18, height=18, bg=COLORS["surface"], highlightthickness=0, bd=0)
        icon.pack(side="left", padx=(0, 6))
        draw_icon(icon, icon_name, 9, 9, COLORS["primary"], size=10, width=1)
        tk.Label(row, text=title, bg=COLORS["surface"], fg=COLORS["muted"],
                 font=(FONT, 8, "bold")).pack(side="left")
        tk.Frame(row, bg=COLORS["border_soft"], height=1).pack(side="left", fill="x", expand=True, padx=(8, 0))

    def _form_label(self, parent: tk.Frame, text: str, pady: tuple[int, int] = (0, 4)) -> None:
        tk.Label(parent, text=text, bg=COLORS["surface"], fg=COLORS["muted"], font=(FONT, 8)).pack(anchor="w", pady=pady)

    def _field_entry(self, parent: tk.Frame, label: str, variable: tk.StringVar, command: Callable[[], None], button_text: str) -> None:
        self._form_label(parent, label, pady=(2, 3))
        line = tk.Frame(parent, bg=COLORS["surface"])
        line.pack(fill="x")
        AntEntry(line, variable).pack(side="left", fill="x", expand=True)
        AntButton(line, text=button_text, icon="folder", kind="default", width=78, height=36,
                  command=command).pack(side="left", padx=(6, 0))

    def _field_combo(self, parent: tk.Frame, label: str, variable: tk.StringVar, values: tuple[str, ...]) -> None:
        self._form_label(parent, label, pady=(9, 4))
        AntSelect(parent, variable, values).pack(fill="x")

    def _field_spin(self, parent: tk.Frame, label: str, variable: tk.IntVar, start: int, end: int, increment: int) -> None:
        self._form_label(parent, label, pady=(9, 4))
        AntSpinbox(parent, variable, start, end, increment).pack(fill="x")

    def _alert(self, parent: tk.Frame, text: str, kind: str = "info") -> tk.Canvas:
        palette = {
            "info": (COLORS["primary_soft"], COLORS["primary_border"], COLORS["primary"]),
            "success": (COLORS["success_soft"], "#B7EB8F", COLORS["success"]),
            "warning": (COLORS["warning_soft"], "#FFE58F", COLORS["warning"]),
        }
        bg, border, icon_color = palette.get(kind, palette["info"])
        frame = tk.Canvas(parent, height=32, bg=COLORS["surface"], highlightthickness=0, bd=0)
        frame.pack(fill="x", pady=(5, 0))

        def redraw(_event=None) -> None:
            frame.delete("all")
            width = max(120, frame.winfo_width())
            frame.create_polygon(rounded_points(1, 1, width - 2, 30, RADIUS_CONTROL), smooth=True, splinesteps=24,
                                 fill=bg, outline=border, width=1)
            draw_icon(frame, "info", 15, 16, icon_color, size=11, width=2)
            frame.create_text(28, 16, anchor="w", text=text, fill=COLORS["muted"], font=(FONT, 7))

        frame.bind("<Configure>", redraw, add="+")
        redraw()
        return frame

    def _dismiss_toast(self) -> None:
        """Remove the active notification and its scheduled callback safely."""
        if self._toast_after_id:
            try:
                self.root.after_cancel(self._toast_after_id)
            except tk.TclError:
                pass
            self._toast_after_id = None
        if self._toast:
            try:
                if self._toast.winfo_exists():
                    self._toast.destroy()
            except tk.TclError:
                pass
        self._toast = None

    def _show_toast(self, title: str, message: str, state: str = "success") -> None:
        self._dismiss_toast()
        palette = {
            "success": (COLORS["success"], "✓"),
            "warning": (COLORS["warning"], "!"),
            "error": (COLORS["danger"], "×"),
            "info": (COLORS["primary"], "i"),
        }
        color, icon = palette.get(state, palette["info"])
        toast = Card(self.root, radius=RADIUS_CARD, shadow=True, width=380, height=76)
        self._toast = toast
        surface = toast.content
        icon_label = tk.Canvas(surface, width=28, height=28, bg=COLORS["surface"], highlightthickness=0, bd=0)
        icon_label.pack(side="left", padx=(10, 8), pady=13)
        icon_label.create_oval(2, 2, 26, 26, fill=color, outline="")
        if icon == "✓":
            draw_icon(icon_label, "check", 14, 14, "#FFFFFF", size=11, width=2)
        else:
            icon_label.create_text(14, 14, text=icon, fill="#FFFFFF", font=("Segoe UI", 9, "bold"))
        content = tk.Frame(surface, bg=COLORS["surface"])
        content.pack(side="left", fill="both", expand=True, padx=(0, 12), pady=10)
        tk.Label(content, text=title, bg=COLORS["surface"], fg=COLORS["text"], font=(FONT, 9, "bold")).pack(anchor="w")
        tk.Label(content, text=message, bg=COLORS["surface"], fg=COLORS["muted"], font=(FONT, 8), wraplength=330, justify="left").pack(anchor="w", pady=(3, 0))
        # Do not obstruct the step bar or current task; stack feedback above the status bar.
        toast.place(relx=1.0, rely=1.0, x=-18, y=-58, anchor="se")
        # Card is canvas-based, whose .lift() expects a canvas item. Raise the
        # widget itself through Tcl so the timer below is always registered.
        self.root.tk.call("raise", toast._w)
        self._toast_after_id = self.root.after(3200, self._dismiss_toast)

    def _show_log_dialog(self) -> None:
        source = self.ppt_log if self.active_mode == "ppt" else self.image_log
        title_text = "PPT 导出处理记录" if self.active_mode == "ppt" else "图片合成处理记录"
        window = tk.Toplevel(self.root)
        window.title(title_text)
        window.geometry("720x440")
        window.minsize(560, 340)
        window.configure(bg=COLORS["app"])
        window.transient(self.root)

        shell = Card(window, radius=RADIUS_CARD, shadow=True)
        shell.pack(fill="both", expand=True, padx=16, pady=16)
        surface = shell.content
        header = tk.Frame(surface, bg=COLORS["surface"])
        header.pack(fill="x", padx=14, pady=(12, 9))
        icon = tk.Canvas(header, width=30, height=30, bg=COLORS["surface"], highlightthickness=0, bd=0)
        icon.pack(side="left", padx=(0, 8))
        icon.create_polygon(rounded_points(1, 1, 29, 29, RADIUS_CONTROL), smooth=True, splinesteps=24,
                            fill=COLORS["primary_soft"], outline="")
        draw_icon(icon, "history", 15, 15, COLORS["primary"], size=14, width=2)
        tk.Label(header, text=title_text, bg=COLORS["surface"], fg=COLORS["text"],
                 font=(FONT, 11, "bold")).pack(side="left")
        close_btn = AntButton(header, text="关闭", kind="default", width=68, height=32, command=window.destroy)
        close_btn.pack(side="right")

        body = tk.Frame(surface, bg=COLORS["surface"])
        body.pack(fill="both", expand=True, padx=14, pady=(0, 10))
        text = tk.Text(body, wrap="word", bd=0, bg=COLORS["code"], fg="#D9D9D9", insertbackground="#FFFFFF",
                       font=("Consolas", 9), padx=14, pady=12, relief="flat")
        scroll = AutoHideScrollbar(body, orient="vertical", command=text.yview)
        text.configure(yscrollcommand=scroll.set)
        text.pack(side="left", fill="both", expand=True)
        scroll.pack(side="right", fill="y", padx=(4, 0), before=text)
        content = source.get("1.0", "end-1c").strip()
        text.insert("1.0", content or "暂无处理记录。完成导入或转换后，日志会显示在这里。")
        text.configure(state="disabled")

        footer = tk.Frame(surface, bg=COLORS["surface"])
        footer.pack(fill="x", padx=14, pady=(0, 12))

        def clear_logs() -> None:
            self._clear_log(source)
            text.configure(state="normal")
            text.delete("1.0", "end")
            text.insert("1.0", "暂无处理记录。")
            text.configure(state="disabled")

        AntButton(footer, text="清空记录", icon="trash", kind="danger_text", width=92, height=32,
                  command=clear_logs).pack(side="right")
        window.grab_set()
        window.focus_force()

    def _open_help(self) -> None:
        help_path = resource_path("README_使用说明.md")
        if help_path.exists():
            self._open_path(str(help_path))
        else:
            messagebox.showinfo("使用说明", "README_使用说明.md 未找到。")

    def _refresh_action_state(self) -> None:
        if not hasattr(self, "ppt_start_button"):
            return
        ppt_files = len(self.ppt_tree.get_children()) if hasattr(self, "ppt_tree") else 0
        ppt_output = bool(self.ppt_output_var.get().strip()) if hasattr(self, "ppt_output_var") else False
        if self.busy:
            self.ppt_start_button.configure(state="disabled")
            self.image_start_button.configure(state="disabled")
        else:
            self.ppt_start_button.configure(state="normal" if ppt_files and ppt_output else "disabled")
            image_files = len(self.image_tree.get_children()) if hasattr(self, "image_tree") else 0
            image_output = bool(self.pptx_output_var.get().strip()) if hasattr(self, "pptx_output_var") else False
            self.image_start_button.configure(state="normal" if image_files and image_output else "disabled")
        if not ppt_files:
            self.ppt_ready_var.set("请先导入 PPT 文件")
        elif not ppt_output:
            self.ppt_ready_var.set("请选择图片输出文件夹")
        else:
            self.ppt_ready_var.set("设置已完成，可以开始导出")
        image_files = len(self.image_tree.get_children()) if hasattr(self, "image_tree") else 0
        image_output = bool(self.pptx_output_var.get().strip()) if hasattr(self, "pptx_output_var") else False
        if not image_files:
            self.image_ready_var.set("请先导入图片")
        elif not image_output:
            self.image_ready_var.set("请选择 PPTX 保存位置")
        else:
            self.image_ready_var.set("设置已完成，可以生成 PPT")
        self._update_steps()

    def _update_steps(self, mode: Mode | None = None, finished: bool = False, error: bool = False) -> None:
        modes = (mode,) if mode else ("ppt", "images")
        for current_mode in modes:
            if current_mode == "ppt" and hasattr(self, "ppt_steps"):
                has_files = bool(self.ppt_tree.get_children())
                has_output = bool(self.ppt_output_var.get().strip())
                index = 0 if not has_files else (1 if not has_output else 2)
                self.ppt_steps.set_current(index, finished=finished and index == 2, error=error and index == 2)
            elif current_mode == "images" and hasattr(self, "image_steps"):
                has_files = bool(self.image_tree.get_children())
                has_output = bool(self.pptx_output_var.get().strip())
                index = 0 if not has_files else (1 if not has_output else 2)
                self.image_steps.set_current(index, finished=finished and index == 2, error=error and index == 2)

    def _apply_initial_mode(self) -> None:
        requested = self.config.get("active_mode", "ppt")
        self._switch_mode("images" if requested == "images" else "ppt")

    def _switch_mode(self, mode: Mode) -> None:
        self.active_mode = mode
        self.ppt_page.pack_forget()
        self.image_page.pack_forget()
        page = self.ppt_page if mode == "ppt" else self.image_page
        page.pack(fill="both", expand=True)
        if mode == "ppt":
            self.header_breadcrumb.configure(text="转换工具 / PPT 导出图片")
            self.header_title.configure(text="PPT 导出图片")
            self.header_subtitle.configure(text="批量将演示文稿导出为高清图片")
        else:
            self.header_breadcrumb.configure(text="转换工具 / 图片合成 PPT")
            self.header_title.configure(text="图片合成 PPT")
            self.header_subtitle.configure(text="保持图片比例，快速生成可编辑演示文稿")
        self._refresh_nav()
        self._update_steps(mode)

    def _refresh_nav(self) -> None:
        for mode, holder in (("ppt", self.nav_ppt), ("images", self.nav_images)):
            holder.set_selected(mode == self.active_mode)

    def _bind_shortcuts(self) -> None:
        self.root.bind_all("<Control-o>", lambda _event: self._add_ppts() if self.active_mode == "ppt" else self._add_images())
        self.root.bind_all("<Control-O>", lambda _event: self._add_ppts() if self.active_mode == "ppt" else self._add_images())
        self.root.bind_all("<Delete>", lambda _event: self._remove_ppts() if self.active_mode == "ppt" else self._remove_images())
        self.root.bind_all("<Control-Return>", lambda _event: self._start_ppt_export() if self.active_mode == "ppt" else self._start_images_to_ppt())
        self.root.bind_all("<Control-Up>", lambda _event: self._move_image(-1) if self.active_mode == "images" else None)
        self.root.bind_all("<Control-Down>", lambda _event: self._move_image(1) if self.active_mode == "images" else None)

    def _on_root_resize(self, event) -> None:
        if event.widget is not self.root:
            return
        compact_width = event.width < 1120
        self.ppt_dropzone.configure(height=76 if event.height < 720 else 80)
        self.image_dropzone.configure(height=76 if event.height < 720 else 80)
        self.sidebar.configure(width=178 if compact_width else 194)

    def _add_ppts(self) -> None:
        patterns = " ".join(f"*{ext}" for ext in sorted(SUPPORTED_PRESENTATIONS))
        paths = filedialog.askopenfilenames(title="选择 PowerPoint 文件", filetypes=(("PowerPoint 文件", patterns), ("所有文件", "*.*")))
        self._append_ppts([Path(path) for path in paths])

    def _add_images(self) -> None:
        patterns = " ".join(f"*{ext}" for ext in sorted(SUPPORTED_IMAGES))
        paths = filedialog.askopenfilenames(title="选择图片", filetypes=(("图片文件", patterns), ("所有文件", "*.*")))
        self._append_images([Path(path) for path in paths])

    def _add_image_folder(self) -> None:
        folder = filedialog.askdirectory(title="选择图片文件夹")
        if folder:
            self._append_images(self._scan_folder(Path(folder), SUPPORTED_IMAGES))

    def _handle_native_drop(self, paths: list[Path], _x: int, _y: int) -> None:
        ppt_files: list[Path] = []
        image_files: list[Path] = []
        skipped = 0
        for path in unique_existing(paths):
            candidates = [path]
            if path.is_dir():
                try:
                    candidates = sorted((item for item in path.iterdir() if item.is_file()), key=natural_sort_key)
                except OSError:
                    candidates = []
            for candidate in candidates:
                suffix = candidate.suffix.lower()
                if suffix in SUPPORTED_PRESENTATIONS:
                    ppt_files.append(candidate)
                elif suffix in SUPPORTED_IMAGES:
                    image_files.append(candidate)
                else:
                    skipped += 1

        added_ppt = self._append_ppts(ppt_files)
        added_images = self._append_images(image_files)
        if added_ppt and not added_images:
            self._switch_mode("ppt")
            self.ppt_dropzone.pulse()
        elif added_images and not added_ppt:
            self._switch_mode("images")
            self.image_dropzone.pulse()
        elif added_ppt and added_images:
            (self.ppt_dropzone if self.active_mode == "ppt" else self.image_dropzone).pulse()

        parts: list[str] = []
        if added_ppt:
            parts.append(f"{added_ppt} 个 PPT")
        if added_images:
            parts.append(f"{added_images} 张图片")
        if parts:
            suffix = f"；忽略 {skipped} 个不支持文件" if skipped else ""
            message = f"已加入 {'、'.join(parts)}{suffix}"
            self._set_status(message, "success")
            self._show_toast("导入完成", message, "success")
        elif ppt_files or image_files:
            self._set_status("文件已在列表中，未重复添加", "warning")
            self._show_toast("未重复添加", "拖入的文件已经存在于任务列表中。", "warning")
        else:
            self._set_status("未发现支持的 PPT 或图片文件", "warning")
            self._show_toast("无法导入", "没有发现受支持的 PowerPoint 或图片文件。", "warning")

    def _scan_folder(folder: Path, extensions: set[str]) -> list[Path]:
        try:
            return sorted((path for path in folder.iterdir() if path.is_file() and path.suffix.lower() in extensions), key=natural_sort_key)
        except OSError:
            return []

    def _append_ppts(self, paths: list[Path]) -> int:
        existing = {os.path.normcase(str(Path(self.ppt_tree.item(item, "values")[4]))) for item in self.ppt_tree.get_children()}
        added = 0
        first_parent: Path | None = None
        for path in unique_existing(paths):
            if path.is_dir():
                for child in self._scan_folder(path, SUPPORTED_PRESENTATIONS):
                    added += self._append_ppts([child])
                continue
            if path.suffix.lower() not in SUPPORTED_PRESENTATIONS:
                continue
            key = os.path.normcase(str(path))
            if key in existing:
                continue
            try:
                size = human_size(path.stat().st_size)
            except OSError:
                size = "—"
            self.ppt_tree.insert("", "end", values=(path.name, path.suffix.upper().lstrip("."), size, path.parent.name or str(path.parent), str(path)))
            existing.add(key)
            first_parent = first_parent or path.parent
            added += 1
        if added and not self.ppt_output_var.get().strip() and first_parent:
            self.ppt_output_var.set(str(first_parent / "PPT导出图片"))
        self._refresh_ppt_state()
        return added

    def _append_images(self, paths: list[Path]) -> int:
        existing = {os.path.normcase(str(Path(self.image_tree.item(item, "values")[5]))) for item in self.image_tree.get_children()}
        added = 0
        first_parent: Path | None = None
        for path in unique_existing(paths):
            if path.is_dir():
                for child in self._scan_folder(path, SUPPORTED_IMAGES):
                    added += self._append_images([child])
                continue
            if path.suffix.lower() not in SUPPORTED_IMAGES:
                continue
            key = os.path.normcase(str(path))
            if key in existing:
                continue
            try:
                with Image.open(path) as opened:
                    image = ImageOps.exif_transpose(opened)
                    dimensions = f"{image.width}×{image.height}"
            except Exception:
                dimensions = "无法读取"
            try:
                size = human_size(path.stat().st_size)
            except OSError:
                size = "—"
            thumbnail = self._create_image_thumbnail(path)
            item = self.image_tree.insert("", "end", text="", image=thumbnail,
                                          values=(0, path.name, dimensions, size, path.parent.name or str(path.parent), str(path)))
            if thumbnail is not None:
                self._image_thumbnails[item] = thumbnail
            existing.add(key)
            first_parent = first_parent or path.parent
            added += 1
        self._renumber_images()
        if added and not self.pptx_output_var.get().strip() and first_parent:
            self.pptx_output_var.set(str(first_parent / "图片合成.pptx"))
        self._refresh_image_state()
        return added

    # ---------- 图片缩略图与大图预览 ----------
    def _create_image_thumbnail(self, path: Path, size: int = 54) -> ImageTk.PhotoImage | None:
        try:
            with Image.open(path) as opened:
                source = ImageOps.exif_transpose(opened).convert("RGBA")
                source.thumbnail((size - 8, size - 8), Image.Resampling.LANCZOS)
            thumb = Image.new("RGBA", (size, size), (255, 255, 255, 0))
            card = Image.new("RGBA", (size - 2, size - 2), (255, 255, 255, 255))
            mask = Image.new("L", card.size, 0)
            ImageDraw.Draw(mask).rounded_rectangle((0, 0, card.width - 1, card.height - 1), radius=8, fill=255)
            thumb.paste(card, (1, 1), mask)
            x = (size - source.width) // 2
            y = (size - source.height) // 2
            thumb.alpha_composite(source, (x, y))
            draw = ImageDraw.Draw(thumb)
            draw.rounded_rectangle((1, 1, size - 2, size - 2), radius=8, outline=(217, 217, 217, 255), width=1)
            return ImageTk.PhotoImage(thumb)
        except Exception:
            return None

    def _image_tree_motion(self, event) -> None:
        item = self.image_tree.identify_row(event.y)
        column = self.image_tree.identify_column(event.x)
        region = self.image_tree.identify_region(event.x, event.y)
        if item and column == "#0" and region in ("tree", "cell"):
            bbox = self.image_tree.bbox(item, "#0")
            if bbox:
                self._hover_image_item = item
                x, y, width, _height = bbox
                self.image_zoom_button.place(x=x + width - 31, y=y + 5)
                self.image_zoom_button.lift()
                self._schedule_hover_preview(item)
                return
        self._schedule_hover_hide(40)

    def _image_tree_leave(self, _event=None) -> None:
        self._schedule_hover_hide(100)

    def _cancel_hover_hide(self) -> None:
        if self._hover_after_id:
            try:
                self.root.after_cancel(self._hover_after_id)
            except Exception:
                pass
            self._hover_after_id = None

    def _schedule_hover_hide(self, delay: int = 100) -> None:
        self._cancel_hover_hide()
        self._hover_after_id = self.root.after(delay, self._maybe_hide_image_hover_ui)

    def _maybe_hide_image_hover_ui(self) -> None:
        self._hover_after_id = None
        try:
            widget = self.root.winfo_containing(self.root.winfo_pointerx(), self.root.winfo_pointery())
        except Exception:
            widget = None
        if widget is self.image_zoom_button:
            return
        if self._hover_preview_window and widget is not None:
            try:
                if widget.winfo_toplevel() is self._hover_preview_window:
                    return
            except Exception:
                pass
        self._hide_image_hover_ui()

    def _hide_image_hover_ui(self) -> None:
        self._cancel_hover_hide()
        self._hover_image_item = None
        if hasattr(self, "image_zoom_button"):
            self.image_zoom_button.place_forget()
        if self._hover_preview_window:
            try:
                self._hover_preview_window.destroy()
            except Exception:
                pass
        self._hover_preview_window = None
        self._hover_preview_photo = None

    def _schedule_hover_preview(self, item: str) -> None:
        self._cancel_hover_hide()
        if self._hover_preview_window and self._hover_image_item == item:
            return
        self._hover_after_id = self.root.after(170, lambda: self._show_image_hover_preview(item))

    def _show_image_hover_preview(self, item: str) -> None:
        self._hover_after_id = None
        if item != self._hover_image_item or not self.image_tree.exists(item):
            return
        values = self.image_tree.item(item, "values")
        path = Path(values[5])
        try:
            with Image.open(path) as opened:
                image = ImageOps.exif_transpose(opened).convert("RGBA")
                image.thumbnail((250, 205), Image.Resampling.LANCZOS)
                background = Image.new("RGBA", image.size, "white")
                background.alpha_composite(image)
                self._hover_preview_photo = ImageTk.PhotoImage(background.convert("RGB"))
        except Exception:
            return
        if self._hover_preview_window:
            try:
                self._hover_preview_window.destroy()
            except Exception:
                pass
        popup = tk.Toplevel(self.root)
        self._hover_preview_window = popup
        popup.overrideredirect(True)
        popup.attributes("-topmost", True)
        popup.configure(bg=COLORS["border"])
        card = tk.Frame(popup, bg=COLORS["surface"], highlightbackground=COLORS["border_soft"], highlightthickness=1)
        card.pack(fill="both", expand=True, padx=1, pady=1)
        image_label = tk.Label(card, image=self._hover_preview_photo, bg="#FFFFFF", bd=0)
        image_label.pack(padx=10, pady=(10, 7))
        tk.Label(card, text=path.name, bg=COLORS["surface"], fg=COLORS["text"], font=(FONT, 9, "bold"),
                 anchor="w", justify="left", wraplength=250).pack(fill="x", padx=11)
        tk.Label(card, text=f"第 {values[0]} 页 · {values[2]} · {values[3]}", bg=COLORS["surface"],
                 fg=COLORS["subtle"], font=(FONT, 8), anchor="w").pack(fill="x", padx=11, pady=(2, 10))
        popup.update_idletasks()
        bbox = self.image_tree.bbox(item, "#0")
        x = self.image_tree.winfo_rootx() + (bbox[0] + bbox[2] + 12 if bbox else 90)
        y = self.image_tree.winfo_rooty() + (bbox[1] if bbox else 40)
        popup_w = popup.winfo_reqwidth()
        popup_h = popup.winfo_reqheight()
        screen_w = self.root.winfo_screenwidth()
        screen_h = self.root.winfo_screenheight()
        if x + popup_w > screen_w - 10:
            x = self.image_tree.winfo_rootx() + (bbox[0] if bbox else 0) - popup_w - 12
        y = min(max(8, y), max(8, screen_h - popup_h - 42))
        popup.geometry(f"+{int(x)}+{int(y)}")
        popup.bind("<Enter>", lambda _e: self._cancel_hover_hide(), add="+")
        popup.bind("<Leave>", lambda _e: self._schedule_hover_hide(100), add="+")

    def _open_hovered_image_preview(self) -> None:
        if self._hover_image_item:
            self._open_image_preview(self._hover_image_item)

    def _open_image_preview(self, item: str | None = None) -> None:
        items = list(self.image_tree.get_children())
        if not items:
            return
        if item is None:
            selection = self.image_tree.selection()
            item = selection[0] if selection else items[0]
        if item not in items:
            return
        paths = [Path(self.image_tree.item(row, "values")[5]) for row in items]
        index = items.index(item)
        self._hide_image_hover_ui()
        if self._preview_dialog:
            try:
                self._preview_dialog.close()
            except Exception:
                pass
        self._preview_dialog = ImagePreviewDialog(self, paths, index)

    def _on_image_tree_double_click(self, event) -> str | None:
        item = self.image_tree.identify_row(event.y)
        if item:
            self.image_tree.selection_set(item)
            self._open_image_preview(item)
            return "break"
        return None

    def _image_tree_context_menu(self, event) -> str:
        item = self.image_tree.identify_row(event.y)
        if item:
            if item not in self.image_tree.selection():
                self.image_tree.selection_set(item)
            menu = tk.Menu(self.root, tearoff=False, font=(FONT, 9), bg=COLORS["surface"], fg=COLORS["text"],
                           activebackground=COLORS["primary_soft"], activeforeground=COLORS["primary_active"], bd=1)
            menu.add_command(label="查看大图", command=lambda: self._open_image_preview(item))
            menu.add_command(label="打开所在位置", command=lambda: self._open_path(str(Path(self.image_tree.item(item, "values")[5]).parent)))
            menu.add_separator()
            menu.add_command(label="上移", command=lambda: self._move_image(-1))
            menu.add_command(label="下移", command=lambda: self._move_image(1))
            menu.add_command(label="删除", command=self._remove_images)
            try:
                menu.tk_popup(event.x_root, event.y_root)
            finally:
                menu.grab_release()
        return "break"

    # ---------- 列表交互 ----------
    def _remove_ppts(self) -> None:
        for item in self.ppt_tree.selection():
            self.ppt_tree.delete(item)
        self._refresh_ppt_state()
        self._preview_selected_ppt()

    def _clear_ppts(self) -> None:
        items = self.ppt_tree.get_children()
        if not items:
            return
        if not messagebox.askyesno("确认清空", f"确定清空列表中的 {len(items)} 个 PPT 文件吗？"):
            return
        for item in items:
            self.ppt_tree.delete(item)
        self._refresh_ppt_state()
        self._preview_selected_ppt()
        self._show_toast("列表已清空", "PPT 文件已从任务列表移除。", "info")

    def _remove_images(self) -> None:
        self._hide_image_hover_ui()
        for item in self.image_tree.selection():
            self._image_thumbnails.pop(item, None)
            self.image_tree.delete(item)
        self._renumber_images()
        self._refresh_image_state()
        self._preview_selected_image()

    def _clear_images(self) -> None:
        items = self.image_tree.get_children()
        if not items:
            return
        if not messagebox.askyesno("确认清空", f"确定清空列表中的 {len(items)} 张图片吗？"):
            return
        self._hide_image_hover_ui()
        for item in items:
            self._image_thumbnails.pop(item, None)
            self.image_tree.delete(item)
        self._renumber_images()
        self._refresh_image_state()
        self._preview_selected_image()
        self._show_toast("列表已清空", "图片已从任务列表移除。", "info")

    def _sort_images(self) -> None:
        items = list(self.image_tree.get_children())
        if len(items) < 2:
            self._show_toast("无需排序", "当前图片数量不足两张。", "info")
            return
        items.sort(key=lambda item: natural_sort_key(self.image_tree.item(item, "values")[1]))
        for index, item in enumerate(items):
            self.image_tree.move(item, "", index)
        self._renumber_images()
        self._set_status("图片已按文件名自然排序", "success")
        self._show_toast("排序完成", "已按数字与文件名的自然顺序排列。", "success")

    def _move_image(self, direction: int) -> None:
        selected = list(self.image_tree.selection())
        if not selected:
            return
        children = list(self.image_tree.get_children())
        indices = [children.index(item) for item in selected]
        for index, item in sorted(zip(indices, selected), reverse=direction > 0):
            target = index + direction
            if 0 <= target < len(children):
                self.image_tree.move(item, "", target)
                children = list(self.image_tree.get_children())
        self._renumber_images()
        for item in selected:
            self.image_tree.selection_add(item)

    def _image_drag_press(self, event) -> None:
        self._image_drag_item = self.image_tree.identify_row(event.y) or None
        self._image_drag_start_y = event.y

    def _image_drag_motion(self, event) -> None:
        if not self._image_drag_item or abs(event.y - self._image_drag_start_y) < 5:
            return
        target = self.image_tree.identify_row(event.y)
        if not target or target == self._image_drag_item:
            return
        index = self.image_tree.index(target)
        self.image_tree.move(self._image_drag_item, "", index)
        self._renumber_images()
        self.image_tree.selection_set(self._image_drag_item)

    def _image_drag_release(self, _event) -> None:
        if self._image_drag_item:
            self._set_status("已更新幻灯片顺序", "success")
        self._image_drag_item = None

    def _renumber_images(self) -> None:
        for index, item in enumerate(self.image_tree.get_children(), start=1):
            values = list(self.image_tree.item(item, "values"))
            values[0] = index
            self.image_tree.item(item, values=values)

    def _refresh_ppt_state(self) -> None:
        items = list(self.ppt_tree.get_children())
        total = 0
        for item in items:
            path = Path(self.ppt_tree.item(item, "values")[4])
            try:
                total += path.stat().st_size
            except OSError:
                pass
        self.ppt_count_var.set(f"{len(items)} 个文件 · {human_size(total)}" if items else "0 个文件")
        if items:
            self.ppt_empty.place_forget()
        else:
            self.ppt_empty.place(relx=0.5, rely=0.5, anchor="center")
        self._last_task_success["ppt"] = False
        self._refresh_action_state()

    def _refresh_image_state(self) -> None:
        items = list(self.image_tree.get_children())
        total = 0
        for item in items:
            path = Path(self.image_tree.item(item, "values")[5])
            try:
                total += path.stat().st_size
            except OSError:
                pass
        self.image_count_var.set(f"{len(items)} 张图片 · {human_size(total)}" if items else "0 张图片")
        if items:
            self.image_empty.place_forget()
        else:
            self.image_empty.place(relx=0.5, rely=0.5, anchor="center")
        self._last_task_success["images"] = False
        self._refresh_action_state()

    def _preview_selected_ppt(self, _event=None) -> None:
        selection = self.ppt_tree.selection()
        if not selection:
            return
        values = self.ppt_tree.item(selection[0], "values")
        self._set_status(f"已选择：{values[0]} · {values[2]}（双击打开所在文件夹）", "success")

    def _preview_selected_image(self, _event=None) -> None:
        selection = self.image_tree.selection()
        if not selection:
            return
        values = self.image_tree.item(selection[0], "values")
        self._set_status(f"已选择：第 {values[0]} 页 · {values[1]} · {values[2]}（双击查看大图）", "success")

    def _open_selected_source(self, mode: Mode) -> None:
        tree = self.ppt_tree if mode == "ppt" else self.image_tree
        selection = tree.selection()
        if not selection:
            return
        values = tree.item(selection[0], "values")
        path = Path(values[4] if mode == "ppt" else values[5])
        self._open_path(str(path.parent))

    # ---------- 设置与预设 ----------
    def _apply_ppt_preset(self, preset: str) -> None:
        if preset == "web":
            self.ppt_format_var.set("PNG")
            self.ppt_dpi_var.set(200)
        elif preset == "print":
            self.ppt_format_var.set("PNG")
            self.ppt_dpi_var.set(300)
        else:
            self.ppt_format_var.set("JPG")
            self.ppt_dpi_var.set(160)
            self.jpg_quality_var.set(90)
        for key, choice in self.ppt_presets.items():
            choice.set_selected(key == preset)
        self._set_status("已应用导出预设，可继续微调", "success")
        self._show_toast("预设已应用", {"web": "电商 PNG · 200 DPI", "print": "高清 PNG · 300 DPI", "light": "轻量 JPG · 160 DPI"}[preset], "success")

    def _apply_image_preset(self, preset: str) -> None:
        if preset == "wide":
            self.slide_size_var.set("16:9")
            self.fit_mode_var.set("完整显示（不裁切）")
        elif preset == "source":
            self.slide_size_var.set("按首图比例")
            self.fit_mode_var.set("完整显示（不裁切）")
        else:
            self.slide_size_var.set("16:9")
            self.fit_mode_var.set("铺满页面（居中裁切）")
        for key, choice in self.image_presets.items():
            choice.set_selected(key == preset)
        self._set_status("已应用 PPT 页面预设", "success")
        self._show_toast("预设已应用", {"wide": "16:9 完整显示", "source": "按首图比例完整显示", "cover": "16:9 居中裁切铺满"}[preset], "success")

    def _choose_ppt_output(self) -> None:
        path = filedialog.askdirectory(title="选择图片输出文件夹", initialdir=self.ppt_output_var.get() or None)
        if path:
            self.ppt_output_var.set(path)

    def _choose_pptx_output(self) -> None:
        current = self.pptx_output_var.get().strip()
        initialdir = str(Path(current).parent) if current else None
        initialfile = Path(current).name if current else "图片合成.pptx"
        path = filedialog.asksaveasfilename(title="保存 PPTX", defaultextension=".pptx", filetypes=(("PowerPoint 演示文稿", "*.pptx"),), initialfile=initialfile, initialdir=initialdir)
        if path:
            self.pptx_output_var.set(path)

    # ---------- 环境检测 ----------
    def _detect_environment_async(self) -> None:
        def worker() -> None:
            status = detect_export_backends()
            self._backend_status = status
            try:
                self.root.after(0, lambda: self._render_environment_status(status))
            except (RuntimeError, tk.TclError):
                pass

        threading.Thread(target=worker, daemon=True).start()

    def _render_environment_status(self, status: dict[str, object]) -> None:
        if status.get("powerpoint"):
            self.env_tag.set("PowerPoint 引擎可用", "success")
        elif status.get("libreoffice"):
            self.env_tag.set("LibreOffice 引擎可用", "success")
        else:
            self.env_tag.set("未检测到导出引擎", "warning")

    def _show_backend_status(self) -> None:
        status = self._backend_status or detect_export_backends()
        drag_text = "Windows 原生拖拽已启用" if getattr(self, "_drop_adapter", None) and self._drop_adapter.supported else "拖拽未启用，可使用选择按钮导入"
        lines = [
            f"Microsoft PowerPoint：{'已检测到' if status.get('powerpoint') else '未检测到'}",
            f"LibreOffice：{'已检测到' if status.get('libreoffice') else '未检测到'}",
            f"文件拖拽：{drag_text}",
            "",
            "自动选择会优先使用 PowerPoint，失败后再尝试 LibreOffice。",
        ]
        messagebox.showinfo("运行环境", "\n".join(lines))

    # ---------- 转换任务 ----------
    def _start_ppt_export(self) -> None:
        paths = [self.ppt_tree.item(item, "values")[4] for item in self.ppt_tree.get_children()]
        output = self.ppt_output_var.get().strip()
        if not paths:
            self._set_status("请先拖入或添加 PPT 文件", "warning")
            self._show_toast("缺少文件", "请先拖入或添加 PPT/PPTX 文件。", "warning")
            return
        if not output:
            self._show_toast("缺少输出位置", "请选择图片输出文件夹。", "warning")
            return

        backend_map = {"自动选择": "auto", "Microsoft PowerPoint": "powerpoint", "LibreOffice": "libreoffice"}
        options = PptExportOptions(
            output_dir=Path(output),
            image_format=self.ppt_format_var.get(),
            dpi=self.ppt_dpi_var.get(),
            jpg_quality=self.jpg_quality_var.get(),
            backend=backend_map.get(self.backend_var.get(), "auto"),
            separate_folder=self.separate_folder_var.get(),
        )
        if not self._begin_task(self.ppt_start_button, self.ppt_cancel_button, self.ppt_log):
            return

        def worker() -> None:
            try:
                outputs = export_presentations_to_images(paths, options, log=lambda message: self._append_log(self.ppt_log, message), progress=self._update_progress, should_cancel=self.cancel_event.is_set)
                self._task_success(f"已导出 {len(outputs)} 张图片", self.ppt_start_button, self.ppt_cancel_button, output_path=Path(output))
            except InterruptedError:
                self._task_cancelled(self.ppt_start_button, self.ppt_cancel_button)
            except Exception as exc:
                self._task_failed(str(exc), self.ppt_start_button, self.ppt_cancel_button, self.ppt_log)

        threading.Thread(target=worker, daemon=True).start()

    def _start_images_to_ppt(self) -> None:
        paths = [self.image_tree.item(item, "values")[5] for item in self.image_tree.get_children()]
        output = self.pptx_output_var.get().strip()
        if not paths:
            self._set_status("请先拖入或添加图片", "warning")
            self._show_toast("缺少图片", "请先拖入或添加图片。", "warning")
            return
        if not output:
            self._show_toast("缺少保存位置", "请选择 PPTX 保存位置。", "warning")
            return

        options = ImagesToPptOptions(
            output_path=Path(output),
            slide_size=self.slide_size_var.get(),
            fit_mode="cover" if self.fit_mode_var.get().startswith("铺满") else "contain",
            background="black" if self.background_var.get() == "黑色" else "white",
        )
        if not self._begin_task(self.image_start_button, self.image_cancel_button, self.image_log):
            return

        def worker() -> None:
            try:
                result = images_to_presentation(paths, options, log=lambda message: self._append_log(self.image_log, message), progress=self._update_progress, should_cancel=self.cancel_event.is_set)
                self._task_success(f"PPT 已生成：{result.name}", self.image_start_button, self.image_cancel_button, output_path=result)
            except InterruptedError:
                self._task_cancelled(self.image_start_button, self.image_cancel_button)
            except Exception as exc:
                self._task_failed(str(exc), self.image_start_button, self.image_cancel_button, self.image_log)

        threading.Thread(target=worker, daemon=True).start()

    def _begin_task(self, start_button: ttk.Button, cancel_button: ttk.Button, log_widget: tk.Text) -> bool:
        if self.busy:
            self._show_toast("任务进行中", "已有转换任务正在执行。", "warning")
            return False
        self.busy = True
        self.cancel_event.clear()
        start_button.configure(state="disabled")
        cancel_button.configure(state="normal")
        self.progress.configure(value=0, running=True)
        self._set_status("正在准备转换…", "busy")
        self._clear_log(log_widget)
        mode: Mode = "ppt" if start_button is self.ppt_start_button else "images"
        self._last_task_success[mode] = False
        self._update_steps(mode)
        self._refresh_action_state()
        return True

    def _cancel_task(self) -> None:
        if self.busy:
            self.cancel_event.set()
            self._set_status("正在取消任务…", "warning")

    def _task_success(self, message: str, start_button: ttk.Button, cancel_button: ttk.Button, output_path: Path | None = None) -> None:
        def finish() -> None:
            self.busy = False
            cancel_button.configure(state="disabled")
            self.progress.configure(value=100, running=False)
            self._set_status(message, "success")
            mode: Mode = "ppt" if start_button is self.ppt_start_button else "images"
            self._last_task_success[mode] = True
            self._refresh_action_state()
            self._update_steps(mode, finished=True)
            detail = message if not output_path else f"{message}\n{output_path}"
            self._show_toast("转换完成", detail, "success")

        self.root.after(0, finish)

    def _task_cancelled(self, start_button: ttk.Button, cancel_button: ttk.Button) -> None:
        def finish() -> None:
            self.busy = False
            cancel_button.configure(state="disabled")
            self.progress.configure(running=False)
            self._set_status("任务已取消", "warning")
            mode: Mode = "ppt" if start_button is self.ppt_start_button else "images"
            self._refresh_action_state()
            self._update_steps(mode)
            self._show_toast("任务已取消", "未完成的转换已停止。", "warning")

        self.root.after(0, finish)

    def _task_failed(self, message: str, start_button: ttk.Button, cancel_button: ttk.Button, log_widget: tk.Text) -> None:
        self._append_log(log_widget, f"错误：{message}")

        def finish() -> None:
            self.busy = False
            cancel_button.configure(state="disabled")
            self.progress.configure(running=False)
            self._set_status("处理失败，请查看处理记录", "error")
            mode: Mode = "ppt" if start_button is self.ppt_start_button else "images"
            self._refresh_action_state()
            self._update_steps(mode, error=True)
            self._show_toast("处理失败", message, "error")
            messagebox.showerror("处理失败", message)

        self.root.after(0, finish)

    def _update_progress(self, current: int, total: int, name: str) -> None:
        percent = 0 if total <= 0 else max(0, min(100, current / total * 100))

        def update() -> None:
            self.progress.configure(value=percent)
            self._set_status(f"正在处理：{name}（{current}/{total}）", "busy")

        self.root.after(0, update)

    def _set_status(self, message: str, state: str = "success") -> None:
        color = {"success": COLORS["success"], "warning": COLORS["warning"], "error": COLORS["danger"], "busy": COLORS["primary"]}.get(state, COLORS["success"])
        self.status_dot.configure(fg=color)
        self.status_var.set(message)

    def _append_log(self, widget: tk.Text, message: str) -> None:
        def append() -> None:
            widget.configure(state="normal")
            widget.insert("end", message + "\n")
            widget.see("end")
            widget.configure(state="disabled")

        self.root.after(0, append)

    @staticmethod
    def _clear_log(widget: tk.Text) -> None:
        widget.configure(state="normal")
        widget.delete("1.0", "end")
        widget.configure(state="disabled")

    # ---------- 系统与配置 ----------
    def _open_path(self, value: str) -> None:
        if not value:
            messagebox.showwarning("路径为空", "尚未设置输出路径。")
            return
        path = Path(value)
        if not path.exists():
            messagebox.showwarning("路径不存在", "该路径尚未生成，请先完成转换。")
            return
        try:
            if sys.platform.startswith("win"):
                os.startfile(path)  # type: ignore[attr-defined]
            elif sys.platform == "darwin":
                subprocess.Popen(["open", str(path)])
            else:
                subprocess.Popen(["xdg-open", str(path)])
        except Exception as exc:
            messagebox.showerror("无法打开", str(exc))

    def _open_pptx_parent(self) -> None:
        value = self.pptx_output_var.get().strip()
        if not value:
            messagebox.showwarning("路径为空", "尚未设置 PPT 保存位置。")
            return
        self._open_path(str(Path(value).parent))

    def _load_config(self) -> dict:
        try:
            if CONFIG_PATH.exists():
                return json.loads(CONFIG_PATH.read_text(encoding="utf-8"))
        except Exception:
            pass
        return {}

    def _restore_config(self) -> None:
        values = self.config
        self.ppt_output_var.set(values.get("ppt_output", ""))
        self.ppt_format_var.set(values.get("ppt_format", "PNG"))
        self.ppt_dpi_var.set(values.get("ppt_dpi", 200))
        self.jpg_quality_var.set(values.get("jpg_quality", 92))
        self.backend_var.set(values.get("backend", "自动选择"))
        self.separate_folder_var.set(values.get("separate_folder", True))
        self.pptx_output_var.set(values.get("pptx_output", ""))
        self.slide_size_var.set(values.get("slide_size", "16:9"))
        self.fit_mode_var.set(values.get("fit_mode", "完整显示（不裁切）"))
        self.background_var.set(values.get("background", "白色"))
        geometry = values.get("geometry")
        if isinstance(geometry, str) and "x" in geometry:
            try:
                self.root.geometry(geometry)
            except tk.TclError:
                pass
        # Sync preset-card selection without overwriting restored custom values.
        ppt_key = "light" if self.ppt_format_var.get() == "JPG" and self.ppt_dpi_var.get() <= 180 else ("print" if self.ppt_dpi_var.get() >= 280 else "web")
        for key, choice in self.ppt_presets.items():
            choice.set_selected(key == ppt_key)
        if self.slide_size_var.get() == "按首图比例":
            image_key = "source"
        elif self.fit_mode_var.get().startswith("铺满"):
            image_key = "cover"
        else:
            image_key = "wide"
        for key, choice in self.image_presets.items():
            choice.set_selected(key == image_key)
        self._refresh_ppt_state()
        self._refresh_image_state()
        self._preview_selected_image()
        self._refresh_action_state()

    def _save_config(self) -> None:
        data = {
            "ppt_output": self.ppt_output_var.get(),
            "ppt_format": self.ppt_format_var.get(),
            "ppt_dpi": self.ppt_dpi_var.get(),
            "jpg_quality": self.jpg_quality_var.get(),
            "backend": self.backend_var.get(),
            "separate_folder": self.separate_folder_var.get(),
            "pptx_output": self.pptx_output_var.get(),
            "slide_size": self.slide_size_var.get(),
            "fit_mode": self.fit_mode_var.get(),
            "background": self.background_var.get(),
            "active_mode": self.active_mode,
            "geometry": self.root.geometry(),
        }
        try:
            CONFIG_PATH.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
        except Exception:
            pass

    def _on_close(self) -> None:
        if self.busy and not messagebox.askyesno("确认退出", "任务正在执行，确定退出吗？"):
            return
        self.cancel_event.set()
        self._dismiss_toast()
        self._hide_image_hover_ui()
        if self._preview_dialog:
            try:
                self._preview_dialog.close()
            except Exception:
                pass
        self._save_config()
        try:
            self._drop_adapter.close()
        except Exception:
            pass
        self.root.destroy()


def main() -> None:
    root = TkinterDnD.Tk()
    PptImageConverterApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
