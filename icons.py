"""
Simple line-icon set drawn with PIL instead of relying on emoji glyphs,
which render inconsistently across Windows font configurations and don't
carry a consistent stroke weight/style the way a real icon set does.

Icons are drawn at 4x supersampling then downsampled with LANCZOS for
clean anti-aliased edges, and cached per (kind, color, size) so repeated
calls (e.g. rebuilding a page) don't redraw the same icon twice.
"""

from PIL import Image, ImageDraw
import customtkinter as ctk

_SUPERSAMPLE = 4
_CACHE = {}


def _canvas(size):
    dim = size * _SUPERSAMPLE
    img = Image.new("RGBA", (dim, dim), (0, 0, 0, 0))
    return img, ImageDraw.Draw(img), dim


def _finish(img, size):
    return img.resize((size, size), Image.LANCZOS)


def _draw_search(color, size):
    img, d, dim = _canvas(size)
    w = max(2, dim // 14)
    r = dim * 0.30
    cx, cy = dim * 0.42, dim * 0.42
    d.ellipse([cx - r, cy - r, cx + r, cy + r], outline=color, width=w)
    d.line([cx + r * 0.72, cy + r * 0.72, dim * 0.86, dim * 0.86], fill=color, width=w)
    return _finish(img, size)


def _draw_box(color, size):
    img, d, dim = _canvas(size)
    w = max(2, dim // 16)
    pad = dim * 0.14
    top = dim * 0.30
    d.line([pad, top, dim / 2, pad, dim - pad, top], fill=color, width=w, joint="curve")
    d.line([pad, top, pad, dim - pad, dim - pad, dim - pad, dim - pad, top], fill=color, width=w, joint="curve")
    d.line([pad, top, dim / 2, dim * 0.52, dim - pad, top], fill=color, width=w, joint="curve")
    d.line([dim / 2, dim * 0.52, dim / 2, dim - pad], fill=color, width=w)
    return _finish(img, size)


def _draw_truck(color, size):
    img, d, dim = _canvas(size)
    w = max(2, dim // 16)
    body = [dim * 0.10, dim * 0.35, dim * 0.62, dim * 0.68]
    d.rectangle(body, outline=color, width=w)
    cab = [dim * 0.62, dim * 0.46, dim * 0.90, dim * 0.68]
    d.line([cab[0], cab[1], cab[2], cab[1], cab[2], cab[3]], fill=color, width=w, joint="curve")
    for wx in (dim * 0.28, dim * 0.74):
        r = dim * 0.09
        d.ellipse([wx - r, dim * 0.68 - r, wx + r, dim * 0.68 + r], outline=color, width=w)
    return _finish(img, size)


def _draw_calendar(color, size):
    img, d, dim = _canvas(size)
    w = max(2, dim // 16)
    box = [dim * 0.14, dim * 0.22, dim * 0.86, dim * 0.86]
    d.rounded_rectangle(box, radius=dim * 0.06, outline=color, width=w)
    d.line([box[0], dim * 0.40, box[2], dim * 0.40], fill=color, width=w)
    for tx in (dim * 0.32, dim * 0.68):
        d.line([tx, dim * 0.12, tx, dim * 0.28], fill=color, width=w)
    return _finish(img, size)


def _draw_clock(color, size):
    img, d, dim = _canvas(size)
    w = max(2, dim // 14)
    cx, cy, r = dim * 0.5, dim * 0.5, dim * 0.36
    d.ellipse([cx - r, cy - r, cx + r, cy + r], outline=color, width=w)
    d.line([cx, cy, cx, cy - r * 0.55], fill=color, width=w)
    d.line([cx, cy, cx + r * 0.45, cy], fill=color, width=w)
    return _finish(img, size)


def _draw_doc_plus(color, size):
    img, d, dim = _canvas(size)
    w = max(2, dim // 16)
    doc = [dim * 0.22, dim * 0.12, dim * 0.68, dim * 0.88]
    fold = dim * 0.14
    d.line([
        doc[0], doc[1], doc[2] - fold, doc[1], doc[2], doc[1] + fold,
        doc[2], doc[3], doc[0], doc[3], doc[0], doc[1],
    ], fill=color, width=w, joint="curve")
    cx, cy, r = dim * 0.76, dim * 0.72, dim * 0.16
    d.line([cx - r, cy, cx + r, cy], fill=color, width=w)
    d.line([cx, cy - r, cx, cy + r], fill=color, width=w)
    return _finish(img, size)


def _draw_refresh(color, size):
    img, d, dim = _canvas(size)
    w = max(2, dim // 14)
    box = [dim * 0.16, dim * 0.16, dim * 0.84, dim * 0.84]
    d.arc(box, start=-30, end=200, fill=color, width=w)
    d.arc(box, start=150, end=380, fill=color, width=w)
    # arrowheads
    d.polygon([
        (dim * 0.84, dim * 0.30), (dim * 0.98, dim * 0.30), (dim * 0.91, dim * 0.16),
    ], fill=color)
    d.polygon([
        (dim * 0.16, dim * 0.70), (dim * 0.02, dim * 0.70), (dim * 0.09, dim * 0.84),
    ], fill=color)
    return _finish(img, size)


def _draw_list(color, size):
    img, d, dim = _canvas(size)
    w = max(2, dim // 14)
    r = dim * 0.045
    for y in (dim * 0.24, dim * 0.5, dim * 0.76):
        d.ellipse([dim * 0.12 - r, y - r, dim * 0.12 + r, y + r], fill=color)
        d.line([dim * 0.24, y, dim * 0.88, y], fill=color, width=w)
    return _finish(img, size)


def _draw_plug(color, size):
    img, d, dim = _canvas(size)
    w = max(2, dim // 14)
    d.line([dim * 0.5, dim * 0.10, dim * 0.5, dim * 0.40], fill=color, width=w)
    d.rounded_rectangle([dim * 0.30, dim * 0.36, dim * 0.70, dim * 0.62], radius=dim * 0.06, outline=color, width=w)
    d.line([dim * 0.5, dim * 0.62, dim * 0.5, dim * 0.78], fill=color, width=w)
    d.arc([dim * 0.30, dim * 0.66, dim * 0.70, dim * 0.94], start=20, end=160, fill=color, width=w)
    return _finish(img, size)


def _draw_inventory(color, size):
    img, d, dim = _canvas(size)
    w = max(2, dim // 14)
    bars = [
        (dim * 0.18, dim * 0.55, dim * 0.34),
        (dim * 0.42, dim * 0.35, dim * 0.34),
        (dim * 0.66, dim * 0.68, dim * 0.34),
    ]
    for x, top, height in bars:
        d.rectangle([x, dim - dim * 0.14 - height, x + dim * 0.16, dim - dim * 0.14], outline=color, width=w)
    return _finish(img, size)


def _draw_check(color, size):
    img, d, dim = _canvas(size)
    w = max(2, dim // 12)
    d.line([dim * 0.18, dim * 0.52, dim * 0.42, dim * 0.76], fill=color, width=w, joint="curve")
    d.line([dim * 0.42, dim * 0.76, dim * 0.84, dim * 0.26], fill=color, width=w, joint="curve")
    return _finish(img, size)


def _draw_warning(color, size):
    img, d, dim = _canvas(size)
    w = max(2, dim // 14)
    d.polygon([
        (dim * 0.5, dim * 0.12), (dim * 0.90, dim * 0.84), (dim * 0.10, dim * 0.84),
    ], outline=color, width=w)
    d.line([dim * 0.5, dim * 0.40, dim * 0.5, dim * 0.62], fill=color, width=w)
    r = dim * 0.03
    d.ellipse([dim * 0.5 - r, dim * 0.72 - r, dim * 0.5 + r, dim * 0.72 + r], fill=color)
    return _finish(img, size)


_DRAWERS = {
    "search": _draw_search,
    "box": _draw_box,
    "truck": _draw_truck,
    "calendar": _draw_calendar,
    "clock": _draw_clock,
    "doc_plus": _draw_doc_plus,
    "refresh": _draw_refresh,
    "list": _draw_list,
    "plug": _draw_plug,
    "inventory": _draw_inventory,
    "check": _draw_check,
    "warning": _draw_warning,
}


def make_icon(kind, color="#FFFFFF", size=20):
    """
    Returns a ctk.CTkImage for the given icon kind ('search', 'box',
    'truck', 'calendar', 'doc_plus', 'refresh', 'list', 'plug',
    'inventory'), drawn fresh at the requested size/color and cached.
    """

    key = (kind, color, size)

    if key in _CACHE:
        return _CACHE[key]

    drawer = _DRAWERS.get(kind)

    if drawer is None:
        return None

    pil_image = drawer(color, size)
    ctk_image = ctk.CTkImage(light_image=pil_image, dark_image=pil_image, size=(size, size))

    _CACHE[key] = ctk_image

    return ctk_image
