"""
Template library for status effect icon images.

Templates are stored as PNG files in the ``templates/`` directory alongside
this module.  A JSON sidecar (``templates.json``) maps each slug to a
human-readable display name and the screen region it was originally captured
from.
"""

import json
import os
import re

from PIL import Image

_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "templates")
_META_PATH = os.path.join(_DIR, "templates.json")


def _read_meta() -> dict:
    try:
        with open(_META_PATH, "r", encoding="utf-8") as f:
            return json.load(f)
    except (FileNotFoundError, json.JSONDecodeError):
        return {}


def _write_meta(meta: dict):
    os.makedirs(_DIR, exist_ok=True)
    with open(_META_PATH, "w", encoding="utf-8") as f:
        json.dump(meta, f, indent=2)


def _slugify(name: str) -> str:
    slug = re.sub(r"[^a-z0-9]+", "_", name.lower()).strip("_")
    return slug or "template"


def list_templates() -> list[dict]:
    """Return ``[{slug, name, region}, ...]`` sorted by name."""
    meta = _read_meta()
    items = []
    for slug, info in meta.items():
        png = os.path.join(_DIR, f"{slug}.png")
        if os.path.isfile(png):
            items.append({"slug": slug, "name": info.get("name", slug), "region": info.get("region")})
    items.sort(key=lambda t: t["name"].lower())
    return items


def save_template(name: str, pil_image: Image.Image, region: tuple | None = None) -> str:
    """Save *pil_image* as a new template and return its slug."""
    meta = _read_meta()
    slug = _slugify(name)
    base = slug
    counter = 2
    while slug in meta:
        slug = f"{base}_{counter}"
        counter += 1

    os.makedirs(_DIR, exist_ok=True)
    pil_image.save(os.path.join(_DIR, f"{slug}.png"))

    meta[slug] = {"name": name}
    if region is not None:
        meta[slug]["region"] = list(region)
    _write_meta(meta)
    return slug


def load_template(slug: str) -> Image.Image | None:
    """Load template PNG as a PIL RGB image, or ``None`` if missing."""
    path = os.path.join(_DIR, f"{slug}.png")
    if not os.path.isfile(path):
        return None
    return Image.open(path).convert("RGB")


def delete_template(slug: str):
    """Remove a template from disk and metadata."""
    meta = _read_meta()
    meta.pop(slug, None)
    _write_meta(meta)
    path = os.path.join(_DIR, f"{slug}.png")
    if os.path.isfile(path):
        os.remove(path)
