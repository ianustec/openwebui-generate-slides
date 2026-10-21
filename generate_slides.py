"""
title: Generate Slides
author: IANUSTEC
author_url: https://ianustec.com
funding_url: https://github.com/ianustec
description: Generate native, editable PowerPoint (.pptx) decks from JSON. v2 blocks format with 44 consulting components (KPIs, native charts incl. combo, layers, gantt, swot, matrix, waterfall, heatmap, RACI, journey...), 14 themes and auto-fit text; classic layouts and Template Mode kept. With an attached .pptx used as base/template, ALWAYS call inspect_slides first, then generate_slides with reference_file_id + per-slide reuse.
requirements: python-pptx, pillow
required_open_webui_version: 0.4.0
version: 2.0.0
license: MIT
"""

# ============================================================================
# Native PPTX Slides Engine — open-source OpenWebUI tool
# ----------------------------------------------------------------------------
# Server-side OpenWebUI tool `generate_slides`. The model emits a structured
# JSON spec ({title, subtitle, author, theme, slides:[{layout, ...}]}) and this
# engine renders a NATIVE .pptx with python-pptx: coherent visual system,
# layered decorative shapes, native charts, icons-in-circles, rich layouts.
#
# The deck is saved via the OpenWebUI Files API (with a /cache/files fallback)
# and a clickable download link is emitted in chat.
#
# License: MIT — Copyright (c) IANUSTEC.
# ============================================================================

import inspect
import logging
import time
import mimetypes
import os
import re
import json
import uuid
import base64
import unicodedata
from zipfile import ZipFile
from datetime import datetime, timezone
from copy import deepcopy
from dataclasses import dataclass, field, replace
from io import BytesIO
from pathlib import Path
from typing import Any, Iterator, Optional

from pydantic import BaseModel, Field

# --- python-pptx (rendering engine) -----------------------------------------
try:
    from pptx import Presentation  # type: ignore
    from pptx.util import Inches, Pt  # type: ignore
    from pptx.dml.color import RGBColor  # type: ignore
    from pptx.enum.text import PP_ALIGN, MSO_ANCHOR  # type: ignore
    from pptx.enum.shapes import MSO_SHAPE, MSO_SHAPE_TYPE  # type: ignore
    from pptx.enum.dml import MSO_FILL_TYPE  # type: ignore
    from pptx.oxml.ns import qn  # type: ignore
    from pptx.chart.data import CategoryChartData  # type: ignore
    from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION, XL_LABEL_POSITION  # type: ignore

    _HAS_PPTX = True
except Exception:  # pragma: no cover
    _HAS_PPTX = False

# --- Pillow (icons + image resize) ------------------------------------------
try:
    from PIL import Image  # type: ignore

    _HAS_PIL = True
except Exception:
    _HAS_PIL = False

# --- httpx (image fetch) ----------------------------------------------------
try:
    import httpx  # type: ignore

    _HAS_HTTPX = True
except Exception:
    _HAS_HTTPX = False

# --- OpenWebUI Files API (native download chip) -----------------------------
try:
    from fastapi import UploadFile  # type: ignore
    from starlette.datastructures import Headers  # type: ignore
    from open_webui.routers.files import upload_file_handler  # type: ignore
    from open_webui.models.users import Users  # type: ignore

    _HAS_OWUI_FILES = True
except Exception:
    _HAS_OWUI_FILES = False

# --- OpenWebUI file download (reference template) ---------------------------
_HAS_OWUI_FILE_DOWNLOAD = False
_OwuiFiles = None
_OwuiStorage = None
try:
    from open_webui.models.files import Files as _OwuiFiles  # type: ignore
    from open_webui.storage.provider import Storage as _OwuiStorage  # type: ignore

    _HAS_OWUI_FILE_DOWNLOAD = True
except Exception:
    pass

_REFERENCE_PPTX_MAX_BYTES = 25 * 1024 * 1024
_INSPECT_MEDIA_WARN_BYTES = 50 * 1024 * 1024
_INSPECT_MEDIA_WARN_COUNT = 200
_PPTX_EXTENSIONS = (".pptx", ".potx")

# Presentation edit — `operations[].slide` in JSON (not template reuse / inspect).
# 0 = 0-based (first slide is 0; matches inspect slides[].index and python-pptx).
# 1 = 1-based (first slide is 1; matches PowerPoint / LibreOffice slide labels).
EDIT_SLIDE_INDEX_ORIGIN = 1

# Open WebUI Files API ids are UUIDs. Anything else (a bare number, a
# filename, a slide index) was invented by the model and must be rejected
# before download, with the real attachment id when the chat has one.
_FILES_API_ID_RE = re.compile(
    r"^[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}$"
)
_PPTX_MIMES = (
    "application/vnd.openxmlformats-officedocument.presentationml.presentation",
    "application/vnd.ms-powerpoint",
)

# --- OpenWebUI AI image generation (optional) -------------------------------
try:
    from open_webui.routers.images import image_generations as _owui_image_generations  # type: ignore
    from open_webui.routers.images import GenerateImageForm as _OwuiImageForm  # type: ignore

    _HAS_OWUI_IMAGES = True
except Exception:
    _HAS_OWUI_IMAGES = False

log = logging.getLogger(__name__)


async def _maybe_await(value):
    """Await coroutines; pass through sync return values (OWUI v0.6 vs v0.11)."""
    if inspect.isawaitable(value):
        return await value
    return value


def _extract_file_id(item) -> Optional[str]:
    """FileModel.id or dict['id'] from upload_file_handler."""
    if not item:
        return None
    fid = getattr(item, "id", None)
    if fid:
        return str(fid)
    if isinstance(item, dict):
        fid = item.get("id")
        return str(fid) if fid else None
    return None


async def _upload_bytes_to_files_api(
    data: bytes,
    *,
    filename: str,
    content_type: str,
    request,
    user_dict,
    export_dir: str,
    log_prefix: str = "generate_slides",
) -> tuple[Optional[str], Optional[str], Optional[str]]:
    """Upload bytes via OWUI Files API or cache fallback. Returns (url, error, file_id)."""
    user_id = (user_dict or {}).get("id") if isinstance(user_dict, dict) else None
    if _HAS_OWUI_FILES and request is not None and user_id:
        try:
            user_model = await _maybe_await(Users.get_user_by_id(user_id))
            if user_model:
                buf = BytesIO(data)
                buf.seek(0)
                upload = UploadFile(
                    file=buf,
                    filename=filename,
                    headers=Headers({"content-type": content_type}),
                )
                item = await _maybe_await(
                    upload_file_handler(
                        request=request,
                        file=upload,
                        metadata={},
                        process=False,
                        user=user_model,
                    )
                )
                fid = _extract_file_id(item)
                if fid:
                    return f"/api/v1/files/{fid}/content", None, fid
                log.warning(
                    "[%s] Files API returned no file id (type=%s)",
                    log_prefix,
                    type(item).__name__,
                )
            else:
                log.warning("[%s] user not found for Files API: %s", log_prefix, user_id)
        except Exception:
            log.exception("[%s] Files API upload failed", log_prefix)
    elif not _HAS_OWUI_FILES:
        log.debug("[%s] Open WebUI Files API unavailable", log_prefix)
    elif request is None or not user_id:
        log.debug("[%s] missing request or user id for Files API", log_prefix)

    export_dir = (export_dir or "").strip() or "/app/backend/data/cache/files"
    try:
        os.makedirs(export_dir, mode=0o775, exist_ok=True)
        path = os.path.join(export_dir, filename)
        with open(path, "wb") as fh:
            fh.write(data)
        if os.path.isfile(path) and os.path.getsize(path) > 0:
            log.warning("[%s] saved via cache fallback: %s", log_prefix, filename)
            return f"/cache/files/{filename}", None, None
    except Exception as exc:
        return None, str(exc), None
    return None, "impossibile salvare", None


# ============================================================================
# Geometry (LAYOUT_WIDE 16:9)
# ============================================================================
EMU_IN = 914400
SLIDE_W_IN = 13.333
SLIDE_H_IN = 7.5
MARGIN = 0.62          # outer margin (in)
CONTENT_TOP = 1.9      # y where content starts on standard slides (in)
FOOTER_Y = 7.02        # footer band y (in)
BODY_FONT = "Calibri"
HEAD_FONT = "Georgia"  # serif display for a premium, editorial feel


# ============================================================================
# Colour utilities
# ============================================================================

def _hex(value: str) -> str:
    """Return a clean 6-char uppercase hex (no #)."""
    if not isinstance(value, str):
        return "1E2761"
    h = value.strip().lstrip("#")
    if len(h) == 3:
        h = "".join(c * 2 for c in h)
    if len(h) != 6:
        return "1E2761"
    try:
        int(h, 16)
        return h.upper()
    except ValueError:
        return "1E2761"


def _to_rgb(value: str) -> tuple[int, int, int]:
    h = _hex(value)
    return int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16)


def _from_rgb(r: int, g: int, b: int) -> str:
    def clamp(x):
        return max(0, min(255, int(round(x))))
    return f"{clamp(r):02X}{clamp(g):02X}{clamp(b):02X}"


def _mix(a: str, b: str, t: float) -> str:
    ra, ga, ba = _to_rgb(a)
    rb, gb, bb = _to_rgb(b)
    t = max(0.0, min(1.0, t))
    return _from_rgb(ra + (rb - ra) * t, ga + (gb - ga) * t, ba + (bb - ba) * t)


def _darken(c: str, t: float) -> str:
    return _mix(c, "000000", t)


def _lighten(c: str, t: float) -> str:
    return _mix(c, "FFFFFF", t)


def _luminance(c: str) -> float:
    def ch(x: float) -> float:
        x /= 255.0
        return x / 12.92 if x <= 0.03928 else ((x + 0.055) / 1.055) ** 2.4

    r, g, b = _to_rgb(c)
    return 0.2126 * ch(r) + 0.7152 * ch(g) + 0.0722 * ch(b)


def _on(bg: str) -> str:
    """Readable ink (near-black or white) for text on ``bg``."""
    return "1A1A1A" if _luminance(bg) > 0.42 else "FFFFFF"


def _rgb(value: str) -> "RGBColor":
    return RGBColor(*_to_rgb(value))


# ============================================================================
# Named palettes (curated theme presets)
# ============================================================================
_PALETTES: dict[str, dict[str, str]] = {
    "midnight":   {"primary": "1E2761", "dark": "141A45", "accent": "C99A3B", "light_accent": "CADCFC"},
    "forest":     {"primary": "2C5F2D", "dark": "1A3D1B", "accent": "97BC62", "light_accent": "B7D48A"},
    "coral":      {"primary": "2F3C7E", "dark": "1E2A5C", "accent": "F96167", "light_accent": "F9E795"},
    "terracotta": {"primary": "B85042", "dark": "8A3328", "accent": "C9A227", "light_accent": "A7BEAE"},
    "ocean":      {"primary": "21295C", "dark": "0A0F33", "accent": "1C7293", "light_accent": "7FBADC"},
    "charcoal":   {"primary": "2B3640", "dark": "1A2128", "accent": "C08A2D", "light_accent": "9FB0BB"},
    "teal":       {"primary": "013F46", "dark": "012A2F", "accent": "02C39A", "light_accent": "7FE3CD"},
    "berry":      {"primary": "6D2E46", "dark": "3F1A28", "accent": "C99A3B", "light_accent": "D6A5A6"},
    "sage":       {"primary": "2E4D58", "dark": "1F343C", "accent": "84B59F", "light_accent": "A4C7BD"},
    "cherry":     {"primary": "2F3C7E", "dark": "1A2356", "accent": "C0392B", "light_accent": "E8636E"},
    "slate":      {"primary": "27303F", "dark": "171C26", "accent": "3B82F6", "light_accent": "93C5FD"},
    # legacy aliases
    "blue":       {"primary": "1E2761", "dark": "141A45", "accent": "C99A3B", "light_accent": "CADCFC"},
    "green":      {"primary": "2C5F2D", "dark": "1A3D1B", "accent": "97BC62", "light_accent": "B7D48A"},
    "dark":       {"primary": "1F2937", "dark": "111827", "accent": "60A5FA", "light_accent": "93C5FD"},
}
_DEFAULT_PALETTE = "midnight"

_THEME_KEYWORDS: dict[str, tuple[str, ...]] = {
    "forest": ("sostenib", "ambient", "green", "eco", "natur", "agricol", "environment"),
    "ocean": ("data", "analy", "research", "scienz", "ricerca", "dati", "tech"),
    "coral": ("creativ", "design", "marketing", "brand"),
    "terracotta": ("food", "hospitality", "travel", "ristorante", "cucina", "viaggio"),
    "teal": ("health", "medic", "wellness", "fintech", "salute", "benessere"),
    "berry": ("luxury", "fashion", "retail", "lusso", "moda"),
    "sage": ("yoga", "mindfulness", "spa"),
    "cherry": ("urgent", "alert", "energy", "urgente", "energia"),
    "charcoal": ("legal", "consult", "professional", "corporate", "legale", "avvocat"),
    "midnight": ("finance", "banking", "executive", "enterprise", "business", "strategy",
                 "finanz", "strategia", "startup", "invest", "plan"),
}


def _pick_palette_name(requested: Optional[str], title: str, slides: list) -> str:
    if requested:
        r = str(requested).strip().lower()
        if r in _PALETTES:
            return r
        if r not in ("auto", ""):
            return _DEFAULT_PALETTE
    hay = " ".join([title or ""] + [str(s.get("title", "")) for s in slides]).lower()
    for name, kws in _THEME_KEYWORDS.items():
        if any(k in hay for k in kws):
            return name
    return _DEFAULT_PALETTE


def _resolve_theme(spec: dict, slides: list) -> dict:
    """Build the full design-token palette for the deck."""
    name = _pick_palette_name(spec.get("theme"), spec.get("title", ""), slides)
    base = dict(_PALETTES.get(name, _PALETTES[_DEFAULT_PALETTE]))

    # Explicit overrides from the spec (palette{} or accent).
    pal = spec.get("palette") if isinstance(spec.get("palette"), dict) else {}
    for k in ("primary", "dark", "accent", "light_accent"):
        if pal.get(k):
            base[k] = _hex(pal[k])
    if spec.get("accent"):
        base["accent"] = _hex(spec["accent"])
    if spec.get("primary"):
        base["primary"] = _hex(spec["primary"])

    primary = _hex(base["primary"])
    dark = _hex(base.get("dark") or _darken(primary, 0.35))
    accent = _hex(base["accent"])

    theme = {
        "name": name,
        # dark surfaces (cover / section / closing)
        "bg_dark": dark,
        "bg_dark2": _lighten(dark, 0.06),
        "panel_dark": _lighten(dark, 0.10),
        "on_dark": "FFFFFF",
        "on_dark_soft": _mix(dark, "FFFFFF", 0.72),
        "on_dark_faint": _mix(dark, "FFFFFF", 0.50),
        # light surfaces (content)
        "bg_light": "FFFFFF",
        "bg_soft": _lighten(primary, 0.955),
        "ink": _darken(primary, 0.15) if _luminance(primary) > 0.12 else "1A1A2E",
        "muted": _mix(primary, "6B7280", 0.5),
        "faint": "9AA0AB",
        # brand
        "primary": primary,
        "accent": accent,
        "accent_soft": _lighten(accent, 0.82),
        "light_accent": _hex(base.get("light_accent") or _lighten(accent, 0.5)),
        # cards
        "card_bg": "FFFFFF",
        "card_border": _lighten(primary, 0.86),
        "card_soft": _lighten(primary, 0.955),
        "hairline": "E4E7EC",
        # fonts
        "head_font": _hex_font(spec.get("heading_font")) or HEAD_FONT,
        "body_font": _hex_font(spec.get("body_font")) or BODY_FONT,
    }
    return theme


def _hex_font(v):
    if isinstance(v, str) and v.strip():
        return v.strip()
    return None


# ============================================================================
# Layout vocabulary (canonical layout names + aliases)
# ============================================================================
_LAYOUT_ALIASES: dict[str, str] = {
    "title": "cover", "copertina": "cover",
    "content": "content",
    "image": "image_full_caption", "full_image": "image_full_caption",
    "hero_image": "image_full_caption",
    "stat": "kpi_row", "kpi": "kpi_row", "kpis": "kpi_row", "stats": "kpi_row",
    "metrics": "kpi_row", "big_number": "kpi_row", "bignumber": "kpi_row",
    "timeline": "timeline_horizontal", "tempistica": "timeline_horizontal",
    "process": "process_flow", "flow": "process_flow",
    "comparison": "comparison_two", "compare": "comparison_two", "vs": "comparison_two",
    "confronto": "comparison_two", "tabella": "table",
    "conclusion": "closing", "chiusura": "closing", "fine": "closing",
    "icon-list": "icon_list_vertical", "icon_list": "icon_list_vertical",
    "iconlist": "icon_list_vertical", "vertical_list": "icon_list_vertical",
    "icon-grid": "icon_grid_3", "icon_grid": "icon_grid_3",
    "image-grid": "image_grid", "gallery": "image_grid",
    "two_column": "two_column_text", "two_col": "two_column_text",
    "due_colonne": "two_column_text",
    "image_right": "text_image_right", "image_left": "image_left_text_right",
    "bullet_list": "title_bullets", "bulletlist": "title_bullets",
    "bullets": "title_bullets", "bullet-list": "title_bullets",
    "list": "title_bullets", "elenco": "title_bullets", "punti": "title_bullets",
    "highlights": "title_bullets", "key_points": "title_bullets",
    "keypoints": "title_bullets", "ranking": "title_bullets",
    "sezione": "section", "titolo": "title_only",
    "citazione": "quote", "grafico": "chart", "vuoto": "blank",
    "alert": "alert", "callout": "alert", "note": "alert", "warning": "alert",
    "info": "alert", "tip": "alert", "danger": "alert",
    "scheda": "title_body", "paragrafo": "title_body",
    # diagrams
    "diagram": "diagram", "diagramma": "diagram", "schema": "diagram",
    "imbuto": "funnel", "piramide": "pyramid", "ciclo": "cycle",
    "occhio_di_bue": "bullseye", "pilastri": "pillars", "pillar": "pillars",
    "strada": "roadmap", "road": "roadmap", "matrix": "quadrant",
    "matrice": "quadrant", "quadrante": "quadrant", "venn_diagram": "venn",
}

_DIAGRAM_LAYOUTS = ("funnel", "pyramid", "cycle", "venn", "quadrant",
                    "bullseye", "pillars", "iceberg", "roadmap")

_LIST_FIELD_ALIASES = (
    "bullets", "bullet_points", "bullet_list", "points", "items", "key_points",
    "highlights", "features", "benefits", "challenges", "problems", "risks",
    "advantages", "pros", "cons", "steps", "actions", "tasks", "takeaways",
    "conclusions", "recommendations", "details", "notes", "facts", "objectives",
    "goals", "results", "reasons", "examples", "entries", "lines",
)
_TEXT_FIELD_ALIASES = (
    "body", "content", "text", "description", "summary", "paragraph",
    "detail", "note", "explanation", "overview", "lead", "intro",
)


# ============================================================================
# Text / helpers
# ============================================================================
_MD_STRONG = re.compile(r"\*\*(.+?)\*\*|__(.+?)__")
_MD_EM = re.compile(r"(?<!\*)\*(?!\*)(.+?)(?<!\*)\*(?!\*)|(?<!_)_(?!_)(.+?)(?<!_)_(?!_)")
_MARKER_RE = re.compile(r"^\s*(?:[\u2022\u25E6\u2013\u2014\-\*\u2611\u2610\u2713\u2714\u2192>]\s+)")


def _smart(text: str) -> str:
    if not text:
        return ""
    return str(text).replace("--", "\u2013")


def _strip_md(text: str) -> str:
    """Remove **bold**/*em*/`code` markers, returning plain text."""
    if not text:
        return ""
    t = str(text)
    t = re.sub(r"\*\*(.+?)\*\*", r"\1", t)
    t = re.sub(r"__(.+?)__", r"\1", t)
    t = re.sub(r"`(.+?)`", r"\1", t)
    t = re.sub(r"(?<!\*)\*(?!\*)(.+?)\*(?!\*)", r"\1", t)
    return t.strip()


def _clean_bullet(text: str) -> str:
    t = _strip_md(str(text or ""))
    t = _MARKER_RE.sub("", t)
    return t.strip()


def _md_runs(text: str) -> list[tuple[str, bool, bool]]:
    """Split text into (chunk, bold, italic) runs from lightweight markdown."""
    if text is None:
        return [("", False, False)]
    s = str(text)
    runs: list[tuple[str, bool, bool]] = []
    # handle **bold** first via tokenization
    pattern = re.compile(r"\*\*(.+?)\*\*|`(.+?)`")
    pos = 0
    for m in pattern.finditer(s):
        if m.start() > pos:
            runs.append((s[pos:m.start()], False, False))
        if m.group(1) is not None:
            runs.append((m.group(1), True, False))
        else:
            runs.append((m.group(2), False, True))
        pos = m.end()
    if pos < len(s):
        runs.append((s[pos:], False, False))
    return runs or [("", False, False)]


def _slugify(text: str, *, max_len: int = 60) -> str:
    t = unicodedata.normalize("NFKD", text or "")
    t = "".join(c for c in t if not unicodedata.combining(c))
    t = t.lower()
    t = re.sub(r"[^a-z0-9]+", "-", t)
    t = re.sub(r"-{2,}", "-", t).strip("-")
    return (t or "presentation")[:max_len]


def _as_list(val) -> list:
    if val is None:
        return []
    if isinstance(val, list):
        return val
    if isinstance(val, str):
        parts = [p.strip() for p in val.split("\n") if p.strip()]
        return parts or [val]
    return [val]


def _first(d: dict, *keys, default=None):
    for k in keys:
        v = d.get(k)
        if v not in (None, "", [], {}):
            return v
    return default


def _harvest_bullets(slide: dict) -> list[str]:
    """Collect bullet strings from the many alias fields the LLM may use."""
    for key in _LIST_FIELD_ALIASES:
        val = slide.get(key)
        if isinstance(val, list) and val:
            out = []
            for item in val:
                if isinstance(item, dict):
                    t = _first(item, "title", "label", "heading", "name", "text", default="")
                    d = _first(item, "description", "desc", "detail", "body", "text", default="")
                    line = t
                    if d and d != t:
                        line = f"**{t}** — {d}" if t else d
                    if line:
                        out.append(str(line))
                elif item not in (None, ""):
                    out.append(str(item))
            if out:
                return out
        if isinstance(val, str) and val.strip():
            lines = [ln.strip() for ln in val.split("\n") if ln.strip()]
            if lines:
                return lines
    # text fields → split into paragraphs/bullets
    for key in _TEXT_FIELD_ALIASES:
        val = slide.get(key)
        if isinstance(val, str) and val.strip():
            lines = [ln.strip() for ln in val.split("\n") if ln.strip()]
            if len(lines) > 1:
                return lines
    return []


def _resolve_layout(slide: dict) -> str:
    raw = str(_first(slide, "layout", "type", default="content")).strip().lower()
    return _LAYOUT_ALIASES.get(raw, raw)


# ============================================================================
# Low-level PPTX drawing helpers
# ============================================================================

def _set_bg(slide, color: str) -> None:
    slide.background.fill.solid()
    slide.background.fill.fore_color.rgb = _rgb(color)


def _maybe_set_bg(deck, slide, color: str) -> None:
    if not deck.template_mode:
        _set_bg(slide, color)


def _no_line(shape) -> None:
    shape.line.fill.background()


def _set_shadow(shape, *, blur=8, dist=3, alpha=78000, direction=5400000) -> None:
    """Apply a soft outer shadow to a shape via OOXML."""
    try:
        spPr = shape._element.spPr
        # remove existing effectLst
        for el in spPr.findall(qn("a:effectLst")):
            spPr.remove(el)
        effect = spPr.makeelement(qn("a:effectLst"), {})
        shadow = effect.makeelement(qn("a:outerShdw"), {
            "blurRad": str(int(blur * 12700)),
            "dist": str(int(dist * 12700)),
            "dir": str(int(direction)),
            "rotWithShape": "0",
        })
        clr = shadow.makeelement(qn("a:srgbClr"), {"val": "1A1A2E"})
        a = clr.makeelement(qn("a:alpha"), {"val": str(int(alpha))})
        clr.append(a)
        shadow.append(clr)
        effect.append(shadow)
        spPr.append(effect)
    except Exception:
        pass


def _fill_alpha(shape, color: str, alpha_pct: float) -> None:
    """Solid fill with transparency (alpha_pct 0..100 = opacity)."""
    shape.fill.solid()
    shape.fill.fore_color.rgb = _rgb(color)
    try:
        sf = shape.fill.fore_color._xFill.find(qn("a:srgbClr"))
        if sf is not None:
            for el in sf.findall(qn("a:alpha")):
                sf.remove(el)
            a = sf.makeelement(qn("a:alpha"), {"val": str(int(alpha_pct * 1000))})
            sf.append(a)
    except Exception:
        pass


def _card(slide, x, y, w, h, *, fill="FFFFFF", radius=0.09, line=None,
          line_w=1.0, shadow=True):
    shp = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,
                                 Inches(x), Inches(y), Inches(w), Inches(h))
    try:
        shp.adjustments[0] = radius
    except Exception:
        pass
    shp.fill.solid()
    shp.fill.fore_color.rgb = _rgb(fill)
    if line:
        shp.line.color.rgb = _rgb(line)
        shp.line.width = Pt(line_w)
    else:
        _no_line(shp)
    if shadow:
        _set_shadow(shp, blur=10, dist=3, alpha=62000)
    shp.shadow.inherit = False
    return shp


def _rect(slide, x, y, w, h, *, fill, alpha=None):
    shp = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE,
                                 Inches(x), Inches(y), Inches(w), Inches(h))
    if alpha is not None:
        _fill_alpha(shp, fill, alpha)
    else:
        shp.fill.solid()
        shp.fill.fore_color.rgb = _rgb(fill)
    _no_line(shp)
    shp.shadow.inherit = False
    return shp


def _oval(slide, x, y, d, *, fill, alpha=None, line=None, line_w=1.0):
    shp = slide.shapes.add_shape(MSO_SHAPE.OVAL,
                                 Inches(x), Inches(y), Inches(d), Inches(d))
    if alpha is not None:
        _fill_alpha(shp, fill, alpha)
    elif fill is None:
        shp.fill.background()
    else:
        shp.fill.solid()
        shp.fill.fore_color.rgb = _rgb(fill)
    if line:
        shp.line.color.rgb = _rgb(line)
        shp.line.width = Pt(line_w)
    else:
        _no_line(shp)
    shp.shadow.inherit = False
    return shp


def _set_tracking(run, pts: float) -> None:
    """Letter-spacing in points (OOXML spc is in 1/100 pt)."""
    try:
        rPr = run._r.get_or_add_rPr()
        rPr.set("spc", str(int(pts * 100)))
    except Exception:
        pass


def _textbox(slide, x, y, w, h, *, anchor="top", wrap=True):
    tb = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = tb.text_frame
    tf.word_wrap = wrap
    tf.margin_left = 0
    tf.margin_right = 0
    tf.margin_top = 0
    tf.margin_bottom = 0
    tf.vertical_anchor = {
        "top": MSO_ANCHOR.TOP, "middle": MSO_ANCHOR.MIDDLE, "bottom": MSO_ANCHOR.BOTTOM,
    }.get(anchor, MSO_ANCHOR.TOP)
    return tb, tf


def _style_para(p, *, size=None, color=None, bold=None, italic=None, font=None,
                align=None, space_after=None, space_before=None, line=None,
                tracking=None):
    if align is not None:
        p.alignment = {"left": PP_ALIGN.LEFT, "center": PP_ALIGN.CENTER,
                       "right": PP_ALIGN.RIGHT, "justify": PP_ALIGN.JUSTIFY}.get(align, PP_ALIGN.LEFT)
    if space_after is not None:
        p.space_after = Pt(space_after)
    if space_before is not None:
        p.space_before = Pt(space_before)
    if line is not None:
        p.line_spacing = line
    for r in p.runs:
        if size is not None:
            r.font.size = Pt(size)
        if color is not None:
            r.font.color.rgb = _rgb(color)
        if bold is not None:
            r.font.bold = bold
        if italic is not None:
            r.font.italic = italic
        if font is not None:
            r.font.name = font
        if tracking is not None:
            _set_tracking(r, tracking)


def _add_para(tf, text, *, first=False, **style):
    p = tf.paragraphs[0] if (first and not tf.paragraphs[0].runs) else tf.add_paragraph()
    if isinstance(text, list):
        for chunk, b, it in text:
            if chunk == "":
                continue
            r = p.add_run()
            r.text = _smart(chunk)
    else:
        r = p.add_run()
        r.text = _smart(text)
    _style_para(p, **style)
    # apply bold/italic per-run for markdown runs
    if isinstance(text, list):
        runs = p.runs
        idx = 0
        for chunk, b, it in text:
            if chunk == "":
                continue
            if idx < len(runs):
                if b:
                    runs[idx].font.bold = True
                if it:
                    runs[idx].font.italic = True
                idx += 1
    return p


def _add_bullet(tf, text, *, first=False, size=13.5, color="333333",
                accent="1E2761", font=BODY_FONT, space_after=7, indent=0.24,
                line=1.12):
    """Add a real PPTX bullet paragraph (buChar dot in accent colour)."""
    p = tf.paragraphs[0] if (first and not tf.paragraphs[0].runs) else tf.add_paragraph()
    for chunk, b, it in _md_runs(str(text)):
        if chunk == "":
            continue
        r = p.add_run()
        r.text = _smart(chunk)
        r.font.size = Pt(size)
        r.font.name = font
        r.font.color.rgb = _rgb(color)
        if b:
            r.font.bold = True
        if it:
            r.font.italic = True
    p.space_after = Pt(space_after)
    p.line_spacing = line
    # bullet formatting via pPr
    pPr = p._pPr if p._pPr is not None else p.get_or_add_pPr()
    pPr.set("marL", str(int(indent * EMU_IN)))
    pPr.set("indent", str(-int(0.18 * EMU_IN)))
    # remove existing bullet defs
    for tag in ("a:buNone", "a:buChar", "a:buAutoNum", "a:buClr", "a:buSzPct", "a:buFont"):
        for el in pPr.findall(qn(tag)):
            pPr.remove(el)
    buClr = pPr.makeelement(qn("a:buClr"), {})
    srgb = buClr.makeelement(qn("a:srgbClr"), {"val": _hex(accent)})
    buClr.append(srgb)
    buSz = pPr.makeelement(qn("a:buSzPct"), {"val": "90000"})
    buFont = pPr.makeelement(qn("a:buFont"), {"typeface": "Arial"})
    buChar = pPr.makeelement(qn("a:buChar"), {"char": "\u25CF"})
    pPr.append(buClr)
    pPr.append(buSz)
    pPr.append(buFont)
    pPr.append(buChar)
    return p


# ============================================================================
# Page chrome (eyebrow, title, footer) — shared by content slides
# ============================================================================

def _eyebrow(slide, theme, text, *, x=MARGIN, y=0.62, w=None, dark=False):
    if not text:
        return
    w = w or (SLIDE_W_IN - 2 * MARGIN)
    tb, tf = _textbox(slide, x, y, w, 0.32)
    color = theme["accent"] if not dark else theme["accent"]
    _add_para(tf, _strip_md(str(text)).upper(), first=True, size=11,
              color=color, bold=True, font=theme["body_font"],
              align="left", tracking=2.2, space_after=0)


def _title(slide, theme, text, *, x=MARGIN, y=0.98, w=None, size=30, dark=False,
           color=None):
    w = w or (SLIDE_W_IN - 2 * MARGIN)
    tb, tf = _textbox(slide, x, y, w, 0.95, anchor="top")
    c = color or (theme["on_dark"] if dark else theme["ink"])
    _add_para(tf, _strip_md(str(text)), first=True, size=size, color=c,
              bold=True, font=theme["head_font"], align="left",
              line=1.02, space_after=0)
    return tb


def _footer(slide, theme, *, page=None, label="", dark=False, y=None, x=None, w=None):
    color = theme["on_dark_faint"] if dark else theme["faint"]
    fx = MARGIN if x is None else x
    fy = FOOTER_Y if y is None else y
    fw = (SLIDE_W_IN - 2 * MARGIN) if w is None else w
    # left label
    tb, tf = _textbox(slide, fx, fy, fw - 0.6, 0.3, anchor="middle")
    _add_para(tf, _strip_md(label).upper() if label else "", first=True, size=8,
              color=color, bold=False, font=theme["body_font"], align="left",
              tracking=1.2, space_after=0)
    if page is not None:
        tb2, tf2 = _textbox(slide, fx + fw - 0.8, fy, 0.8, 0.3, anchor="middle")
        _add_para(tf2, str(page), first=True, size=9, color=color,
                  font=theme["body_font"], align="right", space_after=0)


# ============================================================================
# Icon system (icon-in-circle motif)
# ----------------------------------------------------------------------------
# Baked monochrome PNGs (white glyph on transparent) live in _ICON_PNG (filled
# by dev-time baking from Lucide SVGs). At render time an icon is drawn as a
# coloured circle with the white glyph centred. When a glyph is unavailable we
# fall back to a clean numbered/lettered circle so slides never break.
# ============================================================================
_ICON_PNG: dict[str, str] = {
    "activity": "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",
    "alert-triangle": "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",
    "award": "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",
    "bar-chart": "iVBORw0KGgoAAAANSUhEUgAAAIAAAACACAYAAADDPmHLAAAABmJLR0QA/wD/AP+gvaeTAAAFBUlEQVR4nO3du49VVRTH8e/mpQkik7EyKAISJTYQGgsfUbS0UAqNQU3GwsRor5V/gCGxMeiEhA5fBdEGK2M0mpAIBgsKCilQNCiRkcI4CPOzmD0JDHeGs+95Muv3SWiGtc9acH/cuXfPPRswMzMzMzMzMzOzlS+Ns0jSBLAX2JW/dBI4klKaaWowGyhJ+yRd1I0uStrX93zWovzgz4148BfMOQQrlKSJJf7lj3ommOh7XqtmVUHtXqDKA7vw+sBuASUB2HXzkrFqrUclAbitpVrrUUkAbAVyAIJzAIJzAIJzADJJOyR9KOm0pAuSjkl6S9L6vmcbBEnTFTaBFkz3PW8JSa9Kml3iz3JG0oN9z9iW8M8Akp4CDgLrlijZChyVtKG7qboTPgDAu9z872Eb8EYHs3QudAAk3Qfsrlj+bJuz9CV0AIAtBbVb2xqiT9EDsNT3/VHWtjZFj6IHIDwHIDgHIDgHIDgHIDgHIDgHIDgHIDgHIDgHIDgHIDgHIDgHIDgHILg1fQ8QkaQ7mf+AycKHUX4EPk8pXep6FgegY5KeBz4AJhf91l+SXk8pfdblPA5AhyS9AHzM6JNZJoFPJKWU0qddzeTXAB3JT/sHWP5YngQcyLWdcAC68xw3Pu2PMkmHH0B1ALqzs6XaWhyA7pTcYnZHa1Ms4gAE5wAE5wAE5wAE5wAE5wAE5wAE5wAE5wAE5wAE5wAE5wAE5wAE5wAE5wAE5wAE5wAE5wAE5wAE5wAEN+gASFqfz+w/ls/wP53P9N/R92wrxWDvDMpn9H/J9Wf03gU8AEzl26gO9TLcCjLIZ4B8Nv9Rlj6geR1wMJ/1bzUMMgDMn82/7SY1q5g/699qGGoAqt4atTuf+W9jGmoASs7m39LWEBEMNQAlZ/OXnPlviww1ANYRByA4ByA4ByA4ByA4ByA4ByA4ByA4ByA4ByA4ByA4ByA4ByA4ByC4oQZABbVzNfqUrC2Zqe76ur0qG2oA/myptq8+XfeqbKgB+KZi3QXgVI0+p/I1qqg6UxPr6/aqbKgBeA+YrVC3P6V0ddwmee3+CqWzeaY6vgKOV6g7nmuHRdK0qptuoN+UpKvL9PhC0uoG+qzO11rKVUlTdfvkXtslnVum1zlJ25vo1biuA5B77pF0YtG1z0t6W1JjN7VIWpOveX5RrxOS9jTVJ/faJOkjSZev6XM5f21Tk70a1UcArum9OYdhZ5MP/Ig+a3KPPZI2t9Un99oo6ZH8a2ObvZYz2FvDrpVSOguc7aDPFeCntvvkXn8D33fRazlDfRFoHXEAgnMAgisJQMn2pIN1iyh5oP4pqL2/dBDrR0kAfiuofbTtt1HWjJIAlOy5rwWm23zPbh2TNCHpSsFmkCQdlXRv37Pb0pb7j4xvIOlr4InCHv8B3wE/U+9n9zbaLHASOJJSmildXBqAKcAHMw3TDPBmSulwyaLSANwOnAHuLllnnRHwckkIit6vp5T+Bd4pnco6k4D3JU1UXTDOhs0h4Nsx1lk3JoC9VYuLA5BSmgNeAv4oXWud2VW1cKwt25TSL8AzwKVx1ttwjL1nn1L6AXgaON/cONaQk1ULi94FjCLpHuAw8Hjda1kjZoCtVfcEav/ULqX0K/Ak8Brwe93rWS1ifi+g8oZQ7WeA67rP7xO8CLwCPAbU/tSuVdb+RlCJ/F70YeAhYBOwoa1ewdXaCjYzMzMzMzMzM7MI/gfYDV5J/ydD+AAAAABJRU5ErkJggg==",
    "bar-chart-3": "iVBORw0KGgoAAAANSUhEUgAAAIAAAACACAYAAADDPmHLAAAABmJLR0QA/wD/AP+gvaeTAAAFBUlEQVR4nO3du49VVRTH8e/mpQkik7EyKAISJTYQGgsfUbS0UAqNQU3GwsRor5V/gCGxMeiEhA5fBdEGK2M0mpAIBgsKCilQNCiRkcI4CPOzmD0JDHeGs+95Muv3SWiGtc9acH/cuXfPPRswMzMzMzMzMzOzlS+Ns0jSBLAX2JW/dBI4klKaaWowGyhJ+yRd1I0uStrX93zWovzgz4148BfMOQQrlKSJJf7lj3ommOh7XqtmVUHtXqDKA7vw+sBuASUB2HXzkrFqrUclAbitpVrrUUkAbAVyAIJzAIJzAIJzADJJOyR9KOm0pAuSjkl6S9L6vmcbBEnTFTaBFkz3PW8JSa9Kml3iz3JG0oN9z9iW8M8Akp4CDgLrlijZChyVtKG7qboTPgDAu9z872Eb8EYHs3QudAAk3Qfsrlj+bJuz9CV0AIAtBbVb2xqiT9EDsNT3/VHWtjZFj6IHIDwHIDgHIDgHIDgHIDgHIDgHIDgHIDgHIDgHIDgHIDgHIDgHIDgHILg1fQ8QkaQ7mf+AycKHUX4EPk8pXep6FgegY5KeBz4AJhf91l+SXk8pfdblPA5AhyS9AHzM6JNZJoFPJKWU0qddzeTXAB3JT/sHWP5YngQcyLWdcAC68xw3Pu2PMkmHH0B1ALqzs6XaWhyA7pTcYnZHa1Ms4gAE5wAE5wAE5wAE5wAE5wAE5wAE5wAE5wAE5wAE5wAE5wAE5wAE5wAE5wAE5wAE5wAE5wAE5wAE5wAEN+gASFqfz+w/ls/wP53P9N/R92wrxWDvDMpn9H/J9Wf03gU8AEzl26gO9TLcCjLIZ4B8Nv9Rlj6geR1wMJ/1bzUMMgDMn82/7SY1q5g/699qGGoAqt4atTuf+W9jGmoASs7m39LWEBEMNQAlZ/OXnPlviww1ANYRByA4ByA4ByA4ByA4ByA4ByA4ByA4ByA4ByA4ByA4ByA4ByA4ByC4oQZABbVzNfqUrC2Zqe76ur0qG2oA/myptq8+XfeqbKgB+KZi3QXgVI0+p/I1qqg6UxPr6/aqbKgBeA+YrVC3P6V0ddwmee3+CqWzeaY6vgKOV6g7nmuHRdK0qptuoN+UpKvL9PhC0uoG+qzO11rKVUlTdfvkXtslnVum1zlJ25vo1biuA5B77pF0YtG1z0t6W1JjN7VIWpOveX5RrxOS9jTVJ/faJOkjSZev6XM5f21Tk70a1UcArum9OYdhZ5MP/Ig+a3KPPZI2t9Un99oo6ZH8a2ObvZYz2FvDrpVSOguc7aDPFeCntvvkXn8D33fRazlDfRFoHXEAgnMAgisJQMn2pIN1iyh5oP4pqL2/dBDrR0kAfiuofbTtt1HWjJIAlOy5rwWm23zPbh2TNCHpSsFmkCQdlXRv37Pb0pb7j4xvIOlr4InCHv8B3wE/U+9n9zbaLHASOJJSmildXBqAKcAHMw3TDPBmSulwyaLSANwOnAHuLllnnRHwckkIit6vp5T+Bd4pnco6k4D3JU1UXTDOhs0h4Nsx1lk3JoC9VYuLA5BSmgNeAv4oXWud2VW1cKwt25TSL8AzwKVx1ttwjL1nn1L6AXgaON/cONaQk1ULi94FjCLpHuAw8Hjda1kjZoCtVfcEav/ULqX0K/Ak8Brwe93rWS1ifi+g8oZQ7WeA67rP7xO8CLwCPAbU/tSuVdb+RlCJ/F70YeAhYBOwoa1ewdXaCjYzMzMzMzMzM7MI/gfYDV5J/ydD+AAAAABJRU5ErkJggg==",
    "book": "iVBORw0KGgoAAAANSUhEUgAAAIAAAACACAYAAADDPmHLAAAABmJLR0QA/wD/AP+gvaeTAAAEoElEQVR4nO3cPYhcZRSH8eesIa5sFI0QTdQmbmO7mBRRJEYSibCN2thExKjxAyyjKYLBxsKPJn6AWhix0AgpVCImIUrEiJV2amHjBysjC7ILIevqsZgJblLs7p333HfGOf9ffec9l73P3rlz7zAgIiIiIulY7YHuPgZsBXYAU8AksAEYr70vQ24W+AE4DbxvZj+3MaRaAO5+E/AksAfYWGvuiFgE3gOeMbOZyIVbD8Dd1wOHgEeBtW3PG3F/AA+Y2cmoBVsNwN2ngTeB69qck8wCcJ+ZfRyxWCsBuLsBzwMH2pqR3Dywxcy+L10o/OD0LvLeAB6JXlsu8hVwu5l5ySJjQTuz1Mvo4NewDdhdukhoAO7+MPB05JqyrL2lC4S9Bbj7JPAtMBG1pqxoHrjGzBb7XSDyDPAqOvi1rQNuLlkgJAB3vwvYFbGWNHZ9yYvXBO3E/j5fNwMcBc4CnaB9GQXH6P53r8Zgb665+2Z3/8ebOe/uB9z98oHu/JBy99kGf8udJbMizgD30+xi8hwwbWanAmZLoYhrgKbv/ft08IdHUQC9W75bGrzkjJkdKZkpsUrPADcAVzXY/nDhPAlWGsCmhtuHPcaUGKUBrPajCsCcmc0WzpNgpQFc1mDbvm9XSnvaeBoo/yMKIDkFkJwCSE4BJKcAklMAySmA5BRAcgogOQWQnAJITgEkpwCSUwDJKYDkFEByCiA5BZCcAkhOASSnAJJTAMkpgOQUQHIKIDkFkJwCSE4BJKcAklMAySmA5BRAcgogOQWQnAJITgEkpwCSUwDJKYDkFEByCiA5BZCcAkhOASSnAJJTAMkpgOQUQHIKIDkFkJwCSE4BJKcAklMAySmA5BRAcgogOQWQnAJIbk3tge4+BmwFdgBTwCSwARivvS9D7Opag2oGsNbdXwD2ABsrzpVl1AxgAthfcZ6sgq4BklMAySmA5BRAcgogOQWQnAJITgEkV/1W8CVmgKPAWaAz4H0ZJseAdTUGDSqABeAQ8JKZnR/QPgwtd/+r1qxBBHAOmDazUwOYLZcYxDXAPh384VE7gDNmdqTyTFlG7QAOV54nK6gdwMnK82QFNQOYM7PZivNkFUoD+DtkL2RgSgOYb7Dtle5+beE8CVYawK8Nt99ZOE+ClQbwG/Bng+2fKpwnwYoCMDMHvmnwktvc/aGSmRIr4lPAZw23f83d7w6YKwEiAvgQ8AbbjwMfuftBd78iYL4UsIhF3P1ToJ//6g7dgL6mez0hXU0eB+8ysxP9DooKYDtwOmItaWy7mX3R74tD7gSa2efA8Yi1pLGZkheHnAEA3H0z8B2VvskiAMwB681ssd8Fwp4FmNlP6HN+bSdKDj4EPwwys3eAFyPXlGW9VbpA2FvABe5uwOvAY9Fry0W+BO7o3YzrW/jj4N4OPQ48R7P7A7J688De0oMPLZwBlnL33cDb6AchIi0A95rZJxGLtfqFEDM7DtwCvALo69/lOsA9UQcfWj4DLOXum4AngAeBG2vNHRGLwLvAs2b2e+TC1QK4oPcjUVPAncCt/PcjURO192XIdYAf6d5h/cDMfhnw/oiIiIjIqPgXLHIsRkaF8NUAAAAASUVORK5CYII=",
    "book-open": "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",
    "bot": "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",
    "brain": "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",
    "building": "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",
    "building-2": "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",
    "calendar": "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",
    "camera": "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",
    "chart-line": "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",
    "check-circle": "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",
    "circle-check-big": "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",
    "clock": "iVBORw0KGgoAAAANSUhEUgAAAIAAAACACAYAAADDPmHLAAAABmJLR0QA/wD/AP+gvaeTAAAMv0lEQVR4nO2daaxV1RXH/wsFbR8y2AEqVsVUHCg1pB+kFbDg0Da20jZxVtTaSW1qnY1pHGq1qGhF7WATWwsUU2PjgOIHLJRKjbYxJmqraNICUiZFpgcICL9+2PcJPu60zj3n7HveO7/k5RLuPnevtfe6++699t5rSSUlJSUlJSUlJb0Oiy1AVgCfknSEpBGSDpd0oKQOSf0rr4Mrr5K0SdLaymtn5XWZpEWS3pD0upmtyFP+vOgRBgDsL+k4SRMkjVHo8AEpV7NewRielzRP0gIzW5tyHblTSAMA9pZ0QuVvoqSjJfXJWYydkl6SNF/SXEnzzOz9nGVomUIZADBS0rmSzpc0JK40e/CupEckzTCzhbGFaZa2NwDg45IuVOj4kZHFaZZXJc2Q9ICZrYktTCEBDgDuAjopLhuBqYQJaUkzAAcD04DNMXsuZbYC04HDYrdv2wIMAu4DtkftqmzZBtwDDIzd3l20xRwAOFXSvWq/iV1WrJR0jcKEkZiCRDUAYISkXyos53ojf5N0sZn9K5YAea+dJUmAAdcqzJZ7a+dL0nhJLwFXAVG+jLlXCnxC0nRJX8mpylWSXlPw4i2VtFkfdv1KH3YNd0g6SMGFfKSkT+Yk51OSzjezd3KqT1LOBgCMlzRL0rCMquiU9KyCd+5ZBR/+ulY+EBikYAjjFFzN47RrDyFtlkk6s0iOpKapDHNZzPAXAz8DvkhwEWetR19gLHArsCQDfbYDV2StR24Qfu/vSLmR1hPW1ScAUeYxFd36EIzh/opMaXJPTN1SAegHzEqxUVYBN9JG6+gugP7ApcCyFPX9I9A3tm6JADqAOSk1xGJC434ktl6NIBj9ZODNlHR/Btgvtl4uCN+G51JQvhO4FugXWycvBEO4DtiUQjs8B2Q18UwXwiTp6RSUng0cElufVgGGEeYrrfIMsE9sfepCmBS1+pu/CvhabF3SBpgErG6xbWYSyWHUFITt21b4K3BAbD2yAhgCzG2xje6LrUdVCOv8pOwAbgL2yknWAYSJ2t2Vv8lA2mcIa9W9N8FvsaOF9ro8D1mbBhhPcifPJuDkHGU9DVhTRY41wGk5yjGJ5GcetgPH5iVrXYD9Se4RexcYm6OspwM768izEzg9R3mOAd5J2HZvEY7LxYPg5Xs8oQLLgc/lKOsAqn/zu7OGnH4OKnIdVenMJDxFzEkhcE1Cwd8i5yUecJ5Dvsk5yzac5B7EK/OUdXehRxKOOHlZQzjenbe8nhXKnRHkG0X4SfSyFTgyab2JNhsIw869krx+6i2SJkU6AePxpPXPTIoamNkrkr6qXWcUmqWfpN+Q8Kcg6W7TOQp74x52Sjq1R+51p4SZvSDpLIW28jBe0plJ6nQbAGGCdFuCum42s6cSPNerMLMnJE1J8OhdhMMrLpKMALdK8l50mC/p5gR19VZuUDgw6mGIpJu8FbkMABgu6fvOOlZLOsfMdjif67VULpmernB83MNFOFdX3hHgakneY1cXmtly5zO9HjNbKekHzsf6SnItC5s2AGCowq1cD4+a2ZPOZ0oqmNnjkmY7H7sQx4aaZwS4WtK+jvKbJbXXpkUx+aF8S8N9JV3WbOGmDAD4mKTvOoSQwqx/sfOZkm6Y2VL5VwUXEe5fNKTZEeA78jlHFku6y1G+pD53SHrLUb5D0gXNFGzWAM52VC5Jt5vZNuczJTUws62SpjofO6+ZQg0NAPi8pFGOildKetBRvqQ5fivJE6nsKGB0o0LNjADnOiqVpKlmtsX5TEkDzOw9SXc7H2vYd3UNgHApweNj3iDpfkf5Eh+/lrTRUf4sGlyXazQCHC/f7dhHzKzTUb7EgZltVIhE1ixDFMLo1aSRAZzoqEwK175LsmWms/zx9d5sZACeLd8lCleyS7JlvsIyu1mSjQCE8KtHOyqaaWbefewSJ5WYQg85HhkNDK71Zr0RYEKD97szx1G2pDWedpTdSyGOclUaGUCzdEr6p6N8SWu8IN/+QM2+rGcAxzgqWGhm2x3lS1qg4mX1HK0bU+uNegYwwlHBPEfZknSY7yh7eK03qhpAZT/ZczmiPOiZP54V18DKeY49qDUC1LSYGrzuLF/SOq85y1cd0WsZwBGOD17VEzJnFI1Km7/teKRqn6YxApTf/ngscpSt2qe1DMBz7NsjREm6eNq+ap/WMgDP6Z9ljrIl6eI5JVQ10lgtA/CEJfNsTxaFw+q5T9sIT9uXBuBggqSlwJTKnki70hYG0FP3//srJHVY0saGkJkBeOYAPdUAumhnQ8jMAHoirRpqOxtCYmoZgKexcg+mkJCXU/qcdjKEludqtQzAM7QUxQAeVcjumRZdhvAfQozDGKuGtjCAQkSyNrMNki6WlHaWroGSrpf03wiGUBqABzP7k6QzlO5I0EUMQ2gLAzjQUTY6ZvawpM8oRCxZn0EV3Q0hyzDvn3aUdRmA5wqSd+s4Oma21syulzRcwRA2ZFBNlyG8AmTVRp7PrRqko5YBeDYZPFvHbcVuhnCIsjOE4ZLmkE3WD48BVO3TWgbg2eIdkiQ6VTuRgyEcKumSND+wMsdoKgZABZcBeLd4E0eqbCcyNoRvpPhZkr/N36j2n1UNwMxWyKf8OKcwbU1GhnBwCp+xO+MdZddVgk7tQT1XsOdnwBs1tBCkbAhprzg8bV5zRK9nAC84KhhPATN7NUtKhrAgLXkqbe1JGPF8kkq+6Yxa3R4ZLHIAGAz8lOazhb4HpLZaImRo8TApqZKevDa3pqVgUaA5Q9gBNBWwyVHvFEe/vE/SVRrwoqOiJRQ9121CdjOEVd3a5EWg7vXsBHX1AZY6+uUfrVTmTfr8pfRULR6EjGBHAxOBgzKq43hnnySJPP5BZSc5K/tdirqWVAF40NknJ7RS2d7ASkdl64GinA8oHMB+wAZHfyynQS7Gur/ZlbDlsxwyDpA/wnVJ81wi3xbwrJbD9AOjHRYHYcRo+zTvRQPYt/KN9uAJ8VO38pedFae68VEiAT929sGraVbuzQu8hHZPdV4gCN9+b3LJq9IUYH98kw+A61IToJcD3OBs+07STisLTHUKsZmQY6ikBYCDCQm2PdyehSBD8We7fix1QXoZwGxnm2/BkTLGK8yvnMJAko2IEkkS8K0E7X1vlgIdRMhV62E1MCwzoXoowIHA28623gakffBkD8HucQoFsIAGYctLdkHwwC5M0M7Zp+kBBgD/SyBcmTm0SfBt93axAhiYl4BnJxBwB3BKLgIWGMJBnJ0J2veMvAX9SwIhNwM96gBpmgBj8C/5AObGEPYo/BNCgHeBz+YucJsDjALWJmjPrWR386ih0FcmEBhgGaWT6AOA4SSbVwE0nSU0C8ENeCyh4MtJa7eqwAAj8fv5u3gSsNgKDAYWJ1RgLb14TkD4zX8nYdstJaT0jQ9wLLA9oSKb6YWrA8Jsf0vCNtsGfCG2Dh8CuDyhMhCWiLfQC5xFBCfPFJIt9bq4NLYeVQHubEEpCB7DHus2Jrh3n22xjW6LrUdNCJPCmS0quJoeuIFE2Njx+va7M53Yk75GAH2BOS0qCmEbtPBLReBQwmy9VeZSlPuXQH/guRSU3gz8hAIeLyMc47oe/xmKavydbGMNpQ/QQTojAYSTxtcAH42tVyOAfYDvkXxt3525ZBNeJnsIM94HUmoICPODG8lr18sBYdS7lOQevWrMIGRwLy6EieHtKTYKhGH1YeDrNLj5krFufYCxwP34D802Yho96bItcBnJnUX1WAL8nNARmX9bgH7AOOA20hvmd2cb8KOs9egi1yUFIYjEQ/IFOPSwSSGH4XyFvHqvtZrRjBCN60iFOEgTJY2VlNVcZImkM8zMH9EjIbmvKQn+6z9IOjmnKt9WiHe0SCHHziaFqJnrtCv/boekQQr37joUDPRwhRiInlBsrfCEpAvMLIswtjWJ4lQgODOukHSLpGKsbbNjm6RrJd1dSQ2fK1G9SsBhku6TdFJMOSKyQNLFZvbvWAJEnWWa2Ztm9mVJp8iXAq3orJB0nqQJMTtfapOUMWY2W9IoSdMk9eQ09Nsk/ULSEWY2PcaQ3/YQLp9MIx0XarvwHsFXkNXqp+dBuIs4FdgYtetaYyPBCTYkdnvWor23FvXBsvHbkiZLKspp4lckTZf0ezNbE1uYerS9AewOMFLSuQoTqKGRxenOGkl/ljTDzBbGFqZZCmUAXRCOjk2UdKJC0OTRyn9Cu0PSS5LmSXpG0ryWAzJFoJAG0B2Cu/Y4BaMYI2mEQsqWNFmnEHP/eYVOX2Bm61KuI3d6hAFUAxiq4ModoeDWHaaQ669DweU7SLtyHnYqdPDGyr83SVqm0OGLJL1uZqvylL+kpKSkpKSkpKQkM/4PWQDvZsAfjOoAAAAASUVORK5CYII=",
    "cloud": "iVBORw0KGgoAAAANSUhEUgAAAIAAAACACAYAAADDPmHLAAAABmJLR0QA/wD/AP+gvaeTAAAJGklEQVR4nO2dW6ydRRXH/4u2ULC1hZYWSVBa0ZiDVFCCoi2l0tLIpajhYo2RqIlWrUpCRB/E+KBWpUHhwWhiGiNQqKgxFqs9pRx6UeEB7EUaRaQ9JdwptNLS0nN6fj7MVgvZ3z57vm++/e3L+j1+ey5rZtaey5qZNZLjOI7jOI7jOI7jOI7jOI7jOI7jOI7jOI7jOI7jOI7T8VjVAjQCOFnSOZLOlHSGpDdLmi5piqSJksbWgo5I2idpj6RnJQ1K+pekRyRtMbPnWit559BWCgBMk7RQ0nxJsyXNTJT0TkmbJa2TtNYVoo0ATgQ+D2wAjlA+w7W8vgCcVHX5exbgXOA24GALGj2LQ8DtwHlV10fPAFwIDFTY6FkMAPOqrp+uBZgF9Ffbxk3RD8yqur66BmAi8CPC2NspDAO3Am+suv46GmA+MFhtWxZiN7Cg6nosk2PKSBQYB9wkqV9h7d6pnCZpLbAcGFe1MGWQ3A4ATJd0t6Q5qdOumD9JutLMnqlakJQkVQDgLEn3qLP/9Y3YLekyM9tetSCpSKYAwGxJqyVNTpXmUeyV9LCkfyiYeJ+pfXu19vs4BfPwKQrWw3dIOlvSiSXIsk/SIjPbWELanQnwQeBAwsnXYeAPwFKgD4hWVMBqcb8E/BF4NaF8B4CLyqjLjgOYDexPVLHbCY0+pQQ5TwK+WMsjBQeAbpvnxAG8E3gxQWVuBhaS45+eU+6LgU0J5H6JMO/pPYBpwK6CFfhPYFGFZVhUk6EIg4SVT+8AjAU2Fqi0IeA7wPFtUJbxNVmGCpRnE11qJ6gL8L0ClTUIvL/qMrwe4HyK9WjLqy5DSyDM+PPu298PTK26DFkAU4H7cpZtBJhfdRlKhbCxk9e2vxI4tuoyjAZwLOGMQB4G6eYNJOCWnBXzM6CUfYcyAI6pyZyHW6uWvxQI+/l5tnRXdlLj/5eaEtyRo7zDwLuqlj85wNoclbGBDuj2syAMBwM5yt1ftexJAebmqITdhGPdHQ1hYphn3tM9x8uA9ZGFHyJsDnUFhCVirJ1goGq5k0A4vRvLsqrlTg3w3Rz10PanjUe1vQO3SfpERJqPSTrLzA7llqoNAcZL2q5wQ6lZVkvKWhWMSHpJ4SbT02ZGMQlLgLCDFntu//Kq5S4L4PIcvUAzHAT+CqwAPgO8peqySpIIN3Zi2Fy1zGVDml3EZtgCfB04tcrCbogUemFlwrYIYEHypm7MYeBOWm1bAE4mzvDzN1q0n181wLYSGno0RoBVJB4eGlnoLpY0JiKtn7blRKYcflJBnibpakk7gBuAmLaJB/h5hHYepoRjXO0KMKVW5irZDBQ+fd2oB4g573afme0pKkynUCvr+orF+ICkh4ALiiRSVwEIjhpinDOsKSJEh3Jn1QJImiqpH7gmbwJZPUDsjLPqf0MVrJT0UNVCSDpO0h3AJ/NEzlKAvog09krakSfzTsbMhiVdoWAdrJoxklYAH46NODbj+1sj0tjaQ7P/12BmTxLs/Z+VdI2kt6v5ldPxksYnFGeMQk9wgZkV65mA30bMRn+cRv7eA5hE2GxbAtwNvJxgdbCLor6PgL9EZPi1RPXR8wAnAJ8mGNWK8Kuigjwakdm1icrv1CAcR/sU8GwBJbi6iABPRWRU2c2ebodgjl+dUwGeBCaMlkfWKuC4CDkPRoR1IjCz5yUtkvT9HNFPlXR9royB5yM0ret3ANsB4Bs5eoG9wKRG6Wb1AEMRsqVcyjgZmNm3JcWuuCYpLFEzyVKA/RGZlOERxKnPdZIeiIyzhAbb9FkKELOx86Y4eZy8mNmQpGv1f9c4zTBTwfF2XbIUIMYTVnucX+sRzOxRST+MjHZV1g9ZCvBEROIx+wZOGm5S3DB9WdYPWQrwWETisxqNMU56zOxFSbdHRJkBnF7vhywFeCQi8ckKL3o4reUXkeHfW+9jlgJsjUzcXaa1ngclxbx8UveMR10FMLMXFBwyNsslEWGdBJjZiIL72mZ5W72Pjc4EbopIfB5dcBO4A4npqU+r97GRAtwbkfg4SYsjwjtp2BURtu4ftJECrJV0JCKDz/lqoOW8EBG27s5gpgLU5gExw0CfpA9FhHeKE7NnU/eo2mi+e+6KyECSbowM7xTjDRFh65qPR1OAVYrb738f8JGI8E4xTokI++96HxsqgJntVXj9I4bltIEL2B6h7tIug6frfWzGfVus37uZkr4VGcfJx3siwg7mzgVYF3kSZZhe96VfMsAEwsunzXJDkczmRCoAwBOEO4ZOCQCLI9ujmB9jYE0OJdgExBwwdZoEuDeiHYaAiUUzPJN8PvV/SdnODHoM4N2RbZDGdxNwcw4FgOBswpUgEcS77f1qqownADtzKsEqfDgoDPFj/xESeBI5WoB55H8wYiM+McwNMIP4R7rSO64GluVUAAirA18iRgJMJt+Td+kddxIejRoooATDwA+AE5IL14XUGv/POep5G2Xt0BIuLj5eQAmoxf9oKQJ2CcBM8j92eUXZwvUBewoqAcADwKX4eYLXAHyc8DhlHlrjt4ngSz+FZwuAHcCX6fHjZYR1fn+BenwFiNkoKizwXNIpAQSD0zrgOsLztB335lAshCX2YuIsfFksbTbflM/Hny/p9yrvyfYtCs/H71TY2tynuNsx7cZESdMVHEudK+k8xfllyOI3kq5s1nFX0jEX6JN0j6QZKdN1mmarpDlm9nKzEZJ2rWa2Q+EGyoaU6TpNsVPSJTGNLyVWAOl/bk3mS1qm8CyKUz67JF1kZk/FRix12QVcKGmFfEgoky2SLs3T+FIJPcDRmNn9kmZJulnScJl59Si/VhjzczW+VLICSJKZ7Tez6xUuJ/aiV/EyeEXSUklXmVmhlVDLLW+EByW/KWlBq/PuEtZLWmJmMT4cMqnM9AqcI+krCs+g+DHy0dkm6UYz+13KRCu3vQOTFXzYfEzSXMW9U9TtjCj842+RtKYMr+yVK8DREN4dWqjgcGKO4i4+dAvDCq7gVku6y8x2l5lZWynA66kpxNkKF0/PkHS6pGmSpijcdu1kJ5WHFB7beE7S45L+LulhSQ/GGnMcx3Ecx3Ecx3Ecx3Ecx3Ecx3Ecx3Ecx3Ecx3Ecx3GcHuc/r6GJ6QNcniMAAAAASUVORK5CYII=",
    "compass": "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",
    "cpu": "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",
    "credit-card": "iVBORw0KGgoAAAANSUhEUgAAAIAAAACACAYAAADDPmHLAAAABmJLR0QA/wD/AP+gvaeTAAAEhUlEQVR4nO3dT2gcZRzG8e8vtYgiRaRQFDVBEKSikFARPKut+B8FiwhaPShEPHhQhIKCglo8qQj+IUKVSMGDFHuQ6lVQtAd7UDyobdVqbQ1YobTVPB42pzjv7rw7s7NJ3ucDucy88+5veZ+deWd2MwNmZmZmZmZmZmZmZmZmZmZmZmZmZrbqRdMOJE0CdwDXA5cCm5b+rF0CjgKHlv6+Aj6MiL+adDpUACStAx4CZoHpJgVYI6eAj4DXI+LzYTrIDoCk24CXgc3DvKCNhIB3gaci4kTOhrUDsPSp3wU8mVebdeg4sD0iPqu7Qa0ASNoA7AG2DVmYdec08GBE7KnTeGAAlj75+4CtDQuz7iwCOyJi96CGEzU6ewUP/mozAbwl6bpBDfvuASTdDuxtqyrr3GFgpt/EMLkHWNr1vziKqqwzlwPP9mvQ7xCwA7i61XJsHB6TdGVqZb8AzI6gGOveemBnamXlHEDSFPBj5gudBT4A9gO/ZW5rgwVwBfAIMHByt8xJYFNEnKrVWtITynNI0jWZRdkQJE1I2ilpMXOM7s55kfmMjs9IunaE79kqSHohMwBvVPWTmgNcklHLfER8k/8WrKHngO8y2k9VLWwjAPsz2lpLIuIf4M2MTSarFqYCsDGj42MZba1d32a0rfyNRioAOV8TK6OttWsxo23lWNf5LsDWMAegcA5A4RyAwjkAhXMACucAFM4BKNw5LfQxI/la0JjMNO0g9XuABeDCpp3birIQERctX+hDQOFSAVjXaRXWhcoxTQXg3BEWYuNROaapALQxObSVpXJMPQcoR+WEPxUAn9etPZW/HUgF4OwIC7HxqBzTVABOj7AQG4/KMfUhoByVY9rGbP9vfMgYl/XABa33Kmkh4x8Obmy9AKtF0k0Z4/RnVR8+DSycA1A4B6BwDkDhHIDCOQCFcwAK5wAUzgEonANQOAegcA5A4RyAwjkAhXMACucAFM4BKJwDUDgHoHAOQOEcgMI5AIVzAArnABTOASicA1A4B6BwDkDhHIDCOQCFSwWg8bNorBONn+2UGrzfMzq+KqOttavySWAJf1QtTAXgSEbHj6r3qHnr3taMtr9WLUwF4KeMjjcz4Bn11j5J08B9GZv8ktP5PRm3HpF6DzJ+RpLnAx2QNC3pSOYYzVb1lbpd/Pn05gG5NyD6EpgDfsB3GhuFi4Gbge3k3eBLwGRE/O/QnpxFSnoPeCC3QluRvo6ILVUr+u2yn8e3f1srXk2tSAYgIr4H3h5JOdalg8D7qZV9LyRI2ggcAC5ruSjrzraI+CS1su+sPSKOA/fiewevVi/1G3yoeSlR0sPAO3Xb24rwMXBnRPS9rF/rvD0i5uidEZxpoTAbvX3A/YMGHzK+yImIeeBW4ESDwmz0dtH75J+s0zjryl1EfErvy585fKFnpTkI3BIRT0fEv3U3GvqYLukG4HHgLuC8Yfuxxg4ArwG76+zyl2s8qZO0gd6ZwhZgEpiid8nS2ncMOAr8DHwB7I2Iw+MtyczMzMzMzMzMzMzMzMzMzMzMzMzMxu4/1upWe9CuEfwAAAAASUVORK5CYII=",
    "database": "iVBORw0KGgoAAAANSUhEUgAAAIAAAACACAYAAADDPmHLAAAABmJLR0QA/wD/AP+gvaeTAAAMTElEQVR4nO2daaxV1RXHfwtUHAABRcCiDIqPQnFEa4ljqRitVdtYG4dSjEPTJk36wVYT2nRQ2jR+avqhqU3T1rHWqaDiSCyoFQuioKiIlcEyFUWGhwwK/35Y+/oez/fgnnvPOffcd/Yv2bmXxxnW3ed/z917nbXWhkgkEolEIpFI6bBGG1APknoAQ4CjgIGhDQ6v/YDe4bVPeH9Q2NXC39uzEVB4vw1oBbaEv7eG1/XA2vC6HngPWGNmu9P/dPlQeAFI6gmMAEYDnwdawr+H4Rf+gMZZB8BOXAgrgGXAEuDN0Jab2a4G2rZPCiUASfsBY4Hx7do4oFcj7aqDHcAiYH67trhIomioAMK3+1TgbOAs4Ez8dt2d2QI8B8wObX4jBZG7ACQdDEwELgIuAQblbUPB2ADMAh4FppvZpjxPnosAJB0AXAB8G7/wzXpLz5rtuBDuAJ4ws4+zPmGmApB0NPADYApweJbn6oasB/4C/M7M3svqJJkIQNKJwE3AZcB+WZyjRHwMPAD8xswWpn3wVAUgaSjwU+BaoGeax44gXAg3mdmytA6aigDC9O1mYCpwYBrHjHTJduAW/I5Q9+yhbgFIGgH8FZ/CRfLjJWCymb1dz0HqEoCkLwMPA33rOU5KbGJPN+268LfNuCt3a3hfob3rt6NruC/uOj4kvD8Un662dzcfmtHnSMJm4FIze7bWA9QsAEmXAXeR35ROwErgLdzN+g5t7teVZrYlJzvcGKkP7o4eHl6PBcbgruqjyc/HsgO40sweqmXnmoyUdB7wONkN9IRf5Ir79GVgoZltzeh8qSKpN3A8e7q0R5OdKHYB55vZrKQ7JjYojPQX4LfCNFkKPAn8E5hjZutTPn5DkTQQd3mfDZwPjEr5FP8DTjazVSkfd08kzVZ6zJX0Y0mjMzW6gEgaHT773BT7s+axQLVGn5OCkasl3SKpJVNjmwhJLZJulbQmhf49K0tDZ9Zh2OuSviN/LhDpBEm9JE2RtLiOfn40K+OG1mjQKknXyB/9RqpAUk9J18rvlknZLenILIy6vgZj7pXUP3VjSoKkAZLuq6Hfr8vCmOkJjfhF6kaUFPmYKQnTszBiYQIDsp2KlBAl+zmo+qlhjwQ27Eiw7SBJExJsH9kLks4Ajkiwy/ZqN0wigCShSj2Bp+Tu4kgdSPom8ATJvK5VX6skAtiYYFvwByn3S7pT0pCE+5YeSUdKuhv4O96XSaj6WiURwIcJjahwNbBUPpBJ233c7ZA0UNKtuGv8yhoPU/W1SiIA7XuTLjkE+AmwQtIf5CFjkXZIOknS7fgTz6nAwXmcN4kA0uAg4AbgFUmLJN0oaWTONhQGSSNDHyzCH7BdT84RVY0M2BwH3AbcJmkx8AjwFDDXzLY10K7MkHQQ8CVgEh4eP7axFhUnYndsaDcDOyS9BMwhxAKY2X8baVytyB+dn4LHA5wNnEbBciKKIoD29MLTxD59qiVpHbAQWExb4uVSM1vXEAs7IGkQ/nx/DJ7AOgY4gSbIeiqiADpjEH7bnNT+j5K2A8vx0LBVwBraYgLX42lXraFtNbOqpkeS+uED196hDcAdMZWYwCHA5/BQsGE0cSR0swigKw7EQ62qCiiR9pjIbAIqef09KEaQZ+40uwDqoZQXvCN5TwMjBSMrASygds9h5LN8iPdp6mQlgPl4bPx38Tj+SG28C/wQL4UzP4sTZPYTYGatZnY7Pr8/D7gTz86J7J1WvD7AV4BRZvbbLPMhMh8EhgpazwDPSPo+XhXkYjw2Pg7EnI14TsQMYIaZteZ14lxnAeGD3Q3cLWl/3NkzCTgHODlvexrIJ3i202zc/T0nj2ogndGwDg8feFZolVy7M4DTcdfpKTSBJ61K1tGW4vYi8Hye3/K9UZhvXEjufDw0ACQdhbtUx+DOnsprUX86NuGD3sXh9Q08p7GwzzIKI4DOCLVx3sMLJ32KPNR8GD7TGA4Mpc1VOzi8r1QKTYMt+MVdF1r7SqErKs3Mmm7qW2gBdEXo6A+BV/e1bfDrV3z6ldCqfZWK3Up4hlDt84NmpSkFkIRwAbv1RayH6AouOVEAJScKoOREAZScKICSEwVQcqIASk4UQMmJAig5UQAlJwqg5CQRQJLybrEUXPpk0v9JBDAio20j1ZFJ/ycRwGEZbRupjkz6P4kAkmx7tKQirCHQLQgxDcMS7FJ1EfAkF/WjBNv2xYtEJalsFekESYOBp0m2oGbV1yqJADYk2Bbgi8Crkr6WcL9IQNIlwCt4kGwSPqh2wyQCWJ3QCPA06hmSHpV0fA37lxJJJ0qaCfwDj3FMyppqN0wigJ01GFLhq/jdYLqkcyUVatHqIiDJJE2U9AieB3hBHYerOscgz5hAwzOCLgbeknQHcI+ZrcjRhsIhaThwFTAZOC7v8zcqKHQ08CtgmqR5+K3ucTyGvp5ydIUn3P1OBC4ELsUTYBp2R2x0VLDhhZNOwwWxVr7syXPAC8DiRi6tngbyRTXH4FlPZwLnUqCMp0YLoCODgStCA2iVtAAfCS8CXgOWmNnmLvZvKMH3MRovgTcOz3c8ifQSVFKnaALoSG86VAwDkLQGL6W6DC8StRKfpazCV896P+07h3zFk8PxrKOhwJF43v6I0EZR24i9oRRdAF0xJLQuF0iSVMke2oSndu2gLfunq0SRyuom/fBydX3wPMT+7f6vW9GsAqiGbnvR0iTGA5ScrATQdFmyTUAmfZqVAB4ALseXOI/Ux4vAZXifpk5WApCZ3W9mp+Nz/D/hA7FIdWwG/giMN7MJZvYg9a3X0CWZjwHMbJ6ZXYdPka4EpgPdshx8nWzDPaJXAIPN7AYzeznrk+Y2CzCzj4B7gXslHYIXh7ogvCYJduhOLMeLRM0Eng59lCsNmQaGuncPh4akY/F6+pUiUS000D+eEcLrBs0Fnscrg73TWJMK4gcIHfEOPlaohECNx92o44Av4HX4m6Us+zb8Yr+Ou68X4AtfFK5SSSEE0JHQUc+EBoCkHvhPxShgJF4cahjujh2KjzHyWo1jO7AWdz1XCkUtw0u7LgVWhgKZhaeQAuiM0KHLQuuUcOc4Ao+KHYDHJvbBPYL70/ZQph+f/Ylp7yLeggdVbKStQtiG0NaZWZJFNAtN0wigGmJBqOREV3DJiQIoOVEAJScKoOREAZScKICSEwVQcqIASk4UQMmJAig5UQAlJwqg5EQBlJwogJKTRABJolKHJ7Qjsm+OSbBt1cEoSQSQ5Dn7JEk/T7B9ZC9ImgZMTLBL+kkkkq5Rcu4JUTqRGpDUX9J9NfT75CyMOaYGQyRptaTJ8pi+SBVI6iFpiqQ1NfT3bknZhNlLeqlGEUjS65Kuki8aHekESfuHL8sbdfTzC1kaeHkdhlVYLemXkpIMaro1kkZJmqbavvEd+XqWhpqkF1IwssKLkm6UJ4aUCvlF/5GkuSn252wlLMGXOPtG0nH4UuhJSpdWw9vAk8CzwHNm9n7Kx28okgbimU8T8XS4USmfYjNwspn9J8lONaVfSboYeAjoWcv+1ZwCWAL8G5iHZ9a8FpaYLzyS+tBWJOpUvGxuS4an/AS41MweS7pjzfl3kq4G/kx+uQXCM3CWhLYUT65chmfi5CoOeUWwyvL1lSJRLaHlmez6MTDZzP5Wy851JWBKuhDP+C1CafgttKsShhdM3oBn9bTiRaIqDpLOHCXtl5TvT1uRqL54ltFh7FklrAil3zYB3zKzJ2s9QN0ZuJJacBGcVO+xIomYD1xRb4Zx3c4ZM1uC/8b9DE+ajGTLNmAqMCGN9PJUc/AljQB+jdcH6m75/Y1mN3AfMNXMukyQTUomF0nSCbhKv0F2M4WysAt4EJhmZovSPnim31JJI4HvAVPwAVSketbjs6zfm9nyrE6Sy21aUi/gIuBqvC5QXoUcmo3teL2gu4DHzKyeRTqqIvff6TB/vhC4BPeIDcjbhoLxAe4BnQHMzNuf0dCBmrwC93jcPXoWMIH0XcxFYzPwL2AOMAuvHdSwNREKNVIPghiLTytPxVfWGAsc3Ei76uAjYDG+3sE8vHLqG0VaBKNQAugMeSDJSHzVjRbc5XoM7oI9Cq/900h24oWiltNWJOot4E3g3aIXiyq8APZGEMdg3DU7qF07DHfnHor/pPTDhVL5eenLZ6enu/DbM7hbeSdtaw1sxN3HG4B1oVWqhK0t+kWORCKRSCQSiUT24P9hGpDy7eZ/HwAAAABJRU5ErkJggg==",
    "dollar-sign": "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",
    "eye": "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",
    "factory": "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",
    "file": "iVBORw0KGgoAAAANSUhEUgAAAIAAAACACAYAAADDPmHLAAAABmJLR0QA/wD/AP+gvaeTAAAFh0lEQVR4nO3dT4hVZRjH8e/j5DiVUqSbotCoRW1Ks1ArIhtBhhb90UVtXARt22gUhFKbaNWihW2ihSQk1WiQFhWaJlmkDRpSEvbPjZCpkaU25a/FzATpHT3nvO+559x5ns/6nPd9597v3HvuOWfuQAghhBBCcMe6NZGky4FlwP3A7cA84JpuzV+RATOAfmAacA74HTgI7AZ2ATvM7HRjK0xUewCSbgRWA6uAWXXP14BTwNvAK2Y20vRiyqotAElXAC8ATzH2G+TBFmCNmR1ueiFF1RKApFuBYeCWOsZvuT+Bp81sfdMLKSJ7AJLuBrYCV+ceu8dsAJ40s7+aXsjFZA1A0h3AJ0zN9/oqtgIrzOxs0wuZTLYAJM0BRoDrc405RbwPPNLWCKZlHGs98eR3MgRsljSj6YV0kuUVQNIy4KMcY01hrXwlyBXAbuCeHGNNca2LIDmA8QO/fRnW4kWrIrgswxiPVdjnCPAaY+G04oE4z2ZgZk1jTxwTtCaCJJJGVM6b42cJW0vS8ZI/UxXb1NIDw8IkDUgaLfFD75GU41WnVioXwMkS256v8QhSPwbeRLm3kXVm9nfinG3zDHCs4r6Nf0RMDWBOiW1HgZ2J87XR98AgPRpBagBlrvKdavt58arM7AA9GkHOM4GujUdwH3C04hBDwLuSBvKt6tIigIzM7BvgAapHsBzY0s0IIoDMei2CCKAGvRRBBFCTXokgAqhRL0QQAdRsPILlVP+IuBwYrusjYgTQBW0+TxABdEmmCN6QlPU5iwC6KEMEK4G1+VYUAXRdhgjWSlqcaz0RQAMSI+gDXpXUl2MtEUBDEiOYT7U7sS4QATQoMYI1OdYQATQsIYL5khamzt/627MaohLbDksazTBn3/i8Ze7UfpTEO7IjgM5OUPzLK+q6e7iIQeC5lAHiLaCzQ00voKAFqTfZRgCd7Wh6AQX1AzekDBABdLYJ6JW7l5MOBCOADszsCLCx6XUUNC9l5whgcs9S/XRtN/2RsnMEMAkzOwo8DrT9VvYfU3aOAC7CzD4GVpD4W1azpGOVCOASzOw94E7gs6bXUocIoAAz+xa4F3iQsT8dP9XsivKJM4EFmZmAbcC28UuxNwPXUs9jWOf3E/xPBFCBmf3D2NnCWs4YZrq2UEi8BTgXATgXATgXATgXATgXATgXATgXATgXATgXATgXATgXATgXATgXATgXATgXATgXATgXATgXATgXATgXATgXATgXATgXATgXATgXATgXATgXATgXATgXATgXATgXATgXATgXATgXATgXATgXATgXATgXATgXATgXATgXATgXATgXATgXATgXATgXATgXATgXATgXATgXATgXATgXATgXATgXATgXATgXAbSTldhWKROlBlDmX6vPlNSfON+UJ2kGMKvELmdS5ksN4JcS204HlibO58Eg0Fdi+zLPQV6SBiSNqrgvJE1vbMEtJ6lf0t4Sj+do46+qkkZKLFiS3pJ0ZaOLbiFJMyW9U/Kx3Js6b47/Hv4hML/E9iuBJZJeB74ETmdYQy8bABYBTwDXldz3g9TJyxxtdiRpAfBV6jihktvM7OuUAZI/BprZCPBp6jihtJ2pTz7kOw/wfKZxQnHrcgySJQAz2w5syjFWKGSjme3KMVDyMcAESbOBfcDcXGOGjn4AFprZiRyDZTsVbGa/Ag8Bv+UaM1zgJPBwricfMl8LMLP9wBCQbYHhP8eBITM7kHPQ7BeDzGwPsAQ4mHtsxw4Ai83s89wD13I10MwOAXcBLwFn65jDiTPAi8AiM/uujgmyHQRORtJcYDWwCriq7vmmiBPABuBlM/u5zolqD2CCpAHGrgYuZezU8VxgNnFPwjngGPATsB/YDuwws3jlDCGEEEII9fgX+SxGKvAT4BsAAAAASUVORK5CYII=",
    "file-text": "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",
    "filter": "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",
    "fingerprint": "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",
    "flag": "iVBORw0KGgoAAAANSUhEUgAAAIAAAACACAYAAADDPmHLAAAABmJLR0QA/wD/AP+gvaeTAAAH2ElEQVR4nO2dW6wdVRmAvx8KlgJCEW8UKEFuBYoaiRAsSMBKuOgTJMYLBBU1oqKJ0WA0JkpifNAQxJgQ4yX4AhGlBlSg0hSpCCjVQoUiWAoFqkILNqVAL58PazdUUtuz98zes89Z/5fM09lrrX/P+vbMmn/mzA9JkiRJkiRJdUTXAWxD3Rd4BzAXmAMcChwE7AfszyuxbgDWA88ATwKPAQ8D9wMPRMSLIw18ktOZAOruwDzgXOA9wAnA7g273QQsBe4AfgcsjoiNDftM2kSdq16pPu3w2aAuUC9S9+v6u1eLGuo56qIRTPr/40X15+p56rSu90k1qGeq93Q48TviSfWb6qFd75+uGdoaQJ0NXAW8f1hjtMAW4FfA1RFxe1dBWNZDJwEnA7OB6cBzwApgUUSs7Cq2vrEc7i9V13f4Cx+E+9VPqnuPcF/trV6urt5FbHeq540qroFRD1BvGuo0DZ916nfUI4a8r+apK/uM7Ub1gGHGNTDqweqKVqeiW7aqt6kfUKe3uJ8OVa9RtwwY10PqrLbiaWUNoO4D3AUc30Z/Y8jzwALgRmBhRKzvp7H6BuC9wAXAOUDTq5C/AqdExAsN+2lNgO8Dn26jr0nAZuBPwH3AQ8BqSmYSYAawL3AgMAs4gpLgessQ4rgqIi5r2kljAdSjgeU0z+JtYyMlm7cM+DuwBlhHyfJBSQ2/jpImPpKSNj4e2KOl8ScLm4E5EfFIk07aSIh8huaTvxq4nnJJdldEvNxPY3Uv4J3AmcDZlHsKY3OfY0hMAy4FvtCkkzaOAKsph7tBWAh8F7glIrY2jWW7mGYB5wMfpIgxVVkVEYc16aCRAOohwOMDNF0CfDkiljQZfyKoxwIfBS6inJunGrMi4qlBG+/WcPDZfX7+OcpknDqKyQeIiL9FxBeBg4ELgbtHMe4IOaxJ46YC9HN9/AJwQkT8OCJsOG7fRMRLEXFtRJxMSbv+DOhrrTEitva2iTKjyWBNBejnFPJSRDzRcLxWiIh7IuIjlCPY14GBD6EtswI4g5J3mCiNTuNNBZjURMSaiPgG5TB6PnAr/f362uIp4HPA3IhYPMqBqxZgGxGxKSJuiIizgMOBrwEPDnlYKYvhi4DDI+J7EbFpF21aJx+MeBURsQq4ArhCPY5yO/tsyrphz4bdPw/cSTnSLOiN1SkpwE6IiOWULOe31BnAiZQk0zGUI8VBlAdW9+o1eZny0OpaykOrTwMrKef2ZcDDbeY72iAFmCC9Gy939LYpQ64BKicFqJwUoHJSgMpJASonBaicFKByUoDKSQEqJwWonBSgclKAykkBKicFqJwUoHJSgMpJASonBaicFKByUoDKSQEqJwWonBSgclKAykkBKicFqJwUoHJSgMpJASonBaicFKByUoDKSQEqJwWonBSgclKAykkBKicFqJwUoHJSgMpJASpnbN4Uqu5LeQ3rXOBoyhu830x5FevM7T66EfgP8G/gSWAV8DDwAHB/lovvj84EsNTLnUepozefUl5tIsWnZlLe0XvMDv62Wf0L8Hvgdkrd3Q3tRDw1GbkAvTdwf4xS0OmNLXc/jfJC5xMp1bReUhcDvwB+GRH/anm8ulHn91HydKO6aMByqW2wWb1VvdARFogeBHVtH99rfpOxRrkInA6cPsLxXs3ulFPNT4E16o/UUzqMZyyo9SpgH+BiYIm6XL1MnbmrRqNAPYz+inE1KsBVqwDbcyxwJfCUeq16mjryqqPqbuqnKIWh99rV57ejUQHpFOAVpgMfBhYDK9TLLYUxh456OvBH4AfAa/tsvrLJ2E0rh86n1L+ZqmylXFJeR7mKWNNWx+o0Sj2izwOnDtjN4xHRb/HO/2FsEkFjym7Au3vb1eq9wC3AIuDefnMMliLX84D3ARcAb2oY3w0N2+cRoAGbgYcoGchHKVnJZyjlcaWcx/ejlKw9HHgb8FaaVx7bfvyxKB/fNhuBpZQqW49SKm89S/nCUFbwB1Iqlh8JzKGkj/cYcZzTgON7Wxdc3XTyYXwEWE05z94E/CEi+qrpq06nZP/OAM6i1PibSFp5srIU+ErXQfSbCdwRt6nnWu4LtBnX69VL1IWWDOBUYrnadO3QDg4uwBJHlIVTD1K/pD7Y0gR0yQLHJGEFDCTAWvViO0i09OJ9l/oT9YV252XoPK9e0sU+2yn2J8AG9eCuYwZQ91c/qy4bwmS1zXXqrK732Q6xPwHWdh3vjlBPUq+x/MrGiSXqvK73z05xCgiwDXWG+iH11+qmIUzoRNiq/kY9o+v9MSGcQgJsj3qg+nH1ZstzDMPmCfXb6lGj/q6jzASui4gDmozXBZay8adRcgynAm8HXtOw283An4GFwM3A3V2VlR+XRNDY0isb/9vehronr2QAjwJmA4dQspN7U9K/UCZ5PeXh1X8C/6A8vLoMWNrrt3NSgD7pZSnv622TnnweoHJSgMpJASonBaicFKByUoDKSQEqJwWonBSgclKAykkBKicFqJwUoHJSgMpJASonBaicFKByUoDKSQEqJwWonBSgclKAykkBKicFqJwUoHKaCrClj8928r9vyc5pKsCqPj77WMOxkiHQSICIeJTyz44T4cYmYyVjinqWumUX///+mNrvO3CTyYL6CfXlnUz+sV3HmAwZ9Tj1h+oj6rPqvepX85efJEmSJEmSJGPGfwGse6ut/pWuDQAAAABJRU5ErkJggg==",
    "folder": "iVBORw0KGgoAAAANSUhEUgAAAIAAAACACAYAAADDPmHLAAAABmJLR0QA/wD/AP+gvaeTAAAFmElEQVR4nO3db8jdYxzH8fdXcw9zJ9byt4mMIvcDppUUEWOTJ1KEZP41xQMSyVISpfBACLXWIp55wAP/Q4v5UzIRwtjWmsk2zZ+Nm48H51Zav9855/rd55zrnHN9Xg9/v+u6zrff9Tm/P9fvvjtgZmZmZmZmZmZm4y96MYikRcBi4CRgATCnF+NW+B3YAnwOrI+IXX36HOtE0kJJD0jaqDymJb0l6WpJc3Mfj2JIWiDpKUl/Zpr4KpskXSOpJ2c0qyHpEknbs051e29IOiL3cRpLku6W9E/mCe7GFklTuY/XWJF0b+ZJTfWzpFNzH7dR0PGaKekaYE3/S+m5zcDiiNieu5Bh1jYAkk4EPgEOGkw5PfcKsCwilLuQYbVfh/1PMrqTD3AhcEPuIoZZ7RlA0gXAqwOspV92A1MR8X3uQoZRuzPA7QOror8mgdWSOp3tilR5BpB0NLCJzpeIfW0H3gF2zrKuKgEsAs6i2VLzrRHxWG9LGlOSrk987PpL0p2SJgZQ23GS3mzwaPibWu8srBNJqxMP7rUDrm9C0tsNQrBOvhR0JumDhIP6dqYaF0ra1SAEq3LUO1IkbU44oDdlrHNFgwD8I+lhSZO56h4mdTeBPwOHdTnG0oh4rXclpZH0MrC8QdfdtG5Yt/a2or6aBn4CvgQ+iohvZztg3d10ymvV3KtsNwKf0X1g/zMJXNz7cgZH0jfA88DTEbGlyRgjf0MUEVuBW3PXkckJwCrgW0mPS0r9Eox+AAAi4jngxdx1ZDQB3Ax8IWlZSsexCMCMlbSujyU7HHhJ0h3ddhibAETEj8BVwN+5a8lsP+Chbh93xyYAADNPIyvJf2M6DO6TdGWnRmMVAICIeAZYAfyVu5Yh8JSkE9o1GLsAAETEGuAcYGPeSrKbBzzRrsFYBgAgIt4DpoAHgd8yl5PT+ZLOq9s5tgEAiIhfI+Ju4DjgLmBD5pJyua1uR91S8A7g0C4HvyAiXm9SVQ5q/a3DEuBE4EjggLwVJZlP69I2P7Hf38AxEbGtq9aSdiS8XDk/sRibBUlzJd2j1r/GpVhRNd5YXwLGUUTsjYj7aT3upjizaqMDMKJmHnfXJXQ5uWqjAzDank9oe1TVRgdgtH2X0Pbgqo0OwGhLee9ROdcOQOEcgMI5AIVzAArnABTOASicA1A4B6BwDkDhHIDCOQCFcwAK5wAUzgEonANQOAegcA5A4RyAwjkAhXMACucAFM4BKJwDUDgHoHAOQOEcgMI5AIVzAArnABTOASicA1A4B6BwDkDhHIDCOQCFcwAK5wAUzgEonANQOAegcA5A4RyAwjkAhXMACucAFM4BKJwDUDgHoHAOQOEcgMI5AIVzAApXF4CUHyOa6EUh1siBCW2nqzbWBeCXhIHPSGhrvbU4oe3Oqo11AfghYeDrJc1LaG89IGkSuC6hy/dVG+sC8EnCwEcDayTtn9DHZkHSBLCGml8DrfFpygcsT/xlakn6UNJFM8VZH8z8cvgySR83mJ+LqsaMmg86ANgGHNKgzmlgd4N+1tkkMKdBv53AkRGxd98dlYNFxB5Ja4FbGnzYHODQBv2sf9ZWTT7UnAEAJB0LfAXM7VdVNhB7gJMiYlPVztqFoIj4AXi0X1XZwDxSN/nQ5gwAIOlA4CPglF5XZQOxAVgSEXvqGrRdCo6IP4BLqVlEsKG2A7is3eRDF+8CIuIrYDlpq4OW1y5geUR83alhVy+DIuJ94GzSVggtj43A2RGxvpvGXb8NjIhPgdOAFxoWZv33LHB6RGzotkPS6+CI2BERVwDnAu8mFmf98w6tb/3VEZF0v9b2KaATSVPA5cBSYIpmq1SWbprW2v6rwAsR8VnTgWYVgP+beQdwPLAALx71y17gJ+C7iPgzdzFmZmZmZmZmZmZmZmZmZja0/gUcFq7BhjcYPwAAAABJRU5ErkJggg==",
    "gauge": "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",
    "git-branch": "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",
    "globe": "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",
    "graduation-cap": "iVBORw0KGgoAAAANSUhEUgAAAIAAAACACAYAAADDPmHLAAAABmJLR0QA/wD/AP+gvaeTAAAKDklEQVR4nO2de6xdRRWHf6u1tKBQKWAhUktbWhEM5dVWsMhDULQQ8EEIJlSpSoImEJXKKyQlUVMTDVBJwAgoiE2BpIAGRSoQoRrkVaBQaQu0tqXQB719kb77+cc6x9tc76Vnn7NnZp9750vuH03P2bNmzTp7fvNYM1Imk8lkMplMJpPJZDKZTCaTyWQymUwmk8lkMplMJpPJZDKZTKbtsdQGxAIYJOlsSWdIGitpuKQhtf9eK2mFpHmSnpD0uJltTWFnpmSAI4FbgfU0zvrad0altj/TJMAJwD3AjgIN35VdwJ+AU1LXJ9MAQD/gPGBOC43eE88Dk4H+qeuZ6QKwT61xXgvQ8F15A7gS2Dd1vfs8wOBaY7wdoeG7sgqYBhyU2g99DmAkLtLeT9DwXdkMzABGpPZLr4dyhF0o6oLx5NR+6lUABpxVc267MBe4kCwYm4e4wi4Ui6m4YKzcTCBwgKRLJU2V9PHE5pTFakm3SZphZutSG7MnlQkAXET9QNIUSR8OXNwySXdKel7ug3HyoPtE4HI3S7pD0s1m9p/AZbUHwPHEE3Yv4d3KgG7sqE8i/SuCHXXBOCGFz5NDfGE3F2/cht54wMSabburZltbQ6ewezWCY7cD9wPjWrD3WODXwJYI9i7CBeOgMn1eCYADapVbEcGRG4FbgGEl2n8oPuu3LoL979bKGrJ3yyoOcAQwnWJLsc3yTs1xBwasz/54IC+LUJ9NeCCHFqblAxxHPGH3MnAZMDBi/QbgEz3PRqhfXTCOj1W/pqFTPMWgEuKJvi4Y6RR28yM4oGVhFwpgLP7W2x7BD/W3XjrBSKewWx6hwnVhV/n+EDgM1yIdEfwSXPd0V8HhuLDrnRUsCdL8QEob+XRXoaOAmcQTdpOBfYJVKBJ4F3kpceY+tgG/A0aXWQEDridO3zYH+CJVEjklUfPjl4EnIvhxG3Bty34E+gO/D2zsduBe4PiSfF15gJOAWYR/m95LK3sSgF8GNG5j7fmVF3ahAEbgfffmgH6+tVnjziTM2PZt4GrgoyX7s20BhgDX4aI3BOcWNciAeSUbMR/4Fr1A2IUCGAh8B/h3yb5fSJGuADitxMIfB75ELxR2oaBzb8LfS2yHxt8CwE0tFrYDHzKeENBPfQJgPPAAsLPFNnmwSKH/bLKQTXjwDA/okz4JMIrWciE6ihRWdN1+JXANWdgFBzgYuAHPTCrC7iKFrCnw4PXAV8l9fFTwVcii0/GNCUE8AbIorwFTiLgu39fAR2eTaG42sdAboJU1/ZX4NGTbLeBUFXx4OIXWkmR2FCnwRy0UVCcLwhah3AmiLUUKPpzyFn/ykLAgdE4RbyqpDQA2FDXijhILr/M34ByyYOwWYBxwH62P+bujWEoa8DGKDzUa5RV6yXp/qxBm1q87iuckAqcDWwMatRyYiieE9imAQcB3KX/evyeaS0rFN2hsCGzcDuBhQm5pqgj4RM4sfNNGTJrPSsbP2nsygpG7gD8Ax5bo80qAT+XeRpj+vRFaT0sHzgYei2DsbuBR4KwSfJ8UYALlLOa0SnnnEuBJk7H2wPeY0l1VCHteYbOUfzAFnUmTMbaIr6yVVdkFJ3zGbjKwIII/tlNs5BDuZBLiJk1uwCdJDg9WoYLgwu5q4pxX+L8kGbxLbpTwR9MQN01sG94NHRO8Yj3XdyTeGDHOK/y/JBmqFgBdnBM9UTRi3U7Egy+GsOsx+5kqB8AeRtZTxXcFcE5XXsTfQB8KUI/Ywm6vmcC0QwDsYezMIG7qniW4Jmn5tDHiCrs6sxu0reUA6NeqgwqwKWJZR0i6WdLbeB9d+LxB4BDgaklLJN0t6VPlmviBrIlVUMwASMFgSVdIegvvho7e2xfwGbtbJC2VNF3SYWFNTEtvD4A6+0i6RNJ84CFgYtcP4KL1QUmL5EGzX2Qbk1C6WCoJFOYU036Szpd0PvC6pOfUeVLoJwOUV3mqGgD3SeqQ9E2F+yUeVfsLzXZJMyV9RNLXI5RXiKp2ARvN7Hvyq92mKaIoKpH1cg0xwswulVSpQ6LrVDUAJElmttbMbpQHwuWS3khsUiMslR96PczMrjWzlYnt+UAqHQB1zGyLmd0u76e/JumZxCZ1xwuSLpY02sxuNrPNqQ1qhLYIgDpmttvMZpvZyZJOlfSwpMYTHsoHSY9IOtPMTjKzWWa2M6E9hWmrANgTM5trZhdIOlrSbyTFvOp1m6S7JH3azM41sycjll0qbRsAdcxsoZldJp/9+4mk9wIWt07SzyQdYWbfNrMFAcuKQtsHQB0zW2VmN8gF4xXyKdyyWCLpSknDzex6M3u3xGcnpdcEQB0ze9/MfiVptKSL5NfCNMtztWeMNrMZ7SLsitDrAqCOme0ys/vNbJyk0+QLOusb+GpH7bOnmtn42jN2hbQ1JTFnAinw2VLtMrOnJD2F7xM4QdJYeVdxkHwqeI38IqmXJM0LpOSLZEE16qvtBZ7ZbXZwzAAoshx8GtC/7F9erWGfrf1FoxZ4pxf4ysYGP7e0wDO71UQxu4BlBT47StIPQxmSgB/LRymNsrSRD9WunnuxwWc+VKD88gFOKbgrZgdwYVKjSwC4mOJ7Bxu+Tg74PHvfbvcmsH/IejZi6ACKX7q0E08ebbt0cvw4l2so3vhrKbinET89pKdcw7eAaix141ewNcNjwJGp7W8UYAzNnwre1Nm++HH+t+Ongq4FnsFzFULfwto4wDE0vzt4K37kzKGp69ET+G0hM2g+83cnMCZ1PYIC3N2kc+psAe6iQteuAp8BfkvrZyncGdv26H0rMFTSq5IOLuFxCyU9IOmPkl4wsygrg0A/SSdKukC+y6eMX+1q+eJS1M0vScQVfnDxwyp3GPqepKcl/UO+Nv9yWVe147d4Hidv9M/Kl6LLvNlzt6RzzewvJT6zIZKpa3zP/fTAxayStFg+CfJO7d8d8mvcN6hzL0F/SQfI9+0dKGmofDv4CPmawtDAdk41s18ELqNbUl+6+FNJ16W0oQLcaGbTUhWefHwNXCXp5+rFC1M9sEv+y78ppRHJA0CSgEmS7lG5/WqVWSvpEjN7NLUhlfjVmdkj8hW6v6a2JQJ/ljS2Co0vVSQAJMnMVpjZOZK+IWlFansCsFzSRWY2qepbxZMD7IevAaxucWKlCqwCrgL2Te3XtgMPhO8DrydtwuZYAFxOxRu+EiJwb+CrgZ+T5wp+RVJVTwrrkDRbLmifNrMiu6CS0BYBsCf4AdNnSjpP0hckpV4lXCxpjnw6+kkzK7JNKzltFwBdwc8Xnihpgnyq9lj5rF4INkh6Rb4L5xlJc82srQVr2wdAd+D3EY+RNFLSMEmHSzpEvgl0sKSB8infPemQZxdtlK8rrJGPRpZLelPSIjNbHsP+TCaTyWQymUwmk8lkMplMJpPJZDKZTCaTyWQymUwmk8lkMplMphD/BUqcGggUizckAAAAAElFTkSuQmCC",
    "handshake": "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",
    "heart": "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",
    "home": "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",
    "image": "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",
    "info": "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",
    "key": "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",
    "layers": "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",
    "lightbulb": "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",
    "line-chart": "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",
    "lock": "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",
    "mail": "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",
    "map-pin": "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",
    "message-circle": "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",
    "message-square": "iVBORw0KGgoAAAANSUhEUgAAAIAAAACACAYAAADDPmHLAAAABmJLR0QA/wD/AP+gvaeTAAAEiUlEQVR4nO3bTWgcdRjH8e9TmlgsIlrWghURweZSREHwrNjoxbebiiCGehBRQRARjCKCB8GbKNgSvJQeetEWtBGUUEGQHkQQ6ktVKh6USAVbtdrWn4fNRTqz2f9mZieb5/c5zv7nZft8s9mdbMHMzMzMzMzMzMw2vmjiIJJuAG4BZoAesLmJ49pFzgHLwNfAsYj4bq0HHDkASdcCjwMPANet9UJsJCeAA8DbEfHTKAcoDkBSD3gFmAOmRjmpNe4fYB8wHxGnSnYsCkDSPSsn6pXsZ2PzCzAXEe8Pu8OmYRdKeh54Fw9/PdsOHJb07LA7DBWApJeAV2noTaO1ahPwmqT5YRavOlBJjwDvrPGirBsPR8T+QQsGBiBpJ/A5cGmTV2Vj8wdwU0ScqFuw2q+At/DwJ9lW4M1BC2oDkDQL3N70FdnY7ZZ0R92Dg14BnmnhYqwbtbOsfA8gaQfwIwUfE1csA0vAb4X72XCuBG4DthXudwG4JiJ+Hmq1pD0qc07Sc5KmCy/MCkm6RNILks4Xzmiu5CQLhQd/tMXnbBUkPVY4o30lB/+s4MBL7T1NG0TSJwVz+rTqGHW/468uuI4D5ZduDSn5t99RtbEugJLP/j8UrLVmfV+wdmvVxroASu75q2CtNetCwdrKWZd+zLMNxgEk5wCScwDJOYDkHEByDiA5B5CcA0jOASTnAJJzAMk5gOQcQHIOIDkHkJwDSM4BJOcAknMAyTmA5BxAcg4gOQeQnANIzgEk5wCScwDJOYDkHEByDiA5B5CcA0jOASTnAJJzAMk5gOQcQHIOIDkHkJwDSM4BJOcAknMAyTmA5BxAcg4gOQeQnANIzgEk5wCScwDJOYDkHEByDiA5B5CcA0jOASTnAJJzAMk5gOQcQHIOIDkHMNl6BWv/qtroACbbbMHa5aqNDmBCSboReLBgl2+rNjqACbQy/MPAdMFux6o2bm7kitZI0jRwPbAdmOr4ctazHnAn/Z/8kuEDLFZt7CwASTPAQ8BdwM148G06HhFfVD0w9gAk3Qq8TP8NTIz7/Em9UffA2AKQdBnwOrAHD36cTgILdQ+OJQBJO4FDwMw4zmf/81REnK17sPUAJO0CPgKuavtcdpG9EXFo0IJWA1gZ/seU3bGyZhwFnlxtUWv3ATz8Th0F7o6Iv1db2EoAHn6n9gKzEfH7MIsb/xXg4XfmJPB0RLxXslOjAXj4nThO/3P+wqB3+3UaC6Ch4Z8HTjdzRRvSWfp/1fuG/r39xbo7fMNqJIAGhv8lMA8cGaViG10TAewC9jP68BeB+yOi8gsL1q7KW7KSTgFXDHmMfxn908QicJ9/6rvTxMfAUY/xIf2ffA+/Q119IWQRuNcv+93rIgC/7K8jdQGopfN5+OtMXQBD3UYs5OGvQ3UBfNXweTz8daougMovEI7Iw580knqSzmjtjkja0vXzsRFIetHDT0zSlKQlDz8xSZePEMFBD38DWXklmJd0epXB/yrpCUn+yvcEGXpYkrbR/y9Ju+l/vXsL8Cf9P+V+AByMiDNtXKSZmZmZmTXoP8oLREgjsokPAAAAAElFTkSuQmCC",
    "network": "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",
    "package": "iVBORw0KGgoAAAANSUhEUgAAAIAAAACACAYAAADDPmHLAAAABmJLR0QA/wD/AP+gvaeTAAAKwUlEQVR4nO2da6weRRnHf38EhFYxMSSILVdBbhGRywEqlatAIvgJiH6gKqA1MbFqDGAiikZRQAn1g4ZKjNQEIsSQCGhMOGlruLYUjKQtF8ECvaHFAtVeoO3fD7PLKaen7bv7zuzuOe/8kn7oOXtmn5l53vf5zzO780Amk8lkMplMZuBQ2wY0he19gHOAc4GPAwcB+xe/Xgu8DPwNGAbmSdrchp2ZyNg+xPZs2+vcO+ts32r74Lbtz9TE9n62f2p7U4WJH82moo392u5Ppkds72H7Kttr+pj40awp2tyj7f5ldoHts2w/FXHiR7PE9gVt9zMzCtsH2Z6bcOJHc5/tj7Td74HH9mTb19ve2ODkl7zlIC4/0PY4DBxOE+frkvVBkzh9nK9L1gcpcfNxvi5ZH8TE7cb5uowbfdDZVLBDTL0C+BFwQOTmVwG3A4uK/58KXAkcGPk+rwLfBX4jaVvkticutr9qe22CT+YGh8ze+8e45yTb19hen+C+WR/0gkPefmmCCbBDbD60BxumOGiNbYlsyPpgNB6J828nGPQnbJ9Rw6ZTbD+cwJ5xow+S47Cen2F7dYKBXmn7K+5jfW5bti+1vTyBfWttz7L9nphjOm5wuvX8TuN8H7ZmfRALp13PD7uHON+H7Vkf1MXNrOc32r7BET/9O+lL1ge94rRxfmessv0lJ8zPO+uD3eP28/ZP2J6euI9ZH4zG3cvb/94JtUHR56wP3O28/UbbP7b9vsRjMHj6wN3an98dqxzyA8niq4M+mGF7RQL719i+0nY39nEcvu4XJOjoBts/sX2d08TXxU6vDybb/mHRl9jMsz01pf29dHDI9quRO7bN9p3e7pl8p4+vhyUep4Nt35XA/tW2T0pp+646NWT7jcgdetz2tF3cM1V83ewQX5M+/297mu2FkW1/3fbJKe0eqyMHFzeOxQrbl7uHuOaR9fdLEe9fUu4fjDd9sM5NhQMHwRdr8Dc4xMjJNeyY5LDiSBFfm8gfxNYH/3QTD6ba/nkEY3eI833YM9VZH5TcnNJWbO/r/t61s3cT5/uwbcj2I/2O4BhE313cif0x9MFG2+9NaeSVfRi31fZvbe+Z0L4m9EHK/YU9izHa2oedX0xlH7YfijCQf7d9bjIjaUQfVH7CqAebzy3Gpl/mx7atNPCDtrdEMLDkXttHJDF2xObU+uDQCDYeUYxFLN52inSx7c9ENLJks+2bnH793Tl94HB+wU3FGMTm/BSD+O0EhpYkf7/OI+vvVQns71kfuJl9k2+mGMBbEhpc8qTtM6Mb/+5+pNyxXORd6APbZxZ9TE385aDt2xowvOQep9+/L59ZiK0Pttm+e3v7bR9a9KkpftXrOHT1leZLgGedMD8v6RVJM4DTgcdiNg1cCixxSJrdACwj9Kkpep7XrjoAwN7A14FnnDA/L+lxYBrwBWBNxKYnAd8CvgPsE7HdqKRygLcitnUgcBuw0Iny85IsaS7wUeBGoItnBG5J0WgqB7gbmAPEfCP2ROCvjrT+HgtJ6yVdCxwL/CHFPWpg4B7CmEYnlQNskDQTGAIeitz2RcBSJ8zPS3pR0iXA2YTTQ9tiEXCGpMuA/6a4QVINIGmxpOnAZ4HlEZveF7gGWOaE+XlJ84GTgC8T3vVvipUETXKqpEdS3qgRESjpPuA44FrievIUgj54zAl2GAEkbZN0O0Ef3ERafbCRoEGOkTRXkhPeC2hwFSBpg6QbgaOB3xFiWyxOAR4q1t+HRGz3HSS9CcwG7iWu7SX3A8dKulbS+gTtj0njy0BJK4v196lAzK+3cv291CHTt2+shm3vbXsW8AzwOeIerbMYmC7pYknLI7bbE63lASQtAs4ALiMc1R6LScD3gecccv99TZbtiwkTfysQU3SuBmYCQ5JiC+WeaTURVKy/7wGOAX5AiIGxmArcQdAHp1X9Y9sn2l4A/BGI+VjYVuAXwNGS5rR9eFQnMoGFPrieILRi64Mh4BGHvP+Hdnex7f1tzwYWAp+KaAeEldAxkmYVmqJ1OuEAJZJWJMzPXw78o9AHO6Rmbe9VxPkXCCnomKnn14CvSTpM0vMR2+2bTjlASZGf/yTx8/OT2U4flD8s4vwyQpyPufn0GvAN4ABJv4zYbjQ66QDwzvp7LnAUcDNx198HAXc4HCkzTIjzMV+/3kyw+XBJsyVtjdh2VDrrACWS3pR0NSP6ICbnFP9icj9wnKSruxLnd0XnHaBE0suFPjgPeLpte8bgaeC8Yj3/QtvG9Mq4cYASScPACQR98K+WzQH4DyHOn1jYNq4Ydw4AO+iDtvbv3yZseR9VxPkk+/WpGZcOUCLp9WL//mOEPfOmeBD4hKSZktY2eN/ojGsHKJH0fLFnfiFhOZeKZcCFkj4taUnC+zTGhHCAEkl/AY4n5Nj/HbHpdYSt7BOKe0wYJpQDAEjaImkOI/qgn+cTtzAS52+UFPNZx04w4RygRNK6Qh8cDzxQo4lhgrKfKSnmt0mnmLAOUCLpWUkXUc0JHpR0nqQu5huiMuEdYDtWVrj2xWRWdIxBcoDMGGQHGHCyAww42QEGnOwAA052gAEnO8CAkx1gwMkOMOBkBxhwsgMMONkBBpzsAANOdoABJzvAgJMdYMBJ5QDZseJTpdBGz2cOVJmo/1W49vAK12Z648gK1/Z8xlAVB1hV4drpjlAQKhMoDsY8vcKf9DxXVRxgaYVr9wLmOGF9oEHB9l7Ar6kWAnqeqyoO8CjhfJteuQB4IH8T1KcYuz8R3ojulS3A471e3LMDSFoHPFzBEIDzCad91yoOOai4KCpJOJ2syuQDLJD0Rq8XV1Xrd1S8HsKxrtcRzv/v+9i2iYxDWZvLgWcJY1bnrMM6c9QbtvdxqI/TDwtTHeu6G9urVDy5rQX7pjkU1eyHV1yxcGSlbwBJm4DvVerZjpTHut7lrA/K0rF3Ek5VH+qzueskpT0rwaHq1fw+PbWkdvHoGnZ36hvA8YtHD7up8OpQkHF1JMPtUEo9qT5wRxzARZx33PLxK21/uI49tVK2klYQCjf0rDZ3wxRGjnVtXB80RdG3x4C5hD7H4HXgIklVEnXvUDtnL2kx4Yi11XXbGIMhJqA+cNw4vz2rgLMlPVW3gb42bSQ9SRB18/ppZ3SzhCPZx33+wO9ez3+euMfMDwMnS2qzpE2giGtXOE051Cj6wA1qAKeJ8yVrHMa6e/kUpy3L+oT7KBvnhhzA9im2H07Q/7ccCmnGrwweG6cry2qHsnGVz+93YgewPSVxn2OeZdwMtk+z/WiCAalctt2JHMD2JNvX2F6foJ9LbF9Qb/Q7gkPiaIbj5g1KqpRtj+oADnH+UtvLE/Rrre1ZTlQutxXcsj5wRAdwjvP1cUv6wBEcwDnOx8MN6wP34QDOcT4NblAfuIYDOMf5ZnAD+sAVHcA5zjeP0+qDlypc+1wiGwYrztfF6fRBWwxunK+Lgz64ymn2F5pijUMf8ptSdXFafZCKHOdj47T6ICY5zqfE3dUHOc43hbulD3Kcbwvb+zlk/Ta1MPGbinvHrDWcqYOb1wc5zncR22fafjLhxOc433WcZn8h5+3HGx7RB/3kDzY6x/nxjYM+uMX2axU/8T+zPbVt+1PTvUeLE2F7b+Bs4CzC4RVHApOKX28Angf+DMwnvGM/4YpEZjKZTCaTyWQyAPwffKRm0g6K4rMAAAAASUVORK5CYII=",
    "phone": "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",
    "pie-chart": "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",
    "puzzle": "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",
    "refresh-cw": "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",
    "rocket": "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",
    "scale": "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",
    "search": "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",
    "send": "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",
    "server": "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",
    "settings": "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",
    "share-2": "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",
    "shield": "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",
    "shield-check": "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",
    "shopping-cart": "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",
    "sparkles": "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",
    "star": "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",
    "target": "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",
    "thumbs-up": "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",
    "timer": "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",
    "trending-down": "iVBORw0KGgoAAAANSUhEUgAAAIAAAACACAYAAADDPmHLAAAABmJLR0QA/wD/AP+gvaeTAAAEt0lEQVR4nO3dTYtcRRjF8VPJaKIyQQJxkYWzEQ1MstCtoCRBl+LCL2DAvaDgfAjXviAu3TuYhROIOyEQspJAREYUgjoSNRtJBjP8XXQ3zHT65b5U3ap7+/xgNs2k6qnU6erqut13JDMzMzMzMzMzMzMzMzMzMzMzMzMzMzMzMzMzMzMzM+uBUOWXgOclvSDpL0m3QwiPklZlZQAuAbc4ag/YAtZy12cJAe8CB8x3FTiZu05LADgHPFww+RM7DsEAAZ9XmHyHYKiAOzUC4BAMDfB7zQA4BD11bM7jvzZo601J2w5Bv8wLwNcN2xtMCIBTwKvjn1O56+kUsA7sNngZmPi2ryEAzgJfAfuHxrM/fuxs7vo6A2wyOvRZmRCMx3x3wZjuApu56+wMo/OA31qE4BrwVO5xVDEea5XN75/Ahdz1diZCCHZKD0GNyXcIhhaCBpO/siE4z8D2BBHGtAeczz2OzjCglYDmz/xpK7cS9D4EESd/wiGoKVsIiD/5EysXgt7tCSLUvMxK7gnaPJs6Wwki1FrVyq0ExYcgQo11OQQ1JQtBhNqaWrkQFLcniFDTzvinqc72BIyC/hnwI3APuAF8BDzTRf+HiyhiJYhVC3AC+KZFO8lXAuAKR69eHvYz8FLK/qeLyR6C2DVQcAiAyyz+9DaMLu2vp+h/XlHZQpCqbwoNAY9/b2Oerdh9Lyus8xCk7pPCQgBs1Oj7Rqx+6xTYWQi66ouCQgC8XqPfvRh9Niky+cR00cdUf0WEAHijRp9/t+2vTaHJJihl20vGlD0E9CUA42KjT1SKNmuOKWsI6FMAxgVHm7CYbbUcU7YQ0LcASNEm7uUIbUQ7eiZTCOhjACQJuDAedFPLDj4WSfIVNuAkHR8bEyEA874ZlFQI4QdJr0n6o2ETTeu+JuntEMLDhv9+rnGbb0m62rCJ5yR9R8cXkLIEQJJCCHckXVTzENQ1mfwHqToIIexLekfNQ3BG0vUuQ5AtAFKnIUg++RN9DEF2tN8TLJLla+t0sCegr5vAWUjzwY2sn0Am8bsDhhQAKXoIsn/8fDymZCFgaAGQor0cFHW3EhK9HDDEAEitV4IinvnTSLASMNQASI1DUOTkTxA5BAw5AFLtEBQ9+RNEDAFDD4BUeU9Q1Gv+MkTaE7AKAZAk4EXg5oxBHQCfACdy11hXpBC81zYAle4WXgLgmKRLGp0cnpa0K2k7hPBT1sJaYLRqbWt0d7Um7kt6tuLv/hNCON2wH0uF9nuCVitA1msBFuXaQSsOQAFyhsABKESuEDgABckRAgegMF2HwAEoUJch8B9+igQ4J+l9jc4pzkRs+j9JT0Rs7wgHIALgiqRPJT2Zu5a6HICWgMuSvlD5L6czT317cxRcKuCWpFdy11HBgxDC09MPOgAtABuSfsldR0WPQgiP7SVKX7ZKt5G7gBpmzrUD0M793AXUMPPvPTsA7dyWdC93ERXtznrQAWghhHAg6ePcdVT0Ye4CBgk4Dmx3cD2/jS9z/z8NGrAGbJH2TuRN/At8sKh2vw2MCFiTtKm4R8F1HNfoncm6pO9DCN3fGs7MzMzMzMzMzMzMzMzMzMzMzMzMzMzMzMzMzMzMzMxS+B9VDtQWPkbC9wAAAABJRU5ErkJggg==",
    "trending-up": "iVBORw0KGgoAAAANSUhEUgAAAIAAAACACAYAAADDPmHLAAAABmJLR0QA/wD/AP+gvaeTAAAE3ElEQVR4nO3dz2sdVRjG8ec00VY0RYqCdGEXFoymXbl0UUxNF7rwb1DoXhC0f4m4celOwZQuTEBEoVAQ3aioSIVuWpP6ayM0/vq6mLkSx3tz7505M+fMzPOBbobmnPfmeefOnJm5N5KZmZmZmZmZmZmZmZmZmZmZmZmZmZmZmZmZmZmZmZn1QEhdwJAAq5I2JD2aqIQVSWckrUm6HkK4kaiOcQFWgSvAHnn5DXgt9e9n0IAVYDtx0PO8nfr3NFjAG6nTXdAL0+r3OUADwIqkHyQ9krqWBXwdQni6uvFYikoGZEP9CF+Snpi20Q3QzMOpC1jC6rSNboBmbqUuYAl/T9voBmgghHBL0uep61jQH9M2ugGae10z9q7MHEzb6AZoKITwoaTLkn5PXcscTNvoZWAkwLqkVyU9p7iXgh+SdF+EcX4JIZyKMI51ATgOXIt4IejnafP4EJAh4Lik9yS92PZcboDMdBm+5AbIStfhS26AbKQIX3IDZCFV+JIbILmU4UszbhDkCDgmaVPFOvuUpJuStkMI3yUtrAHghKRtSZdqDvGr+nVDqh7gLPDplLXtX8CbwP2pa1wWcALYabCu3wMuL/H/p14HyB5wnvnP2n1AsTf1AnHCPwdsLfEz/WsA4CngzoIvcBd4IHXN89D8Ct8+cL4ca7gNwHLhT2TdBEQMvxxvmA1AvfAnsmwCIodfjjm8BqBZ+BNZNQEthF+OO6wGIE74E1k0AS2FX449nAYgbvgTSZuAFsMvxx9GA1As9fYb/KKOskOCJSKRlnpz5mjcAMkvBVM8SbOr9j5QeUnSVTp8J6C4vPuu6l/huyvp+RDCl/Gqmi5pA5ThfyTpsZan2pK03UUT0Pza/l1JF0MIX8SrarZkDdBh+BOtN0HfwpcSNQDFic0nqh9+3cewtyS938Y5QTnmVdUPf1/SZpfhSwkaIMIxf1fSM5Lu1Pz56OcEfTrmJwWs02yp9+8ZfTnW7QZjRVki0vJSb87c/VkGEjH8ypjJmoCE4Zfz96MBaCH8ytidNwGJwy9ryL8BaDH8yhydNQEZhF/WkXcDECf8hYKhoyYgk/DLWi4sMe9ejDmXKa6z8CtzttYEZBR+Wc+ZJebu7ivjSBB+Ze7oTUBm4R+q67MF578Se+5ZBSULv1JDtCYg0/DL2i5SPBx7lJvAWhvzV4tJHn6llsZNQMbhH3qtrwAHM+b/HniyzfknRZyj2TdlRn+yN0JNO7R8Szfia10H3gK+BX4EblB8h+GDXU2exZ4/o7Ym7wR1tb7nZ4GMw6/U2GUTOPwFtR5+pdYummA04Wd3zO+g5nk6O+YnRY/2/Bm1t/FOMJo9v7fhV15DzCZw+AtKHv4E8ZpgNOH37pg/T4TX5GP+grLZ86uo/04wmj1/sOFP1GgChz+U8CeWaILRhL/BwI7581CcExzVBLcZyTF/jeIW4mjCnwBOA+/w37trB+W206nr6wTF38CrK8mHMWMDTgLPlv9Opq6nUxS3Ekcb/uhR78TP4Q8F8I3DH4dZnw38eIkxdiW9FEK4F6EeywHFmvie9/wRA17m6KdNrzn8gQM2+f9z53sUy8TefNG0zbbQXw0DHpd0VtJPkr4KIfzZalVmZmZmZmZmZmZmZmZmZmZmZmZmZmZmZmZmZmZmZmZmdtg/CTHncu0305IAAAAASUVORK5CYII=",
    "trophy": "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",
    "truck": "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",
    "users": "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",
    "wallet": "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",
    "workflow": "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",
    "wrench": "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",
    "x-circle": "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",
    "zap": "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",
}  # baked Lucide icons (white glyph, transparent)

_ICON_ALIASES = {
    "chat": "message-circle", "message": "message-circle", "comment": "message-square",
    "people": "users", "team": "users", "user": "users",
    "graph": "bar-chart", "chart": "bar-chart", "bar": "bar-chart",
    "analytics": "activity", "metrics": "activity", "data": "database",
    "ai": "cpu", "secure": "shield", "security": "shield-check", "privacy": "lock",
    "ok": "check-circle", "approved": "check-circle", "done": "check-circle",
    "warning": "alert-triangle", "info-circle": "info", "time": "clock",
    "schedule": "calendar", "fast": "zap", "speed": "zap", "quality": "award",
    "premium": "star", "money": "dollar-sign", "price": "dollar-sign",
    "revenue": "dollar-sign", "growth": "trending-up", "shipping": "truck",
    "company": "building", "goal": "target", "idea": "lightbulb", "launch": "rocket",
    "process": "workflow", "gear": "settings", "config": "settings",
    "cloud": "cloud", "global": "globe", "world": "globe", "doc": "file-text",
    "document": "file-text", "book": "book-open", "flag": "flag",
    "shield-check": "shield-check", "puzzle": "puzzle", "layers": "layers",
    "scale": "scale", "balance": "scale", "lock": "lock", "key": "key",
    "eye": "eye", "star": "star", "heart": "heart", "trophy": "trophy",
    "check": "check-circle", "phone": "phone", "mail": "mail",
}

_ICON_FALLBACK_GLYPH = {
    "check-circle": "\u2713", "check": "\u2713", "star": "\u2605", "heart": "\u2665",
    "shield": "\u25C6", "shield-check": "\u25C6", "lock": "\u25AE", "target": "\u25CE",
    "bar-chart": "\u2759", "trending-up": "\u2197", "trending-down": "\u2198",
    "zap": "\u26A1", "rocket": "\u25B2", "lightbulb": "\u25CF", "flag": "\u2691",
    "clock": "\u25F7", "calendar": "\u25A6", "dollar-sign": "$", "users": "\u25CF",
    "cpu": "\u25A3", "cloud": "\u2601", "globe": "\u25CB", "settings": "\u2699",
    "award": "\u2606", "trophy": "\u2605", "layers": "\u2263", "scale": "\u2696",
    "workflow": "\u21C4", "arrow": "\u2192", "info": "i", "alert-triangle": "!",
}


def _icon_name(raw: str) -> str:
    if not raw:
        return ""
    n = str(raw).strip().lower().replace("_", "-").replace(" ", "-")
    return _ICON_ALIASES.get(n, n)


def _icon_png_bytes(name: str, color_hex: str, px: int = 220) -> Optional[bytes]:
    """Return a tinted PNG of the baked icon glyph, or None."""
    if not _HAS_PIL:
        return None
    b64 = _ICON_PNG.get(name)
    if not b64:
        return None
    try:
        raw = base64.b64decode(b64)
        glyph = Image.open(BytesIO(raw)).convert("RGBA")
        glyph = glyph.resize((px, px), Image.LANCZOS)
        r, g, b = _to_rgb(color_hex)
        solid = Image.new("RGBA", glyph.size, (r, g, b, 0))
        alpha = glyph.split()[3]
        solid.putalpha(alpha)
        out = BytesIO()
        solid.save(out, format="PNG")
        return out.getvalue()
    except Exception:
        return None


def _draw_icon(slide, theme, name, x, y, d, *, circle_fill=None, glyph_color=None,
               idx=0):
    """Draw an icon-in-circle at (x,y) with diameter d (inches)."""
    name = _icon_name(name)
    circle_fill = circle_fill or theme["accent_soft"]
    glyph_color = glyph_color or theme["accent"]
    _oval(slide, x, y, d, fill=circle_fill)
    pad = d * 0.28
    png = _icon_png_bytes(name, glyph_color)
    if png:
        try:
            slide.shapes.add_picture(BytesIO(png), Inches(x + pad), Inches(y + pad),
                                     Inches(d - 2 * pad), Inches(d - 2 * pad))
            return
        except Exception:
            pass
    # fallback: glyph char or first letter
    glyph = _ICON_FALLBACK_GLYPH.get(name) or (name[:1].upper() if name else "\u25CF")
    tb, tf = _textbox(slide, x, y - 0.02, d, d, anchor="middle")
    _add_para(tf, glyph, first=True, size=int(d * 26), color=glyph_color,
              bold=True, font=theme["body_font"], align="center", space_after=0)


# ============================================================================
# Image pipeline (base64 / URL / Unsplash / AI-gen → resized bytes)
# ============================================================================

def _cover_crop(raw: bytes, target_ratio: float, max_px: int) -> Optional[bytes]:
    """Center-crop to ``target_ratio`` (w/h) and downscale. Returns PNG bytes."""
    if not _HAS_PIL:
        return raw
    try:
        im = Image.open(BytesIO(raw)).convert("RGB")
        w, h = im.size
        cur = w / h
        if cur > target_ratio:
            nw = int(h * target_ratio)
            x0 = (w - nw) // 2
            im = im.crop((x0, 0, x0 + nw, h))
        else:
            nh = int(w / target_ratio)
            y0 = (h - nh) // 2
            im = im.crop((0, y0, w, y0 + nh))
        if im.width > max_px:
            im = im.resize((max_px, int(max_px / target_ratio)), Image.LANCZOS)
        out = BytesIO()
        im.save(out, format="JPEG", quality=86)
        return out.getvalue()
    except Exception:
        return None


def _decode_data_image(val: str) -> Optional[bytes]:
    if not isinstance(val, str):
        return None
    s = val.strip()
    if s.startswith("data:image"):
        try:
            return base64.b64decode(s.split(",", 1)[1])
        except Exception:
            return None
    # bare base64 (heuristic: long, no spaces)
    if len(s) > 256 and re.fullmatch(r"[A-Za-z0-9+/=\s]+", s):
        try:
            return base64.b64decode(re.sub(r"\s", "", s))
        except Exception:
            return None
    return None


async def _fetch_url(url: str) -> Optional[bytes]:
    if not (_HAS_HTTPX and isinstance(url, str) and url.startswith("http")):
        return None
    try:
        async with httpx.AsyncClient(follow_redirects=True, timeout=25) as c:
            r = await c.get(url)
            if r.status_code == 200 and r.content:
                return r.content
    except Exception:
        return None
    return None


async def _unsplash(query: str, key: str) -> Optional[bytes]:
    if not (_HAS_HTTPX and key and query):
        return None
    try:
        async with httpx.AsyncClient(timeout=25) as c:
            r = await c.get(
                "https://api.unsplash.com/photos/random",
                params={"query": query, "orientation": "landscape"},
                headers={"Authorization": f"Client-ID {key}"},
            )
            if r.status_code == 200:
                data = r.json()
                u = (data.get("urls") or {}).get("regular")
                if u:
                    return await _fetch_url(u)
    except Exception:
        return None
    return None


async def _ai_image(prompt: str, request, user_dict) -> Optional[bytes]:
    if not (_HAS_OWUI_IMAGES and request is not None and user_dict and prompt):
        return None
    try:
        user_model = Users.get_user_by_id(user_dict["id"]) if _HAS_OWUI_FILES else None
        res = await _owui_image_generations(
            request=request, form_data=_OwuiImageForm(prompt=prompt), user=user_model
        )
        if isinstance(res, list) and res:
            u = res[0].get("url") if isinstance(res[0], dict) else None
            if u and u.startswith("http"):
                return await _fetch_url(u)
            if u and u.startswith("/"):
                # local cache path
                path = u.lstrip("/")
                for base in ("/app/backend/data", ""):
                    fp = os.path.join(base, path) if base else path
                    if os.path.isfile(fp):
                        with open(fp, "rb") as fh:
                            return fh.read()
    except Exception:
        return None
    return None


async def _resolve_one_image(item, valves, request, user_dict, ratio=1.4):
    """item may be a dict (image spec) or a string (url / base64 / hint)."""
    if isinstance(item, str):
        item = {"url": item} if item.startswith(("http", "data:")) else {"hint": item}
    if not isinstance(item, dict):
        return None
    raw = None
    for k in ("base64", "b64", "data", "image_base64"):
        if item.get(k):
            raw = _decode_data_image(item[k])
            if raw:
                break
    if not raw:
        url = _first(item, "url", "src", "image_url", "href", default="")
        if url:
            raw = _decode_data_image(url) or await _fetch_url(url)
    if not raw:
        hint = _first(item, "hint", "image_hint", "query", "keywords", default="")
        if hint and valves.unsplash_access_key:
            raw = await _unsplash(str(hint), valves.unsplash_access_key)
    if not raw:
        prompt = _first(item, "generate", "image_generate", "prompt", "ai", default="")
        if prompt and valves.image_generation:
            raw = await _ai_image(str(prompt), request, user_dict)
    if not raw:
        return None
    return _cover_crop(raw, ratio, valves.max_image_px)


# ============================================================================
# Template engine (parse reference .pptx — Fase 1)
# ----------------------------------------------------------------------------
# Dual path geometry:
#   Classic mode: fixed MARGIN / FOOTER_Y / CONTENT_* constants in renderers.
#   Template mode: content frame from deck.frame (parsed reference .pptx).
# Activation requires AND: valves.template_mode_enabled + reference_file_id +
# successful parse/clone (clone wired in later phases).
# ============================================================================

TEMPLATE_MARGIN_IN = 0.15
TEMPLATE_FOOTER_RESERVE_IN = 0.35
TEMPLATE_EDIT_DEFAULT_DROP_TEXT = False
TEMPLATE_EDIT_DEFAULT_DROP_PLACEHOLDERS = False
TEMPLATE_EDIT_DEFAULT_DROP_OFFSLIDE = True
SAFE_ZONE_METHOD = "max_empty_rect"
SAFE_ZONE_QUALITY_COMPUTED = "computed"
SAFE_ZONE_QUALITY_GRID_CAPPED = "grid_capped"
SAFE_ZONE_QUALITY_ADMISSIBLE_FALLBACK = "admissible_fallback"
_SAFE_ZONE_MIN_AREA_SQ_IN = 1.0
# Edge-grid is O(n^4) in decoration count; Slidesgo decks can have 100+ shapes/slide.
_MAX_SAFE_ZONE_GRID_CELLS = 2_000_000
# Grid edge simplification: ignore tiny slivers; cap count for xs/ys only (§6.1.2 Fase 3.3).
_SAFE_ZONE_MIN_DECO_AREA_SQ_IN = 0.08
_SAFE_ZONE_MAX_DECO_FOR_GRID = 48

# Kinds skipped by _clone_decorations (inspect: cloneable=false).
# Tables (p:graphicFrame/a:tbl) are self-contained XML with no package
# relationships, so they clone verbatim and can be refilled via reuse.text.
_CLONE_SKIP_KINDS = frozenset(
    {"chart", "unsupported", "media", "ole", "smartart"}
)

# Kinds that are *content* rather than decoration: never cloned implicitly
# on the classic template path (the renderer draws its own), only via reuse.
_CONTENT_KINDS = frozenset({"table"})

# Max cells summarised per table in inspect_slides output
_INSPECT_TABLE_MAX_ROWS = 8
_INSPECT_TABLE_MAX_COLS = 8
_INSPECT_TABLE_CELL_CHARS = 40


def _shape_cloneable(kind: str) -> bool:
    return kind not in _CLONE_SKIP_KINDS


def _shape_clone_reason(kind: str) -> Optional[str]:
    if _shape_cloneable(kind):
        return None
    return {
        "smartart": "not_cloned_v1_smartart",
        "ole": "not_cloned_v1_ole",
        "chart": "not_cloned_v1_chart",
        "media": "not_cloned_v1_media",
        "unsupported": "not_cloned_v1_unsupported",
    }.get(kind, "not_cloned_v1")


@dataclass
class _BBox:
    x: float
    y: float
    w: float
    h: float
    method: Optional[str] = None
    quality: Optional[str] = None


@dataclass
class _ShapeRecord:
    id: int
    kind: str
    name: str
    bbox: _BBox
    has_text: bool
    text: Optional[str]
    z_order: int
    cloneable: bool = True
    is_placeholder: bool = False
    # kind == "table": {"rows": R, "cols": C, "cells": [[...]]} (truncated)
    table: Optional[dict] = None


@dataclass
class _Decoration:
    shape_id: int
    kind: str
    name: str
    bbox: _BBox
    z_order: int
    source_container: str = "slide"  # slide | layout | master


@dataclass
class _SlideTemplate:
    index: int
    shapes: list[_ShapeRecord] = field(default_factory=list)
    decorations: list[_Decoration] = field(default_factory=list)
    text_verbatim: list[str] = field(default_factory=list)
    safe_zone: Optional[_BBox] = None
    background: dict = field(default_factory=dict)
    shape_count: int = 0
    picture_count: int = 0


@dataclass
class _TemplatePack:
    slide_width_in: float
    slide_height_in: float
    slides: list[_SlideTemplate] = field(default_factory=list)
    theme: dict = field(default_factory=dict)
    decorations_source: str = "slide"


@dataclass
class _MediaUsage:
    slide_index: int
    source: str  # "shape" | "background"
    shape_id: Optional[int] = None
    name: Optional[str] = None


@dataclass
class _InspectImage:
    internal_path: str
    mime: str
    w_px: int
    h_px: int
    usages: list[_MediaUsage] = field(default_factory=list)
    blob: bytes = b""
    cache_file_id: Optional[str] = None
    url: Optional[str] = None
    upload_skipped: bool = False


def _normalize_media_path(partname: str) -> str:
    p = (partname or "").lstrip("/").replace("\\", "/")
    if p.startswith("ppt/"):
        return p
    if "/ppt/" in p:
        return p[p.index("/ppt/") + 1 :]
    return p


def _mime_for_media_path(path: str, blob: bytes) -> str:
    ext = Path(path).suffix.lower()
    guessed, _ = mimetypes.guess_type(path)
    if guessed:
        return guessed
    if blob[:8] == b"\x89PNG\r\n\x1a\n":
        return "image/png"
    if blob[:2] == b"\xff\xd8":
        return "image/jpeg"
    if ext in (".png",):
        return "image/png"
    if ext in (".jpg", ".jpeg"):
        return "image/jpeg"
    if ext in (".gif",):
        return "image/gif"
    if ext in (".svg",):
        return "image/svg+xml"
    return "application/octet-stream"


def _image_dimensions_px(blob: bytes) -> tuple[int, int]:
    if not blob or not _HAS_PIL:
        return 0, 0
    try:
        with Image.open(BytesIO(blob)) as im:
            return int(im.width or 0), int(im.height or 0)
    except Exception:
        return 0, 0


def _blob_from_zip(zip_data: dict[str, bytes], internal_path: str) -> bytes:
    if internal_path in zip_data:
        return zip_data[internal_path]
    alt = internal_path.lstrip("/")
    return zip_data.get(alt, b"")


def _internal_path_from_part(part) -> Optional[str]:
    if part is None:
        return None
    try:
        name = part.partname
        if name:
            return _normalize_media_path(str(name))
    except Exception:
        pass
    return None


def _resolve_rId_to_path(slide_part, r_id: str) -> Optional[str]:
    if not r_id or slide_part is None:
        return None
    try:
        rel = slide_part.related_part(r_id)
        return _internal_path_from_part(rel)
    except Exception:
        return None


def _collect_blips_from_element(
    element,
    slide_part,
    slide_index: int,
    source: str,
    shape_id: Optional[int],
    shape_name: Optional[str],
    out: dict[str, list[_MediaUsage]],
) -> None:
    if element is None:
        return
    for blip in element.iter(qn("a:blip")):
        embed = blip.get(qn("r:embed"))
        if not embed:
            continue
        path = _resolve_rId_to_path(slide_part, embed)
        if not path or not path.startswith("ppt/media/"):
            continue
        usage = _MediaUsage(
            slide_index=slide_index,
            source=source,
            shape_id=shape_id,
            name=shape_name,
        )
        out.setdefault(path, []).append(usage)


def _extract_referenced_media(data: bytes, pack: _TemplatePack) -> list[_InspectImage]:
    """Referenced ppt/media parts from slide shapes (flatten) + slide background blips.

    Inspect-only: generate_slides still clones graphics from reference_file_id (§1).
    """
    if not _HAS_PPTX or not data:
        return []

    zip_data: dict[str, bytes] = {}
    try:
        with ZipFile(BytesIO(data)) as zf:
            for name in zf.namelist():
                if name.startswith("ppt/media/"):
                    zip_data[name] = zf.read(name)
    except Exception as exc:
        log.warning("[inspect_slides] media zip read failed: %s", exc)
        return []

    usages_by_path: dict[str, list[_MediaUsage]] = {}

    try:
        prs = Presentation(BytesIO(data))
    except Exception as exc:
        log.warning("[inspect_slides] media parse failed: %s", exc)
        return []

    for idx, slide in enumerate(prs.slides):
        slide_index = idx
        if idx < len(pack.slides):
            slide_index = pack.slides[idx].index

        bg_el = slide.element.find(qn("p:bg"))
        if bg_el is not None:
            _collect_blips_from_element(
                bg_el,
                slide.part,
                slide_index,
                "background",
                None,
                None,
                usages_by_path,
            )

        for shape, _z in _iter_shapes(slide, flatten_groups=True):
            sid = int(getattr(shape, "shape_id", 0) or 0)
            sname = str(getattr(shape, "name", "") or "") or None
            try:
                _collect_blips_from_element(
                    shape.element,
                    slide.part,
                    slide_index,
                    "shape",
                    sid,
                    sname,
                    usages_by_path,
                )
            except Exception as exc:
                log.debug(
                    "[inspect_slides] media shape id=%s slide=%s: %s",
                    sid,
                    slide_index,
                    exc,
                )

    images: list[_InspectImage] = []
    total_bytes = 0
    for path, usages in sorted(usages_by_path.items()):
        blob = _blob_from_zip(zip_data, path)
        if not blob:
            log.debug("[inspect_slides] media missing in zip: %s", path)
            continue
        w_px, h_px = _image_dimensions_px(blob)
        mime = _mime_for_media_path(path, blob)
        total_bytes += len(blob)
        images.append(
            _InspectImage(
                internal_path=path,
                mime=mime,
                w_px=w_px,
                h_px=h_px,
                usages=usages,
                blob=blob,
            )
        )

    log.info(
        "[inspect_slides] media: %s unique parts, %s usages, %s bytes",
        len(images),
        sum(len(im.usages) for im in images),
        total_bytes,
    )
    if total_bytes > _INSPECT_MEDIA_WARN_BYTES or len(images) > _INSPECT_MEDIA_WARN_COUNT:
        log.warning(
            "[inspect_slides] large media inventory: %s parts, %s MB",
            len(images),
            round(total_bytes / (1024 * 1024), 2),
        )
    return images


def _inspect_images_to_json(images: list[_InspectImage]) -> list[dict]:
    out: list[dict] = []
    for im in images:
        entry: dict = {
            "internal_path": im.internal_path,
            "mime": im.mime,
            "w_px": im.w_px,
            "h_px": im.h_px,
            "usages": [
                {
                    "slide_index": u.slide_index,
                    "source": u.source,
                    **({"shape_id": u.shape_id} if u.shape_id is not None else {}),
                    **({"name": u.name} if u.name else {}),
                }
                for u in im.usages
            ],
        }
        if im.cache_file_id:
            entry["cache_file_id"] = im.cache_file_id
        if im.url:
            entry["url"] = im.url
        if im.upload_skipped:
            entry["upload_skipped"] = True
        out.append(entry)
    return out


def _emu_in(emu: int) -> float:
    return float(emu) / EMU_IN


def _shape_bbox(shape) -> _BBox:
    return _BBox(
        x=_emu_in(shape.left),
        y=_emu_in(shape.top),
        w=_emu_in(shape.width),
        h=_emu_in(shape.height),
    )


def _shape_kind_from_element(shape) -> Optional[str]:
    """XML fallback when MSO_SHAPE_TYPE is ambiguous (e.g. MEDIA + p:pic)."""
    if not _HAS_PPTX:
        return None
    try:
        el = shape.element
    except Exception:
        return None
    tag = el.tag.split("}")[-1] if "}" in el.tag else el.tag
    if tag == "pic":
        return "picture"
    if tag == "cxnSp":
        return "connector"
    if tag == "grpSp":
        return "group"
    if tag == "graphicFrame":
        for node in el.iter():
            if node.tag == qn("c:chart"):
                return "chart"
        for gd in el.iter(qn("a:graphicData")):
            uri = gd.get("uri") or ""
            if "diagram" in uri.lower():
                return "smartart"
        return "unsupported"
    return None


def _placeholder_kind(shape) -> str:
    try:
        from pptx.enum.shapes import PP_PLACEHOLDER  # type: ignore

        ph = shape.placeholder_format
        ph_type = ph.type
        if ph_type in (
            PP_PLACEHOLDER.PICTURE,
            PP_PLACEHOLDER.OBJECT,
            PP_PLACEHOLDER.MEDIA,
        ):
            return "picture"
    except Exception:
        pass
    return "textbox"


def _shape_kind_enum(shape) -> str:
    try:
        st = shape.shape_type
    except Exception:
        return "unsupported"
    if st == MSO_SHAPE_TYPE.PICTURE or st == MSO_SHAPE_TYPE.LINKED_PICTURE:
        return "picture"
    if st == MSO_SHAPE_TYPE.MEDIA:
        xml_kind = _shape_kind_from_element(shape)
        return xml_kind if xml_kind else "media"
    if st == MSO_SHAPE_TYPE.WEB_VIDEO:
        return "media"
    if st == MSO_SHAPE_TYPE.CHART:
        return "chart"
    if st == MSO_SHAPE_TYPE.TABLE:
        return "table"
    if st == MSO_SHAPE_TYPE.LINE:
        return "connector"
    if st == MSO_SHAPE_TYPE.GROUP:
        return "group"
    if st in (MSO_SHAPE_TYPE.TEXT_BOX, MSO_SHAPE_TYPE.TEXT_EFFECT):
        return "textbox"
    if st == MSO_SHAPE_TYPE.AUTO_SHAPE:
        return "autoshape"
    if st == MSO_SHAPE_TYPE.FREEFORM:
        return "autoshape"
    if st == MSO_SHAPE_TYPE.PLACEHOLDER:
        return _placeholder_kind(shape)
    if st in (
        MSO_SHAPE_TYPE.DIAGRAM,
        MSO_SHAPE_TYPE.IGX_GRAPHIC,
    ):
        return "smartart"
    if st in (
        MSO_SHAPE_TYPE.EMBEDDED_OLE_OBJECT,
        MSO_SHAPE_TYPE.LINKED_OLE_OBJECT,
        MSO_SHAPE_TYPE.OLE_CONTROL_OBJECT,
        MSO_SHAPE_TYPE.FORM_CONTROL,
    ):
        return "ole"
    return "unsupported"


def _shape_kind(shape) -> str:
    if not _HAS_PPTX:
        return "unsupported"
    enum_kind = _shape_kind_enum(shape)
    if enum_kind != "unsupported":
        return enum_kind
    xml_kind = _shape_kind_from_element(shape)
    if xml_kind:
        sid = int(getattr(shape, "shape_id", 0) or 0)
        log.debug(
            "[template] shape id=%s reclassified unsupported -> %s",
            sid,
            xml_kind,
        )
        return xml_kind
    return "unsupported"


def _shape_text(shape) -> Optional[str]:
    if not getattr(shape, "has_text_frame", False):
        return None
    try:
        tf = shape.text_frame
        parts = []
        for para in tf.paragraphs:
            t = (para.text or "").strip()
            if t:
                parts.append(t)
        if not parts:
            raw = (tf.text or "").strip()
            return raw or None
        return "\n".join(parts)
    except Exception:
        return None


def _group_child_shapes(shape) -> list:
    try:
        if _HAS_PPTX and shape.shape_type == MSO_SHAPE_TYPE.GROUP:
            return list(shape.shapes)
    except Exception:
        pass
    return []


def _shape_is_placeholder(shape) -> bool:
    if not _HAS_PPTX:
        return False
    try:
        if getattr(shape, "is_placeholder", False):
            return True
    except Exception:
        pass
    try:
        if shape.shape_type == MSO_SHAPE_TYPE.PLACEHOLDER:
            return True
    except Exception:
        pass
    return False


def _is_text_shape(shape) -> bool:
    if not _HAS_PPTX:
        return False
    children = _group_child_shapes(shape)
    if children:
        return any(_is_text_shape(ch) for ch in children)
    try:
        if getattr(shape, "is_placeholder", False):
            return True
    except Exception:
        pass
    try:
        if shape.shape_type == MSO_SHAPE_TYPE.PLACEHOLDER:
            return True
    except Exception:
        pass
    txt = _shape_text(shape)
    if txt:
        return True
    try:
        if shape.shape_type == MSO_SHAPE_TYPE.TEXT_BOX and shape.has_text_frame:
            return True
    except Exception:
        pass
    return False


def _iter_shapes(
    slide, *, flatten_groups: bool = True
) -> Iterator[tuple[Any, int]]:
    """Yield (shape, z_order) in deterministic slide order."""
    z = 0
    for shape in slide.shapes:
        if flatten_groups and _group_child_shapes(shape):
            for ch in _group_child_shapes(shape):
                yield ch, z
                z += 1
        else:
            yield shape, z
            z += 1


def _table_summary(shape) -> Optional[dict]:
    """Compact grid summary of a table graphicFrame for inspect_slides."""
    try:
        if not getattr(shape, "has_table", False):
            return None
        tbl = shape.table
        n_rows = len(tbl.rows)
        n_cols = len(tbl.columns)
        cells: list[list[str]] = []
        for r_idx, row in enumerate(tbl.rows):
            if r_idx >= _INSPECT_TABLE_MAX_ROWS:
                break
            line: list[str] = []
            for c_idx, cell in enumerate(row.cells):
                if c_idx >= _INSPECT_TABLE_MAX_COLS:
                    break
                txt = (cell.text_frame.text or "").strip().replace("\n", " ")
                if len(txt) > _INSPECT_TABLE_CELL_CHARS:
                    txt = txt[: _INSPECT_TABLE_CELL_CHARS - 1] + "…"
                line.append(txt)
            cells.append(line)
        return {"rows": n_rows, "cols": n_cols, "cells": cells}
    except Exception as exc:
        log.debug("[template] table summary failed: %s", exc)
        return None


def _shape_to_record(shape, z_order: int) -> _ShapeRecord:
    txt = _shape_text(shape)
    kind = _shape_kind(shape)
    has_text = bool(txt) or _is_text_shape(shape)
    sid = int(getattr(shape, "shape_id", 0) or 0)
    name = str(getattr(shape, "name", "") or "")
    return _ShapeRecord(
        id=sid,
        kind=kind,
        name=name,
        bbox=_shape_bbox(shape),
        has_text=has_text,
        text=txt,
        z_order=z_order,
        cloneable=_shape_cloneable(kind),
        is_placeholder=_shape_is_placeholder(shape),
        table=_table_summary(shape) if kind == "table" else None,
    )


def _extract_decorations(slide) -> list[_Decoration]:
    """Leaf non-text decorations (flattened groups), aligned with inspect shapes[].

    Text shapes and content kinds (tables) are excluded: on the classic
    template path the renderer draws its own content inside the safe zone,
    so cloning a template table would overlap it. Tables are cloned only
    when addressed explicitly through slides[].reuse.
    """
    out: list[_Decoration] = []
    seen_ids: set[int] = set()
    for shape, z in _iter_shapes(slide, flatten_groups=True):
        if _is_text_shape(shape):
            continue
        kind = _shape_kind(shape)
        if kind == "group" or kind in _CONTENT_KINDS:
            continue
        sid = int(getattr(shape, "shape_id", 0) or 0)
        if sid in seen_ids:
            continue
        seen_ids.add(sid)
        out.append(
            _Decoration(
                shape_id=sid,
                kind=kind,
                name=str(getattr(shape, "name", "") or ""),
                bbox=_shape_bbox(shape),
                z_order=z,
                source_container="slide",
            )
        )
    return out


def _rgb_to_hex(rgb) -> Optional[str]:
    if rgb is None:
        return None
    try:
        return f"{int(rgb[0]):02X}{int(rgb[1]):02X}{int(rgb[2]):02X}"
    except Exception:
        return None


def _theme_color_from_element(el) -> Optional[str]:
    if el is None:
        return None
    try:
        srgb = el.find(qn("a:srgbClr"))
        if srgb is not None and srgb.get("val"):
            return str(srgb.get("val")).upper()
        sys_clr = el.find(qn("a:sysClr"))
        if sys_clr is not None and sys_clr.get("lastClr"):
            return str(sys_clr.get("lastClr")).upper()
    except Exception:
        pass
    return None


# clrScheme slots used when resolving a:schemeClr on cloned slide XML.
_THEME_CLR_SCHEME_SLOTS: tuple[tuple[str, str], ...] = (
    ("dk1", "a:dk1"),
    ("lt1", "a:lt1"),
    ("dk2", "a:dk2"),
    ("lt2", "a:lt2"),
    ("accent1", "a:accent1"),
    ("accent2", "a:accent2"),
    ("accent3", "a:accent3"),
    ("accent4", "a:accent4"),
    ("accent5", "a:accent5"),
    ("accent6", "a:accent6"),
    ("hlink", "a:hlink"),
    ("folHlink", "a:folHlink"),
)


def _theme_color_map_from_theme_element(root) -> dict[str, str]:
    """Map scheme token names to RGB hex from a theme element (a:theme)."""
    out: dict[str, str] = {}
    try:
        clr_scheme = root.find(".//" + qn("a:clrScheme"))
        if clr_scheme is None:
            return out
        for key, tag in _THEME_CLR_SCHEME_SLOTS:
            node = clr_scheme.find(qn(tag))
            val = _theme_color_from_element(node)
            if val:
                out[key] = val
    except Exception as exc:
        log.debug("[template] clrScheme read: %s", exc)
    return out


def _theme_element_from_slide_master_part(master_part):
    """Return the a:theme root for a slide master (python-pptx or rel fallback)."""
    if master_part is None:
        return None
    try:
        theme_part = getattr(master_part, "theme_part", None)
        if theme_part is not None:
            return theme_part.element
    except Exception:
        pass
    try:
        from pptx.opc.constants import RELATIONSHIP_TYPE as RT  # type: ignore
        from pptx.oxml import parse_xml  # type: ignore

        for rel in master_part.rels.values():
            if rel.reltype == RT.THEME:
                return parse_xml(rel.target_part.blob)
    except Exception as exc:
        log.debug("[template] theme rel read: %s", exc)
    return None


def _theme_color_map_for_source_slide(
    source_slide,
    theme_fallback: Optional[dict] = None,
) -> dict[str, str]:
    """Color map for the slide master chain of a template slide."""
    try:
        master_part = source_slide.slide_layout.slide_master.part
        root = _theme_element_from_slide_master_part(master_part)
        if root is not None:
            mapped = _theme_color_map_from_theme_element(root)
            if mapped:
                return mapped
    except Exception as exc:
        log.debug("[template] theme map for source slide: %s", exc)
    if isinstance(theme_fallback, dict):
        out = {
            k: str(v).upper()
            for k, v in theme_fallback.items()
            if k in {s[0] for s in _THEME_CLR_SCHEME_SLOTS} and v
        }
        if out:
            return out
    return {"dk1": "1A1A1A", "lt1": "FFFFFF", "accent1": "4472C4"}


def _materialize_scheme_colors(element, color_map: dict[str, str]) -> None:
    """Replace a:schemeClr with a:srgbClr using resolved theme hex values.

    Only touches schemeClr nodes inside cloned target XML. phClr, sysClr, and
    gradFill stops are left unchanged. Child modifiers on schemeClr (lumMod,
    tint, …) are dropped when materializing to a fixed sRGB.
    """
    if not color_map:
        return
    scheme_tag = qn("a:schemeClr")
    srgb_tag = qn("a:srgbClr")
    for el in element.iter():
        if el.tag != scheme_tag:
            continue
        token = (el.get("val") or "").strip()
        if not token:
            continue
        hex_val = color_map.get(token)
        if not hex_val:
            log.debug("[template] schemeClr token %r not in theme map", token)
            continue
        for child in list(el):
            el.remove(child)
        el.tag = srgb_tag
        el.set("val", hex_val.upper())


def _finalize_cloned_slide(
    target_slide,
    source_slide,
    theme_fallback: Optional[dict] = None,
) -> None:
    """Resolve theme scheme colors on a cloned output slide."""
    if not _HAS_PPTX:
        return
    try:
        color_map = _theme_color_map_for_source_slide(source_slide, theme_fallback)
        _materialize_scheme_colors(target_slide.element.cSld, color_map)
    except Exception as exc:
        log.warning("[template] finalize cloned slide theme: %s", exc)


def _extract_theme_from_prs(prs) -> dict:
    theme: dict = {
        "major_font": "Calibri",
        "minor_font": "Calibri",
    }
    try:
        root = _theme_element_from_slide_master_part(prs.slide_master.part)
        if root is not None:
            theme.update(_theme_color_map_from_theme_element(root))
        font_scheme = root.find(".//" + qn("a:fontScheme"))
        if font_scheme is not None:
            for font_key, tag in (("major_font", "a:majorFont"), ("minor_font", "a:minorFont")):
                block = font_scheme.find(qn(tag))
                if block is not None:
                    latin = block.find(qn("a:latin"))
                    if latin is not None and latin.get("typeface"):
                        theme[font_key] = latin.get("typeface")
    except Exception as exc:
        log.debug("[template] theme XML partial/fallback: %s", exc)
    if not theme.get("dk1"):
        theme["dk1"] = "1A1A1A"
    if not theme.get("lt1"):
        theme["lt1"] = "FFFFFF"
    if not theme.get("accent1"):
        theme["accent1"] = "4472C4"
    return theme


def _extract_slide_background(slide) -> dict:
    if not _HAS_PPTX:
        return {"type": "unknown"}
    try:
        fill = slide.background.fill
        if fill.type == MSO_FILL_TYPE.SOLID:
            rgb = fill.fore_color.rgb
            hx = _rgb_to_hex(rgb)
            if hx:
                return {"type": "solid", "color": hx}
    except Exception as exc:
        log.debug("[template] slide background read: %s", exc)
    return {"type": "unknown"}


def _rects_intersect(a: _BBox, b: _BBox) -> bool:
    return not (
        a.x + a.w <= b.x
        or b.x + b.w <= a.x
        or a.y + a.h <= b.y
        or b.y + b.h <= a.y
    )


def _clip_bbox_to_slide(bbox: _BBox, slide_w: float, slide_h: float) -> _BBox:
    x = max(0.0, min(bbox.x, slide_w))
    y = max(0.0, min(bbox.y, slide_h))
    w = max(0.0, min(bbox.x + bbox.w, slide_w) - x)
    h = max(0.0, min(bbox.y + bbox.h, slide_h) - y)
    return _BBox(x=x, y=y, w=w, h=h)


def _safe_zone_admissible_region(
    x0: float,
    y0: float,
    w0: float,
    h0: float,
    *,
    quality: str,
    clamp_tiny: bool = False,
) -> _BBox:
    w = max(w0, 0.1) if clamp_tiny else w0
    h = max(h0, 0.1) if clamp_tiny else h0
    return _BBox(
        x=x0,
        y=y0,
        w=w,
        h=h,
        method=SAFE_ZONE_METHOD,
        quality=quality,
    )


def _occupied_for_safe_zone_grid(occupied_full: list[_BBox]) -> list[_BBox]:
    """Subset for edge-grid xs/ys only; intersection checks use occupied_full."""
    filtered = [
        o
        for o in occupied_full
        if o.w * o.h >= _SAFE_ZONE_MIN_DECO_AREA_SQ_IN
    ]
    if len(filtered) <= _SAFE_ZONE_MAX_DECO_FOR_GRID:
        return filtered
    ranked = sorted(
        filtered,
        key=lambda o: (-(o.w * o.h), o.x, o.y, o.w, o.h),
    )
    return ranked[:_SAFE_ZONE_MAX_DECO_FOR_GRID]


def _safe_zone_grid_cell_count(xs: list[float], ys: list[float]) -> int:
    nx = len(xs) * (len(xs) - 1) // 2
    ny = len(ys) * (len(ys) - 1) // 2
    return nx * ny


def _compute_safe_zone(
    slide_w: float, slide_h: float, decorations: list[_Decoration]
) -> _BBox:
    x0 = TEMPLATE_MARGIN_IN
    y0 = TEMPLATE_MARGIN_IN
    w0 = slide_w - 2 * TEMPLATE_MARGIN_IN
    h0 = slide_h - TEMPLATE_MARGIN_IN - TEMPLATE_FOOTER_RESERVE_IN
    if w0 <= 0 or h0 <= 0:
        log.warning("[template] safe zone: slide too small for margins")
        return _safe_zone_admissible_region(
            x0,
            y0,
            w0,
            h0,
            quality=SAFE_ZONE_QUALITY_ADMISSIBLE_FALLBACK,
            clamp_tiny=True,
        )

    if not decorations:
        return _safe_zone_admissible_region(
            x0, y0, w0, h0, quality=SAFE_ZONE_QUALITY_COMPUTED
        )

    occupied_full = [
        _clip_bbox_to_slide(d.bbox, slide_w, slide_h)
        for d in decorations
        if d.bbox.w > 0 and d.bbox.h > 0
    ]
    occupied_edges = _occupied_for_safe_zone_grid(occupied_full)

    xs_set = {x0, x0 + w0}
    ys_set = {y0, y0 + h0}
    for occ in occupied_edges:
        xs_set.add(max(x0, min(x0 + w0, occ.x)))
        xs_set.add(max(x0, min(x0 + w0, occ.x + occ.w)))
        ys_set.add(max(y0, min(y0 + h0, occ.y)))
        ys_set.add(max(y0, min(y0 + h0, occ.y + occ.h)))
    xs = sorted(xs_set)
    ys = sorted(ys_set)
    grid_cells = _safe_zone_grid_cell_count(xs, ys)
    if grid_cells > _MAX_SAFE_ZONE_GRID_CELLS:
        log.warning(
            "[template] safe zone: grid too large (%s cells, %s full / %s edge decorations); "
            "quality=%s",
            grid_cells,
            len(occupied_full),
            len(occupied_edges),
            SAFE_ZONE_QUALITY_GRID_CAPPED,
        )
        return _safe_zone_admissible_region(
            x0, y0, w0, h0, quality=SAFE_ZONE_QUALITY_GRID_CAPPED
        )

    cx_slide = slide_w / 2.0
    cy_slide = slide_h / 2.0
    best: Optional[_BBox] = None
    best_area = 0.0
    best_center_dist = float("inf")

    for i, x1 in enumerate(xs):
        for x2 in xs[i + 1 :]:
            rw = x2 - x1
            if rw <= 0:
                continue
            for j, y1 in enumerate(ys):
                for y2 in ys[j + 1 :]:
                    rh = y2 - y1
                    if rh <= 0:
                        continue
                    cand = _BBox(x=x1, y=y1, w=rw, h=rh)
                    if any(_rects_intersect(cand, occ) for occ in occupied_full):
                        continue
                    area = rw * rh
                    ccx = x1 + rw / 2.0
                    ccy = y1 + rh / 2.0
                    dist = (ccx - cx_slide) ** 2 + (ccy - cy_slide) ** 2
                    if area > best_area or (
                        abs(area - best_area) < 1e-9 and dist < best_center_dist
                    ):
                        best_area = area
                        best_center_dist = dist
                        best = cand

    if best is None or best_area < _SAFE_ZONE_MIN_AREA_SQ_IN:
        log.warning(
            "[template] safe zone: no large empty rect (area=%.2f); quality=%s",
            best_area,
            SAFE_ZONE_QUALITY_ADMISSIBLE_FALLBACK,
        )
        return _safe_zone_admissible_region(
            x0, y0, w0, h0, quality=SAFE_ZONE_QUALITY_ADMISSIBLE_FALLBACK
        )
    return _BBox(
        x=best.x,
        y=best.y,
        w=best.w,
        h=best.h,
        method=SAFE_ZONE_METHOD,
        quality=SAFE_ZONE_QUALITY_COMPUTED,
    )


def _decorations_from_container(
    shapes_container,
    *,
    source_container: str,
    z_base: int,
) -> list[_Decoration]:
    out: list[_Decoration] = []
    seen: set[int] = set()
    z = 0
    for shape in shapes_container:
        if _shape_is_placeholder(shape):
            continue
        if _group_child_shapes(shape):
            for ch, cz in zip(
                _group_child_shapes(shape),
                range(len(_group_child_shapes(shape))),
            ):
                if _shape_is_placeholder(ch) or _is_text_shape(ch):
                    continue
                kind = _shape_kind(ch)
                if kind == "group":
                    continue
                sid = int(getattr(ch, "shape_id", 0) or 0)
                if sid in seen:
                    continue
                seen.add(sid)
                out.append(
                    _Decoration(
                        shape_id=sid,
                        kind=kind,
                        name=str(getattr(ch, "name", "") or ""),
                        bbox=_shape_bbox(ch),
                        z_order=z_base + z + cz,
                        source_container=source_container,
                    )
                )
            z += max(1, len(_group_child_shapes(shape)))
            continue
        if _is_text_shape(shape):
            continue
        kind = _shape_kind(shape)
        if kind == "group":
            continue
        sid = int(getattr(shape, "shape_id", 0) or 0)
        if sid in seen:
            continue
        seen.add(sid)
        out.append(
            _Decoration(
                shape_id=sid,
                kind=kind,
                name=str(getattr(shape, "name", "") or ""),
                bbox=_shape_bbox(shape),
                z_order=z_base + z,
                source_container=source_container,
            )
        )
        z += 1
    return out


def _inspect_master_decorations(prs, slide) -> list[_Decoration]:
    """Layout / slide-master shapes linked to this slide (P1)."""
    out: list[_Decoration] = []
    try:
        layout = slide.slide_layout
        out.extend(
            _decorations_from_container(
                layout.shapes, source_container="layout", z_base=-2000
            )
        )
        master = layout.slide_master
        out.extend(
            _decorations_from_container(
                master.shapes, source_container="master", z_base=-4000
            )
        )
    except Exception as exc:
        log.debug("[template] master/layout decoration scan: %s", exc)
    return out


def _merge_decorations(
    slide_decs: list[_Decoration], extra: list[_Decoration]
) -> list[_Decoration]:
    seen = {d.shape_id for d in slide_decs}
    merged = list(slide_decs)
    for dec in extra:
        if dec.shape_id in seen:
            continue
        seen.add(dec.shape_id)
        merged.append(dec)
    return merged


def _parse_reference_pptx(data: bytes) -> _TemplatePack:
    if not _HAS_PPTX:
        raise RuntimeError("python-pptx is not installed")
    t_parse = time.monotonic()
    try:
        prs = Presentation(BytesIO(data))
    except Exception as exc:
        raise ValueError(f"Cannot open .pptx: {exc}") from exc

    if not prs.slides:
        raise ValueError("Template .pptx has no slides")

    slide_w_in = _emu_in(prs.slide_width)
    slide_h_in = _emu_in(prs.slide_height)
    theme = _extract_theme_from_prs(prs)
    slides_out: list[_SlideTemplate] = []
    pack_sources: set[str] = set()

    for idx, slide in enumerate(prs.slides):
        try:
            records: list[_ShapeRecord] = []
            # R8: inspect text is verbatim from the file — no placeholder/boilerplate filtering.
            text_verbatim: list[str] = []
            z = 0
            for shape, _z in _iter_shapes(slide, flatten_groups=True):
                rec = _shape_to_record(shape, z)
                records.append(rec)
                if rec.text:
                    text_verbatim.append(rec.text)
                z += 1

            slide_decs = _extract_decorations(slide)
            master_decs = _inspect_master_decorations(prs, slide)
            decorations = _merge_decorations(slide_decs, master_decs)
            if slide_decs and master_decs:
                pack_sources.add("both")
            elif master_decs:
                pack_sources.add("master")
            else:
                pack_sources.add("slide")
            safe = _compute_safe_zone(slide_w_in, slide_h_in, slide_decs or decorations)
            bg = _extract_slide_background(slide)
            pic_n = sum(1 for r in records if r.kind == "picture")

            st = _SlideTemplate(
                index=idx,
                shapes=records,
                decorations=decorations,
                text_verbatim=text_verbatim,
                safe_zone=safe,
                background=bg,
                shape_count=len(records),
                picture_count=pic_n,
            )
            slides_out.append(st)
            log.info(
                "[template] slide %s: shape_count=%s decorations=%s safe_zone=(%.2f,%.2f,%.2f,%.2f)",
                idx,
                st.shape_count,
                len(decorations),
                safe.x,
                safe.y,
                safe.w,
                safe.h,
            )
        except Exception as exc:
            log.warning("[template] slide %s parse partial error: %s", idx, exc)

    if not slides_out:
        raise ValueError("No slides could be parsed from template .pptx")

    if "both" in pack_sources or (
        "master" in pack_sources and "slide" in pack_sources
    ):
        decorations_source = "both"
    elif "master" in pack_sources:
        decorations_source = "master"
    else:
        decorations_source = "slide"

    parse_ms = (time.monotonic() - t_parse) * 1000.0
    log.info(
        "[template] parse_ms=%.1f slides=%s decorations_source=%s",
        parse_ms,
        len(slides_out),
        decorations_source,
    )

    return _TemplatePack(
        slide_width_in=slide_w_in,
        slide_height_in=slide_h_in,
        slides=slides_out,
        theme=theme,
        decorations_source=decorations_source,
    )


def _bbox_to_dict(b: _BBox, *, include_method: bool = False) -> dict:
    out = {"x": round(b.x, 4), "y": round(b.y, 4), "w": round(b.w, 4), "h": round(b.h, 4)}
    if include_method and b.method:
        out["method"] = b.method
    if include_method and b.quality:
        out["quality"] = b.quality
    return out


def _inspect_uncloneable_summary(pack: _TemplatePack) -> list[dict]:
    out: list[dict] = []
    for st in pack.slides:
        kinds: dict[str, int] = {}
        for sh in st.shapes:
            if sh.cloneable:
                continue
            kinds[sh.kind] = kinds.get(sh.kind, 0) + 1
        if kinds:
            notes: list[str] = []
            if kinds.get("smartart"):
                notes.append("SmartArt is not cloned in v1 (use alternate layout content).")
            if kinds.get("ole"):
                notes.append("OLE/embed objects are not cloned in v1.")
            if kinds.get("unsupported"):
                notes.append("Some shapes are unsupported for clone in v1.")
            entry: dict = {
                "slide_index": st.index,
                "count": sum(kinds.values()),
                "kinds": kinds,
            }
            if notes:
                entry["note"] = " ".join(notes)
            out.append(entry)
    return out


def _inspect_safe_zone_caution(
    pack: _TemplatePack, role_by_index: dict[int, str]
) -> list[dict]:
    out: list[dict] = []
    for st in pack.slides:
        if role_by_index.get(st.index) != "content":
            continue
        sz = st.safe_zone
        if sz is None or sz.quality == SAFE_ZONE_QUALITY_COMPUTED:
            continue
        out.append(
            {
                "slide_index": st.index,
                "role": "content",
                "quality": sz.quality,
            }
        )
    return out


def _reference_indices_from_mapping(
    pack: _TemplatePack, mapping: Optional[dict]
) -> set[int]:
    mapping = mapping if isinstance(mapping, dict) else {}
    if not mapping:
        return {0} if pack.slides else set()
    indices: set[int] = set()
    for val in mapping.values():
        try:
            indices.add(int(val))
        except (TypeError, ValueError):
            continue
    return indices


def _template_strict_issues(
    pack: _TemplatePack, mapping: Optional[dict]
) -> list[dict]:
    issues: list[dict] = []
    for idx in sorted(_reference_indices_from_mapping(pack, mapping)):
        if idx < 0 or idx >= len(pack.slides):
            issues.append(
                {
                    "slide_index": idx,
                    "reason": "mapping_index_out_of_range",
                }
            )
            continue
        st = pack.slides[idx]
        sz = st.safe_zone
        if sz is None:
            issues.append({"slide_index": idx, "reason": "missing_safe_zone"})
            continue
        if sz.quality != SAFE_ZONE_QUALITY_COMPUTED:
            issues.append(
                {
                    "slide_index": idx,
                    "reason": "safe_zone_quality",
                    "quality": sz.quality,
                }
            )
        area = sz.w * sz.h
        if area < _SAFE_ZONE_MIN_AREA_SQ_IN:
            issues.append(
                {
                    "slide_index": idx,
                    "reason": "safe_zone_too_small",
                    "area_sq_in": round(area, 4),
                }
            )
    return issues


def _validate_template_strict(pack: _TemplatePack, mapping: Optional[dict]) -> None:
    issues = _template_strict_issues(pack, mapping)
    if not issues:
        return
    first = issues[0]
    idx = first.get("slide_index", "?")
    reason = first.get("reason", "unknown")
    if reason == "safe_zone_quality":
        raise ValueError(
            f"template_strict_mode: reference slide {idx} safe_zone quality "
            f"is {first.get('quality')!r} (computed required)"
        )
    if reason == "safe_zone_too_small":
        raise ValueError(
            f"template_strict_mode: reference slide {idx} safe_zone area "
            f"({first.get('area_sq_in')} sq in) below minimum"
        )
    if reason == "missing_safe_zone":
        raise ValueError(
            f"template_strict_mode: reference slide {idx} has no safe_zone"
        )
    raise ValueError(
        f"template_strict_mode: reference slide {idx} failed check ({reason})"
    )


def _inspect_master_decoration_hint(pack: _TemplatePack) -> list[dict]:
    out: list[dict] = []
    for st in pack.slides:
        n_master = sum(
            1 for d in st.decorations if d.source_container in ("layout", "master")
        )
        if n_master:
            out.append({"slide_index": st.index, "merged_from_master": n_master})
    return out


def _inspect_hints(pack: _TemplatePack) -> dict:
    roles: list[dict] = []
    n = len(pack.slides)
    if n == 0:
        return {"likely_roles": roles, "uncloneable": []}
    if n == 1:
        roles.append({"index": 0, "role": "content", "confidence": "low"})
    else:
        roles.append({"index": 0, "role": "cover", "confidence": "low"})
        if n >= 2:
            roles.append({"index": 1, "role": "content", "confidence": "low"})
        if n >= 3:
            roles.append({"index": n - 1, "role": "closing", "confidence": "low"})
    bg0 = pack.slides[0].background if pack.slides else {}
    if bg0.get("type") == "solid" and bg0.get("color"):
        try:
            if _luminance(str(bg0["color"])) < 0.25:
                for r in roles:
                    if r["index"] == 0:
                        r["role"] = "cover"
        except Exception:
            pass
    role_by_index = {r["index"]: r["role"] for r in roles}
    hints: dict = {
        "likely_roles": roles,
        "uncloneable": _inspect_uncloneable_summary(pack),
    }
    caution = _inspect_safe_zone_caution(pack, role_by_index)
    if caution:
        hints["safe_zone_caution"] = caution
    master_hint = _inspect_master_decoration_hint(pack)
    if master_hint:
        hints["master_decorations"] = master_hint
    hints["strict_would_fail"] = _template_strict_issues(pack, {"default": 0})
    return hints


def _serialize_inspect_payload(
    pack: Optional[_TemplatePack],
    *,
    file_id: str = "",
    filename: str = "",
    ok: bool = True,
    error: Optional[str] = None,
    images: Optional[list[dict]] = None,
) -> dict:
    if not ok:
        out: dict = {"ok": False, "error": error or "Unknown error"}
        if file_id:
            out["file_id"] = file_id
        return out

    assert pack is not None
    slides_json = []
    for st in pack.slides:
        shapes = []
        for sh in st.shapes:
            entry = {
                "id": sh.id,
                "kind": sh.kind,
                "name": sh.name,
                "bbox": _bbox_to_dict(sh.bbox),
                "has_text": sh.has_text,
                "text": sh.text,
                "cloneable": sh.cloneable,
            }
            reason = _shape_clone_reason(sh.kind)
            if reason:
                entry["clone_reason"] = reason
            if sh.table is not None:
                entry["table"] = sh.table
            shapes.append(entry)
        sz = st.safe_zone
        slide_entry = {
            "index": st.index,
            "shape_count": st.shape_count,
            "background": st.background,
            "shapes": shapes,
            "text_verbatim": list(st.text_verbatim),
            "picture_count": st.picture_count,
        }
        if sz is not None:
            slide_entry["safe_zone"] = _bbox_to_dict(sz, include_method=True)
        slides_json.append(slide_entry)

    payload = {
        "ok": True,
        "file_id": file_id,
        "filename": filename or "template.pptx",
        "slide_width_in": round(pack.slide_width_in, 4),
        "slide_height_in": round(pack.slide_height_in, 4),
        "slide_count": len(pack.slides),
        "theme": dict(pack.theme),
        "decorations_source": pack.decorations_source,
        "slides": slides_json,
        "images": list(images) if images is not None else [],
        "hints": _inspect_hints(pack),
    }
    return payload


def _retry_tool_result(message: str) -> str:
    """Tool result that tells the model to call again, not to stop the turn."""
    return (
        "[TOOL_RESULT — do not show this to the user and do not give up. "
        "Do not generate a classic deck. Call the tool again now, correcting "
        "only what is stated below.]\n\n"
        + message
    )


def _compact_inspect_payload(payload: dict) -> dict:
    """Model-facing inventory: text and tables only.

    A Slidesgo slide can contain hundreds of decorative autoshapes. Listing
    them makes the model lose `file_id` and invent another one. Decorations
    are cloned automatically when `reuse` omits `keep_ids`, so the model only
    needs the ids it must rewrite.
    """
    if not payload.get("ok"):
        return payload
    slides = []
    for st in payload.get("slides") or []:
        texts = []
        tables = []
        for sh in st.get("shapes") or []:
            if sh.get("kind") == "table" or sh.get("table"):
                entry = {"id": sh.get("id")}
                if isinstance(sh.get("table"), dict):
                    entry.update(sh["table"])
                tables.append(entry)
                continue
            text = (sh.get("text") or "").strip()
            if sh.get("has_text") and text:
                texts.append({"id": sh.get("id"), "text": sh.get("text")})
        listed = len(texts) + len(tables)
        slides.append(
            {
                "index": st.get("index"),
                "text": texts,
                "tables": tables,
                "decoration_count": max(0, int(st.get("shape_count") or 0) - listed),
            }
        )
    file_id = payload.get("file_id") or ""
    return {
        "ok": True,
        "file_id": file_id,
        "filename": payload.get("filename"),
        "slide_width_in": payload.get("slide_width_in"),
        "slide_height_in": payload.get("slide_height_in"),
        "slide_count": payload.get("slide_count"),
        "slides": slides,
        "next": (
            "Call generate_slides exactly once. "
            f'reference_file_id must be "{file_id}" copied character for '
            "character. For every template slide you fill, emit one output "
            'slide {"reuse": {"slide": <index>, "text": {"<id>": "new text"}}}. '
            "Table ids take {\"headers\": [...], \"rows\": [[...]]}. "
            "Omit keep_ids so decorations clone automatically. "
            "Do not call inspect_slides again. Do not invent a file_id. "
            "Do not fall back to classic layouts."
        ),
    }


def _inspect_tool_result(payload: dict) -> str:
    body = json.dumps(payload, indent=2, ensure_ascii=False)
    if not payload.get("ok"):
        return (
            "[TOOL_RESULT — the inspect failed. Do not invent another file_id "
            "and do not generate a classic deck. Follow the error below.]\n\n"
            + body
        )
    return (
        "[TOOL_RESULT — do not paste this JSON to the user. "
        'Your next action is one generate_slides call, as stated in "next". '
        "Do not call inspect_slides again.]\n\n"
        + body
    )


def _is_presentation_file(name: str, content_type: Optional[str] = None) -> bool:
    n = (name or "").strip().lower()
    if n.endswith(_PPTX_EXTENSIONS):
        return True
    ct = (content_type or "").strip().lower()
    if ct in _PPTX_MIMES:
        return True
    if "presentationml" in ct or "powerpoint" in ct:
        return True
    return False


def _is_files_api_id(file_id: str) -> bool:
    return bool(_FILES_API_ID_RE.match((file_id or "").strip()))


def _wrong_reference_id_message(sent_id: str, attached: str) -> str:
    """The model sent an id that is not the .pptx actually attached."""
    sent = (sent_id or "").strip()
    if not sent:
        why = "generate_slides was called without reference_file_id."
    elif not _is_files_api_id(sent):
        why = f"reference_file_id {sent!r} is not a Files API id."
    else:
        why = (
            f"reference_file_id {sent!r} is not the attached template "
            "(it does not exist, or it is a different file)."
        )
    return (
        f"{why} The attached template file_id is \"{attached}\". "
        "Do not tell the user the template is missing. "
        "Do not generate a classic deck. "
        "Do not call inspect_slides with a new id. "
        f'Call generate_slides with reference_file_id="{attached}" '
        "and per-slide reuse from the inspect inventory of that file "
        "(reuse.slide = slides[].index, reuse.text = the text and table ids). "
        "Omit keep_ids so decorations are cloned."
    )


def _invented_reference_id_message(ref_id: str, messages: Any) -> str:
    """Tell the model to retry generate_slides with the real Files API id."""
    attached = _find_pptx_attachment(messages)
    lines = [
        f'reference_file_id {ref_id!r} is not a Files API id, so the template '
        "was not loaded. Do not show this message to the user.",
        'Call generate_slides again and copy "file_id" exactly from the '
        "inspect_slides JSON you already have. It is a UUID (8-4-4-4-12 hex "
        "digits). Never invent, shorten, or replace it with a number, a "
        "filename, or a slide index.",
    ]
    if attached:
        lines.append(
            f'The attached template file_id is "{attached}". '
            "Use that exact string as reference_file_id, and build each "
            "output slide with reuse from inspect slides[].shapes[].id."
        )
    else:
        lines.append(
            "Build each output slide with reuse from inspect "
            'slides[].shapes[].id, and set reference_file_id to that '
            '"file_id" — not the value you just sent.'
        )
    return " ".join(lines)


def _find_pptx_attachment(messages: Any) -> Optional[str]:
    """Return Files API id of the most recent .pptx attachment in chat (inspect only)."""
    msgs = _as_list(messages)
    for msg in reversed(msgs):
        if not isinstance(msg, dict):
            continue
        for key in ("files", "attachments", "documents"):
            for item in _as_list(msg.get(key)):
                if not isinstance(item, dict):
                    continue
                fid = (item.get("id") or item.get("file_id") or "").strip()
                if not fid:
                    continue
                name = (
                    item.get("name")
                    or item.get("filename")
                    or item.get("file_name")
                    or ""
                )
                ctype = item.get("content_type") or item.get("type") or ""
                if _is_presentation_file(str(name), str(ctype) if ctype else None):
                    return fid
                if str(name).lower().endswith(_PPTX_EXTENSIONS):
                    return fid
    return None


def _file_model_name_and_type(file_obj) -> tuple[str, str]:
    name = (
        getattr(file_obj, "filename", None)
        or getattr(file_obj, "name", None)
        or ""
    )
    meta = getattr(file_obj, "meta", None) or {}
    if isinstance(meta, dict):
        name = meta.get("name") or name
        ctype = meta.get("content_type") or ""
    else:
        ctype = getattr(meta, "content_type", "") if meta else ""
        if hasattr(meta, "get"):
            name = meta.get("name") or name
            ctype = meta.get("content_type") or ctype
    return str(name or "template.pptx"), str(ctype or "")


async def _load_reference_pptx(
    file_id: str,
    request: Any,
    user_dict: Optional[dict],
) -> tuple[Optional[bytes], Optional[str], Optional[str]]:
    """Download reference .pptx bytes. Returns (data, filename, error_message)."""
    fid = (file_id or "").strip()
    if not fid:
        return None, None, "Missing file_id."

    user_id = (user_dict or {}).get("id") if isinstance(user_dict, dict) else None
    if not user_id:
        return None, None, "Missing user context for file access."

    if _HAS_OWUI_FILE_DOWNLOAD and _OwuiFiles is not None and _OwuiStorage is not None:
        try:
            file_obj = None
            if hasattr(_OwuiFiles, "get_file_by_id_and_user_id"):
                file_obj = await _maybe_await(
                    _OwuiFiles.get_file_by_id_and_user_id(fid, user_id)
                )
            if file_obj is None and hasattr(_OwuiFiles, "get_file_by_id"):
                candidate = await _maybe_await(_OwuiFiles.get_file_by_id(fid))
                owner = getattr(candidate, "user_id", None)
                if candidate and str(owner) == str(user_id):
                    file_obj = candidate
            if file_obj is None:
                return None, None, "File not found or not accessible."

            fname, ctype = _file_model_name_and_type(file_obj)
            if not _is_presentation_file(fname, ctype):
                return None, None, "File is not a PowerPoint template (.pptx/.potx)."

            rel_path = getattr(file_obj, "path", None) or getattr(file_obj, "file_path", None)
            if not rel_path:
                return None, None, "File storage path missing."

            disk_path = await _maybe_await(_OwuiStorage.get_file(rel_path))
            path = Path(disk_path)
            if not path.is_file():
                return None, None, "File not found or not accessible."

            data = path.read_bytes()
            if len(data) > _REFERENCE_PPTX_MAX_BYTES:
                log.warning(
                    "[inspect_slides] reference file %s exceeds %s MB",
                    fid,
                    _REFERENCE_PPTX_MAX_BYTES // (1024 * 1024),
                )
            return data, fname, None
        except Exception as exc:
            log.exception("[inspect_slides] in-process file load failed")
            return None, None, f"Could not read template file: {exc}"

    if _HAS_HTTPX and request is not None:
        try:
            base = str(getattr(request, "base_url", "") or "").rstrip("/")
            if not base:
                return None, None, "Cannot download file without Open WebUI request context."
            url = f"{base}/api/v1/files/{fid}/content"
            headers = {}
            req_headers = getattr(request, "headers", None)
            if req_headers is not None:
                for key in ("cookie", "authorization"):
                    val = req_headers.get(key)
                    if val:
                        headers[key] = val
            async with httpx.AsyncClient(timeout=30.0) as client:
                resp = await client.get(url, headers=headers)
            if resp.status_code in (403, 404):
                return None, None, "File not found or not accessible."
            if resp.status_code >= 400:
                return None, None, f"File download failed (HTTP {resp.status_code})."
            data = resp.content
            disp = resp.headers.get("content-disposition", "")
            fname = "template.pptx"
            if "filename" in disp.lower():
                m = re.search(r"filename\*?=(?:UTF-8''|\"?)([^\";]+)", disp, re.I)
                if m:
                    fname = m.group(1).strip()
            ctype = resp.headers.get("content-type", "")
            if not _is_presentation_file(fname, ctype.split(";")[0].strip()):
                return None, None, "File is not a PowerPoint template (.pptx/.potx)."
            if len(data) > _REFERENCE_PPTX_MAX_BYTES:
                log.warning(
                    "[inspect_slides] reference file %s exceeds %s MB (HTTP)",
                    fid,
                    _REFERENCE_PPTX_MAX_BYTES // (1024 * 1024),
                )
            return data, fname, None
        except Exception as exc:
            log.exception("[inspect_slides] HTTP file download failed")
            return None, None, f"Could not download template file: {exc}"

    return (
        None,
        None,
        "Open WebUI file access unavailable (install httpx or run inside Open WebUI).",
    )


# ============================================================================
# Template clone (Fase 3)
# ============================================================================

_BLANK_LAYOUT_NAMES = frozenset({"blank", "vuota", "empty", "leer"})


def _find_blank_layout(prs):
    if not _HAS_PPTX:
        raise RuntimeError("python-pptx is not installed")
    for layout in prs.slide_layouts:
        name = (getattr(layout, "name", None) or "").strip().lower()
        if name in _BLANK_LAYOUT_NAMES:
            return layout
    return prs.slide_layouts[-1]


def _shape_by_id_in_tree(shape, shape_id: int):
    sid = int(shape_id)
    if int(getattr(shape, "shape_id", 0) or 0) == sid:
        return shape
    for ch in _group_child_shapes(shape):
        found = _shape_by_id_in_tree(ch, sid)
        if found is not None:
            return found
    return None


def _shape_by_id_on_container(container, shape_id: int):
    for sh in container.shapes:
        found = _shape_by_id_in_tree(sh, int(shape_id))
        if found is not None:
            return found
    return None


def _shape_by_id(slide, shape_id: int):
    return _shape_by_id_on_container(slide, shape_id)


def _shape_by_id_recursive(slide, shape_id: int):
    return _shape_by_id_on_container(slide, int(shape_id))


# --- Template clone engine ---------------------------------------------------
#
# The reference .pptx is the graphic source of truth. Cloning works on raw
# OOXML so that every property python-pptx does not model (freeform geometry,
# blip fills, SVG fallbacks, group transforms, letter spacing, ...) survives.
#
# Two invariants:
#   1. Shapes are cloned as *top-level spTree elements*. A shape nested in a
#      p:grpSp is never lifted out of its group: the group element is copied
#      and pruned instead, so grpSpPr/a:xfrm (off/ext/chOff/chExt/rot) keeps
#      mapping child-space coordinates onto the slide.
#   2. Every r:embed found in the copied XML (a:blip, asvg:svgBlip, ...) is
#      re-pointed to a media part copied into the target package. Media parts
#      are copied as opaque blobs (no PIL round-trip) so SVG/EMF are supported.

_SP_TREE_SHAPE_TAGS = frozenset(
    {"sp", "grpSp", "pic", "graphicFrame", "cxnSp", "contentPart"}
)
_GROUP_TAG = "grpSp"


def _oxml_local_name(el) -> str:
    tag = el.tag if isinstance(el.tag, str) else ""
    return tag.rsplit("}", 1)[-1]


def _oxml_shape_id(el) -> Optional[int]:
    """Return the cNvPr id of a spTree element (sp, grpSp, pic, ...)."""
    for child in el:
        name = _oxml_local_name(child)
        if name.startswith("nv") and name.endswith("Pr"):
            cnv = child.find(qn("p:cNvPr"))
            if cnv is not None:
                try:
                    return int(cnv.get("id"))
                except (TypeError, ValueError):
                    return None
    return None


def _oxml_prune_to_ids(el, keep_ids: set[int]):
    """Deep-copy `el` keeping only leaf shapes in `keep_ids`.

    Groups are kept whenever at least one descendant survives, so the group
    transform is preserved. Returns None when nothing survives.
    """
    name = _oxml_local_name(el)
    if name not in _SP_TREE_SHAPE_TAGS:
        return None
    sid = _oxml_shape_id(el)
    if name != _GROUP_TAG:
        return deepcopy(el) if sid is not None and sid in keep_ids else None
    if sid is not None and sid in keep_ids:
        # whole group explicitly requested: keep it verbatim
        return deepcopy(el)
    copy = deepcopy(el)
    survivors = 0
    for child in list(copy):
        if _oxml_local_name(child) not in _SP_TREE_SHAPE_TAGS:
            continue
        pruned = _oxml_prune_to_ids(child, keep_ids)
        if pruned is None:
            copy.remove(child)
        else:
            copy.replace(child, pruned)
            survivors += 1
    return copy if survivors else None


@dataclass
class _CloneContext:
    """Per-target-slide relationship state for media remapping.

    `part_cache` maps a source media partname to the Part already copied into
    the target package, so the same PNG/SVG referenced by several shapes (or
    several slides) is stored once.
    """

    target_part: Any
    part_cache: dict[str, Any] = field(default_factory=dict)
    rid_map: dict[tuple[str, str], str] = field(default_factory=dict)

    def remap_rid(self, source_part, embed: str) -> Optional[str]:
        key = (str(source_part.partname), embed)
        if key in self.rid_map:
            return self.rid_map[key]
        rel = source_part.rels.get(embed)
        if rel is None:
            raise KeyError(f"relationship {embed} not found on {source_part.partname}")
        if rel.is_external:
            return None
        src_media = rel.target_part
        new_part = self._copy_media_part(src_media)
        new_rid = self.target_part.relate_to(new_part, rel.reltype)
        self.rid_map[key] = new_rid
        return new_rid

    def _copy_media_part(self, src_media):
        from pptx.opc.package import Part  # type: ignore

        key = str(src_media.partname)
        cached = self.part_cache.get(key)
        if cached is not None:
            return cached
        package = self.target_part.package
        ext = src_media.partname.ext or "bin"
        partname = package.next_partname(f"/ppt/media/image%d.{ext}")
        new_part = Part(partname, src_media.content_type, package, src_media.blob)
        self.part_cache[key] = new_part
        return new_part


def _remap_media_rels(element, source_part, ctx: _CloneContext) -> None:
    """Re-point every r:embed in `element` to media copied into the target."""
    embed_attr = qn("r:embed")
    for el in element.iter():
        embed = el.get(embed_attr)
        if not embed:
            continue
        try:
            new_rid = ctx.remap_rid(source_part, embed)
        except Exception as exc:
            log.warning(
                "[template] clone media rel %s (%s): %s",
                embed,
                _oxml_local_name(el),
                exc or type(exc).__name__,
            )
            continue
        if new_rid is not None:
            el.set(embed_attr, new_rid)


_TABLE_STYLES_PARTNAME = "/ppt/tableStyles.xml"


def _find_part(package, partname: str):
    for part in package.iter_parts():
        if str(part.partname) == partname:
            return part
    return None


def _copy_table_styles(element, source_part, ctx: _CloneContext) -> None:
    """Copy custom a:tblStyle definitions referenced by cloned tables.

    Built-in PowerPoint style GUIDs render without an entry in
    tableStyles.xml; custom styles do not, so they are appended to the target
    part when missing. Best effort: failures only log.
    """
    style_ids = {
        (el.text or "").strip()
        for el in element.iter(qn("a:tableStyleId"))
        if (el.text or "").strip()
    }
    if not style_ids:
        return
    try:
        from lxml import etree  # type: ignore

        src_part = _find_part(source_part.package, _TABLE_STYLES_PARTNAME)
        tgt_part = _find_part(ctx.target_part.package, _TABLE_STYLES_PARTNAME)
        if src_part is None or tgt_part is None:
            return
        src_root = etree.fromstring(src_part.blob)
        tgt_root = etree.fromstring(tgt_part.blob)
        existing = {s.get("styleId") for s in tgt_root.iter(qn("a:tblStyle"))}
        changed = False
        for style in src_root.iter(qn("a:tblStyle")):
            sid = style.get("styleId")
            if sid in style_ids and sid not in existing:
                tgt_root.append(deepcopy(style))
                existing.add(sid)
                changed = True
        if changed:
            tgt_part._blob = etree.tostring(
                tgt_root, xml_declaration=True, encoding="UTF-8", standalone=True
            )
    except Exception as exc:
        log.debug("[template] table styles copy skipped: %s", exc)


def _bg_element_meaningful(bg_el) -> bool:
    """True when p:bg carries a direct fill (not bgRef-only or empty)."""
    if bg_el is None:
        return False
    try:
        bg_pr = bg_el.find(qn("p:bgPr"))
        if bg_pr is None:
            if bg_el.find(qn("p:bgRef")) is not None:
                log.debug("[template] p:bgRef-only background skipped for clone flatten")
            return False
        if bg_pr.find(qn("a:noFill")) is not None:
            return False
        for tag in ("a:solidFill", "a:blipFill", "a:gradFill"):
            if bg_pr.find(qn(tag)) is not None:
                return True
    except Exception as exc:
        log.debug("[template] bg meaningful check: %s", exc)
    return False


def _resolve_effective_background(source_slide):
    """Effective page background: slide cSld, then layout, then master."""
    if not _HAS_PPTX:
        return None, None, ""
    try:
        chain = [("slide", source_slide.element.cSld, source_slide.part)]
        layout = source_slide.slide_layout
        chain.append(("layout", layout.element.cSld, layout.part))
        master = layout.slide_master
        chain.append(("master", master.element.cSld, master.part))
        for origin, c_sld, part in chain:
            bg = c_sld.find(qn("p:bg"))
            if _bg_element_meaningful(bg):
                return bg, part, origin
    except Exception as exc:
        log.debug("[template] resolve effective background: %s", exc)
    return None, None, ""


def _clone_background(target_slide, source_slide, ctx: Optional[_CloneContext] = None) -> None:
    try:
        tgt_cSld = target_slide.element.cSld
        src_bg, source_part, origin = _resolve_effective_background(source_slide)
        if src_bg is None or source_part is None:
            return
        tgt_bg = tgt_cSld.find(qn("p:bg"))
        if tgt_bg is not None:
            tgt_cSld.remove(tgt_bg)
        new_bg = deepcopy(src_bg)
        tgt_cSld.insert(0, new_bg)
        if ctx is not None:
            _remap_media_rels(new_bg, source_part, ctx)
        log.debug("[template] clone background from %s", origin or "unknown")
    except Exception as exc:
        log.debug("[template] clone background: %s", exc)


def _decoration_container(source_slide, container_name: str):
    if container_name == "layout":
        return source_slide.slide_layout
    if container_name == "master":
        return source_slide.slide_layout.slide_master
    return source_slide


def _clone_shapes_from_container(
    target_slide, container, keep_ids: set[int], ctx: _CloneContext
) -> int:
    """Clone spTree elements of `container` keeping only `keep_ids` leaves."""
    src_tree = container.shapes._spTree
    tgt_tree = target_slide.shapes._spTree
    cloned = 0
    for el in list(src_tree):
        if _oxml_local_name(el) not in _SP_TREE_SHAPE_TAGS:
            continue
        try:
            new_el = _oxml_prune_to_ids(el, keep_ids)
            if new_el is None:
                continue
            _remap_media_rels(new_el, container.part, ctx)
            _copy_table_styles(new_el, container.part, ctx)
            tgt_tree.insert_element_before(new_el, "p:extLst")
            cloned += 1
        except Exception as exc:
            log.warning(
                "[template] clone element %s id=%s failed: %s",
                _oxml_local_name(el),
                _oxml_shape_id(el),
                exc,
            )
    return cloned


def _clone_decorations(
    target_slide,
    source_slide,
    decorations: list[_Decoration],
    ctx: Optional[_CloneContext] = None,
) -> None:
    if not _HAS_PPTX:
        return
    if ctx is None:
        ctx = _CloneContext(target_part=target_slide.part)
    keep_by_container: dict[str, set[int]] = {"master": set(), "layout": set(), "slide": set()}
    for dec in decorations:
        if dec.kind in _CLONE_SKIP_KINDS:
            log.warning(
                "[template] skip clone unsupported kind=%s id=%s",
                dec.kind,
                dec.shape_id,
            )
            continue
        name = dec.source_container if dec.source_container in keep_by_container else "slide"
        keep_by_container[name].add(int(dec.shape_id))
    # z-order: master beneath layout beneath slide; spTree order inside each.
    for name in ("master", "layout", "slide"):
        ids = keep_by_container[name]
        if not ids:
            continue
        container = _decoration_container(source_slide, name)
        _clone_shapes_from_container(target_slide, container, ids, ctx)


def _clone_template_slide_to_prs(
    target_prs,
    source_prs,
    pack: _TemplatePack,
    *,
    source_slide_index: int = 0,
    decorations: list[_Decoration],
    part_cache: Optional[dict] = None,
):
    if not _HAS_PPTX:
        raise RuntimeError("python-pptx is not installed")
    target_prs.slide_width = Inches(pack.slide_width_in)
    target_prs.slide_height = Inches(pack.slide_height_in)
    source_slide = source_prs.slides[source_slide_index]
    target_slide = target_prs.slides.add_slide(_find_blank_layout(target_prs))
    ctx = _CloneContext(
        target_part=target_slide.part,
        part_cache=part_cache if part_cache is not None else {},
    )
    _clone_background(target_slide, source_slide, ctx)
    _clone_decorations(target_slide, source_slide, decorations, ctx)
    _finalize_cloned_slide(target_slide, source_slide, pack.theme)
    return target_slide


def _decoration_visible_on_slide(
    dec: _Decoration, slide_w: float, slide_h: float
) -> bool:
    b = dec.bbox
    ix0 = max(0.0, b.x)
    iy0 = max(0.0, b.y)
    ix1 = min(slide_w, b.x + b.w)
    iy1 = min(slide_h, b.y + b.h)
    iw = max(0.0, ix1 - ix0)
    ih = max(0.0, iy1 - iy0)
    inter = iw * ih
    area = max(b.w * b.h, 1e-9)
    if inter / area < 0.01:
        return False
    cx = b.x + b.w / 2.0
    cy = b.y + b.h / 2.0
    return 0 <= cx <= slide_w and 0 <= cy <= slide_h


def _template_edit_defaults(edits: Optional[dict]) -> dict:
    raw: dict = {}
    if isinstance(edits, dict) and isinstance(edits.get("defaults"), dict):
        raw = edits["defaults"]
    return {
        "drop_text": raw.get("drop_text") is True,
        "drop_placeholders": raw.get("drop_placeholders") is True,
        "drop_offslide": raw.get("drop_offslide", TEMPLATE_EDIT_DEFAULT_DROP_OFFSLIDE),
    }


def _text_decorations_for_slide(st: _SlideTemplate) -> list[_Decoration]:
    out: list[_Decoration] = []
    for sh in st.shapes:
        if not sh.has_text and sh.kind not in ("textbox",):
            continue
        if not sh.cloneable:
            continue
        out.append(
            _Decoration(
                shape_id=sh.id,
                kind=sh.kind,
                name=sh.name,
                bbox=sh.bbox,
                z_order=sh.z_order,
            )
        )
    return out


def _apply_template_edits(
    pack: _TemplatePack, edits: Optional[dict]
) -> _TemplatePack:
    """Filter decorations per template_edits (§6.3). Mutates a logical copy of pack."""
    defaults = _template_edit_defaults(edits)
    drop_text = defaults["drop_text"]
    drop_placeholders = defaults["drop_placeholders"]
    drop_offslide = defaults["drop_offslide"]
    slide_rules: dict[int, dict] = {}
    if isinstance(edits, dict):
        for rule in _as_list(edits.get("slides")):
            if isinstance(rule, dict) and "index" in rule:
                slide_rules[int(rule["index"])] = rule

    placeholder_ids: dict[int, set[int]] = {}
    for st in pack.slides:
        placeholder_ids[st.index] = {sh.id for sh in st.shapes if sh.is_placeholder}

    new_slides: list[_SlideTemplate] = []
    for st in pack.slides:
        decs = list(st.decorations)
        if not drop_text:
            seen = {d.shape_id for d in decs}
            for td in _text_decorations_for_slide(st):
                if td.shape_id not in seen:
                    decs.append(td)
                    seen.add(td.shape_id)
        if drop_placeholders:
            ph = placeholder_ids.get(st.index) or set()
            decs = [d for d in decs if d.shape_id not in ph]
        if drop_offslide:
            decs = [
                d
                for d in decs
                if _decoration_visible_on_slide(
                    d, pack.slide_width_in, pack.slide_height_in
                )
            ]

        rule = slide_rules.get(st.index)
        if rule is not None:
            drop_ids = set(rule.get("drop_ids") or [])
            keep_ids = rule.get("keep_ids", None)
            if keep_ids is not None:
                if keep_ids == []:
                    log.warning(
                        "[template] keep_ids empty for slide index %s — "
                        "no decorations will be cloned (inspect hint)",
                        st.index,
                    )
                    decs = []
                else:
                    keep_set = set(keep_ids)
                    decs = [d for d in decs if d.shape_id in keep_set]
                    decs = [d for d in decs if d.shape_id not in drop_ids]
            elif drop_ids:
                for did in drop_ids:
                    if not any(d.shape_id == did for d in st.decorations):
                        log.debug(
                            "[template] drop_ids %s not on slide %s",
                            did,
                            st.index,
                        )
                decs = [d for d in decs if d.shape_id not in drop_ids]

        new_slides.append(replace(st, decorations=decs))

    return replace(pack, slides=new_slides)


# --- Per-slide reuse libretto (clone template slide by inspect shape ids) ----

@dataclass
class _ReuseTable:
    """New content for a template table: `headers` → row 0, `rows` → body."""

    headers: Optional[list[str]]
    rows: list[list[str]]

    @property
    def grid(self) -> list[list[str]]:
        return ([self.headers] if self.headers is not None else []) + self.rows

    @property
    def n_cols(self) -> int:
        return max((len(r) for r in self.grid), default=0)


@dataclass
class _ReuseSpec:
    slide_index: int
    keep_ids: Optional[list[int]]
    drop_ids: list[int]
    text: dict[int, "str | _ReuseTable"]


def _coerce_cells(raw_rows: Any, where: str) -> list[list[str]]:
    if not isinstance(raw_rows, list):
        raise ValueError(f"{where} must be a list of rows")
    out: list[list[str]] = []
    for row in raw_rows:
        if isinstance(row, dict):
            row = list(row.values())
        if not isinstance(row, list):
            raise ValueError(f"{where} rows must be lists of cell values")
        out.append(["" if c is None else str(c) for c in row])
    return out


def _parse_reuse_table(raw: Any, where: str) -> _ReuseTable:
    """Accept {"headers": [...], "rows": [[...]]} or a plain list of rows."""
    if isinstance(raw, list):
        return _ReuseTable(headers=None, rows=_coerce_cells(raw, where))
    if not isinstance(raw, dict):
        raise ValueError(
            f"{where} must be a string (text shape) or an object "
            '{"headers": [...], "rows": [[...]]} (table shape)'
        )
    headers = raw.get("headers")
    if headers is not None:
        if not isinstance(headers, list):
            raise ValueError(f"{where}.headers must be a list")
        headers = ["" if h is None else str(h) for h in headers]
    rows = _coerce_cells(raw.get("rows") or [], f"{where}.rows")
    if headers is None and not rows:
        raise ValueError(f"{where} table has neither headers nor rows")
    return _ReuseTable(headers=headers, rows=rows)


def _parse_reuse_spec(raw: Any, pack: _TemplatePack) -> _ReuseSpec:
    if not isinstance(raw, dict):
        raise ValueError(
            "slides[].reuse must be an object with "
            '"slide", optional "keep_ids", "drop_ids", and "text" '
            "(inspect shapes[].id) — not a list or string."
        )
    if "slide" not in raw:
        raise ValueError('slides[].reuse requires "slide" (0-based template index)')
    try:
        slide_index = int(raw["slide"])
    except (TypeError, ValueError) as exc:
        raise ValueError(
            f'slides[].reuse.slide must be an integer, got {raw["slide"]!r}'
        ) from exc
    n = len(pack.slides)
    if slide_index < 0 or slide_index >= n:
        raise ValueError(
            f"slides[].reuse.slide {slide_index} out of range "
            f"(template has {n} slides, indices 0..{n - 1})"
        )

    keep_ids: Optional[list[int]] = None
    if "keep_ids" in raw and raw["keep_ids"] is not None:
        if not isinstance(raw["keep_ids"], list):
            raise ValueError("slides[].reuse.keep_ids must be a list of shape ids")
        keep_ids = [int(x) for x in raw["keep_ids"]]

    drop_ids: list[int] = []
    if raw.get("drop_ids") is not None:
        if not isinstance(raw["drop_ids"], list):
            raise ValueError("slides[].reuse.drop_ids must be a list of shape ids")
        drop_ids = [int(x) for x in raw["drop_ids"]]

    text: dict[int, "str | _ReuseTable"] = {}
    raw_text = raw.get("text")
    if raw_text is not None:
        if not isinstance(raw_text, dict):
            raise ValueError(
                "slides[].reuse.text must be an object mapping shape id → string "
                '(text shape) or {"headers": [...], "rows": [[...]]} (table shape)'
            )
        for k, v in raw_text.items():
            try:
                sid = int(k)
            except (TypeError, ValueError) as exc:
                raise ValueError(
                    f"slides[].reuse.text key {k!r} must be a shape id (integer)"
                ) from exc
            if isinstance(v, (dict, list)):
                text[sid] = _parse_reuse_table(v, f"slides[].reuse.text[{sid}]")
            else:
                text[sid] = "" if v is None else str(v)

    return _ReuseSpec(
        slide_index=slide_index,
        keep_ids=keep_ids,
        drop_ids=drop_ids,
        text=text,
    )


def _decoration_from_record(rec: _ShapeRecord) -> _Decoration:
    return _Decoration(
        shape_id=rec.id,
        kind=rec.kind,
        name=rec.name,
        bbox=rec.bbox,
        z_order=rec.z_order,
        source_container="slide",
    )


def _decorations_for_reuse(
    st: _SlideTemplate,
    keep_ids: Optional[list[int]],
    drop_ids: list[int],
) -> list[_Decoration]:
    drop = set(drop_ids)
    known = {rec.id for rec in st.shapes}
    if keep_ids is not None:
        for kid in keep_ids:
            if kid not in known:
                log.warning(
                    "[reuse] keep_ids %s not on template slide %s (inspect shapes[].id)",
                    kid,
                    st.index,
                )
    for tid in drop:
        if tid not in known:
            log.warning(
                "[reuse] drop_ids %s not on template slide %s",
                tid,
                st.index,
            )

    out: list[_Decoration] = []
    if keep_ids is not None:
        keep = set(keep_ids)
        for rec in st.shapes:
            if rec.id not in keep or rec.id in drop:
                continue
            if not rec.cloneable:
                log.warning(
                    "[reuse] skip non-cloneable shape id=%s kind=%s on slide %s",
                    rec.id,
                    rec.kind,
                    st.index,
                )
                continue
            out.append(_decoration_from_record(rec))
    else:
        for rec in st.shapes:
            if rec.id in drop:
                continue
            if not rec.cloneable:
                continue
            out.append(_decoration_from_record(rec))
    return out


def _first_run_rPr_template(txBody):
    """Return a deep copy of the first run properties found in `txBody`.

    Falls back to the first paragraph's a:endParaRPr (renamed to a:rPr) so a
    cloned empty placeholder still keeps size/spacing/color. Returns None when
    no run-level formatting is available.
    """
    for p in txBody.findall(qn("a:p")):
        for r in p.findall(qn("a:r")):
            rPr = r.find(qn("a:rPr"))
            if rPr is not None:
                return deepcopy(rPr)
    for p in txBody.findall(qn("a:p")):
        end = p.find(qn("a:endParaRPr"))
        if end is not None:
            rPr = deepcopy(end)
            rPr.tag = qn("a:rPr")
            return rPr
    return None


def _set_shape_text_preserve_font(shape, text: str) -> None:
    """Replace the text of a text shape keeping the original run formatting."""
    if not getattr(shape, "has_text_frame", False):
        return
    _replace_txBody_text(shape.text_frame._txBody, text)


def _replace_txBody_text(txBody, text: str) -> None:
    """Rewrite the paragraphs of an a:txBody / p:txBody with `text`.

    The first paragraph's a:pPr, the first run's full a:rPr (font, size,
    bold, color, letter spacing `spc`, caps, lang, ...) and a:endParaRPr are
    copied verbatim onto every new paragraph/run. `\\n` creates a paragraph.
    """
    from lxml import etree  # type: ignore

    paras = txBody.findall(qn("a:p"))
    first_p = paras[0] if paras else None
    pPr_tpl = None
    end_tpl = None
    if first_p is not None:
        pPr_src = first_p.find(qn("a:pPr"))
        end_src = first_p.find(qn("a:endParaRPr"))
        pPr_tpl = deepcopy(pPr_src) if pPr_src is not None else None
        end_tpl = deepcopy(end_src) if end_src is not None else None
    rPr_tpl = _first_run_rPr_template(txBody)

    for p in paras:
        txBody.remove(p)
    lines = text.split("\n") if text else [""]
    for line in lines:
        p = etree.SubElement(txBody, qn("a:p"))
        if pPr_tpl is not None:
            p.append(deepcopy(pPr_tpl))
        r = etree.SubElement(p, qn("a:r"))
        if rPr_tpl is not None:
            r.append(deepcopy(rPr_tpl))
        t = etree.SubElement(r, qn("a:t"))
        t.text = line
        if end_tpl is not None:
            p.append(deepcopy(end_tpl))


# --- Presentation edit engine (post-generate; does not touch reuse/clone) ----

_EDIT_PROGRESS_OPS_EVERY = 5
_EDIT_LARGE_DECK_SLIDES = 10


def _edit_slide_json_to_internal(slide_json: int) -> int:
    """Map JSON operations[].slide to 0-based python-pptx slide index."""
    origin = EDIT_SLIDE_INDEX_ORIGIN
    if origin == 1:
        if slide_json < 1:
            raise ValueError(
                "slide must be >= 1 (EDIT_SLIDE_INDEX_ORIGIN=1, same numbering "
                "as PowerPoint / LibreOffice)."
            )
        return slide_json - 1
    if origin == 0:
        if slide_json < 0:
            raise ValueError(
                "slide must be >= 0 (EDIT_SLIDE_INDEX_ORIGIN=0, 0-based index)."
            )
        return slide_json
    raise ValueError(
        f"EDIT_SLIDE_INDEX_ORIGIN must be 0 or 1, got {origin!r}."
    )


def _edit_slide_internal_to_json(slide_internal: int) -> int:
    """User-facing slide number for error messages."""
    if EDIT_SLIDE_INDEX_ORIGIN == 1:
        return slide_internal + 1
    return slide_internal


def _edit_slide_range_hint(deck_slide_count: int) -> str:
    n = deck_slide_count
    if EDIT_SLIDE_INDEX_ORIGIN == 1:
        return f"valid slides 1..{n}" if n else "deck has no slides"
    if n <= 0:
        return "deck has no slides"
    return f"valid slides 0..{n - 1}"


def _presentation_id_is_template_attachment(file_id: str, messages: Any) -> bool:
    attached = _find_pptx_attachment(messages)
    fid = (file_id or "").strip()
    return bool(attached and fid and attached == fid)


def _bad_presentation_file_id_message(sent_id: str) -> str:
    sent = (sent_id or "").strip()
    if not sent:
        why = "edit_presentation requires presentation_file_id."
    elif sent.startswith("http") or "/cache/files/" in sent:
        why = (
            "presentation_file_id must be a Files API UUID from the last "
            "generate_slides or edit_presentation tool message, not a cache URL."
        )
    elif not _is_files_api_id(sent):
        why = f"presentation_file_id {sent!r} is not a Files API UUID."
    else:
        why = f"presentation_file_id {sent!r} is invalid."
    return (
        f"{why} Copy the file id character for character from the file block "
        "returned by generate_slides or edit_presentation. Never use "
        "reference_file_id or the template inspect file_id."
    )


def _parse_edit_json_content(content: Any) -> dict:
    if isinstance(content, dict):
        return content
    if not isinstance(content, str):
        raise ValueError("The `content` parameter must be a JSON object.")
    try:
        return json.loads(content)
    except json.JSONDecodeError:
        cleaned = (content or "").strip()
        for pre in ("```json", "```"):
            if cleaned.startswith(pre):
                cleaned = cleaned[len(pre):]
        if cleaned.endswith("```"):
            cleaned = cleaned[:-3]
        try:
            return json.loads(cleaned.strip())
        except json.JSONDecodeError as exc:
            raise ValueError(f"Invalid JSON: {exc}") from exc


_EDIT_KNOWN_OPS = (
    "set_text",
    "set_font_pt",
    "shrink_font",
    "fit_text",
    "enable_autofit",
    "replace_text_and_fit",
    "split_text",
    "resize_shape",
    "set_table_cell",
)

_EDIT_SHRINK_MODES = ("fit", "to_min", "target")


def _edit_shrink_params(op: dict, i: int, valves: Any) -> tuple[float, float, int]:
    min_pt = op.get("min_pt")
    if min_pt is None:
        min_pt = getattr(valves, "presentation_edit_min_font_pt", 14.0)
    step_pt = op.get("step_pt")
    if step_pt is None:
        step_pt = getattr(valves, "presentation_edit_shrink_step_pt", 2.0)
    max_iter = op.get("max_iterations", 20)
    return float(min_pt), float(step_pt), int(max_iter)


def _edit_shrink_mode(op: dict, i: int) -> str:
    mode = (op.get("mode") or "fit").strip()
    if mode not in _EDIT_SHRINK_MODES:
        raise ValueError(
            f"operations[{i}] shrink mode must be one of: "
            + ", ".join(_EDIT_SHRINK_MODES)
            + "."
        )
    return mode


def _edit_apply_shrink_fields(entry: dict, op: dict, i: int, valves: Any) -> None:
    min_pt, step_pt, max_iter = _edit_shrink_params(op, i, valves)
    entry["min_pt"] = min_pt
    entry["step_pt"] = step_pt
    entry["max_iterations"] = max_iter
    mode = _edit_shrink_mode(op, i)
    entry["shrink_mode"] = mode
    if mode == "target":
        if op.get("target_pt") is None:
            raise ValueError(f"operations[{i}] shrink mode target requires target_pt.")
        entry["target_pt"] = float(op["target_pt"])


def _normalize_edit_op(op: dict, i: int, valves: Any) -> dict:
    if not isinstance(op, dict):
        raise ValueError(f"operations[{i}] must be an object.")
    name = (op.get("op") or "").strip()
    if not name:
        raise ValueError(f"operations[{i}] missing op.")
    try:
        slide_json = int(op["slide"])
    except (KeyError, TypeError, ValueError) as exc:
        raise ValueError(f"operations[{i}] requires integer slide.") from exc
    try:
        slide_idx = _edit_slide_json_to_internal(slide_json)
    except ValueError as exc:
        raise ValueError(f"operations[{i}]: {exc}") from exc

    if name == "split_text":
        try:
            from_id = int(op["from_shape_id"])
            to_id = int(op["to_shape_id"])
        except (KeyError, TypeError, ValueError) as exc:
            raise ValueError(
                f"operations[{i}] split_text requires from_shape_id and to_shape_id."
            ) from exc
        mode = (op.get("mode") or "").strip()
        if mode not in ("first_paragraph", "first_line"):
            raise ValueError(
                f"operations[{i}] split_text mode must be "
                "first_paragraph or first_line."
            )
        return {
            "op": name,
            "slide": slide_idx,
            "from_shape_id": from_id,
            "to_shape_id": to_id,
            "mode": mode,
        }

    if name == "set_table_cell":
        try:
            shape_id = int(op["shape_id"])
            row = int(op["row"])
            col = int(op["col"])
        except (KeyError, TypeError, ValueError) as exc:
            raise ValueError(
                f"operations[{i}] set_table_cell requires shape_id, row, col."
            ) from exc
        if "text" not in op or not isinstance(op["text"], str):
            raise ValueError(f"operations[{i}] set_table_cell requires text string.")
        return {
            "op": name,
            "slide": slide_idx,
            "shape_id": shape_id,
            "row": row,
            "col": col,
            "text": op["text"],
        }

    try:
        shape_id = int(op["shape_id"])
    except (KeyError, TypeError, ValueError) as exc:
        raise ValueError(f"operations[{i}] requires integer shape_id.") from exc

    entry: dict = {"op": name, "slide": slide_idx, "shape_id": shape_id}

    if name == "set_text":
        if "text" not in op or not isinstance(op["text"], str):
            raise ValueError(f"operations[{i}] set_text requires text.")
        entry["text"] = op["text"]
    elif name == "set_font_pt":
        has_pt = op.get("font_pt") is not None
        has_delta = op.get("font_pt_delta") is not None
        if has_pt and has_delta:
            raise ValueError(
                f"operations[{i}] set_font_pt: use font_pt or font_pt_delta, not both."
            )
        if not has_pt and not has_delta:
            raise ValueError(
                f"operations[{i}] set_font_pt requires font_pt or font_pt_delta."
            )
        if has_pt:
            entry["font_pt"] = float(op["font_pt"])
        else:
            entry["font_pt_delta"] = float(op["font_pt_delta"])
    elif name in ("shrink_font", "fit_text"):
        _edit_apply_shrink_fields(entry, op, i, valves)
    elif name == "enable_autofit":
        pass
    elif name == "replace_text_and_fit":
        if "text" not in op or not isinstance(op["text"], str):
            raise ValueError(f"operations[{i}] replace_text_and_fit requires text.")
        entry["text"] = op["text"]
        _edit_apply_shrink_fields(entry, op, i, valves)
        entry["shrink_mode"] = "fit"
    elif name == "resize_shape":
        delta_emu = op.get("delta_height_emu")
        delta_in = op.get("delta_height_in")
        if delta_emu is None and delta_in is None:
            raise ValueError(
                f"operations[{i}] resize_shape requires delta_height_in or "
                "delta_height_emu."
            )
        if delta_emu is None:
            delta_emu = int(float(delta_in) * EMU_IN)
        else:
            delta_emu = int(delta_emu)
        entry["delta_height_emu"] = delta_emu
    else:
        raise ValueError(
            f"operations[{i}] unknown op {name!r}; known: "
            + ", ".join(_EDIT_KNOWN_OPS)
            + "."
        )
    return entry


def _parse_edit_spec(content: Any, valves: Any) -> dict:
    """Validate edit_presentation JSON (structure only; shape checks at apply)."""
    spec = _parse_edit_json_content(content)
    if not isinstance(spec, dict):
        raise ValueError("The `content` parameter must be a JSON object.")

    fid = (spec.get("presentation_file_id") or "").strip()
    if not fid:
        raise ValueError(_bad_presentation_file_id_message(""))
    if not _is_files_api_id(fid):
        raise ValueError(_bad_presentation_file_id_message(fid))

    ops_raw = spec.get("operations")
    if not isinstance(ops_raw, list) or not ops_raw:
        raise ValueError("operations must be a non-empty array.")

    normalized = [_normalize_edit_op(op, i, valves) for i, op in enumerate(ops_raw)]

    spec = dict(spec)
    spec["presentation_file_id"] = fid
    spec["operations"] = normalized
    return spec


def _edit_txBody_runs(txBody):
    from lxml import etree  # type: ignore

    for p in txBody.findall(qn("a:p")):
        for r in p.findall(qn("a:r")):
            yield r


def _edit_rPr_sz_centipoints(r_pr) -> int:
    if r_pr is None:
        return 0
    sz = r_pr.get("sz")
    if sz is None:
        return 0
    try:
        return max(0, int(sz))
    except (TypeError, ValueError):
        return 0


def _edit_shape_font_pt_effective(shape) -> float:
    """Largest font size (pt) from runs, else paragraph defaults; else 18.0."""
    if not getattr(shape, "has_text_frame", False):
        return 18.0
    tx_body = shape.text_frame._txBody
    max_sz = 0
    for r in _edit_txBody_runs(tx_body):
        max_sz = max(max_sz, _edit_rPr_sz_centipoints(r.find(qn("a:rPr"))))
    if max_sz <= 0:
        paras = tx_body.findall(qn("a:p"))
        if paras:
            first_p = paras[0]
            max_sz = max(
                max_sz,
                _edit_rPr_sz_centipoints(first_p.find(qn("a:endParaRPr"))),
            )
            p_pr = first_p.find(qn("a:pPr"))
            if p_pr is not None:
                max_sz = max(
                    max_sz,
                    _edit_rPr_sz_centipoints(p_pr.find(qn("a:defRPr"))),
                )
    if max_sz <= 0:
        return 18.0
    return max_sz / 100.0


def _edit_shape_font_pt_max(shape) -> float:
    """Alias for effective font size (edit engine)."""
    return _edit_shape_font_pt_effective(shape)


def _edit_set_all_runs_sz_pt(shape, pt: float) -> None:
    if not getattr(shape, "has_text_frame", False):
        return
    sz_val = str(max(100, int(round(pt * 100))))
    tx_body = shape.text_frame._txBody
    for r in _edit_txBody_runs(tx_body):
        rPr = r.find(qn("a:rPr"))
        if rPr is None:
            from lxml import etree  # type: ignore

            rPr = etree.SubElement(r, qn("a:rPr"))
        rPr.set("sz", sz_val)


def _text_fits_shape_heuristic(shape) -> bool:
    """Rough v1 fit check — not a PowerPoint layout engine (§5.1 #4A)."""
    if not getattr(shape, "has_text_frame", False):
        return True
    text = shape.text or ""
    w_emu = int(getattr(shape, "width", 0) or 0)
    h_emu = int(getattr(shape, "height", 0) or 0)
    if w_emu <= 0 or h_emu <= 0:
        return True
    font_pt = _edit_shape_font_pt_effective(shape)
    pt_emu = 914400.0 / 72.0
    line_h = font_pt * 1.25 * pt_emu
    char_w = max(font_pt * 0.55 * pt_emu, 1.0)
    chars_per_line = max(1, int(w_emu / char_w))
    flat = text.replace("\n", " ")
    wrapped_lines = max(1, (len(flat) + chars_per_line - 1) // chars_per_line)
    explicit_lines = text.count("\n") + 1
    est_lines = max(wrapped_lines, explicit_lines)
    return est_lines * line_h <= h_emu * 1.05


def _shrink_shape_font(
    shape,
    *,
    min_pt: float,
    step_pt: float,
    max_iterations: int,
    mode: str = "fit",
    target_pt: Optional[float] = None,
) -> bool:
    """Reduce (or adjust to target) a:rPr/@sz — edit-only helper."""
    if not getattr(shape, "has_text_frame", False):
        raise ValueError(f"shape {shape.shape_id} has no text frame")
    shrink_mode = mode or "fit"
    if shrink_mode not in _EDIT_SHRINK_MODES:
        raise ValueError(f"unsupported shrink mode {shrink_mode!r}.")
    if shrink_mode == "target" and target_pt is None:
        raise ValueError("shrink mode target requires target_pt.")
    changed = False
    for _ in range(max(1, max_iterations)):
        cur = _edit_shape_font_pt_effective(shape)
        if shrink_mode == "fit":
            if _text_fits_shape_heuristic(shape):
                break
            if cur <= min_pt:
                break
            new_pt = max(min_pt, cur - step_pt)
        elif shrink_mode == "to_min":
            if cur <= min_pt:
                break
            new_pt = max(min_pt, cur - step_pt)
        else:
            tgt = float(target_pt)
            if abs(cur - tgt) < 0.05:
                break
            if cur > tgt:
                new_pt = max(tgt, min_pt, cur - step_pt)
            else:
                new_pt = min(tgt, cur + step_pt)
        if new_pt >= cur and shrink_mode in ("fit", "to_min"):
            break
        if abs(new_pt - cur) < 0.05:
            break
        _edit_set_all_runs_sz_pt(shape, new_pt)
        changed = True
    return changed


def _edit_set_shape_font_pt(shape, pt: float) -> bool:
    """Set all runs to pt; return True if effective size changed."""
    if not getattr(shape, "has_text_frame", False):
        raise ValueError(f"shape {shape.shape_id} has no text frame")
    before = _edit_shape_font_pt_effective(shape)
    pt = max(1.0, float(pt))
    _edit_set_all_runs_sz_pt(shape, pt)
    after = _edit_shape_font_pt_effective(shape)
    return abs(after - before) > 0.05 or abs(after - pt) < 0.05


def _resolve_edit_shape(
    prs,
    slide_idx: int,
    shape_id: int,
    *,
    allow_table: bool = False,
):
    n = len(prs.slides)
    if slide_idx < 0 or slide_idx >= n:
        shown = _edit_slide_internal_to_json(slide_idx)
        raise ValueError(
            f"slide {shown} out of range (deck has {n} slide(s); "
            f"{_edit_slide_range_hint(n)}; EDIT_SLIDE_INDEX_ORIGIN="
            f"{EDIT_SLIDE_INDEX_ORIGIN})."
        )
    slide = prs.slides[slide_idx]
    shape = _shape_by_id_recursive(slide, shape_id)
    if shape is None:
        slide_label = _edit_slide_internal_to_json(slide_idx)
        raise ValueError(
            f"shape_id {shape_id} not found on slide {slide_label}."
        )
    if getattr(shape, "has_table", False) and not allow_table:
        raise ValueError(
            f"shape_id {shape_id} is a table; use set_table_cell for tables."
        )
    return shape


def _edit_enable_norm_autofit(shape) -> bool:
    """Set a:bodyPr normAutofit so PowerPoint shrinks text on open (edit-only)."""
    from lxml import etree  # type: ignore

    if not getattr(shape, "has_text_frame", False):
        raise ValueError(f"shape {shape.shape_id} has no text frame")
    tx_body = shape.text_frame._txBody
    body = tx_body.find(qn("a:bodyPr"))
    if body is None:
        body = etree.Element(qn("a:bodyPr"))
        tx_body.insert(0, body)
    had_norm = body.find(qn("a:normAutofit")) is not None
    for tag in ("a:noAutofit", "a:spAutoFit"):
        for el in list(body.findall(qn(tag))):
            body.remove(el)
    if body.find(qn("a:normAutofit")) is None:
        etree.SubElement(body, qn("a:normAutofit"))
    return not had_norm or body.find(qn("a:noAutofit")) is None


def _replace_text_and_fit(
    shape,
    text: str,
    *,
    min_pt: float,
    step_pt: float,
    max_iterations: int,
) -> bool:
    before = (shape.text or "") if getattr(shape, "has_text_frame", False) else ""
    _set_shape_text_preserve_font(shape, text)
    text_changed = (shape.text or "") != before
    shrunk = _shrink_shape_font(
        shape,
        min_pt=min_pt,
        step_pt=step_pt,
        max_iterations=max_iterations,
        mode="fit",
    )
    return text_changed or shrunk


def _edit_split_text(from_shape, to_shape, mode: str) -> bool:
    if not getattr(from_shape, "has_text_frame", False):
        raise ValueError(f"shape {from_shape.shape_id} has no text frame")
    if not getattr(to_shape, "has_text_frame", False):
        raise ValueError(f"shape {to_shape.shape_id} has no text frame")
    raw = from_shape.text or ""
    paras = [
        (p.text or "").strip()
        for p in from_shape.text_frame.paragraphs
        if (p.text or "").strip()
    ]
    if mode == "first_paragraph":
        if "\n\n" in raw:
            keep, move = raw.split("\n\n", 1)
        elif len(paras) >= 2:
            keep, move = paras[0], "\n\n".join(paras[1:])
        elif "\n" in raw:
            keep, move = raw.split("\n", 1)
        else:
            raise ValueError("split_text: no paragraph boundary in source shape")
    else:
        if "\n" in raw:
            keep, move = raw.split("\n", 1)
        elif len(paras) >= 2:
            keep, move = paras[0], "\n".join(paras[1:])
        else:
            raise ValueError("split_text: no line break in source shape")
    keep = keep.strip()
    move = move.strip()
    if not move:
        raise ValueError("split_text: nothing to move after split")
    _set_shape_text_preserve_font(from_shape, keep)
    existing = (to_shape.text or "").strip()
    dest = f"{existing}\n{move}" if existing else move
    _set_shape_text_preserve_font(to_shape, dest)
    return True


_EDIT_MAX_HEIGHT_DELTA_FRAC = 0.20


def _edit_resize_shape_height(prs, shape, delta_emu: int) -> bool:
    max_delta = int(prs.slide_height * _EDIT_MAX_HEIGHT_DELTA_FRAC)
    if abs(delta_emu) > max_delta:
        raise ValueError(
            f"resize_shape delta exceeds max ±{max_delta} EMU (~20% of slide height)."
        )
    min_h = int(EMU_IN * 0.25)
    cur = int(shape.height)
    new_h = cur + int(delta_emu)
    if new_h < min_h:
        raise ValueError("resize_shape would make the shape too small.")
    shape.height = new_h
    return True


def _edit_set_table_cell(shape, row: int, col: int, text: str) -> bool:
    if not getattr(shape, "has_table", False):
        raise ValueError(f"shape {shape.shape_id} is not a table")
    tbl = shape.table
    n_rows = len(tbl.rows)
    n_cols = len(tbl.columns)
    if row < 0 or row >= n_rows or col < 0 or col >= n_cols:
        raise ValueError(
            f"set_table_cell row={row} col={col} out of range "
            f"({n_rows}x{n_cols})."
        )
    cell = tbl.cell(row, col)
    before = (cell.text or "").strip()
    _replace_txBody_text(cell.text_frame._txBody, text)
    return (cell.text or "").strip() != before


def _apply_edit_op(prs, op: dict, valves: Any, op_index: int) -> bool:
    name = op["op"]
    if name == "split_text":
        from_sh = _resolve_edit_shape(
            prs, op["slide"], op["from_shape_id"]
        )
        to_sh = _resolve_edit_shape(prs, op["slide"], op["to_shape_id"])
        return _edit_split_text(from_sh, to_sh, op["mode"])

    if name == "set_table_cell":
        shape = _resolve_edit_shape(
            prs, op["slide"], op["shape_id"], allow_table=True
        )
        return _edit_set_table_cell(shape, op["row"], op["col"], op["text"])

    shape = _resolve_edit_shape(prs, op["slide"], op["shape_id"])
    if name == "set_text":
        if not getattr(shape, "has_text_frame", False):
            raise ValueError(
                f"shape_id {op['shape_id']} has no text frame (set_text)."
            )
        before = shape.text or ""
        _set_shape_text_preserve_font(shape, op["text"])
        return (shape.text or "") != before
    if name == "set_font_pt":
        if not getattr(shape, "has_text_frame", False):
            raise ValueError(
                f"shape_id {op['shape_id']} has no text frame (set_font_pt)."
            )
        if "font_pt" in op:
            return _edit_set_shape_font_pt(shape, op["font_pt"])
        cur = _edit_shape_font_pt_effective(shape)
        return _edit_set_shape_font_pt(shape, cur + float(op["font_pt_delta"]))
    if name in ("shrink_font", "fit_text"):
        return _shrink_shape_font(
            shape,
            min_pt=op["min_pt"],
            step_pt=op["step_pt"],
            max_iterations=op["max_iterations"],
            mode=op.get("shrink_mode", "fit"),
            target_pt=op.get("target_pt"),
        )
    if name == "enable_autofit":
        return _edit_enable_norm_autofit(shape)
    if name == "replace_text_and_fit":
        return _replace_text_and_fit(
            shape,
            op["text"],
            min_pt=op["min_pt"],
            step_pt=op["step_pt"],
            max_iterations=op["max_iterations"],
        )
    if name == "resize_shape":
        return _edit_resize_shape_height(prs, shape, op["delta_height_emu"])
    raise ValueError(f"unsupported op {name!r}.")


def _apply_edit_operations(prs, operations: list, valves: Any) -> int:
    mutated = 0
    for i, op in enumerate(operations):
        try:
            if _apply_edit_op(prs, op, valves, i):
                mutated += 1
        except ValueError as exc:
            raise ValueError(f"operations[{i}]: {exc}") from exc
    return mutated


def _fill_table_preserve_format(shape, data: _ReuseTable) -> None:
    """Refill a cloned template table in place, keeping its cell formatting.

    - Row 0 of the template is the header row; body rows are style templates
      reused cyclically (so alternating fills survive when rows are added).
    - Columns are added by copying the last column, removed from the right;
      the remaining gridCol widths are rescaled to the original total width.
    - The graphicFrame height follows the new row count.
    """
    if not getattr(shape, "has_table", False):
        raise ValueError(f"shape id {shape.shape_id} is not a table")
    tbl = shape.table._tbl
    grid = data.grid
    if not grid:
        return
    n_rows = len(grid)
    n_cols = max(data.n_cols, 1)

    # --- columns ------------------------------------------------------------
    tblGrid = tbl.tblGrid
    grid_cols = list(tblGrid.gridCol_lst)
    total_w = sum(int(g.w) for g in grid_cols) or int(shape.width)
    trs = list(tbl.tr_lst)
    while len(grid_cols) > n_cols:
        tblGrid.remove(grid_cols.pop())
        for tr in trs:
            tcs = tr.tc_lst
            if tcs:
                tr.remove(tcs[-1])
    while len(grid_cols) < n_cols:
        new_col = deepcopy(grid_cols[-1])
        tblGrid.append(new_col)
        grid_cols.append(new_col)
        for tr in trs:
            tcs = tr.tc_lst
            if tcs:
                new_tc = deepcopy(tcs[-1])
                tcs[-1].addnext(new_tc)
    cur_total = sum(int(g.w) for g in grid_cols) or 1
    for g in grid_cols:
        g.w = max(1, int(round(int(g.w) * total_w / cur_total)))

    # --- rows ---------------------------------------------------------------
    header_tpl = trs[0]
    body_tpls = trs[1:] or [trs[0]]
    for tr in trs:
        tbl.remove(tr)
    out_rows = []
    for r_idx in range(n_rows):
        if r_idx == 0 and data.headers is not None:
            tpl = header_tpl
        else:
            body_idx = r_idx - (1 if data.headers is not None else 0)
            tpl = body_tpls[body_idx % len(body_tpls)]
        new_tr = deepcopy(tpl)
        tbl.append(new_tr)
        out_rows.append(new_tr)

    # --- cell text ----------------------------------------------------------
    for r_idx, tr in enumerate(out_rows):
        values = grid[r_idx]
        for c_idx, tc in enumerate(tr.tc_lst):
            txt = values[c_idx] if c_idx < len(values) else ""
            txBody = tc.find(qn("a:txBody"))
            if txBody is None:
                continue
            _replace_txBody_text(txBody, txt)

    # --- frame height -------------------------------------------------------
    try:
        shape.height = sum(int(tr.h) for tr in out_rows)
    except Exception as exc:
        log.debug("[reuse] table frame height not updated: %s", exc)


def _apply_reuse_text_map(
    target_slide,
    text_map: dict[int, "str | _ReuseTable"],
    st: _SlideTemplate,
) -> None:
    known = {rec.id for rec in st.shapes}
    for sid, new_text in text_map.items():
        if sid not in known:
            log.warning(
                "[reuse] text id %s not on template slide %s",
                sid,
                st.index,
            )
        shape = _shape_by_id(target_slide, sid)
        if shape is None:
            log.warning(
                "[reuse] shape id %s missing on cloned output slide",
                sid,
            )
            continue
        is_table = bool(getattr(shape, "has_table", False))
        if isinstance(new_text, _ReuseTable):
            if not is_table:
                raise ValueError(
                    f"slides[].reuse.text[{sid}] is a table payload but shape "
                    f"{sid} is not a table (kind from inspect shapes[])"
                )
            _fill_table_preserve_format(shape, new_text)
        elif is_table:
            raise ValueError(
                f"slides[].reuse.text[{sid}] targets a table: pass "
                '{"headers": [...], "rows": [[...]]} instead of a string'
            )
        else:
            _set_shape_text_preserve_font(shape, new_text)


def _reuse_table_from_slide_spec(
    slide_dict: dict, reuse: _ReuseSpec, st: _SlideTemplate
) -> None:
    """Convenience: fill the single template table from slide-level headers/rows.

    When a reuse slide contains exactly one cloneable table and the slide spec
    carries the classic `headers`/`rows` payload without addressing that table
    id in reuse.text, the payload is routed to the table automatically.
    """
    if not isinstance(slide_dict, dict):
        return
    tables = [rec for rec in st.shapes if rec.kind == "table" and rec.cloneable]
    if len(tables) != 1:
        return
    tid = tables[0].id
    if tid in reuse.text:
        return
    if reuse.keep_ids is not None and tid not in reuse.keep_ids:
        return
    if tid in reuse.drop_ids:
        return
    headers = slide_dict.get("headers")
    rows = _first(slide_dict, "rows", "data", "table", default=None)
    if headers is None and not rows:
        return
    try:
        reuse.text[tid] = _parse_reuse_table(
            {"headers": headers, "rows": rows or []}, "slides[].headers/rows"
        )
        log.info(
            "[reuse] slide %s: routed slide-level headers/rows to template table id %s",
            st.index,
            tid,
        )
    except ValueError as exc:
        log.warning("[reuse] slide-level table payload ignored: %s", exc)


def _render_reuse_slide(
    deck: "_Deck", reuse: _ReuseSpec, slide_dict: Optional[dict] = None
) -> None:
    if deck.template_pack is None or deck.source_prs is None:
        raise RuntimeError("reuse slide requires template_pack and source_prs")
    pack = deck.template_pack
    st = pack.slides[reuse.slide_index]
    if slide_dict is not None:
        _reuse_table_from_slide_spec(slide_dict, reuse, st)
    decs = _decorations_for_reuse(st, reuse.keep_ids, reuse.drop_ids)
    t_clone = time.monotonic()
    target_slide = _clone_template_slide_to_prs(
        deck.prs,
        deck.source_prs,
        pack,
        source_slide_index=reuse.slide_index,
        decorations=decs,
        part_cache=deck.media_part_cache,
    )
    deck.template_clone_ms += (time.monotonic() - t_clone) * 1000.0
    source_slide = deck.source_prs.slides[reuse.slide_index]
    if reuse.text:
        _apply_reuse_text_map(target_slide, reuse.text, st)
    # Text/table refill reintroduces schemeClr in a:rPr; materialize again.
    _finalize_cloned_slide(target_slide, source_slide, pack.theme)


# --- Template mode: deck frame + mapping (§5.1 dual path) --------------------

def _default_frame() -> _BBox:
    return _BBox(
        x=MARGIN,
        y=CONTENT_TOP,
        w=SLIDE_W_IN - 2 * MARGIN,
        h=FOOTER_Y - CONTENT_TOP - 0.35,
    )


def _layout_mapping_role(layout: str) -> str:
    if layout == "cover":
        return "cover"
    if layout == "section":
        return "section"
    if layout == "closing":
        return "closing"
    return "content"


def _pick_template_slide(
    pack: _TemplatePack, layout: str, mapping: Optional[dict]
) -> int:
    mapping = mapping if isinstance(mapping, dict) else {}
    if layout == "default":
        for key in ("default", "content"):
            if key in mapping:
                return int(mapping[key])
        return 0
    # An explicit per-layout key (e.g. "table": 9, "team": 7) wins over the
    # coarse role fallback (cover/section/closing/content).
    if layout in mapping:
        return int(mapping[layout])
    role = _layout_mapping_role(layout)
    if role in mapping:
        return int(mapping[role])
    if "default" in mapping:
        return int(mapping["default"])
    if "content" in mapping:
        return int(mapping["content"])
    return 0


def _validate_template_mapping(pack: _TemplatePack, mapping: Optional[dict]) -> None:
    if not isinstance(mapping, dict):
        return
    n = len(pack.slides)
    for key, val in mapping.items():
        try:
            idx = int(val)
        except (TypeError, ValueError) as exc:
            raise ValueError(
                f"Invalid template_mapping index for {key!r}: {val!r}"
            ) from exc
        if idx < 0 or idx >= n:
            raise ValueError(
                f"Invalid template_mapping index {idx} for {key!r} "
                f"(template has {n} slide(s))"
            )


def _content_bounds(deck, top_y: float) -> tuple[float, float, float, float]:
    if deck.template_mode:
        f = deck.frame
        return (f.x, top_y, f.w, f.y + f.h - top_y)
    return (
        MARGIN,
        top_y,
        SLIDE_W_IN - 2 * MARGIN,
        FOOTER_Y - top_y - 0.35,
    )


def _frame_x_w(deck) -> tuple[float, float]:
    if deck.template_mode:
        return deck.frame.x, deck.frame.w
    return MARGIN, SLIDE_W_IN - 2 * MARGIN


def _frame_bottom_y(deck, margin: float) -> float:
    if deck.template_mode:
        return deck.frame.y + deck.frame.h - margin
    return FOOTER_Y - margin


def _slide_ref_bg_hex(deck, layout_role: str) -> str:
    if not deck.template_mode or deck.template_pack is None:
        return _hex("FFFFFF")
    idx = _pick_template_slide(deck.template_pack, layout_role, deck.template_mapping)
    st = deck.template_pack.slides[idx]
    bg = st.background or {}
    if bg.get("type") == "solid" and bg.get("color"):
        return _hex(bg["color"])
    return _hex(deck.template_pack.theme.get("lt1") or "FFFFFF")


def _template_chrome_dark(deck, *, legacy_dark: bool) -> bool:
    if deck.template_mode:
        return False
    return legacy_dark


def _template_text_colors(deck) -> dict:
    bg = getattr(deck, "frame_bg_hex", None) or _slide_ref_bg_hex(
        deck, getattr(deck, "frame_role", "content")
    )
    return _derive_readable_colors(bg)


def _footer_y_template(deck) -> float:
    f = deck.frame
    foot_y = f.y + f.h - 0.35
    pack = deck.template_pack
    idx = getattr(deck, "frame_slide_index", None)
    if pack is None or idx is None:
        return foot_y
    st = pack.slides[idx]
    band = _BBox(x=f.x, y=foot_y, w=f.w, h=0.35)
    for dec in st.decorations:
        if _rects_intersect(band, dec.bbox):
            foot_y = max(f.y + 0.05, foot_y - 0.2)
            log.debug("[template] footer band lifted for decoration overlap")
            break
    return foot_y


def _content_layout_band(
    deck,
    top: float,
    *,
    y_floor_legacy: float = 1.95,
    bottom_margin: float = 0.25,
) -> tuple[float, float, float, float]:
    """Return fx, total_w, y0, avail_h for rows below slide chrome."""
    if deck.template_mode:
        fx, total_w = _frame_x_w(deck)
        y0 = max(top + 0.25, deck.frame.y + 0.15)
        avail = _frame_bottom_y(deck, bottom_margin) - y0
        return fx, total_w, y0, avail
    total_w = SLIDE_W_IN - 2 * MARGIN
    y0 = max(top + 0.25, y_floor_legacy)
    avail = FOOTER_Y - y0 - bottom_margin
    return MARGIN, total_w, y0, avail


def _diagram_region(deck, top: float) -> tuple[float, float, float, float]:
    """Bounding box x, y, w, h for native diagram renderers."""
    if deck.template_mode:
        y = top + 0.25
        x, yb, w, h = _content_bounds(deck, y)
        return x, yb, w, h
    y = max(top + 0.3, 2.1)
    avail = FOOTER_Y - y - 0.3
    return MARGIN, y, SLIDE_W_IN - 2 * MARGIN, avail


def _derive_readable_colors(bg_hex: str) -> dict:
    bg = _hex(bg_hex)
    ink = _on(bg)
    muted = _mix(ink, bg, 0.45)
    on_dark = _on(_hex("1A1A1A") if _luminance(bg) > 0.42 else bg)
    return {
        "ink": ink,
        "muted": muted,
        "on_dark": on_dark,
        "on_dark_soft": _mix(on_dark, bg if _luminance(bg) > 0.42 else "000000", 0.72),
        "on_dark_faint": _mix(on_dark, bg if _luminance(bg) > 0.42 else "000000", 0.50),
    }


def _content_slide_bg_hex(pack: _TemplatePack, mapping: Optional[dict]) -> str:
    idx = _pick_template_slide(pack, "content", mapping)
    st = pack.slides[idx]
    bg = st.background or {}
    if bg.get("type") == "solid" and bg.get("color"):
        return _hex(bg["color"])
    return _hex(pack.theme.get("lt1") or "FFFFFF")


def _merge_theme_from_template(
    resolved: dict,
    pack: _TemplatePack,
    *,
    content_bg_hex: Optional[str] = None,
    palette_override: Optional[dict] = None,
) -> dict:
    theme = dict(resolved)
    txml = pack.theme or {}
    if txml.get("major_font"):
        theme["head_font"] = txml["major_font"]
    if txml.get("minor_font"):
        theme["body_font"] = txml["minor_font"]
    if txml.get("accent1"):
        theme["accent"] = _hex(txml["accent1"])
        theme["primary"] = _hex(txml["accent1"])
    bg = content_bg_hex or _hex(txml.get("lt1") or "FFFFFF")
    readable = _derive_readable_colors(bg)
    theme["ink"] = readable["ink"]
    theme["muted"] = readable["muted"]
    if isinstance(palette_override, dict):
        for k in ("ink", "muted", "primary", "accent"):
            if palette_override.get(k):
                theme[k] = _hex(palette_override[k])
    return theme


# ============================================================================
# Deck builder + renderers
# ============================================================================

class _Deck:
    def __init__(
        self,
        prs,
        theme,
        footer_label,
        image_resolver=None,
        *,
        template_pack: Optional[_TemplatePack] = None,
        source_prs=None,
        template_mapping: Optional[dict] = None,
    ):
        self.prs = prs
        self.theme = theme
        self.footer_label = footer_label
        self.image_resolver = image_resolver
        self.total = 0
        self.template_pack = template_pack
        self.template_mode = template_pack is not None
        self.source_prs = source_prs
        self.template_mapping = template_mapping if isinstance(template_mapping, dict) else {}
        self.frame = _default_frame()
        self.frame_role = "default"
        self.frame_bg_hex: Optional[str] = None
        self.frame_slide_index: Optional[int] = None
        self.template_clone_ms: float = 0.0
        # source media partname -> Part copied into self.prs (shared by all slides)
        self.media_part_cache: dict[str, Any] = {}

    def blank(self, layout: str = "default"):
        if not self.template_mode:
            return self.prs.slides.add_slide(self.prs.slide_layouts[6])
        if self.template_pack is None or self.source_prs is None:
            raise RuntimeError("template_mode requires template_pack and source_prs")
        idx = _pick_template_slide(self.template_pack, layout, self.template_mapping)
        st = self.template_pack.slides[idx]
        t_clone = time.monotonic()
        slide = _clone_template_slide_to_prs(
            self.prs,
            self.source_prs,
            self.template_pack,
            source_slide_index=idx,
            decorations=st.decorations,
            part_cache=self.media_part_cache,
        )
        self.template_clone_ms += (time.monotonic() - t_clone) * 1000.0
        self.frame_role = layout
        self.frame_slide_index = idx
        bg = st.background or {}
        if bg.get("type") == "solid" and bg.get("color"):
            self.frame_bg_hex = _hex(bg["color"])
        else:
            self.frame_bg_hex = _hex(self.template_pack.theme.get("lt1") or "FFFFFF")
        if st.safe_zone:
            self.frame = st.safe_zone
        else:
            log.warning(
                "[template] slide reference %s has no safe_zone; using default frame",
                idx,
            )
            self.frame = _default_frame()
        return slide

    # -- chrome helpers --------------------------------------------------
    def _content_head(self, slide, slide_dict, page, *, dark=False):
        eb = _first(slide_dict, "eyebrow", "kicker", "category", "section")
        title = _first(slide_dict, "title", "heading", default="")
        if self.template_mode:
            tc = _template_text_colors(self)
            chrome_dark = _template_chrome_dark(self, legacy_dark=dark)
            fx, fw = self.frame.x, self.frame.w
            fy = self.frame.y
            y_title = max(fy + 0.08, fy + (0.1 if eb else 0.05))
            if eb:
                _eyebrow(
                    slide, self.theme, eb, x=fx, y=fy, w=fw, dark=chrome_dark
                )
            if title:
                _title(
                    slide,
                    self.theme,
                    title,
                    x=fx,
                    y=y_title,
                    w=fw,
                    dark=chrome_dark,
                    color=tc["ink"],
                )
            foot_y = _footer_y_template(self)
            _footer(
                slide,
                self.theme,
                page=page,
                label=self.footer_label,
                dark=chrome_dark,
                y=foot_y,
                x=fx,
                w=fw,
            )
            bottom = y_title + (0.95 if len(_strip_md(str(title))) < 46 else 1.35)
            return max(bottom, self.frame.y + 0.05) if title else max(fy, self.frame.y + 0.05)
        y_title = 0.98 if eb else 0.8
        if eb:
            _eyebrow(slide, self.theme, eb, dark=dark)
        if title:
            _title(slide, self.theme, title, y=y_title, dark=dark)
        _footer(slide, self.theme, page=page, label=self.footer_label, dark=dark)
        return (y_title + (0.95 if len(_strip_md(str(title))) < 46 else 1.35)) if title else CONTENT_TOP


# ---- Cover -----------------------------------------------------------------

def _r_cover(deck, slide_dict, page):
    t = deck.theme
    slide = deck.blank("cover")
    _maybe_set_bg(deck, slide, t["bg_dark"])
    if deck.template_mode:
        tc = _template_text_colors(deck)
        fx, fw = _frame_x_w(deck)
        fy, fh = deck.frame.y, deck.frame.h
        tw = fw * 0.92
        _draw_icon(
            slide,
            t,
            _first(slide_dict, "icon", default="sparkles") or "sparkles",
            fx,
            fy + fh * 0.02,
            min(0.85, fh * 0.12),
            circle_fill=t["accent_soft"],
            glyph_color=t["accent"],
        )
        eb = _first(slide_dict, "eyebrow", "kicker", "category", default="")
        y_cur = fy + fh * 0.14
        if eb:
            _eyebrow(slide, t, eb, x=fx, y=y_cur, w=tw, dark=False)
            y_cur += 0.38
        title = _first(slide_dict, "title", "heading", default="")
        if title:
            tb, tf = _textbox(slide, fx, y_cur, tw, fh * 0.28, anchor="top")
            _add_para(
                tf,
                _strip_md(str(title)),
                first=True,
                size=min(42, max(28, int(36 * fh / 5.5))),
                color=tc["ink"],
                bold=True,
                font=t["head_font"],
                align="left",
                line=1.02,
                space_after=0,
            )
            y_cur += fh * 0.28
        sub = _first(slide_dict, "subtitle", "lead", "description", "summary", default="")
        if sub:
            tb2, tf2 = _textbox(slide, fx, y_cur, tw, fh * 0.14, anchor="top")
            _add_para(
                tf2,
                _strip_md(str(sub)),
                first=True,
                size=15,
                color=tc["muted"],
                font=t["body_font"],
                align="left",
                line=1.3,
                space_after=0,
            )
        chips = _as_list(_first(slide_dict, "chips", "tags", "authors", default=[]))
        author = _first(slide_dict, "author", default="")
        if author and not chips:
            chips = [c.strip() for c in re.split(r"[,/|]", str(author)) if c.strip()]
        chip_y = _footer_y_template(deck) - 0.55
        cx = fx
        for chip in chips[:4]:
            txt = _strip_md(str(chip))
            cw = min(3.2, 0.62 + len(txt) * 0.085)
            if cx + cw > fx + fw:
                break
            _card(
                slide,
                cx,
                chip_y,
                cw,
                0.5,
                fill=t["card_soft"],
                radius=0.5,
                shadow=False,
                line=t["card_border"],
                line_w=1.0,
            )
            tbc, tfc = _textbox(slide, cx, chip_y, cw, 0.5, anchor="middle")
            _add_para(
                tfc,
                txt,
                first=True,
                size=11,
                color=tc["muted"],
                font=t["body_font"],
                align="center",
                space_after=0,
            )
            cx += cw + 0.18
        date = _first(slide_dict, "date", "footer", default=deck.footer_label)
        if date:
            tbd, tfd = _textbox(slide, fx, chip_y + 0.55, tw, 0.35, anchor="middle")
            _add_para(
                tfd,
                _strip_md(str(date)),
                first=True,
                size=9.5,
                color=tc["on_dark_faint"],
                font=t["body_font"],
                align="left",
                space_after=0,
            )
        _footer(
            slide,
            t,
            page=page,
            label=deck.footer_label,
            dark=False,
            y=_footer_y_template(deck),
            x=fx,
            w=fw,
        )
        return
    # layered decorative circles (right side)
    _oval(slide, SLIDE_W_IN - 3.2, -2.0, 5.6, fill=_lighten(t["bg_dark"], 0.08))
    _oval(slide, SLIDE_W_IN - 1.1, 4.2, 3.4, fill=_lighten(t["bg_dark"], 0.05))
    _oval(slide, SLIDE_W_IN - 4.4, 5.2, 1.8, fill=t["accent"], alpha=22)
    # brand mark (icon circle) top-left
    _draw_icon(slide, t, _first(slide_dict, "icon", default="sparkles") or "sparkles",
               MARGIN, 0.7, 0.95, circle_fill=_lighten(t["bg_dark"], 0.12),
               glyph_color=t["accent"])
    eb = _first(slide_dict, "eyebrow", "kicker", "category", default="")
    if eb:
        _eyebrow(slide, t, eb, x=MARGIN, y=1.95, dark=True)
    title = _first(slide_dict, "title", "heading", default="")
    tb, tf = _textbox(slide, MARGIN, 2.35, 8.7, 2.3, anchor="top")
    _add_para(tf, _strip_md(str(title)), first=True, size=42, color=t["on_dark"],
              bold=True, font=t["head_font"], align="left", line=1.02, space_after=0)
    sub = _first(slide_dict, "subtitle", "lead", "description", "summary", default="")
    if sub:
        tb2, tf2 = _textbox(slide, MARGIN, 4.75, 8.4, 1.2, anchor="top")
        _add_para(tf2, _strip_md(str(sub)), first=True, size=15,
                  color=t["on_dark_soft"], font=t["body_font"], align="left",
                  line=1.3, space_after=0)
    # chips row (author / tags)
    chips = _as_list(_first(slide_dict, "chips", "tags", "authors", default=[]))
    author = _first(slide_dict, "author", default="")
    if author and not chips:
        chips = [c.strip() for c in re.split(r"[,/|]", str(author)) if c.strip()]
    cx = MARGIN
    for chip in chips[:4]:
        txt = _strip_md(str(chip))
        w = min(3.2, 0.62 + len(txt) * 0.085)
        _card(slide, cx, 6.05, w, 0.5, fill=_lighten(t["bg_dark"], 0.12),
              radius=0.5, shadow=False, line=_lighten(t["bg_dark"], 0.22), line_w=1.0)
        tbc, tfc = _textbox(slide, cx, 6.05, w, 0.5, anchor="middle")
        _add_para(tfc, txt, first=True, size=11, color=t["on_dark_soft"],
                  font=t["body_font"], align="center", space_after=0)
        cx += w + 0.18
    date = _first(slide_dict, "date", "footer", default=deck.footer_label)
    if date:
        tbd, tfd = _textbox(slide, MARGIN, 6.95, 6.0, 0.35, anchor="middle")
        _add_para(tfd, _strip_md(str(date)), first=True, size=9.5,
                  color=t["on_dark_faint"], font=t["body_font"], align="left",
                  space_after=0)


# ---- Section divider -------------------------------------------------------

def _r_section(deck, slide_dict, page):
    t = deck.theme
    slide = deck.blank("section")
    _maybe_set_bg(deck, slide, t["bg_dark"])
    if deck.template_mode:
        tc = _template_text_colors(deck)
        fx, fw = _frame_x_w(deck)
        fy, fh = deck.frame.y, deck.frame.h
        num = _first(slide_dict, "number", "chapter", "index", default="")
        y_cur = fy + fh * 0.08
        num_size = 68 if fh >= 4.5 else 58
        if num:
            tbn, tfn = _textbox(slide, fx, y_cur, fw * 0.4, 1.2, anchor="top")
            _add_para(
                tfn,
                str(num),
                first=True,
                size=num_size,
                color=t["accent"],
                bold=True,
                font=t["head_font"],
                align="left",
                space_after=0,
            )
            y_cur += 1.25
        eb = _first(slide_dict, "eyebrow", "kicker", "category", default="")
        if eb:
            _eyebrow(slide, t, eb, x=fx, y=y_cur, w=fw, dark=False)
            y_cur += 0.38
        title = _first(slide_dict, "title", "heading", default="")
        if title:
            tb, tf = _textbox(slide, fx, y_cur, fw, min(1.6, fh * 0.22), anchor="top")
            _add_para(
                tf,
                _strip_md(str(title)),
                first=True,
                size=min(38, max(28, int(34 * fh / 5.5))),
                color=tc["ink"],
                bold=True,
                font=t["head_font"],
                align="left",
                line=1.03,
                space_after=0,
            )
            y_cur += min(1.6, fh * 0.22)
        lead = _first(slide_dict, "lead", "subtitle", "description", "summary", default="")
        if lead:
            tbl, tfl = _textbox(slide, fx, y_cur, fw, fh - (y_cur - fy) - 0.5, anchor="top")
            _add_para(
                tfl,
                _strip_md(str(lead)),
                first=True,
                size=14,
                color=tc["muted"],
                font=t["body_font"],
                line=1.3,
                space_after=0,
            )
        _footer(
            slide,
            t,
            page=page,
            label=deck.footer_label,
            dark=False,
            y=_footer_y_template(deck),
            x=fx,
            w=fw,
        )
        return
    _oval(slide, SLIDE_W_IN - 2.6, -1.6, 4.6, fill=_lighten(t["bg_dark"], 0.07))
    num = _first(slide_dict, "number", "chapter", "index", default="")
    if num:
        tbn, tfn = _textbox(slide, MARGIN, 1.7, 3.0, 1.6, anchor="top")
        _add_para(tfn, str(num), first=True, size=76, color=t["accent"],
                  bold=True, font=t["head_font"], align="left", space_after=0)
    eb = _first(slide_dict, "eyebrow", "kicker", "category", default="")
    if eb:
        _eyebrow(slide, t, eb, x=MARGIN, y=(3.4 if num else 2.6), dark=True)
    title = _first(slide_dict, "title", "heading", default="")
    tb, tf = _textbox(slide, MARGIN, (3.75 if num else 3.0), 10.5, 1.6, anchor="top")
    _add_para(tf, _strip_md(str(title)), first=True, size=38, color=t["on_dark"],
              bold=True, font=t["head_font"], align="left", line=1.03, space_after=0)
    lead = _first(slide_dict, "lead", "subtitle", "description", "summary", default="")
    if lead:
        tbl, tfl = _textbox(slide, MARGIN, (5.35 if num else 4.6), 9.5, 1.4)
        _add_para(tfl, _strip_md(str(lead)), first=True, size=14,
                  color=t["on_dark_soft"], font=t["body_font"], line=1.3, space_after=0)
    _footer(slide, t, page=page, label=deck.footer_label, dark=True)


# ---- Title only / title + body --------------------------------------------

def _r_title_body(deck, slide_dict, page, *, blank_layout: str = "content"):
    t = deck.theme
    slide = deck.blank(blank_layout)
    _maybe_set_bg(deck, slide, t["bg_light"])
    top = deck._content_head(slide, slide_dict, page)
    body = _first(slide_dict, "body", "content", "text", "description", "summary",
                  "paragraph", "lead", default="")
    bullets = _harvest_bullets(slide_dict)
    y = top + 0.15
    if deck.template_mode:
        x, yb, w, h = _content_bounds(deck, y)
        body_h, bullet_h = h, h
    else:
        x, yb, w = MARGIN, y, SLIDE_W_IN - 2 * MARGIN
        body_h, bullet_h = 4.4, 4.6
    if body and isinstance(body, str) and not bullets:
        tb, tf = _textbox(slide, x, yb, w, body_h)
        for i, para in enumerate([p for p in str(body).split("\n") if p.strip()]):
            _add_para(tf, _md_runs(para), first=(i == 0), size=15, color=t["muted"],
                      font=t["body_font"], line=1.36, space_after=10, align="left")
    elif bullets:
        _bullet_block(slide, t, bullets, x, yb, w, bullet_h)


def _r_title_only(deck, slide_dict, page):
    t = deck.theme
    slide = deck.blank("content")
    _maybe_set_bg(deck, slide, t["bg_light"])
    eb = _first(slide_dict, "eyebrow", "kicker", default="")
    if eb:
        _eyebrow(slide, t, eb, y=2.6)
    title = _first(slide_dict, "title", "heading", default="")
    tb, tf = _textbox(slide, MARGIN, 2.95, SLIDE_W_IN - 2 * MARGIN, 1.6, anchor="top")
    _add_para(tf, _strip_md(str(title)), first=True, size=40, color=t["ink"],
              bold=True, font=t["head_font"], line=1.03, space_after=0)
    sub = _first(slide_dict, "subtitle", "lead", "body", default="")
    if sub:
        tbs, tfs = _textbox(slide, MARGIN, 4.55, 10.5, 1.4)
        _add_para(tfs, _strip_md(str(sub)), first=True, size=15, color=t["muted"],
                  font=t["body_font"], line=1.35, space_after=0)
    _footer(slide, t, page=page, label=deck.footer_label)


def _bullet_block(slide, theme, bullets, x, y, w, h, *, size=14, anchor="top",
                  fill=False):
    tb, tf = _textbox(slide, x, y, w, h, anchor=anchor)
    items = [b for b in bullets if str(b).strip()]
    n = len(items)
    if fill and n:
        # distribute so the block breathes and fills the available height
        line_pt = size * 1.35 * max(1, 1)
        avail_pt = h * 72
        leftover = avail_pt - n * line_pt
        sp = int(max(8, min(30, leftover / max(1, n))))
    else:
        sp = 9 if n <= 5 else (6 if n <= 8 else 3)
    first = True
    for b in items:
        txt = _clean_bullet(b) if not str(b).strip().startswith("**") else _strip_md(b)
        if not txt:
            continue
        _add_bullet(tf, txt, first=first, size=size, color=theme["muted"],
                    accent=theme["accent"], font=theme["body_font"], space_after=sp,
                    line=1.2)
        first = False


def _r_title_bullets(deck, slide_dict, page):
    t = deck.theme
    slide = deck.blank("content")
    _maybe_set_bg(deck, slide, t["bg_light"])
    top = deck._content_head(slide, slide_dict, page)
    bullets = _harvest_bullets(slide_dict)
    if not bullets:
        body = _first(slide_dict, "body", "content", "text", default="")
        bullets = [ln for ln in str(body).split("\n") if ln.strip()] if body else []
    n = len([b for b in bullets if str(b).strip()])
    # sparse lists get larger type and breathe to fill the canvas
    size = 18 if n <= 4 else (16 if n <= 6 else 14)
    y = top + 0.3
    if deck.template_mode:
        x, yb, w, h = _content_bounds(deck, y)
    else:
        x, yb, w = MARGIN, y, SLIDE_W_IN - 2 * MARGIN
        h = FOOTER_Y - y - 0.35
    _bullet_block(slide, t, bullets, x, yb, w, h,
                  size=size, anchor=("middle" if n <= 6 else "top"), fill=(n <= 8))


# ---- Two column / comparison ----------------------------------------------

def _card_column(slide, theme, x, y, w, h, card, *, highlight=False, badge=None):
    t = theme
    fill = t["primary"] if highlight else t["card_soft"]
    heading_color = t["on_dark"] if highlight else t["ink"]
    body_color = t["on_dark_soft"] if highlight else t["muted"]
    _card(slide, x, y, w, h, fill=fill, radius=0.06,
          line=None if highlight else t["card_border"], line_w=1.0,
          shadow=highlight)
    if badge:
        bw = 1.7
        _card(slide, x + (w - bw) / 2, y - 0.18, bw, 0.42, fill=t["accent"],
              radius=0.5, shadow=False)
        tbb, tfb = _textbox(slide, x + (w - bw) / 2, y - 0.18, bw, 0.42, anchor="middle")
        _add_para(tfb, _strip_md(str(badge)).upper(), first=True, size=9,
                  color=_on(t["accent"]), bold=True, font=t["body_font"],
                  align="center", tracking=1.2, space_after=0)
    pad = 0.36
    cy = y + pad
    heading = _first(card, "heading", "title", "label", "name", default="")
    icon = _first(card, "icon", default="")
    if icon:
        idd = 0.66
        _draw_icon(slide, t, icon, x + pad, cy, idd,
                   circle_fill=(_lighten(t["primary"], 0.14) if highlight else t["accent_soft"]),
                   glyph_color=(t["light_accent"] if highlight else t["accent"]))
        if heading:
            tbh, tfh = _textbox(slide, x + pad + idd + 0.22, cy, w - 2 * pad - idd - 0.22,
                                idd, anchor="middle")
            _add_para(tfh, _strip_md(str(heading)), first=True, size=19,
                      color=heading_color, bold=True, font=t["head_font"], space_after=0)
        cy += idd + 0.22
    elif heading:
        tbh, tfh = _textbox(slide, x + pad, cy, w - 2 * pad, 0.6)
        _add_para(tfh, _strip_md(str(heading)), first=True, size=19,
                  color=heading_color, bold=True, font=t["head_font"], space_after=0)
        cy += 0.62
    sub = _first(card, "subtitle", "value", "price", "tagline", default="")
    if sub:
        tbs, tfs = _textbox(slide, x + pad, cy, w - 2 * pad, 0.5)
        _add_para(tfs, _strip_md(str(sub)), first=True, size=15,
                  color=(t["accent"] if not highlight else t["light_accent"]),
                  bold=True, font=t["body_font"], space_after=0)
        cy += 0.5
    desc = _first(card, "description", "detail", "note", default="")
    if desc:
        tbd, tfd = _textbox(slide, x + pad, cy, w - 2 * pad, 0.6)
        _add_para(tfd, _strip_md(str(desc)), first=True, size=11.5,
                  color=body_color, font=t["body_font"], line=1.25, space_after=0)
        cy += 0.5
    points = card.get("points") or card.get("bullets") or card.get("items") or []
    pts = _harvest_bullets({"points": points}) if points else []
    if pts:
        tbp, tfp = _textbox(slide, x + pad, cy + 0.1, w - 2 * pad, h - (cy - y) - pad)
        first = True
        for p in pts:
            _add_bullet(tfp, _clean_bullet(p), first=first, size=12,
                        color=body_color, accent=(t["light_accent"] if highlight else t["accent"]),
                        font=t["body_font"], space_after=6)
            first = False


def _split_cards(slide_dict):
    """Return a list of card dicts for two_column/comparison."""
    cards = []
    for side in ("left", "right"):
        c = slide_dict.get(side)
        if isinstance(c, dict):
            cards.append(c)
    if cards:
        return cards
    # columns/cards/items array
    for key in ("columns", "cards", "options", "tiers", "plans", "items"):
        arr = slide_dict.get(key)
        if isinstance(arr, list) and arr and isinstance(arr[0], dict):
            return arr
    # flat left_*/right_*
    flat = {}
    for side in ("left", "right"):
        d = {}
        for f in ("heading", "title", "points", "bullets", "subtitle", "description"):
            v = slide_dict.get(f"{side}_{f}")
            if v is not None:
                d["heading" if f in ("heading", "title") else f] = v
        if d:
            flat[side] = d
    if flat:
        return list(flat.values())
    # fallback: split bullets in half
    bl = _harvest_bullets(slide_dict)
    if bl:
        mid = (len(bl) + 1) // 2
        return [{"heading": "", "points": bl[:mid]}, {"heading": "", "points": bl[mid:]}]
    return []


def _r_comparison(deck, slide_dict, page):
    t = deck.theme
    slide = deck.blank("content")
    _maybe_set_bg(deck, slide, t["bg_light"])
    top = deck._content_head(slide, slide_dict, page)
    cards = _split_cards(slide_dict)
    if not cards:
        return
    n = min(len(cards), 3)
    cards = cards[:n]
    gap = 0.4
    has_points = any(
        (c.get("points") or c.get("bullets") or c.get("items"))
        for c in cards if isinstance(c, dict)
    )
    if deck.template_mode:
        fx, total_w, y0, avail = _content_layout_band(deck, top)
        cw = (total_w - gap * (n - 1)) / n
        if has_points:
            h, y = avail, y0
        else:
            h = min(avail, 3.0)
            y = y0 + max(0.25, (avail - h) / 2)
    else:
        fx = MARGIN
        total_w = SLIDE_W_IN - 2 * MARGIN
        cw = (total_w - gap * (n - 1)) / n
        y = max(top + 0.25, 1.95)
        if has_points:
            h = FOOTER_Y - y - 0.25
        else:
            h = min(FOOTER_Y - y - 0.25, 3.0)
            y = top + max(0.25, (FOOTER_Y - top - h) / 2)
    for i, card in enumerate(cards):
        x = fx + i * (cw + gap)
        hl = bool(card.get("highlight") or card.get("featured") or card.get("recommended"))
        badge = card.get("badge")  # only explicit badges (never auto)
        _card_column(slide, t, x, y, cw, h, card, highlight=hl, badge=badge)


# ---- KPI row ---------------------------------------------------------------

def _r_kpi(deck, slide_dict, page):
    t = deck.theme
    slide = deck.blank("content")
    _maybe_set_bg(deck, slide, t["bg_light"])
    top = deck._content_head(slide, slide_dict, page)
    stats = _first(slide_dict, "stats", "kpis", "metrics", "numbers", "items",
                   "cards", default=[])
    stats = [s for s in _as_list(stats) if isinstance(s, dict)] or _as_list(stats)
    if not stats:
        return
    n = min(len(stats), 4)
    stats = stats[:n]
    gap = 0.4
    fx, total_w, y0, avail = _content_layout_band(deck, top, y_floor_legacy=2.2)
    cw = (total_w - gap * (n - 1)) / n
    if deck.template_mode:
        f = deck.frame
        foot = _footer_y_template(deck)
        y = min(max(top, f.y + 0.08), foot - 1.0)
        h = min(2.7, max(0.85, foot - 0.2 - y))
    else:
        y = max(top + 0.35, 2.2)
        h = 2.7
    for i, s in enumerate(stats):
        x = fx + i * (cw + gap)
        _card(slide, x, y, cw, h, fill=t["card_soft"], radius=0.06,
              line=t["card_border"], line_w=1.0, shadow=False)
        if isinstance(s, dict):
            value = _first(s, "value", "number", "stat", "amount", default="")
            label = _first(s, "label", "name", "title", "caption", default="")
            change = _first(s, "change", "delta", "trend", default="")
        else:
            value, label, change = str(s), "", ""
        pad = 0.3
        vstr = _strip_md(str(value))
        # scale the big number down when long so it stays on one line
        vsize = 40 if len(vstr) <= 5 else (33 if len(vstr) <= 7 else 26)
        tbv, tfv = _textbox(slide, x + pad, y + 0.45, cw - 2 * pad, 1.0,
                            anchor="middle", wrap=False)
        _add_para(tfv, vstr, first=True, size=vsize, color=t["accent"],
                  bold=True, font=t["head_font"], align="left", space_after=0)
        tbl, tfl = _textbox(slide, x + pad, y + 1.55, cw - 2 * pad, 0.9)
        _add_para(tfl, _strip_md(str(label)), first=True, size=12.5, color=t["muted"],
                  font=t["body_font"], align="left", line=1.2, space_after=0)
        if change:
            tbc, tfc = _textbox(slide, x + pad, y + h - 0.5, cw - 2 * pad, 0.35)
            up = not str(change).strip().startswith("-")
            _add_para(tfc, ("\u25B2 " if up else "\u25BC ") + _strip_md(str(change)),
                      first=True, size=11, color=(t["accent"] if up else "C0392B"),
                      bold=True, font=t["body_font"], space_after=0)


# ---- Timeline / process ----------------------------------------------------

def _r_timeline(deck, slide_dict, page, *, process=False):
    t = deck.theme
    slide = deck.blank("content")
    _maybe_set_bg(deck, slide, t["bg_light"])
    top = deck._content_head(slide, slide_dict, page)
    steps = _first(slide_dict, "steps", "phases", "milestones", "stages", "items",
                   "events", default=[])
    steps = _as_list(steps)
    if not steps:
        return
    n = min(len(steps), 5)
    steps = steps[:n]
    gap = 0.35
    has_desc = any(isinstance(s, dict) and _first(s, "description", "detail", "text",
                   "body", default="") for s in steps)
    h = 2.5 if has_desc else 1.7
    if deck.template_mode:
        fx, total_w, y0, avail = _content_layout_band(deck, top)
        cw = (total_w - gap * (n - 1)) / n
        y = y0 + max(0.35, (avail - h) / 2)
    else:
        fx = MARGIN
        total_w = SLIDE_W_IN - 2 * MARGIN
        cw = (total_w - gap * (n - 1)) / n
        y = top + max(0.6, (FOOTER_Y - top - h) / 2)
    _rect(slide, fx + cw / 2, y - 0.32, total_w - cw, 0.03, fill=t["card_border"])
    for i, s in enumerate(steps):
        x = fx + i * (cw + gap)
        # node circle with number
        nd = 0.5
        _oval(slide, x + cw / 2 - nd / 2, y - 0.32 - nd / 2, nd, fill=t["accent"])
        tbn, tfn = _textbox(slide, x + cw / 2 - nd / 2, y - 0.32 - nd / 2, nd, nd, anchor="middle")
        _add_para(tfn, str(i + 1), first=True, size=15, color=_on(t["accent"]),
                  bold=True, font=t["body_font"], align="center", space_after=0)
        _card(slide, x, y, cw, h, fill=t["card_soft"], radius=0.06,
              line=t["card_border"], line_w=1.0, shadow=False)
        if isinstance(s, dict):
            label = _first(s, "title", "label", "heading", "name", "phase", default=f"Step {i+1}")
            when = _first(s, "date", "when", "time", "period", default="")
            desc = _first(s, "description", "detail", "text", "body", default="")
        else:
            label, when, desc = str(s), "", ""
        pad = 0.26
        cy = y + pad
        if when:
            tbw, tfw = _textbox(slide, x + pad, cy, cw - 2 * pad, 0.3)
            _add_para(tfw, _strip_md(str(when)).upper(), first=True, size=9,
                      color=t["accent"], bold=True, font=t["body_font"],
                      tracking=1.0, space_after=0)
            cy += 0.34
        tbt, tft = _textbox(slide, x + pad, cy, cw - 2 * pad, 0.6)
        _add_para(tft, _strip_md(str(label)), first=True, size=14, color=t["ink"],
                  bold=True, font=t["head_font"], line=1.05, space_after=0)
        cy += 0.6
        if desc:
            tbd, tfd = _textbox(slide, x + pad, cy, cw - 2 * pad, h - (cy - y) - pad)
            _add_para(tfd, _strip_md(str(desc)), first=True, size=11, color=t["muted"],
                      font=t["body_font"], line=1.22, space_after=0)


# ---- Quote -----------------------------------------------------------------

def _r_quote(deck, slide_dict, page):
    t = deck.theme
    slide = deck.blank("content")
    _maybe_set_bg(deck, slide, t["bg_dark"])
    if deck.template_mode:
        tc = _template_text_colors(deck)
        fx, fw = _frame_x_w(deck)
        fy, fh = deck.frame.y, deck.frame.h
        tbq, tfq = _textbox(slide, fx, fy + 0.05, min(2.0, fw * 0.2), fh * 0.18)
        _add_para(
            tfq,
            "\u201C",
            first=True,
            size=min(110, int(90 * fh / 5.5)),
            color=t["accent"],
            bold=True,
            font=t["head_font"],
            space_after=0,
        )
        quote = _first(slide_dict, "quote", "text", "content", "message", default="")
        tb, tf = _textbox(slide, fx, fy + fh * 0.22, fw * 0.92, fh * 0.42, anchor="top")
        _add_para(
            tf,
            _strip_md(str(quote)),
            first=True,
            size=24,
            color=tc["ink"],
            italic=True,
            font=t["head_font"],
            line=1.22,
            space_after=0,
        )
        author = _first(slide_dict, "author", "by", "source", default="")
        role = _first(slide_dict, "role", "title", "position", default="")
        if author:
            tba, tfa = _textbox(slide, fx, fy + fh * 0.68, fw, 0.8)
            line = f"\u2014 {_strip_md(str(author))}"
            _add_para(
                tfa,
                line,
                first=True,
                size=15,
                color=t["accent"],
                bold=True,
                font=t["body_font"],
                space_after=2,
            )
            if role:
                _add_para(
                    tfa,
                    _strip_md(str(role)),
                    size=12,
                    color=tc["on_dark_faint"],
                    font=t["body_font"],
                    space_after=0,
                )
        _footer(
            slide,
            t,
            page=page,
            label=deck.footer_label,
            dark=False,
            y=_footer_y_template(deck),
            x=fx,
            w=fw,
        )
        return
    _oval(slide, -1.4, -1.4, 3.4, fill=_lighten(t["bg_dark"], 0.06))
    # big quotation mark
    tbq, tfq = _textbox(slide, MARGIN, 0.9, 2.0, 1.6)
    _add_para(tfq, "\u201C", first=True, size=110, color=t["accent"], bold=True,
              font=t["head_font"], space_after=0)
    quote = _first(slide_dict, "quote", "text", "content", "message", default="")
    tb, tf = _textbox(slide, MARGIN, 2.4, SLIDE_W_IN - 2 * MARGIN - 1.5, 3.0, anchor="top")
    _add_para(tf, _strip_md(str(quote)), first=True, size=27, color=t["on_dark"],
              italic=True, font=t["head_font"], line=1.22, space_after=0)
    author = _first(slide_dict, "author", "by", "source", default="")
    role = _first(slide_dict, "role", "title", "position", default="")
    if author:
        tba, tfa = _textbox(slide, MARGIN, 5.7, 9.0, 0.8)
        line = f"\u2014 {_strip_md(str(author))}"
        _add_para(tfa, line, first=True, size=15, color=t["accent"], bold=True,
                  font=t["body_font"], space_after=2)
        if role:
            _add_para(tfa, _strip_md(str(role)), size=12, color=t["on_dark_faint"],
                      font=t["body_font"], space_after=0)
    _footer(slide, t, page=page, label=deck.footer_label, dark=True)


# ---- Closing ---------------------------------------------------------------

def _r_closing(deck, slide_dict, page):
    t = deck.theme
    slide = deck.blank("closing")
    _maybe_set_bg(deck, slide, t["bg_dark"])
    if deck.template_mode:
        tc = _template_text_colors(deck)
        fx, fw = _frame_x_w(deck)
        fy, fh = deck.frame.y, deck.frame.h
        y_cur = fy + fh * 0.06
        eb = _first(slide_dict, "eyebrow", "kicker", default="")
        if eb:
            _eyebrow(slide, t, eb, x=fx, y=y_cur, w=fw, dark=False)
            y_cur += 0.38
        title = _first(slide_dict, "title", "heading", default="Thank you")
        tb, tf = _textbox(slide, fx, y_cur, fw, min(1.4, fh * 0.2), anchor="top")
        _add_para(
            tf,
            _strip_md(str(title)),
            first=True,
            size=min(40, max(28, int(36 * fh / 5.5))),
            color=tc["ink"],
            bold=True,
            font=t["head_font"],
            line=1.03,
            space_after=0,
        )
        y_cur += min(1.4, fh * 0.2)
        takeaways = _first(
            slide_dict, "takeaways", "key_takeaways", "next_steps", "points", default=[]
        )
        takeaways = (
            _harvest_bullets({"points": takeaways})
            if takeaways
            else _harvest_bullets(slide_dict)
        )
        if takeaways:
            tbk, tfk = _textbox(
                slide, fx, y_cur, fw, _footer_y_template(deck) - y_cur - 0.55
            )
            first = True
            for b in takeaways[:5]:
                _add_bullet(
                    tfk,
                    _clean_bullet(b),
                    first=first,
                    size=15,
                    color=tc["muted"],
                    accent=t["accent"],
                    font=t["body_font"],
                    space_after=10,
                )
                first = False
        contact = _first(slide_dict, "contact", "email", "cta", "footer", default="")
        if contact:
            tbc, tfc = _textbox(
                slide, fx, _footer_y_template(deck) - 0.5, fw, 0.45, anchor="middle"
            )
            _add_para(
                tfc,
                _strip_md(str(contact)),
                first=True,
                size=13,
                color=t["accent"],
                bold=True,
                font=t["body_font"],
                space_after=0,
            )
        _footer(
            slide,
            t,
            page=page,
            label=deck.footer_label,
            dark=False,
            y=_footer_y_template(deck),
            x=fx,
            w=fw,
        )
        return
    _oval(slide, SLIDE_W_IN - 3.0, -2.0, 5.2, fill=_lighten(t["bg_dark"], 0.07))
    _oval(slide, SLIDE_W_IN - 1.0, 4.6, 3.0, fill=t["accent"], alpha=18)
    eb = _first(slide_dict, "eyebrow", "kicker", default="")
    if eb:
        _eyebrow(slide, t, eb, y=1.4, dark=True)
    title = _first(slide_dict, "title", "heading", default="Thank you")
    tb, tf = _textbox(slide, MARGIN, 1.85, 9.5, 1.4)
    _add_para(tf, _strip_md(str(title)), first=True, size=40, color=t["on_dark"],
              bold=True, font=t["head_font"], line=1.03, space_after=0)
    takeaways = _first(slide_dict, "takeaways", "key_takeaways", "next_steps",
                       "points", default=[])
    takeaways = _harvest_bullets({"points": takeaways}) if takeaways else _harvest_bullets(slide_dict)
    y = 3.5
    if takeaways:
        tbk, tfk = _textbox(slide, MARGIN, y, 8.5, 2.6)
        first = True
        for b in takeaways[:5]:
            _add_bullet(tfk, _clean_bullet(b), first=first, size=15,
                        color=t["on_dark_soft"], accent=t["accent"],
                        font=t["body_font"], space_after=10)
            first = False
    contact = _first(slide_dict, "contact", "email", "cta", "footer", default="")
    if contact:
        tbc, tfc = _textbox(slide, MARGIN, 6.4, 10.0, 0.5)
        _add_para(tfc, _strip_md(str(contact)), first=True, size=13,
                  color=t["accent"], bold=True, font=t["body_font"], space_after=0)
    _footer(slide, t, page=page, label=deck.footer_label, dark=True)


# ---- Alert / callout -------------------------------------------------------

_ALERT_COLORS = {
    "info": "2563EB", "tip": "059669", "success": "059669", "note": "2563EB",
    "warning": "D97706", "danger": "DC2626", "error": "DC2626",
}


_ALERT_ICONS = {
    "info": "info", "note": "info", "tip": "lightbulb", "success": "check-circle",
    "warning": "alert-triangle", "danger": "alert-octagon", "error": "alert-octagon",
}


def _r_alert(deck, slide_dict, page):
    import math
    t = deck.theme
    slide = deck.blank("content")
    _maybe_set_bg(deck, slide, t["bg_light"])
    top = deck._content_head(slide, slide_dict, page)
    level = str(_first(slide_dict, "level", "severity", "variant", default="info")).lower()
    color = _ALERT_COLORS.get(level, t["accent"])
    body = _first(slide_dict, "body", "content", "text", "message", "description", default="")
    bullets = _harvest_bullets(slide_dict)
    heading = _first(slide_dict, "callout_title", "alert_title", "heading", default="")

    if deck.template_mode:
        x, card_w, y0, band = _content_layout_band(deck, top, y_floor_legacy=2.2)
        y = y0 + 0.1
        avail = band - 0.15
    else:
        y = max(top + 0.35, 2.2)
        avail = FOOTER_Y - y - 0.4
        x = MARGIN
        card_w = SLIDE_W_IN - 2 * MARGIN
    icon_d = 0.95
    text_x = x + 0.55 + icon_d + 0.5
    text_w = x + card_w - 0.6 - text_x

    # Size the card to its content instead of a fixed height, so short
    # callouts don't leave a large empty coloured area.
    def _est_lines(s: str, size: float) -> int:
        cpl = max(12, int(text_w / (size * 0.0072)))
        return max(1, math.ceil(len(str(s)) / cpl))

    lines = 0
    if heading:
        lines += 1
    if body and not bullets:
        lines += _est_lines(body, 16)
    else:
        for b in bullets:
            lines += _est_lines(_clean_bullet(b), 15)
    content_h = lines * 0.34 + (0.4 if heading else 0.0) + 0.8
    h = max(1.5, min(avail, content_h))

    _card(slide, x, y, card_w, h, fill=_lighten(color, 0.9),
          radius=0.05, shadow=False)
    _rect(slide, x, y, 0.1, h, fill=color)
    _draw_icon(slide, t, _ALERT_ICONS.get(level, "info"),
               x + 0.55, y + h / 2 - icon_d / 2, icon_d,
               circle_fill=color, glyph_color=_on(color))

    tb, tf = _textbox(slide, text_x, y, text_w, h, anchor="middle")
    first = True
    if heading:
        _add_para(tf, _strip_md(str(heading)), first=True, size=18, color=t["ink"],
                  bold=True, font=t["head_font"], space_after=6)
        first = False
    if body and not bullets:
        _add_para(tf, _md_runs(str(body)), first=first, size=16, color=t["ink"],
                  font=t["body_font"], line=1.35, space_after=0)
    else:
        for b in bullets:
            _add_bullet(tf, _clean_bullet(b), first=first, size=15, color=t["ink"],
                        accent=color, font=t["body_font"], space_after=8)
            first = False


# ---- Table -----------------------------------------------------------------

def _r_table(deck, slide_dict, page):
    t = deck.theme
    slide = deck.blank("content")
    _maybe_set_bg(deck, slide, t["bg_light"])
    top = deck._content_head(slide, slide_dict, page)
    rows = _first(slide_dict, "rows", "data", "table", default=[])
    headers = _first(slide_dict, "headers", "columns", "header", default=[])
    rows = _as_list(rows)
    # normalize rows of dicts
    if rows and isinstance(rows[0], dict):
        if not headers:
            headers = list(rows[0].keys())
        rows = [[r.get(h, "") for h in headers] for r in rows]
    if not rows:
        return
    ncols = max(len(headers) if headers else 0, max(len(r) for r in rows))
    nrows = len(rows) + (1 if headers else 0)
    if deck.template_mode:
        fx, fw, y0, avail = _content_layout_band(deck, top, y_floor_legacy=2.0)
        y = y0
        h = min(avail - 0.05, 0.5 * nrows + 0.2)
        tbl_x, tbl_w = fx, fw
    else:
        y = max(top + 0.25, 2.0)
        h = min(FOOTER_Y - y - 0.2, 0.5 * nrows + 0.2)
        tbl_x, tbl_w = MARGIN, SLIDE_W_IN - 2 * MARGIN
    tbl_shape = slide.shapes.add_table(nrows, ncols, Inches(tbl_x), Inches(y),
                                       Inches(tbl_w), Inches(h))
    table = tbl_shape.table
    r0 = 0
    if headers:
        for c in range(ncols):
            cell = table.cell(0, c)
            cell.fill.solid()
            cell.fill.fore_color.rgb = _rgb(t["primary"])
            cell.text = _strip_md(str(headers[c])) if c < len(headers) else ""
            para = cell.text_frame.paragraphs[0]
            for run in para.runs:
                run.font.size = Pt(12)
                run.font.bold = True
                run.font.color.rgb = _rgb(_on(t["primary"]))
                run.font.name = t["body_font"]
        r0 = 1
    for ri, row in enumerate(rows):
        for c in range(ncols):
            cell = table.cell(r0 + ri, c)
            cell.fill.solid()
            cell.fill.fore_color.rgb = _rgb(t["card_soft"] if ri % 2 else "FFFFFF")
            cell.text = _strip_md(str(row[c])) if c < len(row) else ""
            para = cell.text_frame.paragraphs[0]
            for run in para.runs:
                run.font.size = Pt(11.5)
                run.font.color.rgb = _rgb(t["ink"])
                run.font.name = t["body_font"]


# ---- Icon list / icon grid -------------------------------------------------

def _icon_items(slide_dict):
    items = _first(slide_dict, "items", "cards", "features", "points", "bullets",
                   "pillars", default=[])
    out = []
    for it in _as_list(items):
        if isinstance(it, dict):
            out.append(it)
        elif isinstance(it, str):
            out.append({"title": it})
    return out


def _r_icon_list(deck, slide_dict, page):
    t = deck.theme
    slide = deck.blank("content")
    _maybe_set_bg(deck, slide, t["bg_light"])
    top = deck._content_head(slide, slide_dict, page)
    items = _icon_items(slide_dict)[:5]
    if not items:
        return
    if deck.template_mode:
        fx, fw, y0, avail = _content_layout_band(deck, top, y_floor_legacy=2.05)
        y, avail = y0, avail - 0.1
    else:
        fx = MARGIN
        fw = SLIDE_W_IN - 2 * MARGIN
        y = max(top + 0.25, 2.05)
        avail = FOOTER_Y - y - 0.1
    rh = min(1.0, avail / max(1, len(items)))
    for i, it in enumerate(items):
        ry = y + i * rh
        d = min(0.72, rh - 0.18)
        _draw_icon(slide, t, _first(it, "icon", default=""), fx, ry, d,
                   circle_fill=t["accent_soft"], glyph_color=t["accent"], idx=i)
        tx = fx + d + 0.35
        title = _first(it, "title", "heading", "label", "name", default="")
        desc = _first(it, "description", "detail", "text", "body", default="")
        tbt, tft = _textbox(slide, tx, ry, fx + fw - tx, rh)
        _add_para(tft, _strip_md(str(title)), first=True, size=16, color=t["ink"],
                  bold=True, font=t["head_font"], space_after=2)
        if desc:
            _add_para(tft, _strip_md(str(desc)), size=12, color=t["muted"],
                      font=t["body_font"], line=1.22, space_after=0)


def _r_icon_grid(deck, slide_dict, page, cols=3):
    t = deck.theme
    slide = deck.blank("content")
    _maybe_set_bg(deck, slide, t["bg_light"])
    top = deck._content_head(slide, slide_dict, page)
    items = _icon_items(slide_dict)
    if not items:
        return
    n = len(items)
    cols = 2 if n <= 4 and cols != 3 else cols
    cols = min(cols, 3)
    rows = (n + cols - 1) // cols
    gap = 0.35
    if deck.template_mode:
        fx, total_w, y0, avail_h = _content_layout_band(deck, top, y_floor_legacy=2.0)
        avail_h -= 0.1
    else:
        fx = MARGIN
        total_w = SLIDE_W_IN - 2 * MARGIN
        y0 = max(top + 0.25, 2.0)
        avail_h = FOOTER_Y - y0 - 0.1
    cw = (total_w - gap * (cols - 1)) / cols
    ch = (avail_h - gap * (rows - 1)) / rows
    for i, it in enumerate(items[:cols * rows]):
        r, c = divmod(i, cols)
        x = fx + c * (cw + gap)
        yy = y0 + r * (ch + gap)
        _card(slide, x, yy, cw, ch, fill=t["card_soft"], radius=0.06,
              line=t["card_border"], line_w=1.0, shadow=False)
        pad = 0.3
        _draw_icon(slide, t, _first(it, "icon", default=""), x + pad, yy + pad, 0.66,
                   circle_fill=t["accent_soft"], glyph_color=t["accent"], idx=i)
        title = _first(it, "title", "heading", "label", "name", default="")
        desc = _first(it, "description", "detail", "text", "body", default="")
        tbt, tft = _textbox(slide, x + pad, yy + pad + 0.82, cw - 2 * pad, ch - 1.2)
        _add_para(tft, _strip_md(str(title)), first=True, size=15, color=t["ink"],
                  bold=True, font=t["head_font"], space_after=3)
        if desc:
            _add_para(tft, _strip_md(str(desc)), size=11.5, color=t["muted"],
                      font=t["body_font"], line=1.24, space_after=0)


# ============================================================================
# The Tools class (OpenWebUI native function calling) Valves
# ============================================================================

# >>> KIT ENGINE START

# ============================================================================
# DECK KIT ENGINE (v2 "blocks" format)
# ----------------------------------------------------------------------------
# A second rendering engine, ported from the Deck Kit renderer (pptxgenjs) to
# python-pptx: 44 consulting-grade components, 14 theme presets, measured
# auto-fit text (shared font size across sibling items), a schema validator
# with precise, model-readable error paths and an LLM-facing catalogue.
#
# A deck uses this engine when its slides carry `rows` / `blocks` / `type`
# (cover | section | content | closing) instead of the classic `layout` field,
# or when the spec sets "engine": "kit".
#
# Icons: Font Awesome Free solid glyphs (CC BY 4.0, https://fontawesome.com/license/free)
# pre-rendered to white PNGs, tinted at render time.
# ============================================================================

try:
    from pptx.util import Emu as _KEmu  # type: ignore
    from pptx.enum.text import MSO_AUTO_SIZE as _K_AUTO  # type: ignore
    from pptx.enum.shapes import MSO_CONNECTOR as _K_CONN  # type: ignore
    from pptx.enum.dml import MSO_LINE_DASH_STYLE as _K_DASH  # type: ignore
    from pptx.oxml.ns import qn as _kqn  # type: ignore
    from lxml import etree as _K_ET  # type: ignore
except Exception:  # pragma: no cover
    pass

import math as _kmath

_K_W, _K_H, _K_M = 13.333, 7.5, 0.45
_K_CW = _K_W - 2 * _K_M
_K_TOP, _K_BOTTOM, _K_ROW_GAP, _K_BLOCK_GAP, _K_HEAD_H = 1.5, 6.88, 0.28, 0.25, 0.32

# ---------------------------------------------------------------- themes ----
_K_PRESETS = {
    "midnight_violet": {"label": "Midnight Violet — dark navy with violet/teal, gold highlights (tech, AI, strategy)",
        "dark_bg": "0E1B33", "light_bg": "FFFFFF", "primary": "5B3DF5", "secondary": "0FA3A3", "accent": "F2A900",
        "info": "2F6FD6", "success": "1E9E5A", "danger": "D64550", "neutral": "7A8599",
        "text": "1A1F2B", "muted": "5F6B80", "on_primary": "FFFFFF", "on_secondary": "FFFFFF", "on_accent": "0E1B33"},
    "graphite_gold": {"label": "Graphite Gold — charcoal with yellow accents (consulting, executive)",
        "dark_bg": "2E2E38", "light_bg": "FFFFFF", "primary": "3D3D4D", "secondary": "747480", "accent": "FFD400",
        "info": "1777CF", "success": "168736", "danger": "C4212F", "neutral": "9A9AA6",
        "text": "1F1F28", "muted": "5E5E6B", "on_primary": "FFFFFF", "on_secondary": "FFFFFF", "on_accent": "1F1F28"},
    "ocean_teal": {"label": "Ocean Teal — deep sea blue with teal and orange (operations, supply chain)",
        "dark_bg": "062A3A", "light_bg": "FFFFFF", "primary": "0B7A75", "secondary": "1C6E9E", "accent": "F28C28",
        "info": "2E86C1", "success": "2E9B5F", "danger": "C8443A", "neutral": "7D8B94",
        "text": "13232C", "muted": "586973", "on_primary": "FFFFFF", "on_secondary": "FFFFFF", "on_accent": "062A3A"},
    "forest_moss": {"label": "Forest Moss — green with moss and amber (sustainability, ESG, utilities)",
        "dark_bg": "15291C", "light_bg": "FFFFFF", "primary": "2C6E3F", "secondary": "7E9A2F", "accent": "D9A21B",
        "info": "3D7EA6", "success": "2C6E3F", "danger": "B5403A", "neutral": "7F8A80",
        "text": "17231A", "muted": "5A675C", "on_primary": "FFFFFF", "on_secondary": "FFFFFF", "on_accent": "15291C"},
    "crimson_navy": {"label": "Crimson Navy — navy with crimson and gold (finance, banking, risk)",
        "dark_bg": "1B1F3B", "light_bg": "FFFFFF", "primary": "B3123E", "secondary": "2F4B8C", "accent": "F2B134",
        "info": "2F6FD6", "success": "1E8C55", "danger": "B3123E", "neutral": "7C8195",
        "text": "1A1C2E", "muted": "5C6079", "on_primary": "FFFFFF", "on_secondary": "FFFFFF", "on_accent": "1B1F3B"},
    "mono_slate": {"label": "Mono Slate — restrained slate greys with one teal accent (board, audit, legal)",
        "dark_bg": "1F2933", "light_bg": "FFFFFF", "primary": "3E4C59", "secondary": "7B8794", "accent": "0F9D8A",
        "info": "2D6CDF", "success": "0F9D8A", "danger": "C2413A", "neutral": "9AA5B1",
        "text": "1F2933", "muted": "616E7C", "on_primary": "FFFFFF", "on_secondary": "FFFFFF", "on_accent": "FFFFFF"},
    "azure_corporate": {"label": "Azure Corporate — bright enterprise blue with sky and amber (ERP, IT, enterprise software)",
        "dark_bg": "0A1F44", "light_bg": "FFFFFF", "primary": "0070F2", "secondary": "0FA3B1", "accent": "F5A623",
        "info": "4A90E2", "success": "2E9E5A", "danger": "D0362B", "neutral": "7B8798",
        "text": "14213D", "muted": "5A6578", "on_primary": "FFFFFF", "on_secondary": "FFFFFF", "on_accent": "0A1F44"},
    "emerald_finance": {"label": "Emerald Finance — deep green with emerald and gold (banking, wealth, insurance)",
        "dark_bg": "0D2B24", "light_bg": "FFFFFF", "primary": "0F7B5F", "secondary": "3A5A80", "accent": "C9A227",
        "info": "3A7CA5", "success": "0F7B5F", "danger": "B23A3A", "neutral": "7A8A84",
        "text": "13241F", "muted": "56675F", "on_primary": "FFFFFF", "on_secondary": "FFFFFF", "on_accent": "0D2B24"},
    "civic_navy": {"label": "Civic Navy — sober navy with steel blue and red accent (public sector, government, multilaterals)",
        "dark_bg": "13213C", "light_bg": "FFFFFF", "primary": "1F3F77", "secondary": "4F7CAC", "accent": "D9472B",
        "info": "2B6CB0", "success": "2F855A", "danger": "C53030", "neutral": "8090A6",
        "text": "1A2233", "muted": "5B6578", "on_primary": "FFFFFF", "on_secondary": "FFFFFF", "on_accent": "FFFFFF"},
    "clinic_teal": {"label": "Clinic Teal — calm teal with soft blue and coral (healthcare, life sciences, pharma)",
        "dark_bg": "0B2E33", "light_bg": "FFFFFF", "primary": "1A8C8C", "secondary": "3D6FB6", "accent": "F27C63",
        "info": "3D8FD1", "success": "2E9B6A", "danger": "C94A4A", "neutral": "82979A",
        "text": "132A2D", "muted": "58706F", "on_primary": "FFFFFF", "on_secondary": "FFFFFF", "on_accent": "0B2E33"},
    "coral_energy": {"label": "Coral Energy — navy with coral and sunflower (marketing, retail, consumer, startups)",
        "dark_bg": "1E2348", "light_bg": "FFFFFF", "primary": "E8505B", "secondary": "2F3C7E", "accent": "F9C74F",
        "info": "4D7CFE", "success": "2BA36B", "danger": "C0392B", "neutral": "8A8FA8",
        "text": "1E2348", "muted": "5E6385", "on_primary": "FFFFFF", "on_secondary": "FFFFFF", "on_accent": "1E2348"},
    "royal_plum": {"label": "Royal Plum — plum with dusty rose and champagne (luxury, fashion, hospitality, premium brands)",
        "dark_bg": "2A1530", "light_bg": "FFFFFF", "primary": "6D2E6B", "secondary": "A26769", "accent": "D8B26E",
        "info": "5B6FB5", "success": "3C8D6A", "danger": "B5384B", "neutral": "8E8090",
        "text": "24172A", "muted": "65586A", "on_primary": "FFFFFF", "on_secondary": "FFFFFF", "on_accent": "2A1530"},
    "print_mono": {"label": "Print Mono — greyscale only, survives black-and-white printing and photocopies",
        "dark_bg": "262626", "light_bg": "FFFFFF", "primary": "3A3A3A", "secondary": "7A7A7A", "accent": "A8A8A8",
        "info": "5C5C5C", "success": "4A4A4A", "danger": "1A1A1A", "neutral": "9C9C9C",
        "text": "1A1A1A", "muted": "555555", "on_primary": "FFFFFF", "on_secondary": "FFFFFF", "on_accent": "1A1A1A"},
    "high_contrast": {"label": "High Contrast — WCAG-friendly black, white, cobalt and yellow (accessibility, large rooms, projectors)",
        "dark_bg": "000000", "light_bg": "FFFFFF", "primary": "0038A8", "secondary": "006E5A", "accent": "FFD100",
        "info": "0057D9", "success": "00703C", "danger": "B00020", "neutral": "595959",
        "text": "000000", "muted": "333333", "on_primary": "FFFFFF", "on_secondary": "FFFFFF", "on_accent": "000000"},
}
# classic-engine theme names → closest kit preset (so "theme":"ocean" also works in v2 decks)
_K_THEME_ALIASES = {
    "midnight": "midnight_violet", "auto": "midnight_violet", "default": "midnight_violet",
    "ocean": "ocean_teal", "teal": "ocean_teal", "forest": "forest_moss", "sage": "forest_moss",
    "coral": "coral_energy", "terracotta": "coral_energy", "berry": "royal_plum", "cherry": "crimson_navy",
    "charcoal": "graphite_gold", "slate": "mono_slate", "azure": "azure_corporate", "finance": "emerald_finance",
    "public": "civic_navy", "healthcare": "clinic_teal", "print": "print_mono", "accessible": "high_contrast",
}
_K_HEX_RE = re.compile(r"^[0-9A-Fa-f]{6}$")


def _kmix(a: str, b: str, t: float) -> str:
    ia, ib = int(a, 16), int(b, 16)
    out = []
    for s in (16, 8, 0):
        x, y = (ia >> s) & 255, (ib >> s) & 255
        out.append("%02X" % int(round(x * (1 - t) + y * t)))
    return "".join(out)


def _klum(h: str) -> float:
    v = int(h, 16)
    cs = []
    for s in (16, 8, 0):
        x = ((v >> s) & 255) / 255
        cs.append(x / 12.92 if x <= 0.03928 else ((x + 0.055) / 1.055) ** 2.4)
    return 0.2126 * cs[0] + 0.7152 * cs[1] + 0.0722 * cs[2]


def _kread(bg: str, dark: str = "1A1F2B", light: str = "FFFFFF") -> str:
    return dark if _klum(bg) > 0.32 else light


def _k_resolve_theme(spec, warn) -> dict:
    name, ov = "midnight_violet", {}
    if isinstance(spec, str) and spec.strip():
        name = spec.strip()
    elif isinstance(spec, dict):
        name = str(spec.get("preset") or name)
        ov = spec
    name = _K_THEME_ALIASES.get(name.lower(), name.lower()) if isinstance(name, str) else "midnight_violet"
    if name not in _K_PRESETS:
        warn(f"theme: unknown preset '{name}', using midnight_violet. Valid: {', '.join(_K_PRESETS)}")
        name = "midnight_violet"
    t = dict(_K_PRESETS[name])
    for k, v in ov.items():
        if k in ("preset", "fonts", "label"):
            continue
        if isinstance(v, str) and _K_HEX_RE.match(v.replace("#", "")):
            t[k] = v.replace("#", "").upper()
        elif k in t:
            warn(f"theme.{k}: '{v}' is not a 6-digit hex color, ignored")
    fonts = {"head": "Calibri", "body": "Calibri"}
    if isinstance(ov.get("fonts"), dict):
        fonts.update({k: str(v) for k, v in ov["fonts"].items() if k in ("head", "body") and v})
    t["fonts"] = fonts
    t["name"] = name
    t["primary_tint"] = _kmix(t["primary"], "FFFFFF", 0.88)
    t["secondary_tint"] = _kmix(t["secondary"], "FFFFFF", 0.87)
    t["accent_tint"] = _kmix(t["accent"], "FFFFFF", 0.82)
    t["danger_tint"] = _kmix(t["danger"], "FFFFFF", 0.86)
    t["success_tint"] = _kmix(t["success"], "FFFFFF", 0.86)
    t["light_surface"] = _kmix(t["dark_bg"], "FFFFFF", 0.955)
    t["light_line"] = _kmix(t["dark_bg"], "FFFFFF", 0.82)
    t["dark_surface"] = _kmix(t["dark_bg"], "FFFFFF", 0.09)
    t["dark_surface2"] = _kmix(t["dark_bg"], "FFFFFF", 0.16)
    t["dark_line"] = _kmix(t["dark_bg"], "FFFFFF", 0.26)
    t["dark_text"] = "FFFFFF"
    t["dark_muted"] = _kmix(t["dark_bg"], "FFFFFF", 0.72)
    t["series"] = [t["primary"], t["secondary"], t["accent"], t["info"], t["neutral"], t["danger"]]
    t["on_primary"] = t.get("on_primary") or _kread(t["primary"])
    return t


def _k_mode_colors(t: dict, dark: bool) -> dict:
    if dark:
        return {"bg": t["dark_bg"], "surface": t["dark_surface"], "surface2": t["dark_surface2"], "line": t["dark_line"],
                "text": t["dark_text"], "muted": t["dark_muted"], "heading": t["accent"], "title": "FFFFFF",
                "card": t["dark_surface"], "cardLine": t["dark_line"], "cardText": "FFFFFF", "cardMuted": t["dark_muted"],
                "strong": t["accent"]}
    return {"bg": t["light_bg"], "surface": t["light_surface"], "surface2": t["primary_tint"], "line": t["light_line"],
            "text": t["text"], "muted": t["muted"], "heading": t["primary"], "title": t["dark_bg"], "card": "FFFFFF",
            "cardLine": t["light_line"], "cardText": t["text"], "cardMuted": t["muted"], "strong": t["primary"]}


def _k_role(t: dict, role, fallback: str) -> str:
    if role and role != "dark" and isinstance(role, str) and role in t and isinstance(t[role], str):
        return t[role]
    return t["dark_bg"] if role == "dark" else fallback


# ------------------------------------------------------------ text fit ----
# Approximates Calibri/Carlito advance widths (em). Slightly conservative so
# text that "fits" here also fits in PowerPoint.
_K_CW_TABLE = {
    " ": 0.226, "a": 0.479, "b": 0.525, "c": 0.423, "d": 0.525, "e": 0.498, "f": 0.305, "g": 0.471, "h": 0.525,
    "i": 0.229, "j": 0.239, "k": 0.455, "l": 0.229, "m": 0.799, "n": 0.525, "o": 0.527, "p": 0.525, "q": 0.525,
    "r": 0.349, "s": 0.391, "t": 0.335, "u": 0.525, "v": 0.452, "w": 0.715, "x": 0.433, "y": 0.453, "z": 0.395,
    "A": 0.579, "B": 0.544, "C": 0.533, "D": 0.615, "E": 0.488, "F": 0.459, "G": 0.631, "H": 0.623, "I": 0.252,
    "J": 0.319, "K": 0.52, "L": 0.42, "M": 0.855, "N": 0.646, "O": 0.662, "P": 0.517, "Q": 0.673, "R": 0.543,
    "S": 0.459, "T": 0.487, "U": 0.642, "V": 0.567, "W": 0.89, "X": 0.519, "Y": 0.487, "Z": 0.468,
    ".": 0.252, ",": 0.25, ":": 0.268, ";": 0.268, "'": 0.221, '"': 0.401, "!": 0.326, "?": 0.463, "-": 0.306,
    "(": 0.303, ")": 0.303, "/": 0.386, "&": 0.682, "%": 0.715, "+": 0.498, "€": 0.507, "$": 0.507, "→": 0.9,
    "—": 0.9, "–": 0.5, "·": 0.27,
}


def _k_char_w(ch: str) -> float:
    w = _K_CW_TABLE.get(ch)
    if w is not None:
        return w
    if "0" <= ch <= "9":
        return 0.507
    return 1.0 if ord(ch) > 0x2E80 else 0.55


def _k_str_em(s: str, bold: bool = False) -> float:
    return sum(_k_char_w(c) for c in str(s)) * (1.05 if bold else 1.0) * 1.03


def _k_count_lines(text: str, size: float, width_in: float, bold: bool) -> int:
    max_pt = width_in * 72
    if max_pt <= 4:
        return 99
    lines = 0
    sp = _k_char_w(" ") * size
    for raw in str(text).split("\n"):
        words = raw.split()
        if not words:
            lines += 1
            continue
        cur, n = 0.0, 1
        for wd in words:
            ww = _k_str_em(wd, bold) * size
            if ww > max_pt:
                n += 3
                cur = max_pt
                continue
            if cur == 0:
                cur = ww
            elif cur + sp + ww <= max_pt:
                cur += sp + ww
            else:
                n += 1
                cur = ww
        lines += n
    return lines


def _k_measure(paras: list, width_in: float, scale: float = 1.0) -> float:
    h = 0.0
    for i, p in enumerate(paras):
        size = p["size"] * scale
        indent = (size * 1.0) / 72 if p.get("bullet") else 0
        n = _k_count_lines(p["text"], size, width_in - indent, p.get("bold"))
        h += (n * size * 1.2) / 72
        if i < len(paras) - 1:
            sa = p.get("spaceAfter")
            h += ((sa if sa is not None else 0) * scale) / 72
    return h


def _k_fit_scale(paras, width_in, height_in, min_size=9, max_scale=1.0):
    smallest = min(p["size"] for p in paras)
    min_scale = min(1.0, min_size / smallest)
    s = max_scale
    while s > min_scale and _k_measure(paras, width_in, s) > height_in:
        s -= 0.03
    if s < min_scale:
        s = min_scale
    need = _k_measure(paras, width_in, s)
    return s, need <= height_in + 0.01, need


def _k_single_line_size(text, width_in, max_size, bold=True):
    em = _k_str_em(str(text), bold)
    if em == 0:
        return max_size
    return max(8.0, min(max_size, (width_in * 72) / em))


def _kfmt(v, d=0) -> str:
    try:
        return f"{float(v):,.{int(d)}f}"
    except Exception:
        return str(v)


def _ks(v) -> str:
    """JS-like String(v): 3.0 -> '3'."""
    if isinstance(v, bool):
        return "true" if v else "false"
    if isinstance(v, float) and v.is_integer():
        return str(int(v))
    return "" if v is None else str(v)


def _knum(v, default=0.0) -> float:
    try:
        if isinstance(v, bool):
            return float(default)
        return float(v)
    except Exception:
        return float(default)


# ----------------------------------------------------------------- icons ----
_K_ICON_CACHE: dict = {}
_K_UNKNOWN_ICONS: set = set()


def _k_icon_key(name) -> Optional[str]:
    if not name:
        return None
    key = str(name).strip()
    if key in _KIT_ICON_PNG:
        return key
    a = _KIT_ICON_ALIASES.get(re.sub(r"[\s-]+", "_", key.lower()))
    if a and a in _KIT_ICON_PNG:
        return a
    pas = "Fa" + re.sub(r"[_\s-]+(.)", lambda m: m.group(1).upper(), key)
    pas = pas[:2] + pas[2:3].upper() + pas[3:]
    if pas in _KIT_ICON_PNG:
        return pas
    return None


def _k_icon_png(name, color: str) -> Optional[bytes]:
    if not _HAS_PIL:
        return None
    ck = (str(name), color)
    if ck in _K_ICON_CACHE:
        return _K_ICON_CACHE[ck]
    key = _k_icon_key(name)
    b64 = _KIT_ICON_PNG.get(key) if key else None
    if b64 is None:
        # fall back to the classic engine's Lucide set (e.g. "rocket", "cpu")
        try:
            lname = _icon_name(str(name))
            b64 = _ICON_PNG.get(lname)
        except Exception:
            b64 = None
    if b64 is None:
        if name:
            _K_UNKNOWN_ICONS.add(str(name))
        b64 = _KIT_ICON_PNG.get("FaCircle")
    png = None
    try:
        glyph = Image.open(BytesIO(base64.b64decode(b64))).convert("RGBA")
        r, g, b = int(color[0:2], 16), int(color[2:4], 16), int(color[4:6], 16)
        solid = Image.new("RGBA", glyph.size, (r, g, b, 0))
        solid.putalpha(glyph.split()[3])
        out = BytesIO()
        solid.save(out, format="PNG")
        png = out.getvalue()
    except Exception:
        png = None
    _K_ICON_CACHE[ck] = png
    return png


# ------------------------------------------------------------ primitives ----
class _KCtx:
    def __init__(self, prs, slide, t, dark, where, warnings, errors, valves=None):
        self.prs, self.slide, self.t, self.dark = prs, slide, t, dark
        self.c = _k_mode_colors(t, dark)
        self.slide_where = where
        self.where = where
        self._warnings, self._errors = warnings, errors
        self.dry = False
        self.group_scales = None
        self.valves = valves

    def warn(self, m):
        if not self.dry:
            self._warnings.append(m)

    def error(self, m):
        self._errors.append(m)


def _ke(v: float) -> int:
    return int(round(v * 914400))


_K_SHAPES = None


def _k_shape_map():
    global _K_SHAPES
    if _K_SHAPES is None:
        _K_SHAPES = {
            "rect": MSO_SHAPE.RECTANGLE, "roundRect": MSO_SHAPE.ROUNDED_RECTANGLE, "ellipse": MSO_SHAPE.OVAL,
            "homePlate": MSO_SHAPE.PENTAGON, "chevron": MSO_SHAPE.CHEVRON, "diamond": MSO_SHAPE.DIAMOND,
            "pie": MSO_SHAPE.PIE, "donut": MSO_SHAPE.DONUT, "blockArc": MSO_SHAPE.BLOCK_ARC,
            "rightArrow": MSO_SHAPE.RIGHT_ARROW,
        }
    return _K_SHAPES


def _k_strip_style(shape):
    st = shape._element.find(_kqn("p:style"))
    if st is not None:
        shape._element.remove(st)


def _k_set_adj(shape, vals: dict):
    """Set raw <a:gd> adjust values (OOXML integer units)."""
    geom = shape._element.spPr.find(_kqn("a:prstGeom"))
    if geom is None:
        return
    av = geom.find(_kqn("a:avLst"))
    if av is None:
        av = _K_ET.SubElement(geom, _kqn("a:avLst"))
    for g in list(av):
        av.remove(g)
    for k, v in vals.items():
        gd = _K_ET.SubElement(av, _kqn("a:gd"))
        gd.set("name", k)
        gd.set("fmla", f"val {int(round(v))}")


def _k_shadow(shape):
    sp = shape._element.spPr
    old = sp.find(_kqn("a:effectLst"))
    if old is not None:
        sp.remove(old)
    eff = _K_ET.SubElement(sp, _kqn("a:effectLst"))
    sh = _K_ET.SubElement(eff, _kqn("a:outerShdw"))
    sh.set("blurRad", "76200"); sh.set("dist", "19050"); sh.set("dir", "5400000")
    sh.set("algn", "ctr"); sh.set("rotWithShape", "0")
    clr = _K_ET.SubElement(sh, _kqn("a:srgbClr")); clr.set("val", "0E1B33")
    al = _K_ET.SubElement(clr, _kqn("a:alpha")); al.set("val", "12000")


def _k_apply_line(ln, line):
    if line is None:
        ln.fill.background()
        return
    ln.color.rgb = RGBColor.from_string(line["color"])
    ln.width = Pt(line.get("width", 1))
    dash = line.get("dash")
    if dash:
        ln.dash_style = {"dash": _K_DASH.DASH, "sysDash": _K_DASH.SQUARE_DOT}.get(dash, _K_DASH.DASH)


def _kshp(ctx, kind, box, fill=None, line=None, shadow=False, transparency=None, name=None,
          radius=None, adj=None, no_fill=False):
    if ctx.dry:
        return None
    w, h = max(0.01, box["w"]), max(0.01, box["h"])
    shp = ctx.slide.shapes.add_shape(_k_shape_map()[kind], _ke(box["x"]), _ke(box["y"]), _ke(w), _ke(h))
    _k_strip_style(shp)
    if no_fill or fill is None:
        shp.fill.background()
    else:
        shp.fill.solid()
        shp.fill.fore_color.rgb = RGBColor.from_string(fill)
        if transparency:
            sf = shp.fill._xPr.find(_kqn("a:solidFill"))
            clr = sf[0] if sf is not None and len(sf) else None
            if clr is not None:
                a = _K_ET.SubElement(clr, _kqn("a:alpha"))
                a.set("val", str(int((100 - transparency) * 1000)))
    _k_apply_line(shp.line, line)
    if kind == "roundRect":
        r = radius if radius is not None else min(0.07, h / 2, w / 2)
        _k_set_adj(shp, {"adj": min(50000, r / max(0.01, min(w, h)) * 100000)})
    if adj:
        _k_set_adj(shp, adj)
    if shadow:
        _k_shadow(shp)
    if name:
        try:
            shp.name = str(name)[:80]
        except Exception:
            pass
    return shp


def _krr(ctx, box, fill, **o):
    return _kshp(ctx, "roundRect", box, fill, **o)


def _krect(ctx, box, fill, **o):
    return _kshp(ctx, "rect", box, fill, **o)


def _kell(ctx, box, fill, **o):
    return _kshp(ctx, "ellipse", box, fill, **o)


def _kline(ctx, x1, y1, x2, y2, color, width=1.0, dash=None, arrow=False, start_arrow=False):
    if ctx.dry:
        return None
    cn = ctx.slide.shapes.add_connector(_K_CONN.STRAIGHT, _ke(x1), _ke(y1), _ke(x2), _ke(y2))
    _k_strip_style(cn)
    _k_apply_line(cn.line, {"color": color, "width": width, "dash": dash})
    ln = cn.line._get_or_add_ln()
    for tag, on in (("a:headEnd", start_arrow), ("a:tailEnd", arrow)):
        if on:
            e = _K_ET.SubElement(ln, _kqn(tag))
            e.set("type", "triangle")
    return cn


def _k_image(ctx, png: bytes, x, y, w, h, ellipse=False):
    if ctx.dry or not png:
        return None
    pic = ctx.slide.shapes.add_picture(BytesIO(png), _ke(x), _ke(y), _ke(w), _ke(h))
    if ellipse:
        g = pic._element.spPr.find(_kqn("a:prstGeom"))
        if g is not None:
            g.set("prst", "ellipse")
    return pic


def _k_icon_dot(ctx, x, y, d, fill, icon, icon_color=None):
    _kell(ctx, {"x": x, "y": y, "w": d, "h": d}, fill)
    if icon and not ctx.dry:
        s = d * 0.54
        png = _k_icon_png(icon, icon_color or _kread(fill))
        _k_image(ctx, png, x + (d - s) / 2, y + (d - s) / 2, s, s)


def _k_icon(ctx, icon, color, x, y, w, h):
    if ctx.dry:
        return
    _k_image(ctx, _k_icon_png(icon, color), x, y, w, h)


_K_GROW = {"card body": 1.15, "process step": 1.2, "timeline item": 1.2, "comparison bullets": 1.2, "agenda item": 1.12,
           "hub node": 1.12, "quadrant": 1.15, "layer item": 1.08, "side item": 1.12, "bullets": 1.1,
           "pyramid level": 1.12, "funnel stage": 1.12, "leaf": 1.08, "kpi label": 1.1}


def _k_bullet_xml(p, size, numbered, level):
    pPr = p._p.get_or_add_pPr()
    ind = int(round(size * (1.2 if numbered else 0.95))) * 12700
    marl = ind * (1 + (level or 0))
    pPr.set("marL", str(marl))
    pPr.set("indent", str(-ind))
    if level:
        pPr.set("lvl", str(min(8, int(level))))
    for tag in ("a:buNone", "a:buChar", "a:buAutoNum"):
        e = pPr.find(_kqn(tag))
        if e is not None:
            pPr.remove(e)
    sz = _K_ET.SubElement(pPr, _kqn("a:buSzPct")); sz.set("val", "100000")
    if numbered:
        b = _K_ET.SubElement(pPr, _kqn("a:buAutoNum")); b.set("type", "arabicPeriod")
    else:
        b = _K_ET.SubElement(pPr, _kqn("a:buChar")); b.set("char", "•")


def _ktext(ctx, paras, box, valign="top", field=None, group=None, grow=None, min_size=9, quiet=False,
           align=None, rotate=None, char_spacing=None):
    """paras: [{text|runs, size, bold, italic, color, bullet, numbered, align, spaceAfter, indent, font}]"""
    ps = []
    for p in paras:
        if not p:
            continue
        p = dict(p)
        if p.get("runs"):
            p["text"] = "".join(_ks(r.get("text", "")) for r in p["runs"])
        if p.get("text") is None or _ks(p.get("text")) == "":
            continue
        p["text"] = _ks(p["text"])
        ps.append(p)
    if not ps:
        return 1.0
    if box["w"] <= 0.05 or box["h"] <= 0.05:
        if not ctx.dry and not quiet:
            ctx.warn(f"{ctx.where}: no room left for text \"{ps[0]['text'][:40]}…\" — give this row more height or use fewer items")
        return 1.0
    key = None if group is False else (group or field)
    g = grow if grow is not None else (_K_GROW.get(key, 1) if key else 1)
    if ctx.dry:
        s, _, _ = _k_fit_scale(ps, box["w"], box["h"], min_size, g)
        if key and ctx.group_scales is not None:
            ctx.group_scales[key] = min(ctx.group_scales.get(key, 9), s)
        return s
    if key and ctx.group_scales and key in ctx.group_scales:
        s = ctx.group_scales[key]
        need = _k_measure(ps, box["w"], s)
        fits = need <= box["h"] + 0.01
    else:
        s, fits, need = _k_fit_scale(ps, box["w"], box["h"], min_size, g)
    if not fits and not quiet:
        pct = round((need / box["h"] - 1) * 100) if box["h"] else 0
        ctx.warn(f"{ctx.where}: text overflows its box{' (' + field + ')' if field else ''} by ~{pct}% even at minimum size — "
                 f"shorten \"{' '.join(p['text'] for p in ps)[:60]}…\"")
    tb = ctx.slide.shapes.add_textbox(_ke(box["x"]), _ke(box["y"]), _ke(max(0.02, box["w"])), _ke(max(0.02, box["h"])))
    tf = tb.text_frame
    tf.word_wrap = True
    tf.auto_size = _K_AUTO.NONE
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    tf.vertical_anchor = {"top": MSO_ANCHOR.TOP, "middle": MSO_ANCHOR.MIDDLE, "bottom": MSO_ANCHOR.BOTTOM}.get(valign, MSO_ANCHOR.TOP)
    if rotate:
        tb.rotation = rotate
    amap = {"left": PP_ALIGN.LEFT, "center": PP_ALIGN.CENTER, "right": PP_ALIGN.RIGHT}
    for i, p in enumerate(ps):
        para = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        size = round(p["size"] * s * 2) / 2
        para.alignment = amap.get(p.get("align") or align or "left", PP_ALIGN.LEFT)
        sa = p.get("spaceAfter")
        para.space_after = Pt(round((sa or 0) * s)) if sa else Pt(0)
        if p.get("bullet"):
            _k_bullet_xml(para, size, p.get("numbered"), p.get("indent"))
        runs = p.get("runs") or [{"text": p["text"]}]
        for r in runs:
            if _ks(r.get("text", "")) == "":
                continue
            run = para.add_run()
            run.text = _ks(r.get("text"))
            f = run.font
            f.size = Pt(size)
            f.bold = bool(r["bold"]) if r.get("bold") is not None else bool(p.get("bold"))
            f.italic = bool(p.get("italic"))
            f.color.rgb = RGBColor.from_string(r.get("color") or p.get("color") or ctx.c["text"])
            f.name = p.get("font") or ctx.t["fonts"]["body"]
            if char_spacing:
                run._r.get_or_add_rPr().set("spc", str(int(char_spacing * 100)))
    return s


_KIT_ICON_ALIASES = {
    "agent": "FaRobot",
    "ai": "FaBrain",
    "analytics": "FaChartPie",
    "api": "FaPlug",
    "arrow": "FaArrowRight",
    "assistant": "FaUserTie",
    "automation": "FaCogs",
    "award": "FaAward",
    "balance": "FaBalanceScale",
    "bank": "FaUniversity",
    "brain": "FaBrain",
    "building": "FaBuilding",
    "calendar": "FaCalendarAlt",
    "cart": "FaShoppingCart",
    "chart": "FaChartBar",
    "chat": "FaComments",
    "check": "FaCheckCircle",
    "clock": "FaClock",
    "cloud": "FaCloud",
    "code": "FaCode",
    "coins": "FaCoins",
    "compliance": "FaClipboardCheck",
    "contract": "FaFileSignature",
    "cost": "FaMoneyBillWave",
    "cross": "FaTimesCircle",
    "cube": "FaCube",
    "customer": "FaUserFriends",
    "cycle": "FaSyncAlt",
    "data": "FaDatabase",
    "database": "FaDatabase",
    "desktop": "FaDesktop",
    "diagram": "FaProjectDiagram",
    "document": "FaFileAlt",
    "email": "FaEnvelope",
    "energy": "FaBolt",
    "eye": "FaEye",
    "factory": "FaIndustry",
    "filter": "FaFilter",
    "finance": "FaChartLine",
    "flag": "FaFlag",
    "flow": "FaStream",
    "gauge": "FaTachometerAlt",
    "global": "FaGlobe",
    "globe": "FaGlobeEurope",
    "goal": "FaBullseye",
    "governance": "FaBalanceScale",
    "government": "FaLandmark",
    "graph": "FaProjectDiagram",
    "growth": "FaChartLine",
    "handshake": "FaHandshake",
    "health": "FaHeartbeat",
    "hierarchy": "FaSitemap",
    "hr": "FaIdBadge",
    "idea": "FaLightbulb",
    "info": "FaInfoCircle",
    "innovation": "FaLightbulb",
    "integration": "FaPlug",
    "invoice": "FaFileInvoice",
    "key": "FaKey",
    "kpi": "FaTachometerAlt",
    "laptop": "FaLaptop",
    "launch": "FaRocket",
    "layers": "FaLayerGroup",
    "leader": "FaUserTie",
    "leaf": "FaLeaf",
    "learning": "FaGraduationCap",
    "legal": "FaGavel",
    "link": "FaLink",
    "location": "FaMapMarkerAlt",
    "lock": "FaLock",
    "logistics": "FaShippingFast",
    "magic": "FaMagic",
    "manufacturing": "FaIndustry",
    "map": "FaMap",
    "migration": "FaExchangeAlt",
    "mobile": "FaMobileAlt",
    "modules": "FaCubes",
    "money": "FaCoins",
    "network": "FaNetworkWired",
    "partner": "FaHandshake",
    "people": "FaUsers",
    "pie": "FaChartPie",
    "platform": "FaLayerGroup",
    "plus": "FaPlus",
    "procurement": "FaHandshake",
    "puzzle": "FaPuzzlePiece",
    "question": "FaQuestionCircle",
    "quote": "FaQuoteLeft",
    "recycle": "FaRecycle",
    "report": "FaFileAlt",
    "research": "FaSearch",
    "retail": "FaShoppingCart",
    "risk": "FaExclamationTriangle",
    "robot": "FaRobot",
    "rocket": "FaRocket",
    "search": "FaSearch",
    "security": "FaShieldAlt",
    "server": "FaServer",
    "service": "FaHeadset",
    "settings": "FaCogs",
    "shield": "FaShieldAlt",
    "sparkles": "FaMagic",
    "speed": "FaTachometerAlt",
    "star": "FaStar",
    "store": "FaStore",
    "strategy": "FaChess",
    "success": "FaCheckCircle",
    "supply_chain": "FaTruck",
    "support": "FaHeadset",
    "sustainability": "FaLeaf",
    "sync": "FaSyncAlt",
    "talent": "FaUserGraduate",
    "target": "FaBullseye",
    "team": "FaUsers",
    "telecom": "FaBroadcastTower",
    "time": "FaClock",
    "tools": "FaTools",
    "training": "FaChalkboardTeacher",
    "transform": "FaRandom",
    "trophy": "FaTrophy",
    "truck": "FaTruck",
    "user": "FaUser",
    "users": "FaUsers",
    "utilities": "FaBolt",
    "visibility": "FaEye",
    "warehouse": "FaWarehouse",
    "warning": "FaExclamationTriangle",
    "wrench": "FaWrench",
}
_KIT_ICON_PNG: dict[str, str] = {  # Font Awesome Free (CC BY 4.0), white glyph 96px
    "FaArrowDown": "iVBORw0KGgoAAAANSUhEUgAAAGAAAABgCAMAAADVRocKAAABO1BMVEVMaXH///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////9uRaD0AAAAaHRSTlMAZ/bz/V/+gPtX/Es5XEX5XgH4BIIQ7UJ1uKNB62rwt/rlJhFURGg0fF3oM+KrI+yKKGvSmTeehCfWElWdO+AcAmms75KX2jrkKpEVm96ggR/O8i6VGWGkCQ4TAwiDDLvAxcoGz7C1qi3bu6sAAAAJcEhZcwAAFxIAABcSAWef0lIAAAIKSURBVGje7dhnT8JAGMDx2rJkCAg4mIqoiHvvvffee9/3/wSuVOjd0/buuJqY3POOPL3/TxJi0lMUOXL+cGabNfQ9WvO6E/1aH/od37T4/lQWVcxym3CgCxkmJBwoGoGicMBrBLwSkIAEJCABCUhAAhKQwP8GEosr+53VAKPjJ0MWrw6p9NfZkRgvENv5WvWnzPp16s/h5AIfEJ352dUHrPsIjUV5gOiqvoSFch+hLR5goLyFhHa18nycHYhXrtV2fD2oGQNNrECTcX+Ux/ZJ4x55MmxAxoMF9rAHerA9crlZALcLPz+HAa34A8hfoAcKfuL4BgYsEU+gcIAWCITJ070Y0NlCPqM20gGNKvDXdeO/ghD5LZEWpAGCGnnSH7K+TdEn220PHDcA/YjdfY0+6U07IF+i7cNCacIa6CN/fsgVoblz0qfnwAqY3wX6brpbLX3WTs2B7Um2PiwMX5kBiUPWPiy8JGDg8om9DwuPdxBw8cDTh4X7GxJ4fuPrw8JrDAeu33n7sNB7a/wcP+fvf15hAv+XGiw/svXh72A3LH0ega3PLrD2WQX2PpvA02cR+Pr0Am+fVuDv0wnV9GmE6vr2QrV9O8FToyhOCiL6VoKYvrkgqm8miOvDgsg+JIjtk4LoPi6I7xsFJ/qVgqdDUZwUnOp/vsd9vyn6zhTHJpjTtFxQkSPnb+cDKkqf2HsLAMwAAAAASUVORK5CYII=",
    "FaArrowRight": "iVBORw0KGgoAAAANSUhEUgAAAGAAAABgCAMAAADVRocKAAABI1BMVEVMaXH///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////+GeTxlAAAAYHRSTlMAXV/2Xv3+d4j8aGf7RfJhuBEBA5I186DsKNxBMfpaO+rQ5CBE1Pg+mhlcEK3wx70GCjJ1W+GjhOsnVTYX5SZK741P9Lb5FjqhP0mMTjTjgartbJV2Vzif4oCKeawM6Qst1swEAAAACXBIWXMAABcSAAAXEgFnn9JSAAABoElEQVRo3u3Z104CQRSA4UFckCKCFLEggg2l2Hvvvfd+3v8pJNwoMhfGnP9CsucB/i8h7M7sjDHuuNNsszT/uFZKcv2hlFQnM071271SG6eF7VPCV58Rer/1RQKt2v1YWgQVrkVY4UZgYaEBEEdVWBFaKNBC/MEmaD4PnQlcOKEF48eFGR8thBxa6A/TwulIo+ALaQrHXloYjdDCWZoW+gZoofucFgZ7aGFs2CIENYX9HC0UK7TwMUELU1FamHylhXjXr4WYx/+XeXm3vL07LP27suiNRciGRVAhKsIKEYEFUZ9wFgYkSgOpfw/swkDmnv2bBjzsg/azbw7CbN+Y5U22X31dzz61/WGey7/sKy44mh/okyW2b1v0NftvE2w/VmH7qzm2b9v8avZt23fNfnee7ds+oTT7ySu2fxRh+xtetr89wvZthyGa/aDD9m99bN92pKbY79yC+wm2b1t+VX//grD7hyTcNxdw3yTgvgnCfVNcZPsN11zq/fqLOv1rtHqB6VePfHfYfvW6t7Ya72HXvcasH87lL6eNO+4013wC56ZZDqRZ+qQAAAAASUVORK5CYII=",
    "FaArrowUp": "iVBORw0KGgoAAAANSUhEUgAAAGAAAABgCAMAAADVRocKAAABSlBMVEVMaXH////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////k2/KfAAAAbXRSTlMAYPb7/FhhgP1fQfXlRVT5/jpeATld8gQQ7XXrifD6JmnvEZfa83yB6LuCQiOu0jITW57ho6gnakuT1rdVmDbgHEAhAkr0MLacoT9EyioLkTQV3h/OLjO5mnsZXD3ipAkOAwiDDMDFBqvPsLWqHmh08gAAAAlwSFlzAAAXEgAAFxIBZ5/SUgAAAhdJREFUaN7t2OlXUkEYx/FfgCuKbAq4geKe5lpWllZqi7anlfu+y///tvB0DszMw9yZuXM9vZjnHczw/bBcDtwLuHFzv7M7Gor+SgTXrw+X/k5ve2D9xtLdNDYF2w9KqPSDEar7QQhs377A920LYt+uQPVtCnTfnlCrb0t4ULNvR5D1bQjyvn/Bq+9X8O77E1T6fgS1vrmg2jcV1PtmQl2kVApS0OvrC7p9XUG/ryc0EZ9vTnqzPJGkn+d/ecLeHts3fw1U/zrTwN7RcHRrKlD90WPwAC5uzASqf34KEcDBmYlA9a+yoAAcftQXqP6jKdAAshu6AtUf2EYtAN+e6wlUf+8HagN49V1HoPojTyADMN6iLtSFxb1dnyAHkB9RFah+z094AdjMqQk7xPszmIY3gPQgIQjncYu94q7QQ6gAaA6Jj22b5IAJYs9TqAFYaRMf/ZgDxIMhEoMqgJj4/i5xwKzX5yQFkBR+QQY44AO3Hp+HDoBCnAv85jbkO9j1TugB6GTX+97wG9aZQ2EMugCGq5dX18QvQmuV8AX6ALorq9Fm6qtcEWZSJkCqKO9XhLllmABIvZf3gddd5fWXGZgByBTLx1KL5IJStvBs+Cu9pAIAb7e6k5/N/syrAT7GAQ5wgAMc4AAHOMABDvhfgCEWGLIOtLPAtHWgv4c5iey3DuBF1dWYcAwBTGLh37l0x7s03Li5x/kD32DMYI/uAXAAAAAASUVORK5CYII=",
    "FaAward": "iVBORw0KGgoAAAANSUhEUgAAAGAAAABgCAMAAADVRocKAAACiFBMVEVMaXH///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////+VVzLhAAAA13RSTlMA+gP1BAEI/v0C+c7v9/wGhJ/hT9LihS03IusS3r/olSyHt+kjR/GzVnpB23kXDBE17m0a42JDGWuoiNWSkYv0G0KP9tf7YOatb7J0VVrgokowOmOmaBD4xCSqDxQ9HmEJIE7zWDajvV/nZqnsKHAyuDx/HUbM8JN+5UVpB6zQttoW3+QvNBU5oD6BZwVstDNqwSFZwPILCoquTFETGJ7HS47RUDHWmDtAg8ZkpyaNK26XeJyhgsVEJSrPsLXZd7pdH6WGSO2v03HYdrlXON2+ww6JDeqMKb9++f8AAAAJcEhZcwAAFxIAABcSAWef0lIAAAUoSURBVGjerVn3X1NXFH+QPSBAEjCEJEJAZQmogMpGZCggIAJWmUWgrXvWra2zzqrV7r333nvv3b5/p6YJ5N17X+6554XzE/m87/l+eeOee873SpJI9G19N1jnmpiqD/2on5pw6YIDbRukWYvdm+VIuCcn3dN/b949W/zrkmXVSD43O/xrzXKMMF2eDf42kxwzzIvipl/8tsyNM5VxsQ8/kCADkdA6vFgb+8KDNlkwbOsW4vm33SUjYskcLP/pJBkVSadx/CVuGRlpfSiBH2V0rEYJHMMLHMPwb5E1xCsIgUVaBJYiBOarEXS+cbXba7F4uzueTVO7vkqcv5hdv6ZVBZYowFJxka1Q1uPCAulM8p5SGvMLu85TRPkr6PpvTklkUfps+iaSG8T4ny6j771dHfizlQKWdYvwz82k0lY+GAs6tpKCZs6F+Stb6Ee7lrOZ0tiWRyB+i4/OWcOD99Lo3wyAQB6dESzhwb3085TfAwRuoRNu5+OX0vhlgEAhhb/NyMcf2k8l/AAI0PvMO9hFWQfgl1D4LZDAa/TGA+Cv4eA3424ywwfAU0j4y7DAHbiCVEHC74QF7iEzCgC4kawv5QKtK5HQb4TwTQS+HRZYQCRUgfifCPwCrMB2EP8k9hG9RCQcgeB6HYHfiX3JdXoAPgdXWdjqBTWp+VRTi15obwLwHGrlfwDxZ1AJOQA+C9spZFMJhQC+H1muR/xUgg4Q6EQ2hG00PgsQOEonZHp58I3f0PiPAYFxpl/rRe2w4Np3NDEpBzFd+DMOsO/SMY6BMxb2IcZj0Al0Xl1WZhTuUEd+zjTh/V0iveNTTJ6p9jwLO/8E08EnOMW633J2uHDtoEFnv2NRw6L9+xmV+WV7nqKFT/z6MxVIqvAAsl7VYNn/6FTpKaPxVGlzaouquZMrPkPdq2UI9InzS0VaBD5FCJzUIrANIWDx4/n9FsyovxwvkI7yKobmYfn37sL5Oe1YgQ6sIzVeh6Ff0Yz31N4P3CdKf2tgvSZb0Fix/BOYvSq9wChpjx7AetyzKV5nVp/K40/VayLN61H8MNQo6pnb53MrKmGNcuzuyRMWGDA3Kg8HXo/sotZ9V0I/Pfsie55O6VxvaDQPiPKHusG91YPRux/66kN7kuv+0enfo7WuJLstfSj6HAerQyszQ1AgEP4H7YebE4XeUvlhezgjIGhaR8vEfAcMd1ycgc8Ts7DPKb6S30Ef7uFJcWMjEsSXn3WSDy4l2nGbCD8145i/9ES/WOfNZbXJGf02c3tNuPkmFMzCMr9YNJg7cv2K8x+3/YYk3bCn1Y6NXh/JPZv/mFlDY9Gn4yzcmhCimldVvaDA97zK8H/n+S/n1EjeCgpUcbKvwU1NE+TZFfBuYDyMucrDPI8zZ8hh51BkaQU5IGCyLrZycrNFeo6ESswcTvaeM/uix8yBFXEbLt7Z1tHY4zrhwvEq5OO81zcWxTl5uC7syce0baGcDzo5QM5JyK5kTh5h3+3kAM1fiFqOpMlfT3j0vEPUF4SdEGW0ktBWDvTXmAIrOFkHSOgBXqOn5R0XUjXGUBgb+1FMgfojMy+qzG/in1SuJhehv2xm8dk4B9iGEw2XMjzFJaFh5QJxmruR8VkIn/5CaJWWFHsyLjWcMIi2R/n8c5YJ0fIg5JY9x17dEb3aqLXxfZVrGLumrwa08kuGtyIUl3nHN2ssmgWkxGXhD/sPtYt/hSegb/VSHOHIIXYalVk3xyHFFcZqe/DPGP+jvvZvew00Pv0H9u6OqH6K/XQAAAAASUVORK5CYII=",
    "FaBalanceScale": "iVBORw0KGgoAAAANSUhEUgAAAGAAAABgCAMAAADVRocKAAAB/lBMVEVMaXH////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////unmXlAAAAqXRSTlMA9fnvM4xzZsyZH+XoCvoaF/wI8RD3EgcVm54CAQ8JhBGcgvOd9CKDBoGV1ecjVh3CiZg76trAXunKCy/443GfgCYWKtMnhQSKDnoDS/Bv1tt5X6IhbjCmRDTZ/V2qSPIUmh72l326pMgori3SszLcGOF/YEHmx6hVWuB+HMZkaY22VP46G+RjJWirsHJTdznEYoYkyWftbE202Di+XD3iw46T0M+QdROgZEl6NAAAAAlwSFlzAAAXEgAAFxIBZ5/SUgAAA6xJREFUaN7tmWdb2zAQgB1oEmaAJk0gJBAoKXuPsKGMMsssBcosq3vvvffee0//y1JqR7Z0kvXEoQ8ffN+wdPfmxfbZkgXBCCOMWGMR0zsU6J1breotzQPiSlSeS16F8p7GGTEY9j5PuOvXHhBVcc8V3vquKhGLvR2hVdreGQ0FUV8U34MTO+s06keIOiOCTSgVdUcpExCtHxBtAPQB8vQD8piAgH7AJBOQejlSX/lIc2oIt3cTXMwdtk5UBAMehK/VVUP1C3LC10xHIMClEIulFrem4cc6LpD1P/fgs9Jaz2uf4bqm+eXk3ZPYzOJPeP2zN7DfNfR8+ajJvY9d/47crt941QPPBtT1dxSqx31XpIH6i6z6r1CFk9gz0TmqrL/Lh10GNWjsBb3+nElRYxAfLS8IXj/X8bFuRaK/mApIUf7IrT58OEoeisJHvJFc7S5G/bxs5gfkqhJt23gEAAUqwIs1sJ1sgbgFaV4uL0AWyItlKsgClkTpXNudfABZYEOChaUQFIgXzLACDZApHTYL62WFTSwBQaAoUACOoIAgyAopTAEhqJDJA0ACAkshAwnQFGCAUgApZDAFKAowQCnAUFAJUBRAgFqAqoAJIIVsLUC2SoCqgAkghY2zbAAugBSymAJIIZ0NwAWQQtJtBSALF1AoJLMApABScDMFIAUAQAooFDazBCAFEgAJAAqgAFLYQgekAwKAAigAKBAApx0SQApNTAFSgQDAAoQCRYBUwAE0AaTwmymAFJ7CAJqAQqFdsWqykC1cVpjpggBdkoApkcxUPBccfqoAUhiBAH1UAaQw/yvIWrICsSi/KHpIgOet9NcilLkkDf4U8rlWRuUk4CNXYr6QxDXvCQm4xZVoE0xc886QgA9ciX7Of9EjElDOldgm/OCZ9jKHBLi4FrxfhNYk7i0a7DI1cyT+vdM4dg/iHBDAWa2d+XV5Xvx3zSvhONwqxmxamVEry8mW/exZ9Vdp3fR0PTvzm9Ql01JYs468oz9wDh5iZQ7XBlvHsRPUs+R2sJ7JPjf1GokNqNbQlgpo0sRou9aL155TE1BmdSO+mu8ZG7bjfeRhAs/b9az18TV1omnqNbiP4Rq3lpW0NfhtFZU1U4PTyfyLwJbp7rs1lRU2f0NbyYJ1PNRdEjogTGEA/n2HKFxHjX4Z0E+fU6jx5aK2RO+uZlU8E3Bf/75pIxNQph9QZuy+r3FAs37AUSYg4abe+g2H2Xeat8iqK4q8xidlI4z4//EHs1nUmaiWL9wAAAAASUVORK5CYII=",
    "FaBolt": "iVBORw0KGgoAAAANSUhEUgAAAGAAAABgCAMAAADVRocKAAABdFBMVEVMaXH///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////89ue+RAAAAe3RSTlMABujd9wQBCP6TLQ772wL5eAu1oHH8hdcTN71/YMxQHE9e7vMaPoFDpOdrCa7cLPXWecYHSu/VgsUlTvh8bAqv8tjDQiPIqUf2lFaZejyjxwOoRusnrVqd4H6DZNAMkjW738ChP+mML5bZ/d6/xKWGBe3TD1IUV5oZXICe8r77AAAACXBIWXMAABcSAAAXEgFnn9JSAAACbElEQVRo3rWZB1MTURSFX0hCAiSEUBQIiiCGjiAdRIoKCIjSOwpWOkrV/fPOuAszhJe57+259/sB55vswD1nE6WyESgK7jtZicXXIgridNchGMMMaYfkJSQ4oAV7SH4ySgtCiKCNzndaEcETA8E2Iig1EHxFBHl0fs5HRFBMC35Cf6UPaUE9kp8bIvMXwoigmv4AV9ATekYLJiBBFZlfG4AEvaRgBTvWDaSgCxO8oPK7sSekcijBFyyfPtYDmCCfyh8BC5k81sdYPn2sz0EBdayfJ0EBdazHwXzyWE+B+dSxntwCBdSxXkKfEHWsl1EBcawHC1CBwS7V7e2U8b/3rOOPo05DQZ9PgfHebvQrGOVb1llOCN+y1tNuJhjwLZjhW9b6OVzAtqz11PAtay2PCvmWtY5oBd+y1lLCuKx1PK1jXNY6ihiXtY5ghHFZ63jMuaw1VHIuaw1l+RYCP8f6nU1hJuzzf5XbCFrtBUNWlf/XOv+33StzxLqSbd/5m1PxmE3+sOLijzY/0c8mmNYKHrDlq3akJw24iCI96ffV0LgnDfiO9KQJb5GeNOBVAulJA15DPWlACupJAz5DPUkTHoF6kuYT1pN+Nncea75qufe9Yzlrfu63TEET7wdYzMzvCfMKNjIF87z5qiQjv4U5XwXFetLlWq4nXS7FelL/5lnDna82xXrS5axMrCddusR60uNErCc9xsR60qUzJNaTLu/FetKjQ6wn73//wtyTLoXdYj3p8kOsJz2GxHrSY06sJz1WeX5Dzj4oYvzvk3fZufle/42QYF2qJ29fPT78zz8MSAlUczoeGu3wlf8PB0h5VCuP+qgAAAAASUVORK5CYII=",
    "FaBrain": "iVBORw0KGgoAAAANSUhEUgAAAGAAAABgCAMAAADVRocKAAACN1BMVEVMaXH////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////ZmOKXAAAAvHRSTlMAggH4CYP5VZ/gCrT3pOcO9pmoCPoHDM7Y+x0DifUxIervjRG7P2jzEFOl6MkuV3seQsgtcJTcehi94cKjhEZqDZ0cQAJFaU8w1bbaeBahIKbpBs9tC5Ev/RrjJAWqSPIPlZoZgGGFI6lHiihrTPGPMnWef0Eix2XrJ8xL8I4s0W+yk3S3F7xawaLlxqeIJstK0G70c1T+nMBEhyXKSSqwTtRytdkVulib3nw+H8RiQ2ftiylNUnbdXD3iw+dslLEAAAAJcEhZcwAAFxIAABcSAWef0lIAAANrSURBVGje7ZnlXxtBEIaPhIRACCEkQHAt7sWKS3EpXqSFQqHu7u7u7u7unT+uJOGavduT3GX3E8zXd/Z97reXzc3OMMxczPp40DfQVZT7OwpXot7mFnUN9H31yv5RFriie7OVq/j+iJ+RCr6ptg9bBe4oPoBK2a8RKSRMnX9SFqAxNs8tbUvlSAVpavwXtgI3Po+wUtwunvQiSgWgGvhxipXeY1KKcv+ATMwF1rukJlwJ8lUMyMFd4KFLMgpIforMl73UGdoEXGKcz9kfLCC1GXQ5Zg/tF5n0IBLzHfpiMVVf8c4T/xQtiMZqR8KouB64U94/GiTCudN+Uhl35Pw/gHcA+CSz/1pvAfrtkoAK8BYAnZK/T733AK1Z4elSCpA8cytIAHQSgC0kALUSgFckABMSgDckALni/nFGEgBjnIh96BcbkACALTpUyH+oF4AMACB8Oe6f9xPIAaC4ju9vvQgkARDO/4YmAlkArOH6r7WRBugvcAAWIA2AcU6NMkYesAF9C/lAHgBIqcmU0AC0I4BxGoAjCCCDBiADAdTSANxEAJU0AOiX7ToNQDp6mYkgD4hIQk+akTzgsOcVqTrAecQ+7SmNd2D8fzUcagEaAHg8810zXwY6AOitcQJOAy0AGBz5DUAPALen87NoAs4yTCzQBExfGyvpAnTMArqAZOYMXcAwE0gXEMgE0wVcZYbpAvzpv2QfuoCt4r0TMoBYhrlHE7BjOv8jTUCVY8FSeoAQV/+1hRbAPtO2yHtGB2CPZD/Kx0/QACR3uMsKzbog0oCgRg23D2snC7AH8O+xf8gCyrCLuJYsQIsB4skCMjHAQcI3fQzQSRZgwADfyQKaMEB/PElAdwDez7lCElAt0JDqCScH2LNEqOWVl+nZvS5d/m8iUrhpl+3vUfuhXraSyBbrOnYck1l615G1Wyap9blE47Tq76DE0gTn1vbsk0gZLG+QmVCUTvkYDD4JUt2Hc8J0x7qpUg8HRRsFPUZd4l5Bcb+iQVdkqoDFLVYtFxBT65TN6Z7gFsHNrLhJoCQvVDgI1Jj4DjdOutVfeIGiUTpqtPLKJRtnHtrI656EhKoY95agG+Ffz1Xzy9DNa9cwamJkIoa1v4Y9ofUoe+5jLCtVD91rLk1aTImHmoXV2MJEk2XyfhgzF7M8/gFrMzDoHnFuOAAAAABJRU5ErkJggg==",
    "FaBroadcastTower": "iVBORw0KGgoAAAANSUhEUgAAAGAAAABgCAMAAADVRocKAAACalBMVEVMaXH///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////+mISYoAAAAzXRSTlMA9dsyAgHMZjP8EQcSBB3r+urFh7D9tFzHsvcKR4r2WwiIIAYQH0hxctzz8dmv5wl1E/tgwaIMsfm7nOk0d5uaXfSdegP+SoXsj1GoRolLMUDGc6Hkyi/4OejJBe3y0xniHqV4QnCzVjyt8A2YeVp+5SEmMLYbXmNT3r8+xIYuFDjRVBxrJZRFbE3YI5K5qc3SvZ5/wuaEZW+3Nrw7X6fLrI3Qbk+X2jrfwD+r7ozPbdQVuqBDqpAPlXa+n8Omo4K1UH0k144LImhOkWc94XtFxQAAAAlwSFlzAAAXEgAAFxIBZ5/SUgAABR1JREFUaN7tWfdfU0kQ3wSiJhAITSEgIhBRQDrSLFgAPZATTyxY8NCzne30zt776fXee++99972f7q0+e4r+5KX997n7pd8f9r9zuzMy+7MbAljaaSRxn+DjcOTI2itdMbcjdaouX4/Ecse5DGU73LC/ifr4ua6MuPMTZww1wkHhTD3Spy5E8wLTjiYCnNTY0SoCMywEw7EhHgvRonvQbznccKB5wMY/ClKzEF/hjNRdBQGr0e6P6O7p9QZB/NqYfL3cPcSelecyqtbYPKPcK8EvVVOOZiljPtl+D11G5F5S/MnhTH2arN6YOiJiba2iZdDavbs3rGIdv4+yivWGITNTHZZkmVLifpKaSinPR7PRe05Sv5N0l4AqgtGb7Af0b4G+f3gFMu+Khcsz31L8P4+YieBWwDNP8XX8pMkLm0hao2wsz7IFQieE5KtRAYaiRqC4jY2gTYmfBeoj2HF18BVeLRaEjPPYDZBfcdOUbMBI1ZLMm8612AmRNPA1SCZ64h6hD1AzTaM+Awj7iKqYrPWQV2W/nNXwMTzRL3E3qDm25C2o1gh8Pq5DrMQ1F6i2mHiI6KOs03UXKSvtxskYQE0QdijKc/KPWETC1DzC0hHiNoCaoXegZiPLUSNgHoOkcVof+P3QHqCKLfxGnO+EkI3USdALSHKxd6hZjeku7GQoN7XO7gEIQJgtz50G9hBan4K6UIYqSDqZr2DtchLrl9GrEE9O0/NbyCdiREZ+nkmnIdoPrgCcPlEjbK51DwIaS/Xl6+dWgcXINoGbho45P297Fdq9t1O0kMY0YUR2S1q+y3ZED0N8hD2NBTAv9g1/X7jWY6ahi1itvYXzIaDxYh5nBlOQu1vdlWSOTuI2lwVZzLdWgduynJ/rT7PmqD2D/P1iekivF4ep74kpkwfRWXaNXCthwFxlAvP5C/UXodFYBePRE6vRw+DKNE7KBFbxeMR7SOojKw0KKKUsSsYstNwF6/06h14K02c7SLl4W70njQcUSCm3i2Jei06oNMf7v2GKPDmGAzIyhPBI8IpL8tAfRwhHYzu6d9iyLsGI8Se5fL5XJK8UmMAGj9E+2KOas/KRxRDY4miUPJiuXb1Hk26+3Es4A9LR2RwZXFS9yTYB3lPleacpdiXlehWf7P4Pd1SdXHqmh5nPkQuvigboJj1XnUtdPkShlzgMlHDsdjrLJEOqNHEjSKmamT6nxd2RoVb95u8sNRrI1/sF3MyHT2LF90RY2716guSDYgy1JGAsg7Z5yp+lP3HgQLJhCuW5bRd+1U9kj1McTzOrbDpQJSh8mxFNXA5dvU1SFtRkEbs2T8sCs98Ja84iE2x5UAc8nInqyAubAvt2K9ezJNCuTYpo4abwG023k3qzThwWy9IZdwUhiw76DDnoNCBMpQQXqsF6TQ3iQGLZWiDWQd5FTbLUFJYK0i4BvHBKVLcB4V8K/bXGpQhaUHKsFWGio1UxiT3SytlqNdI51lxQsq2UYaMN62sHusFyXNd8jKkw0rrBWlIdxqSISRy/arlN/od5l7aU/xnIMerf3iSQTxUeUMpORgwefz0vAbFxyyWodWJNZuSX9lkmCF23ObEmorLw35LZeihZKpfQ/WUk2XI6AJnEiJ/RpMrj5rJSA0WYUxrcuVWKO817eBcZ3zI8sbkyo30utT5lPlFODMYUN1DEyL25BkYPJNSpo0fWBNOMlNFuPl4+H+AA+Mp12v/9mPzTP7rdGy7n6WRRhr/I/4FfEukQuxbHgIAAAAASUVORK5CYII=",
    "FaBuilding": "iVBORw0KGgoAAAANSUhEUgAAAGAAAABgCAMAAADVRocKAAAARVBMVEVMaXH///////////////////////////////////////////////////////////////////////////////////////+Gg3WpAAAAFnRSTlMA/NzQzS7OLwIBCTL+gv0Kx9HbxoGA59mFzAAAAAlwSFlzAAAXEgAAFxIBZ5/SUgAAAMRJREFUaN7t1ssOgjAUhGFK0YLcvPb9H9WFicjikBanJib/LM+BfAugQ1V9Zh5rF7+Kq++nysx8joLcbGGMkjxMoNYAVxNwGsCZQBQFAOCnQHPoXqsuNCnzbCAsy5AyzwbaZdmmzLMBa5t+FwBA8de0+IfWhPeRcEyZc1wD0Gg0GgCNRqNxXAPQaDQaAI1GowEAAADIgO1/IQGwfUQLgO2SEQD7rgQAAAAAEAL9ZYjSDL5fAT7K41fApAemIv8T5pP4Y+AJluz9gK9DkqYAAAAASUVORK5CYII=",
    "FaBullseye": "iVBORw0KGgoAAAANSUhEUgAAAGAAAABgCAMAAADVRocKAAAClFBMVEVMaXH////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////DZdiQAAAA23RSTlMABfj3Afn6/tsCsAQI9LF38SRmj5QTIm8x8w/8r9glIwlBS7L1k4gW5CCR7Qry3YCu1m1sRGsoDvsYDTYcwdAMGwaWzrlX1UInYOcymdx/5oQD6xK3fkAHSisRtjXa3h8ptL7XgiZVzdKz9lZlrdFuwMpoq0mMKs9OEDR8oOhITdMzGVNFQ3pUX4Xsii1RuJ4dx6jwjixQF1rlp8uNTzCXeFk/punuCy/UtbpYvz7jgcSlhslnkDif4h5G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",
    "FaCalendarAlt": "iVBORw0KGgoAAAANSUhEUgAAAGAAAABgCAMAAADVRocKAAAAilBMVEVMaXH////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////iVxWkAAAALXRSTlMAzc8uMAHOAi8x5vPtiAaZ1u4k2D1C18FFJvTARCWbQ76J5YfxveuYIZy/PprDvEmCAAAACXBIWXMAABcSAAAXEgFnn9JSAAABHklEQVRo3u2Z6Q6CQAyE8UAUEFDxAgTvu+//eiZqhAU3MbYkJs78nJR+myxdStcwNPJWNt217nucGK369NKGE6OVnT9sc2K0ooI4MQAAAICqWTqKiK1ock3epp9PSUzxopr/MiBBDcLK+kXzE7nDEmBJwpqW9pfElSmAkzzAVwAjecBEATjyAEdbk1ICAAAAvviIQtDnbRYKDQAAAPgFQKNlPkyz2+H4WkA3t9scXwuwctvi+B/98Ur5APwboPbXtFA4PY6vBXTazzVZvSbHx3ENAAAAoC9CXwQA+iL0RTiuvwHUcMWyrfuSKFAAqTzgoAASecBYnWvupPPHpcHpcU/1XpYaoSuaP6wOfxdnufzL4dv5cuYHAvXgBH5xf28IjA3/EI/hUQAAAABJRU5ErkJggg==",
    "FaChalkboardTeacher": "iVBORw0KGgoAAAANSUhEUgAAAGAAAABgCAMAAADVRocKAAAByFBMVEVMaXH////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////GlzqaAAAAl3RSTlMA81MfzIj6ZjOZCwHr2AMw8t1hByiucJPbDF0FLvj3+yKYxZEIt0fXvcJVwQJFK27aWf59G17kgkRoKqApDxS5Do35jP3nUCaVnsTtmlLIBKmKzQnx0vYyE1bcehhbRolqjvUx1hI7X0DlaazvT/SXFrucIGMlBqtt1BC1ljR3WDkavyRIi85sCvx7vlw9ddEhTtM8VN8e+cVWGAAAAAlwSFlzAAAXEgAAFxIBZ5/SUgAAArhJREFUaN7t2eV32zAQAHAnmZ2FGZp0SZM0ZWZut7UrjpmZmZl5a8et/921jpN1syXrnPNbP/i+Wrqf37N1OsscZ4YZZvznGLDadYd1SDP91YdiWfHjOzW9a6NYdmx2UYBnIkJ8Iee3bcAANn0jAjtFlOCJwF4c4CsRsOAAfiIgIsUaACLrdUQEADj01BiHCSAAu7t8wcyigUChov6cFgwFlqNqi8GAvtAEnOGw0zCg68SoVOn8o5deGgCM8Kt2GsFxBBsIXvu7Vlm6UQEvryyHOS8iMKVWcGN4wB31kt6PBYQFQmswiQRESLvSfRzgaQcJ6NiPApwkb6xvUYA2MhDHAGy03uA2AnCOBrxBABpowDwCkKABQQSgmQZYEYAWGuBDAKop/bQ/hLEOLpOBHSgL7SgZaEcBknFS/sA9nGraTgIakMq1K6+e/7ELa0d75FHLXzusmmNE+dkxpLkn3w0o8x96rX6TduXQddpty9m6fyfV3+AwAS4Zda+e4m7dzuECHDf86V3x4q7DY+TnqBtY2XsSvbFYb8JGfVHKAZjCBAwBKppb+UxnTenbxlnTmeEf9FX8GbHVUYp6JeAuXTyjAoTO16Xlz46ZQT4a5Qezcq+avtld3HBCeb2HIU39F2nDeuxN8r43A8s/XvxY3aY1MiuviqVZEDBRmHXhivbQ2QX5MYCAvDTnWIpl7L6+gvAZAkgN1WnG8zl/pQS8AOT3JJcnjMWZhxde2FfM+S3SLeUAR8WkJUZoFqR+bTLNDghVKzMOCCxj3Y3vq6X7mYM8s+vwkuIEncDu+QAGfsFWJvz0aAoG5MBAFgY0goEJGNADBlIw4BQYAP7oSIGBAAxoAwNzMOAJGPCNQ/J/rIQvZdDvsuM62g/vNHv+W15dHc5zD1v62oO6m6gWhp+i1gHz57EZZqzV+A31KonwpP3odwAAAABJRU5ErkJggg==",
    "FaChartBar": "iVBORw0KGgoAAAANSUhEUgAAAGAAAABgCAMAAADVRocKAAAATlBMVEVMaXH///////////////////////////////////////////////////////////////////////////////////////////////////82gLR5AAAAGXRSTlMAiAa3G7jl+/q5HLu6Gs/zxwdtyWzn5omHGwJSDwAAAAlwSFlzAAAXEgAAFxIBZ5/SUgAAAOJJREFUaN7t2dkOgjAQBVDKUqgbimv//0cVLdakxTSdmYTovY83ZM5LGQgUBYIgyAJSXi/2Ga1K+rTVpn1Na03tKmXfUXRg7acZV2lfaTrQ+Wmdq+xH6EBk2j8BkRPCC0ROCC8QOSG8QMqVAAAAYAfC7cQMhNuJGQi3EzOQ0gAAAOALkLJnSEDKniEBKXuGBOSNAwAAAADpZZfXzALG1w2lmQXqZnrgVFtKw/wpIfOtHwAAAAAkM34Pt7JR0sBZGrA/APTSwFEaGKSBw072Vnu8aYgKt/FH12nfC64KBEGQBeQOhVJRBNJFmLAAAAAASUVORK5CYII=",
    "FaChartLine": "iVBORw0KGgoAAAANSUhEUgAAAGAAAABgCAMAAADVRocKAAAAvVBMVEVMaXH///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////87ucsBAAAAPnRSTlMAd4gG4CEe3dwdgOUcB/T5z2xwGH/H8G3zyc7n5sjNa0zxj5nCqCeO+nlfxibvbtr+G14lSXLZfL8kqpqJh0xaVIMAAAAJcEhZcwAAFxIAABcSAWef0lIAAAHNSURBVGje7ZnncsIwEIRlO4EEbAOhl/Tee0/0/o8ViAdcpDsdI98MyWj/Srsflq3zyQjh5OTktAIKvibyV3U/YAH4ciGfBTBJAWMWgMzIARgAe0f7Uqv0mbQCPEtYfgmAxzMEUC8B8CoxlQD45AZ0uAEnDWaA2OIGCI8bgBFKKhUeNwAmEAHRzTCOt/sRTKhZAcJmMjQKYcKpBWBzYz62tg7lp3OWB2S9EAHKpwDyXj0BzCcAil4dAc43A1SvSkDyjQCdt0jA8k0AvTdPQPMNAMibJeD5OKBShWzVCjinekAGYL9tfg3KnOmARwTg154QdPn5ygcD4PVJV0ldn0qh8vVAQFvxqmFQfobwBgI+CHFg/oJw9wQCeoQFgfOnncCs12i04Ob3VuPFCbl8Ia46h+fvSHd9ofNihEK+uX2v6TYu/OjCryJwH9SWKQyGfP1O9uilzZQP1CKPWpyN+VA1fXi5Pm4X56p3Gr+/yzdelJeo3SGQ0gbYnTIpjYzdMZbSitmdk8NRMnkn5DqIR61uHF/uBu5bhQM4gAP8UUDpmn17lLzyuQFjboD8B4ABN2DIDehzA6Im71YT4p6V8D37o6vVHTCWCicnJ6cV0A+1gwn8NrP16QAAAABJRU5ErkJggg==",
    "FaChartPie": "iVBORw0KGgoAAAANSUhEUgAAAGAAAABgCAMAAADVRocKAAACEFBMVEVMaXH///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////9sVt61AAAAr3RSTlMAaf7P/RNjVYczEMv76xIC9QEJ8/AD6vQq8gj87PFgx8GIJvl4FeiGTQ9aDrRChdUoTFFBJ8wN1nkXO+DlRbERFt99G17kIO7U+Ja6xAWLzi5x975cHsOAB5G1WKUkdrlmBKlHzY/Ss/bXuBg8fx3mhCJlqI7RslAxdPqYfkAhxkqNKwwwtpc12oKmRCUGjG0LsNl3OXwaPuNDSK+QFDgZ4sqqoUZuL5vYPSNvHO9yE/3S7gAAAAlwSFlzAAAXEgAAFxIBZ5/SUgAAAztJREFUaN7l2mVXG0EUBuCFBig0CSEJgeDuXijFneJSXAp1d3fq7u7u9v7Fwol0Ezabmc3cDz3M57n3ycyc3b27N5K0BMfbJ5GU6VuSgZM1o11JF4iYl3AN68a8BgLAAtnQXz4WLRqA13jWbKQFAEtqbqBJ41QBwBCk/cAjyooma5EwrAoAPbGaspdkpOidpzmgDiA9I4o3+72kHlmCh34AoIDvsEOyEz3Cl/sF0HScPb2xyO4VzQAgcSVj+qhW26JgFgCdYUz5Y1MUYpkAxLMIrSZoBpC+yV/6VdPKkYwALFfV848UIzAAtVWq2/8OgQIYUrm/Vq9B4AD2+8w/4Ds/D6B75CP/uiwIAbC1VzF/VQ0EAVitlD+6G8KA7S0KQB/EAZhYnD8nWCTQOeKdP7cNIgF0eQOnIBaw5Hvm360TDCDIE5iDaOCIR/5SCAcgLzOi+gkA+cV2DgRAsgyopwB0me78e+MpAOS4gTGQANfdQDENsNMV8QM0QJYr4g8REGx2RswSAUhzXmW/qYCbzicxqIBSR8AUFWDZ5QgoIgJsrgvtIycQzjZK3PXdECfAPa5RAwnUQCI1EEMNrP3vV9DEBlyRhXwJYx11FyXpPeMjNlT2FWAZ8/2i7ZJUAFJhVPoKUsEqfWZ/RMnqzYgVjEHF0neQCtNSvo5UmJGkb+yAhnMolKQJDoB7Dab574mfQChMzk8bt3MJfLv0c2HaB5AJ8b8WZtXxATy7NOsovR6QCXmOSSc4AeZdsjnLU6OVSGh2zdnDC7DtUkyI+0ugDRRrSP03ZRAEQpbsc8KGzRC/S0nyJ3O4Xvgakj2bFqkQvAZTuVcXpR9i1xDkXb60G4QK9YsbIs/1AnepUqkxdRfC1mA4q1jkPRYl2N/4KCOzIWSXTGd8FqqaBF2hbA3z3ZMEtVbOei0n3S1LEPm0TL3d9cLOD1RwlfNpFdxABt8Lw+2+YL4jOMzdKS2s5Mi/o1rDW0/vIOtD1DoTp+3FqnELy7ub6dVr7U3qkNP+7n6GsczA+uDmqaPpvmu2grBxAQ1848FD+5R++43hTEnYKA+9c7/D9bpu6th2Pu8Wxf81zA3taQcazUvxrzZ/AU+g8lIDIAKLAAAAAElFTkSuQmCC",
    "FaCheck": "iVBORw0KGgoAAAANSUhEUgAAAGAAAABgCAMAAADVRocKAAAAjVBMVEVMaXH////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////jNCgdAAAALnRSTlMAEMwT+BHK+U9Jz07NBQOtqKZKBPWnoQEqxKVN8vfsrsLrHMFkG15jh79dhq++y3ZIsgAAAAlwSFlzAAAXEgAAFxIBZ5/SUgAAAYhJREFUaN7tmWtTwjAQRVc0IkK1AvJGFHw/9v//PEUdLG3T7Ca7H3D2fm17ztBJJncLgMVisfzHDB7fp7OOHt/d4jajTIl/dIo/ucx0+UqGAl/FsMdXMJT44oYKX9jgWlhNPtblIz4p87GlzMelMh/vlPmYK/OHE11+e66zv3b8M+Mb/8D4/dl09TqQ2V/d6u3Z6PvSxinxL65+L54fq7yfHZ9qiOfTDCl8ioHLvy7dFDKk8kMGJn+c13UBJ7Y+4b6+bXgNrsfjg6/POCE+LJFj4PPhBhmGCD7kSDfE8GEyJBui+ADzNtEQyQfonHgfLO64hPOLZEg6HwmGxPM3aEg+3wMGgf7QaBDpJ13/au3Frk/ybxDqV3wDt79xDfx+yDPE9E+OIa7f0g2x/ZlqiO/nNENK/6cY0uaLsCF1fgkZ0uejZoPE/NVkEJnvGgwyfL9Biu8zyPHrDZL8OoMsv2qQ5pcN8vx9gwa/aNDh/xm0+F8fYRZb/qKv97fS+vnl4+0BLBaL5ZDzCVInipnXPd96AAAAAElFTkSuQmCC",
    "FaCheckCircle": "iVBORw0KGgoAAAANSUhEUgAAAGAAAABgCAMAAADVRocKAAABv1BMVEVMaXH///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////+2elNiAAAAlHRSTlMASQtN+UfB9/gJvPq7dwEF/rHVrm/WRQRshCw2pxGlCoLYQiZEYXlUI2bxrfUxEkArNXgGqyrzNHyvLjO+HrAlQ0ZtJGvnKC1wszL73Bh/HUEi0bLbF37l7/TaFiDp1BXeH8Ri6Ck4e4BSVV6Fqc0O9td1vcwI8Byi6geNbsDkyi9yEFO6/RrJqvIPdhS5/Cd6UV9grIGGigAAAAlwSFlzAAAXEgAAFxIBZ5/SUgAAA9hJREFUaN7Nmndb2zAQxp2SJgETCNmMMBIg7D0LZc9C2aMts4xC6d57773uA9dJTIhtyZYs8Ty8f1r27+yTdTrpJAgnRtkR/0FZ0OkMlh2URrL5sov8uT0lIqRILOnJLS3iQ2+YslkBKceLpVZWel3lFujK7VxkwJ8pmABDiQWjJvG1n11ApswrJvCrOUChzsu0/6S3Dah0e/8pDT//DlDL10iMz/JawYQmnFlk/K8tYFI5dST83k0wrckKY37XGjBo55wRv74KmNTWr88PO4BR9hnd2HABmFWl46WKHeCgJxFs6NwELloIYcZXJ3DSOHrEvQFueo7iN1r5GRAR8bvZBxzl1mYFe8BV39T8sxf5GqhZVRnIAc6qVs0wwF3K+WeevwGPIga52GDWuWXNNVdeioECNr5tWhC2NXG48Ii/IrLxT8Ugp+3q0baRNDDGgY+wMJA0cJMHX2vhVTJ/ZuFnpB+52qLqh135+hQnviA0KRs75Ms2Hv6Ja1HZ2i2vXxy8+MJHVQKQ+Dw/J/9IKlfd0Be/msuNr+5kKGYKpGr/aAeCHI9Kjo0P7vhaQzw2PlgDUsOv4+MDDEot/WjC8pyVqX8TKpWaosg19rY0TduY+XBVaitDTSBpsafSM9n8I+ma1DiA42MtkPNhSGr9jeVjLFDwoRwxvMFhOXo2PcO8/w8NaFzkTH1aY4GKD0FUJ59XPK/yEo1/5E6O6htQWqDkwx/UQGtSMVK8ROcfeaBF9DpZaYGaHw8VARH/myq9ROsfOdghwjXaAj0/Ea5REw7Kggm+POGgpkxtP1yn9n9yykRO+ppvEOjfHyCsk7YYWCDiy2kLJvHStUDEh3vy3UuYjV0Llm8hy9UOU8dW3F+cxvT+ACOHD9yns0DK/5R8wgs0Fkj58Vid0IZIYYGYL34nWQRqetpCnIunLAKF1g9A+A3E769cxuolwAoL5HzVXoLeVkKKBQo+tCtdq7edluwHC8VaqIVmO0f+Bpr3r9EUFJr0bnfEcsx3NGu5u7Rbaj+evafJ7tcRhbZ2jpuCji5UiBnjZyCI3pj18OLPY0ohoUk+/EtvcfPI7DAP/lovfiYc5bG9X6s3l8/YWfn2eoMSy0M2fpvfqIiT95KFP1xrXIaqWGD4f2ZJCmm3bpjl94QIa42VpraoHF7CUmMsLrnp+etddOXeL5Tb+fvNlBXl6WoafssjEzXx149/EuYP4/kmq/p5hQSbVWJhA8O5hFDllsvg2MMK69GK3Y5uTICyd++N8Dkckt5X7PEpZlSrz1McfsD3gEugty/6b8jpHPobDQ8GTs6hnv/xS+0n6SsRzAAAAABJRU5ErkJggg==",
    "FaChess": "iVBORw0KGgoAAAANSUhEUgAAAGAAAABgCAMAAADVRocKAAABelBMVEVMaXH///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////9xrg87AAAAfXRSTlMAdwrk+eb6iAblMokrtRqvMzwsUDFEENnEdueFcLNRo9C7+AWH7Gvxj9L7qNGY252iAuoH9DDVEf4lKgsVWP2ghinO8nFXmj2KzUwOlNd1uFbcGL2ef2BBhCLHzK0NspN0eRe8WsYmy7HaeBY/IKuwTpFyPkOLTZDdGYAeYUCsph4AAAAJcEhZcwAAFxIAABcSAWef0lIAAALrSURBVGje7ZhnU9tAEIZtuUjC2IZAcBx6IBBKINTQSwiBQArpvffee+6/RyMbI3G7e3foNDEzej/J1rv7zEi3t7eKxSAtJRlLNsfCk5OfsSMhApirCPB/AEtJ5pP2tbQrv/61xDjtN0Doj6h590u+GdVBBKgewPyrKYZqqrAaNP+XBUZqYS0g4AwT6Hi1A9pEgIaAAGuUzn/UCvqWP9GAY4GXaSMNmAheCB1U/j4NlXaaArRoALxJ4/n7H+vYLD7ggLyW3agtvCIoqWs0vCIo6Ud4RSAohQlNAKwU+nTlx0qhRRtgDiyF9Bxozro3C/ecS2Nyxz1pOH9sFtzrVi4or1AE5bs1zuWM117reZtcUINCEZTv1jmXtV77AeePGghgZVLYKrJNSwPApDY7UwMgRQEOawDQDUdmVBFE7X+ATTkTGgDkKurWALBMey91oPhgx3jTGLJxjYvzj/NRTbyrCQGcEgPO8lG/eNdlBPBaDIjzUbd41zsEYPwU5b9k8FEzvG0Z6x7TCTp/YhoIquN9jWh/yr2g8g/+gWK+8sYbeAec9TSoZDGXi3u+Q+RnwZCnPGCe6LEjGf8LjVd+ro/AEZ9VW34lZc59aMT6KembVK/x6Lu/aLd/bWD+dh6wSAJWYcAW5u/kARdJwF8YcB3zF3jAeRLQAwN6sAM2cPJKd1GAIgy4gNiXoYp5TgF+w4BrKjPIRwqQkT7muHoEmYsU4BwMeIDYFyHzMAW4CgOGVZrIHQowAAMGkC2+Hzy+GwTgPgx4Cbvfwm7qW9QVOOSJYFH79ZAA3IZDnsHuu7D7PQEYhEM6YDfSA20CMASHDMUiRdIuYoYlTvAKk68pPskCW7zC5JsSA4BJVmHy3du4om3ErH5AfacM4NBBf3pRVMVf38ok5UXIRLn+k1mmoGx5JJSNcvztTEkn0EMy5l9RA/SWACvyfqYo+VVX9kcAod9Wy5+Q+Ljk95tqgG7p7X3bT3wmItqIZJRtWv8AEVRBwwQtleYAAAAASUVORK5CYII=",
    "FaCircle": "iVBORw0KGgoAAAANSUhEUgAAAGAAAABgCAMAAADVRocKAAABelBMVEVMaXH///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////9xrg87AAAAfXRSTlMABTUBd0Ku+vj5sNX7Ev7WRSMJLbMyhG8XwacrfQousYLYRiQEYVRrZn/1ByAq8+iAHiYnbUN5RM3xcNc33KitLLI220Dq7270ERamBv3ev8Sqr/I43b4lelVsUV7nDvYd8DF+ouUCyzC76cqrL3IQNNk5HylxD/czdntfYATxyuMAAAAJcEhZcwAAFxIAABcSAWef0lIAAAKeSURBVGjezZoJXxJRFMUHGBhmWAWEWFRQFi1TQ9HMvbIySzMz0/Z93/flfvdGJQVkmXlzD3k+wPn/fjPv3XffPU+SjoxsxeBWbluWt3NbE0Ubr3dXMKplFKqQktGiwS4e9+RAXKW6cscfJ626+/qvUVN55GMW7LtnvdRSSqJb0P7Rmp+M6eUNAfvBHjIh7bPZNRkbIlO6P3LCjP/ZKTKtKbthe0dMJQF5R54Z8//2hgS14DPiv7lCwtroa+3/4xVZ0HBnK//zk2RJQ73N/Z1usii3s2ltuECWtd7kK3UME4MWRxuWzhVi0cxSg/11nZgUcdQFvCM2xer521U+gPLpsP+VADHq4+Ha+otYJdf6n7rECzg+WAPoIWa5qv2niV3V50+EH6BV+r/18wP84QrALAGUOPC/rCAA3oOe7wxBdHEfcBoDeP7P/ySBNF8GDKAAqTIgjgLcLN9f3CiAO7QLCBJMex1GFAdIYwppbUnN4ACe3buGggOoWR0wSkDtdNu9SMCEDsgjAbd1QA4JuKMDxpCAgg64hQSMtwMA/0TwnwxfpvCNBi8V8GIHL9f4Awd+ZMIPfXjbAm+88K0jrPlNgtv31+gLyFj7rlBSAgFYrswIroKvsdICqg7tD6vRowRJ4/afa/c4RxrnBfwEj9Q8NvBQsG6o8xUxpWjrYFbybfD4n3vfcDh+l8N/sYgd7092/s+AQo9YHhI0YpGk8FOChkT6n56xsH46jARp99ZE/TWfwayxXyxqlB2G00y7QOXzfDAX966brP/m4t6dE8hl6vx6IJCJv3hisNfwR6YFU/1wwkBHpizPW3iXsIR99rDXe6dWGxSo76upLzyPQ0KltCtQtTfUgCvtDPE+cMlulvJ/CrJc+J0v9WWPzqOev4bY1QraqYVDAAAAAElFTkSuQmCC",
    "FaClipboardCheck": "iVBORw0KGgoAAAANSUhEUgAAAGAAAABgCAMAAADVRocKAAAByFBMVEVMaXH////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////GlzqaAAAAl3RSTlMA7ZnX8knu8/TvyAjWPR22/efxxw3BxkWIB0qNDBZEjBDev+gkydj8Gx9sQi3SDtwYPH9L8I7bF+omETXa/pzAxeklC9Q02ZvEQ0jTD/e5HsP7txxOBKlH7CjNCY/2MhO4N73CQeuJatExEpg7okCDIQLLK9DV3z+CIAEGKm2wtRW6OT7jgYYFKc4utDMUmjjdGb6Aa6pNf8HTfgAAAAlwSFlzAAAXEgAAFxIBZ5/SUgAAAsJJREFUaN7tmVdD2zAUhU2mnWaH0DSUAAkzHbQUKKVQVssspYzuvffee0D3nvq7bWK7ca5jxZGv+uTzFMm653NkS7qyBMG8gtl5pyg6Qz8XBB5qc8wRReFf7fj+iy6ikes3tv+TMClSOIvr724lQOFF1P53EZ2i3xAB20gJOfD8g/86qKa2O6b+nmtDA2QVS2kq/zwkpXgODRBSHKfkYlIpfkADjCn9o5Z9ctmJBlDeoXq1XKuMNja3leMtIikpCIASnZ9TZe23xImhYqCLSmlwhO6/5CEUueVGm2htPAHq/VP9iZTM+0vURlVNFECclJGv+3BzuTZpyvMlKJqsbMqpXMaTVAsOwHgAenEAXkMAQZINsAE2wAbYABtgA/43YGAswxXgTQp7XnMEiKdzUT+4AcSrctgrTgDxnhp3nwsg+rIQeI0DoOuLNvIsOqDrRFHkjWlkgLQVhO7FBXh2gMi+DCpAd/+jaeZnMNw/oN/mHQVx64eY36L6N4Ibbkv2X4D+NczjIJGrveyi9//OGPNITsjVG7X/ofc99Pcxz0UJtV5D6F0LYhqamWfTzsKFgEoYNuNvEtCpvZI6kq/L3AIRFw8yr2g3HxZ/BJj5Wxd5DAKOvWNfMq+Aa/tmSOQtqPt028KaHOmD32YbD4Ca5+ctLfoT1eDqI1CuW2cxqxiqpn67MvY3PQ4maIS6aYS8aPdqw6abt6MkXoYEqn8ls+mp0oQNd9BSx/hoKX8/Ym764pCu2UIPavK7awX82N+DnF0DQrADPX0/riXMdnDYH8wXCLNfuWxAQquUFu2NnHY4CsGcP1Pi9SxHeHCX4x7taYPwcZnrJlDsb7X3yTbABiAAkI5YorwPifyGgHEcwCVDQAoHsMY4HzyJ4T9ISTivn7HuTz0sFQJVlv0D9APlkbQ1/+9NZc/EJx1+xvHg9Tv0z/cPJVlwRcFb2b0AAAAASUVORK5CYII=",
    "FaClock": "iVBORw0KGgoAAAANSUhEUgAAAGAAAABgCAMAAADVRocKAAAByFBMVEVMaXH////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////GlzqaAAAAl3RSTlMAQgGw+P4F+vl3rjXV+28SKwqC1kVGayMJLTKEssECpwd9ICougNgnbSUEVGbxs3+t9RdPFunzEP0av+gPHrEmVSRDYXlEzdc33AOo6/AsNtt+QOrvbvQRpsoG2d7Eqq/y93b8ON2+UXpsYF/nqUdwDva4VhgdxzGi5SHLMHi7G8CrkS9yNBWbOT4f7SnOTZBxM3sZ4lJe0NODjAAAAAlwSFlzAAAXEgAAFxIBZ5/SUgAAA39JREFUaN7NWudf2lAUjQQMQURFRBQUt+Ju3drWXVfVWmu1tXXWWrV77733vv9uo1CSEMgbee9Xz+fDOSQv776bcyMI+wbpvb6+wS1Z3hrsG+5NZ6ud4wuKVRJoIFWJQV8OG/XspiInJIW9aCbbqnpW/TkwhVs+aEG+YiITkJDCFZTyd5ccgIcn5ynky44AAcQPpM+kpx2IcL2yhER/egSIMZKBLW/zOIECmZXFePqHngMlQlk4+pGzQI2NfrT+9zawgJ5ClP6JBrCE9jxz/TQ7WIQ9zbQ2nATLWDG5S/k9wACz1SlLp/nz41zrivK6ZPN9MjafYn+FzP+ZrFJlc+ayLanBK8Sl/1SptxBUTzL9DFR90JJRZ8Q7o/6ZXGBnAIvG2voDWBpoFyyKw6fYGhwoSzDAOL+IDMCl1x8F1gagP38a2RuIWv5LB3sDR4GGPwHsDSCs0lslHgaZas93FXgYwFycfpSPwcN/7HLgYwBTMXYTL4PSGLuIl8GF2PuLnZeBPbBH9gEvA4h2GEF+Bl7cQkprEC2pVfwM3HvvGhI/A2etwq0Gfgaw223n8TQYVrh1PA02Fe4gT4M/CreGp4Ff4V7kaTBEa9A3iW9AdYuEsjvYt2gb10AfD5V8xelEdkge08QXly/H0L/5TbLR/In97NNnWBsNu1R8NiQF4zPriN9ESIod1BjfKt7/Qhc77HIN0maSYOMyslwLIn6e8tiYp9iakQdOEPBxo9WYSiKPTB+BAbQZsogO5KGfQ5RPrFeO6w1aUnNXA0SNV7yd6tKt8mJqZjdZ66jepiuqfnEIo3UsJzQAx4t4OO4142WTte9anH4QW4BOM1L8Oj3kmc3Hlr1rb8Db+a0SuYPjbaTjpmk9ld6oSxUGDljQ7sZO9vq611ghxN5AnyVMszdIiLJF1vrN5HEOEQxxjjDE1mCNIlIjgTudIhQkgJR0qPONncFc8mB2mZV+oy1FtLzBRv94R8pw/BIL/dlevvF+Q+H/HFAoI5ZJ4DpiEYSC+8B1SKSs9JiF5ycfZ5B2bYlWX8Qa1Cmopxs1yjbsaWYGReVzvyYb964Q1n+yce/uCeQiOr9uU8zEH93D7DUcjaOUU/2CMEZHJi1MWfguYZ7vZw/R3ru0O0WBWu0u/cTm45DAgNeVq9sbzlyXNy3A9gOX2shA3bZflv07dQP9tfvno56/cXcMeWf3jIEAAAAASUVORK5CYII=",
    "FaCloud": "iVBORw0KGgoAAAANSUhEUgAAAGAAAABgCAMAAADVRocKAAAB5lBMVEVMaXH////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////es3kSAAAAoXRSTlMADJUE/P4Cmfb3auBAl30BBgv9Gfv1B/Tt3grnKHB/A4nRUHS3Jsv5OkRo7s8v2ZBXgMNaOeOFzbNGMa+KnjwwpMipawmPLQ5RlHUTN9x6vR1gQeaEqOvMCI5vDdb6mDbbF7x+HIOI741U2lkbXoJj6YerSYxO85G1+BWbGl3EYkNnBWxN8tNxUjN2OJ/iYcIiEjvBxrDHsnmdoj/Y8bjwSK7QZnAAAAAJcEhZcwAAFxIAABcSAWef0lIAAAJRSURBVGje7ZnnU1pBFEcv5GF4gBgRFEw0tliisfdoND2m9957772ZmN6b6fX3nwYMEYO893bYvfm05xszzDkDb2f3LhBpNBqN5n9QUOc7tG33uafDewrVy3MCuY0YZ9eW0llK9a5KP9Kpbs5RpZ9SEkUmHlao8c9pgBULVijwzzZhzfSArD6vD7bUtEn6++FAyCcVuAhHQo9lvn8I4F2Z/foxRQKIueLvHQhW+R+Ysdvunmni678BYripojr16kgwXzBQIuiHZ37a4j0m5M/fjGzxHBbZRiohwQ2B/dMvE8Ajx0BAyg/jmVMgVy6AVqdAo2QA9x3OR1k/5toH6qQDkWLbgE86gAu2gdPygeW2gcXygaOsqzTOkO1G1C0fKLXWry2rl/fjrpV+ZLUJFdRa+HvLlehRPzOz/1JEjR/lMzL6FxpQxqork0fkW1CKMZj2MWqvQTHhlomP4vpZqGdNQSrwHBzsPfnXP8/DEsDxm8nABzDR9Wfa6wUbnWOBz3wBoz3uL4zwBfAu8YjByVuiT6yB90TdrIHwG+piDaCNXvAG3OTlDbyiGt5AB0V5A0X0jTdg0EfewE/q5A3EqIc38IOKi1gDfURfWAO/iL5y+r8n7jujjIH+xJF2h8/vGRg7NDexBZJXzn0bmPze18nBZX2YJ7B9fLZbxzJ7nc9LTY/3Qur9rYsmDtjNyg+e/Uv/HeGX7FR8Vk76lXBri6lO37Es0z3qYJOijS96wGVx1dzx8upl6cVfdcpl+7PpkxPBsqlZ0rTxTDvDHzEajUaj+Q1jyskg6bwllQAAAABJRU5ErkJggg==",
    "FaCode": "iVBORw0KGgoAAAANSUhEUgAAAGAAAABgCAMAAADVRocKAAACClBMVEVMaXH///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////9nulMrAAAArXRSTlMA9AQCA/2AAf77NQj88IgG6AXT7OEd5kYsbzYXBysW85FiTI/SVtw8wqiy9TGYX+XvDLHpThD4m8lnkKSpcCJqEvpAISYReKGHyrDU3hqgPqUucfdXexm+HkLnR5lbf2CjQa0Nt1XgfsH52n0/RCVoSc8Ld7o54x/EQ6pI7c4K8g8UuZriZgktsxO4vWWJJ8x5p+qN0E8wtrucOt/AXqtt2XxdgYtsTdjdXD3DYUhFAo8AAAAJcEhZcwAAFxIAABcSAWef0lIAAAQTSURBVGje7Vn3VxNBEL4AaRCK9A4hdKT3JkgVlA4KCIgiiCI2RAFBQey9997b/Y+GSwg7O3slvjsfvJf57ebb+Sa3t7vzzYbjPOYxj21OK+6d9dGS3yeY50M6h1c0S9DBC6a/H5DmrUmCAN5lgxV+GiRI5gl7oD5/GMnP71I/wUmQIEL9BAUkv079bxCbQCYoUP8F9oMZKlc/QQVIcF79BMdI/mT1+Ru0XqTVIEGH+gnOkvz+21TnP1NCJsj4dyKxn1YEZiiH2B5u0cdP6ivZKbJJfn3MujswwzQXo5x/r72i8IeYRSuUTFDmKkEZ9qc9O5Ty1/8QwtuNGOoCM2Rweo2rwmN4pDL+uFxnfCLGmkACX6f3i/O5cLsSft8QF0EfAtNJ/kfOgmnYWLZp8vwHLxEU3RTYaCYTtDucXwlX1gU5/qIeksJkhWgvmKEUwfdJT/qOjkvzH7AACt48DeBmgDWuuSKoCMsfKf4oE09ZJgl7F5JQuuDLpCNM1eL8IzyyIyQeDaAmwTeFY0ZE6I02PLbUixxhAFiXY49l4yibkSkIX6KBJgMc0kKC59ZZovQocNQL8wcuoGHmO9QJBZiyXf7xHhTaisTGxcNoUMIpeokBuIjYO3ko+HY+Jcjn0JBwtO+rSLiEZIgLReFlteDlT6MBu/tpfi/wOy8DrG0JESzFb8D9HxC8rxh9peNgQBQE8/EUfw9bB+/hF1zwk1TtdjtBv18pIgltcECRuQhKZDUWweSIVLyNDIgm6JmwPRdpv6VGXrXbGCOsO2mmvFccN5BFe3VJzI1eLq/ak3SIq467gt7LV4lqZ4sI3yCa7TU3oTBBrE6BascJJri6BGVTFKFAteMpyhoANVjqIytQ7VYzzeSo0XGKlqmsamcs03WVUZ+KoHd+7qr2wFF8mkUCJSd9VMip9rFbiIJUerVlcoedjGpvY5yWbeSA5d/Sx7WMao9Eu4lvuUkVnOeSBeepmGoXbN4fT/EY+khVuGRedaFDbNXusLfoDOIzGaexz6pE0WerdjExxVd6MWXLY1HZwlbtYnWATzSKKKMpMeEFVXu0uFYS7LO4tHuP3rYVq/ZCsM1voHntlhKnOdRxYp7Hqr0ZRExTERartLz+xpLvLNXusp9QvqfINQh1jAYEqHZLLRVR41YDwmqhvEEZmUQRH91qoRhNIFTtMzjirhNajFbYxqbCNhauxOuMHTTkVhuLGnG2agetz0O3GnG7Vn1DXCXEiKh2cLXRaUp34yph7ewj9oaYahcLcNfEVbs6JqXaVTFJ1a7yXTtDtatgMqpd5bt22xa8a8/8r3ftrVvwrh0u0jD1E/RpfNfOXasY3EgQoM3/cnEzL5za8Benla0Md9or9RNt/8CcTVnmPOYxj21O+wurLwBOnsTkKgAAAABJRU5ErkJggg==",
    "FaCogs": "iVBORw0KGgoAAAANSUhEUgAAAGAAAABgCAMAAADVRocKAAACx1BMVEVMaXH////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////b6qweAAAA7HRSTlMA+wM4/QpNAf4CFfUH3aIM/PfRzcz5BDz6G0Ro7j4S+BjStQUqZHZMLVGUuPAsDSEmaRFUFrtez/N3Od5dQ6+AYtjjJPbX68bo04oI9JWWfGsJN55GtzZf5UWs0JLV2nhZnDrfwCCmchBY7S4PtFKTF+rEqUkL4VWbqBrJ5rlWJ4ZChWYos3UT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",
    "FaCoins": "iVBORw0KGgoAAAANSUhEUgAAAGAAAABgCAMAAADVRocKAAACf1BMVEVMaXH///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////8Un7H2AAAA1HRSTlMA3ewbAvnM7vz+Afj2yNr9TjUMEQv3MtkPkXX7vet0g0qQDkTVGkfw9fqsoAprCJ9NFFC4rbdaECQpub5bvNh3MwPEUeZVRe+2KrBTWGftcVdck0OvgNZqbV6iI81MlFYYHYTHZQ0xfhxkVLvKm13jyZoehfNvOEIEjy3ScLPXmaNBIqhGJ9GymDbbF+ppsXOXWZw6fcChxYercpbeYuiGBUjOtFJ2GSxmPPSk56koCROe4cJg4EDGiAcr0JJ43z9j6SUGL7UVOXy/PoEfpary0+JJFo5YIw0AAAAJcEhZcwAAFxIAABcSAWef0lIAAASfSURBVGjeY2AgBTgylxlFFE8zZmNjZGRnv3KFnZ2RkY3NeFpxhFEZsyMDRUBi5QpJp1VX8ABWlSh7g8tkGS4UGyR+hUhgtydWiDTTM6SsWa+QBFitpTKINZ3F2Zb/ChmAP9uIhRjjA7uukA34zLgImb+c7QpFIGcXfvO3iF6hEIhG4jO/l/MKxYDTF7f5DppXqABYxXBaEHKFKsAEpwVM1LFAduAsMKGOBVo4LRBjpUokp+FORr7USKYp+DJCHg+l5vME4s/KnryUmZ9sSagw4vIRJ994uwAuYspTASsOckznMDZgIbZKYJaT4SbNdG4ZOWbSKjURBcUwYk1XUVQQIateVnKIlHQqIVDplyobotW3M0mzhUVVUEBXXtuUl42PkYmdk5OdiZGPjddDXz5cQFAVM9BTT2hcYaAdMAz1A/qLVqY7qjVogAOOJqbrqUm6wGKG6obrrOw/qoEU9dQ0+4zarNwYabS0RbmxLIKXbCJC9W3tsJeNONu6gmZSklrWi/axswMrBlZ29jmLrLUUpcwElTCbT/iqb3VsZYP3fCsX3FpcrOZ7o5UEwbjLxBD0dnRWIR9R7cLCLOQ2tJ4kjtpJE6XZst5mMQmNOdFaMz2EXiOsXQeOuUjGW8S5k1oou8dZIOrwHiw1mitS7aVFXjN9JrzWOuKHLjn1ENx48wnS5FZc0rMlYFGNWjVVIzXflfMpqXovlkONCUWy1sMAqcKM4KCscufQhZizE5paOabXdCMnnr2Ut3/sISbZACMlLDdLGDXpu1KjgbUEbNQCqWAJtEzrxSB8kBpNxHnmWLu/U85fYVhKnUauFKbxRcWgrMGQQx0LCjBqNC9I0mFQp44F7ijNHIN6RI1GbQsMFU7LIHfvqBVEvBDjz2L03KkVyXIQC7DUaObzqGH+GnMcFgBrtCVUyGj8Rgw4LAB1Y+2pmAswGtqtINFwCgs76V4GHBbAarTyTZSYv3UbAw4LNHzgjZQ4skcmNIOQGzIoUpstkGTWnSOvq7S2FaWEQJJqc0Vrz0cHsZNqPHtQNFoRBI+W9lhYjeaN1LoJ8CKha6wREo/UbPFGskB6qj9SjXZlYypy29h5Qh0xptfNdtZB7l/pQ9ugV/gvdLhKoHmK0Q3Vm8uck0zxtALyTZOMlqHqUJsDa+S2YDZdQXqsyzArI7F4/8pmWzZGP3ZgsPGw+zGy2TZX+seLYY66Hj+Jt5kOcVbfKTIb71z72wn0A2A+F7c/Rrrx5iuqCXY0kMoNr8wEUkxPyOzjJqIng1oqFkxpJGpoO7UxvYCfuK4SZubc3XFg8gY83d/JmfVtHER3lRhwJEam6dozamyUq5oShYHJWkI4sanKzcZ/hvZ2XIOFo6OLo6OLo6OLw3p00daSGqOL0Ytxl0Wi8jqUji6KJHHjLezUJRMpGZFKCJUlWJqyNgSTOVlbobxak7jiWrawZQGpph9Wy5UlpT7QcNqhrEe00x3sa9XJqHBEY/RnuakSyDFu/doxohTVaKLi2ZPSdScu7LQoEhYGJjIdYeEii86FE3XTJ2XzcY/WaKM12miNNjpfRl6NRvKsBmk1Gj4AAHnfSiAHl6KiAAAAAElFTkSuQmCC",
    "FaComments": "iVBORw0KGgoAAAANSUhEUgAAAGAAAABgCAMAAADVRocKAAACeVBMVEVMaXH///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////+CY/WAAAAA0nRSTlMAEAfh+/HzAf79RAv8BAIIeBLsdeZv9TEM+U33Vxnn8N/yEbu3zqSFLRijDVCIyyu27i/4GiQF7V3Y0dqHU+OqbNL6yZCzc1EOE1bcvVs8woRGaq2OLLJVF1o74MHq7/ReoSCmBkmMKm2wNHygPmJD6C60ex7DIg/rSyZpnBvpyrofZvae5cZFWYtCqUcJrnC4N3odYMdlJ8zbeV9kp43QbrEw1TUWOsA/5MVjaE6RcrWWFVi/gaWGZymv03GVM3a5mjjdvmHIUhRcTAOSnyMyzz3mOJbuAAAACXBIWXMAABcSAAAXEgFnn9JSAAAD6klEQVRo3u2Z5V8bMRjHi5QKztBSHMZwHe4+ZBvDYQrDYcM3GK5zd3d3d3df/qJx0nJl7V0uub3h09+rNpf7/q65J0+epCKRXnrpNdcVVf1Svr1MIX9b7yg0WhKyOuPShAyoFe2zxm6xpUB009xxW6BN9qENI/jPvrXFBOiW7P5+rOGyUBwGXBLfdEF+eqU7gJF9QgAS36EbwMr2mIT/6Ng5AR7aHcmT77IE8NPgFC/+girAV2YNPPjrTQCCgqD5F5H4ANTBho8HQNQ8KL71O1Q+kNXCTK/XAF02+7gN1gEchXLyP9tgGYBELoMteHyQZ87xhj0wDYCS3cAKlw+G2A3isA3AQjZ+Hz4fWLEZbBLAwIvNoF8AgzC2xWdIAAPQRZCO1ii0BayREAbnCVIWAEmdi/4xEAthUKjOOB4Jsy2MhTAIJEib6RIwxVz4X3BZIx5PnWUauAthEK6ZEiJSGAZeAvClBgSpjtEyf6ZMzhDAII8khTKblhuoDD4JYHCCJLlqtPmoHKzT8A0qCFDkrMYDprRDJjbfOYfgXJvdXEQnkGxsg4ck57auksYiFjeG2gnMhoh/LkSEUw5fMA1KScotLVc+UoMkuYrFHyPXsyitSTOGLqydcQz2kowYrdfiv1EOdhj85VHkvtSItWaSzEfmu1L7nCAdlzPpuWAZj8h3ayLvL9G1+UqzpgyqndD49q3Uxv2czh7ZlIE/Gn9jJXU7ywhnUGEkQ+LH0gXXK5Y+w+Q7zkfib/Ok+AelLJ3ciR7LUPArC+gAqTBjjYLpHj0IG0DZmws03459eI2nQzRdY+8L4yYtdVCdDpzk6LpKJPrBvDP5ieP1VK703xKiPh0Y5nqWCVHKzBePfiq/tq5ZofvhJx+rV1tR7wrOHzspGqQ/OeV/mDlostjzIFhLRZZUs7qccbjhL+UezXHRd2M31/TkzkLT2TXl2vZ7vk+749zDpM5hVc867ljVlmtcfg5Vso0inlxZLoMsmn8i4f2sbCHD+TcC3VDRCF8wZ/Gld402ruIxH8We8OiAkhvNZ2x5TvgELurXJB8v7+RMr+C4aJR81baB87m9sQqOGO6BGdiIwS+CGfpeGTL/CNwbzkLli0MgwycI0eDKDtgAffRfzwsJHUfb6trkwJ/atiE5+PLIbh0oBmkOPJJQgSuCQ6ohD4cAX4TiYwuvTGowj3uReZ/bvJK5gm/ml6zNlUvYJrb4Rf10J8+ljHPYsV98F4SB09722gvVYrlqxA3LvNXVXirKkla5yz/YjVGrxe9cquzR7HM3sDidnD0+yP+HGRxaLA9MlDf1+ekKFfORP1PZ1vr/a/XSay7oL6FSunc7oD0FAAAAAElFTkSuQmCC",
    "FaCube": "iVBORw0KGgoAAAANSUhEUgAAAGAAAABgCAMAAADVRocKAAABfVBMVEVMaXH///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////9hN2cgAAAAfnRSTlMAoogr/jy8+kTEZvV3VQKqrC4N+PxD55kBs8ISYgXMMCQdkRXpBuGCTbsK9xDN5dWrA/3UZy/ka/HSwSHvDCXKv+1x2BTdGNsH0HPyBAkIPzPrWi1CTA77J2q3F4PqNYzzm94ahosPOOKPE/Am+aGH2aDJzmzTvkux9FKKjeMvjh6ZAAAACXBIWXMAABcSAAAXEgFnn9JSAAAC7klEQVRo3u2YWVvTQBSGp2lq09I2UoptKSq4gNCqgID7giIiKAi47/u+7+v8didELJnOPifPw0XPbSffm3knOTMpQjo1vLi8nFpAMdXun9/rmNQP/9uFGOKTzi78v7bs2A6bXsvuxVTtz+0EU7PHT2NGdVXzW6HV0GWtiqEGUBVPDZAqoRprVQpqLFSpqjFUpaVGW9Wt04clER0dkgEnD9Ys1GzLFi/mh1wjVXI1k86JcGh/7qyuqmm5mkR3oTm+0J2Qq2o0u/w9yWi3p7dIz7jY2yNRNX979t+MD4gHDgz2sRetb3BAfOWnmWBY4726Grpkqr4EE3+rp0ZPVRmhFVdbjYaqsdfohYkadVXv0HlDNYqqPqJ5/gulW6wXcAm5Nmrkqt4gutfY7uFUr0pHAOMViFNIZZwLwOfsDzuNbKQxUADSa33PZhqXytQatABIncnVzNKvZlufeRYA47qfMrn5+4woNmD1eTquk76Qz7BzuACMD5UnVOOvO1O8FAFgdRrD8vRZ4SYtBmB8rJyUdFLnjjBABiCV4R+nRr2qZNtUAZAzgnOUFf8o91x+rRKA7BtDeeqzrOD511SuVASQ2uf0N+Nncg8VL1MHkLEjYRepeCNp8ThDQNgMqWbGnKk5gDTDLtGvpeqVUTJHG4DS0xYPYN37EgMg+saDA+ieBQtgdF1IAHPfAAPwdj4gAH/vhgAITx8AgLlp0S4EABBvc21AG9AGtAEbANAZLyCZKEUAJVhA5+bogb6OliABdDw5SaKvcIDWeIxvojkoQNT9Wt1AL1+BAFh3T2rqKUKfAQCcePL3SvA98cEWwI3HD8JvxSeuDYAfP/Z47b+tZ0eMAfx4fOruuttIbTICsJ+c8HPBo9ZJD6Edr4swiNdDGMUHdVlxuQVLuyj5d0Z3uTXu3h6hFG+OUI43Q2jF6yO04/UQRvHqCON4NYRVvBxhHS9GgMTzEWDxbARofCsCPD6KiCU+qInfmck/mV8raCPVX5u58Fwof+yiAAAAAElFTkSuQmCC",
    "FaCubes": "iVBORw0KGgoAAAANSUhEUgAAAGAAAABgCAMAAADVRocKAAAB71BMVEVMaXH////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////joNC+AAAApHRSTlMARMkCr4CEAVVg0d8w63GqkVZmA5kG0/e7zk0zb4j+d+yKMsLMouIJ+/3x9fn4EtoWsdIT5icXp9AE+rM8CEuyWuoHkirz40hSGQ3bEcTlUTEjKPReJQvdefw9yGjPKalM9jfcvVueatbgIctKG8o5H+gFiwouD74ezSDncNeJaUIOfx3wLJh+HMFAg8Ym1ZzA5IJj6e7UENmggWKGJO1s8p8YrEfm72sAAAAJcEhZcwAAFxIAABcSAWef0lIAAAPuSURBVGje7ZjnQ9pAGMYPqAIOQJC6oKhUhqioddWBe+/WvWur1jrqaGu1rXbvvffkDy0QgeRyyV1C6JfyfoLc73mf3HuXy10AiEUs/mmUx1e5XNNPhqOT/cpQ+jEvFcdNJ6TOrqhNPe2lRZo286SUpYlL9LLilfmZ5KWBQ4JSwaWBI8JSlX9tgxKWbW+XQZds8SJn1dbQUyhVodZjAMDg0RZCDb0LIrqxtwtlkS/1BNt6luRw40Oh+YuXmaXJaWG2t+QwS3VW6HDfp4/k7KqBTRhWZ9No0C9h+dcKQsrJgx4uqudgMoQ5OgUZfDmU9e3P8IMz+32H6IIgg9sBTcGUAY8apqjejgkyiKfuypmuwkxAhUzroNg4MQb+IpmvcmO/LeFFSqxBoBvISill2moaJt7AP8stxTDxQf+cyURk4BtvrUxJa89KPQMTERr44qPl8Gm9pU9ANEdu4PW6FucBmF90Ma/u5ktm4IuMDOb/8T8TIFFKA2bk6/Z8cNQMuk0jAThKBr2hpyMaBml598Kw9AZlcYzXi+QGP6AFUHKDJBAziBlgY/BzH7nBxgPq1dAWd4Esu13FuV1nG9SpasZp23l9F8F2vRW7CiWF32s/Wdt5xosPjjuWUZJVlDJYs7iRrW3mb2JLQzNgloZ18mGXSnEzDz7JNHIb1L7Ix8Ctnx7TSzX4ehkCspuzQFZzNtpgkwhODM2qTq2T2eSo7zAGWowd9Q5cvXhgZ2WFv2EYOoeVaiponavQlPKlx8C2QZ/xe/qVZF0ua/Rzdcno7ARw6RaYCv+rLuk3IieYsb+kGs5OCOuBNfgzQVPMd2bTMHZzxPBLkH/Y2wHsGjIQ7L0Q+C6gNG5NEVZTpHELh4+B8ClbxqvISnXRj+TEMKAV9oaF68669NPwIJPCANpP5SIULbpW9OaLBAbshWqEIahTpfOtalgYILbMqZfD31vM2Zj9NQ4GSJm1ocknaGqwkqze/DDgUCXnTORAK8S1sesiYOAljG6THahFwIQG1GFALQImMXDWXKRGUS0CxhsU5p0KThO1CBhn0GimfQdSi4D5DUZNl+jPkVoEzGtwTsFcBtQiYF6DI0CAAQccM/gfDN4SaxQp572C4U1gJdQoUo7yP8Zo+BHQE2mUmRm4ZQgNvwH2brwGe/ecsNvu+/r8DqMhSo+GvwfW1o0dPg2UXi4XAO8EPxrL5FwaqPZVKiWQTRLCCSraKQptUcJOTwwz0vtHfm6FrWGcOtc9YdqzjoFX5hDHZZmcZwiroBviheG7x6uqEApOmDM9p4pLIQjmVvEpBMFoFU4hCGarSBSC4NCkba+0Fdgq25WSw7GIhYj4C2aSofZp0MtsAAAAAElFTkSuQmCC",
    "FaDatabase": "iVBORw0KGgoAAAANSUhEUgAAAGAAAABgCAMAAADVRocKAAABsFBMVEVMaXH///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////9pYNpdAAAAj3RSTlMAEfzqzPL1+e4FTbXUaLsBU91l0DQYF8sQoYfc5g1+ZBac7bSBiC3hAxrEuWtCCWAxogdKl17kIOmGi697gPob0rOeHSdqk53BJpLKC/P4FbpnqgrDWVrNDvYyuJm9fyzRVRxAKwy2VH0/Ktk53nzoDzPYn6SCBk8EN4m3MP1c4sLr8G+8AvTVbZGWlRke05x3GIAAAAAJcEhZcwAAFxIAABcSAWef0lIAAAIxSURBVGje7ZfnV9pQGIcZiUkgYStDxQki4qoLF2hb7aS1VqvV2tYO69574exy3H+5p9+8IZh1+eDx/X3LOeH5nQD3fZ8YDDLZTu5sxTbWpiaoEecCyyLEsgvOEWpiam0jtrWT3DZoz+b+4e5e+gDdmoP03u7h/qZatn1WWOcTSHES1+vCrGL66GRjHGlI3D85Kk9/YBlkkOaYbC0vb8Xndx8jnTn5c5Ud/zeBCCRxkS+JPz/9iQgl9dueyTeeIYKxGcX8y1+IaHjR18RFEeEccViBgIjHhRVQ5Au8WEEe+YI8rADlIFAABXetIOcHLQej4gVW8IZ8QQk+rodI859yon3Mk+X3lGaszAGS/BqfhM3FaFJ4us8u6RWl/QwJPPP8dVYzKug268WbQ49k1LFRx6F71dvyVl5Pix0Pi7TQi544irNCO5P49WOX363iB2Hc/pJnOOFLJz7s6MUMLebC70LWKpOMT1dZA8NhTvzZ78FUxjRdChZIPVuk2mMRKssqbF6q0Fn+/xWq3FlIeW0VZZWCxVMdkfyvfIxLjmtmcN5AIGMf3mffB+7pb/ron4Kf5RbOavtMnTZ43dfAnLKNxkTbm4zq4MamQJRRtzJbxzvqV3wK0J76jppWzTuZ5q3NXS5HbUPb8s37fW0NtT9cXc1Wnia59E0sazazrAm8CAqgAOwa7BrsGuwa7BrsGuwa7BrsGgqgAOwa7BrsGuz63tu1OPrtWkH02LWKaLHrm/kHa2TV6Y7E71UAAAAASUVORK5CYII=",
    "FaDesktop": "iVBORw0KGgoAAAANSUhEUgAAAGAAAABgCAMAAADVRocKAAAA5FBMVEVMaXH///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////8fIpK+AAAAS3RSTlMA6+qMAkfsqlX+SPwtnn9BA9F0nZzpBu5yBa+0n6RwDlY3ejzCo45vDTH62zt+wWluc/k1eBY6wD/PEFM02Vj9OUPobNMPMxQ4gMPjI/2zAAAACXBIWXMAABcSAAAXEgFnn9JSAAABQElEQVRo3u3a507DQBAE4I3t4NjpQIAUeu+9dxLqvf/7oCgCRY6laI8bpCgzv+39pGuWbi3CjEeC1c9W6Dit0l7wU/+5YyCplnv1Fx4NKPmjbv2HyMAS5URuJgwwmVeZNNDMyRcWOJAMFqhJhAUiMeD8JxC6O3tCAgQIECBAgAABAgQIECBAgAABAqMPjOqNF/xS0MMCL7KGBS7xV8tbNWR9b1FkCXm9f9vdFZt5WIOi3dt35Sqm/vXVb5PoolQc0vNpJN9uDHmhuP8UaE6XOAlkHXfCppLAtNv69YFOz+yMUyAenMIsdoQcj1E9pRfmdIzitGX+tzGqfOT60kwDmv1PLFd09d92tdv2fltT/9yiu+a9K4AVm6PnTAHM2wDrCmDDBvAVwKnFR+ikoJnlY7Wwc6hbp3e+7h8BvyDM2OYbNCThkR+Hj0sAAAAASUVORK5CYII=",
    "FaEnvelope": "iVBORw0KGgoAAAANSUhEUgAAAGAAAABgCAMAAADVRocKAAAB3VBMVEVMaXH////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////49NrfAAAAnnRSTlMACfr5/iSAAe5vRPOR/BD9Be1CBLOZwYixJTTeI9cT+x3CA9YmEcqr+Hc+Q6rYFDjdvj1m7EzxcBjhf0Eix+sIrfANVTbboiECRQfvDPQwFrvfwOTPU5s5GuPOpOdHzdIOUfYyuFa9PKPmhCcs0fUSmLzgHOXqyytPVJc1nDobXoIgpumHsE7UtZbZv10fhsmLr03y03EPtFIzmlziwwYVdiYAAAAJcEhZcwAAFxIAABcSAWef0lIAAALxSURBVGje7ZnnWxpBEMaXfjSR3hQUItKMiC0ae01sicYSTe+999577/3+1mBk96jHnew8j0luPvFwc785luV9ZxaEpJBCCinWSmyeb5hkqcZk9Xcdwa9zsCAx417hRxkWKJjkn+cH47OsWZUq0MQChiP1/bKgEUM3YQvYUQNsgWqkhy2gRyxw/AMFjLB8I5qGLZBAD2ELvEYTs5D8pQmEKlxwfFfFspxqtHKgL1i9Ke040WMQ/EucpaEt5+jza7dnmrJVbaKLVxoqc3z/PVVja7uW31m8CtHjz6o4rmefJv1K9owW/7osjdQcGkIM+6gOF/tKZZmYOczzPWWZZTX9uYDfqaktn/+mC9NueLFc+yMaspuMtHaPRqvI8AMXWaaN28rhP9iFOXvrsw2noxFfWW9bPX/rBkz5lsh1NHkEZX828eEn2oOaFQUss76CbN3h1fCHPfj+ur7CntxxBGcMHBXP555vKFHM9P3qAE66I7Jf2q/lVljO01U0EQk8sSSG/+QkEeaD/G1LTyvOHDMoBfNHsDYg3c5SfZHSMIaTH7cIw/cP4jsq81W/QOO1243zLwiarbpHcf69bmGd3YeX+I5ApKRdG53EWQbHBbeOcbLlGr38/MUozpQZxPSm00Q5Psd5+xIfzjvwXFzzq/iEBVbTXnSZFO0kSRsW3V1zwuUpskyLtwQIJE/7ftFC7v9S6LrNJ0TikTDzaD6dezGsxcsT4DUp/gEkVIUf8sqenOFuSqDNlphwxq+SXXg8wyXCd4k2tJrLG6GUb63kzOR+WgdMO8hv3epUlj2j/aohPZSq8zLDhDq5ruq2g8YQ2P+j2DHTXAulKTN+phB+wE5vjD2vy+dbvFTn5HeqbHxXH+1B3GRb4PBTI3KISf+jU1fV21sVc576v84qpAJSgb+gAPCxZhv0wWwQzcMWOIt0sAVSvdVhSP5MSttHe1jYPyhQ0gzGT674k/sFDL+Js9mYPUj596AP2i3S34dSSCHFmonfS8GeRXnvHuoAAAAASUVORK5CYII=",
    "FaExchangeAlt": "iVBORw0KGgoAAAANSUhEUgAAAGAAAABgCAMAAADVRocKAAAAn1BMVEVMaXH///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////8Qg5o/AAAANHRSTlMAZwlQX1j+/WD87Pvcdw/r+tu2EDnxf7fGOoLugQrytDiAuFY3W8dV6u8R+dpZxU5TXe1c9hZ7igAAAAlwSFlzAAAXEgAAFxIBZ5/SUgAAAWFJREFUaN7tmcluwkAQBV9siDFgs4adANn39f3/t8VAJDgkClao27zzuEoaz7S6NVJISEjIXnrP2XsP5L/VbU9jjD9ue52kRgm6Nmuofwt8esIIbNhg2mDaYNpg2mDaYNpg2uBfDfNJMzWSbdWY3xjLxjAxmKQiNUmBG32lqMBDsXyf8QJ4i7r0T15qRh/TaARfNEWrDloquGK33h9ScGS+aL5ovmi+aL5ovmi+aP6uff+T/1DJqj1wAInvi1X18qNW/8ARqpZs1rWXpQ1xq/juIz6Mby/Kb1J+mV3nh/LdQk7Cjm/TfNN803zTfNN803zTfNN80/xjJG3OIpK/zqgwVBKy/V1p3ED7944u2AEk1Tk8Q/GCIb1FffYnP9HHdABftBf0KqedQbQtRtWULXY/GkQbRBtEG0QbRBtEG0QbBBuY9v1/A0gpw+MV9h4WTwv+3Sv4opd/Zre5QkJCQvbyBdWfQ1a+fw8zAAAAAElFTkSuQmCC",
    "FaExclamation": "iVBORw0KGgoAAAANSUhEUgAAAGAAAABgCAMAAADVRocKAAABOFBMVEVMaXH////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////10QvYAAAAZ3RSTlMAAQc66QLo+f7HA+qIVN85UQjg7C0TLCHGIAaghu3TFOH4xEcJ0pR1uHo8YEatk/p5QKyHbVO6H8mLbPe5n+JPd+eKdDvFpIUjKDd/61+Dp42SoeSmjNRyllg+gaUpdlc9gB5h/YlzMiPeDwAAAAlwSFlzAAAXEgAAFxIBZ5/SUgAAAi5JREFUaN7t2edOwmAUxvEXWkrLUAEFcQ9QRMWFe++99972/u/A4BcH79No8pxEE88FnF/S0vffUKW+MaMHYbdswvUPijTXW652wuskwHHBLHP2r7lwUhRgAAMDFMDAQCUFqMZAkgKkMJCmAM0YqKEAfgzUcX6nQbQ/SHrQGhGQIAFxBFSQgBMENEmfRbUkYAQBUySgFwGDJKAeAdMkoAMBBRIwhIAGEtCFgH4SUETABAl4REA3CYgiYJgEtCOghwQYskn2iHKSBKRkk+wR5RoS4BdOMowyK8kwygkaEJdNslJNeiBLAxzZJMMoT9GAXtkkwyhP04CCHtinAUOySYZR7qcBRdkkK9WtB7ZpQFQ2yTDKPTTAkE0yjHKSBqT1wCgNaJZNMowyLckgyrwkgygniEBcNskgylki4MgmWamcDsgRgTMd0EEEDLt8v71BBFREMphvcz42+XH75Nit+p//+YUTujt2zL4+0xnPhwTWZ27M98fA7Iyx90dnPj/HLVXU9b77ssPIjgR4+wO7utP0ykcDVvVvFZe062/rAZt0H2Iz6L+KwwwFmMMfKBYpz5eJAYvxS7pwPWaJAOx5AQuSXw9K00YAwl6AJfkNqjStfwEwpS+R+E0e9wIiBCDvBVQKHxXzlOh0YuCUc1y3wG+Zfk4QqlBwZllJ29EDm7yXiiPd/hVe9JVv7kn0taV0H77c6fisIk9s0frw+3/xK/4ElhaWrWCr1fZs/ODqvAJ9S0QCYYSOSgAAAABJRU5ErkJggg==",
    "FaExclamationTriangle": "iVBORw0KGgoAAAANSUhEUgAAAGAAAABgCAMAAADVRocKAAACDVBMVEVMaXH///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////9xxQJjAAAArnRSTlMAAQUJyczvqv4I++v6AmkM+aGCFf059/z4y8dghfaZPEGiFhur85vefMQK2Di+XGEHN4GAw+cEKK6PcA4y13qj5gOt0fVVNts7X0DlIUWsK9D0MLZUOt/A5CVJzy/UclO6XYYk7c6vcTPiHmdlpMhm7GtM0lGUE7hWvVvhwoQiakuOkzESt3kXfoPGSk8Rl3i7fV4/IMVj6YfKaAuwThA02Rq/PmLoSJDTmt17PTXWaxp3AAAACXBIWXMAABcSAAAXEgFnn9JSAAADfUlEQVRo3u2aZ1cTQRiFlxSyoQiElkJQEAJKk6aAdKSqiApKU0CkKNh777333nub3yhSwszutOzMnuMH7te57705m53ZZ3OiKAtaEKKu7EAgO9G0+N/fwbQOLTYnf3sEmFVEhRn5+QDSSfn56eFwQXi69IIAQPRDdv6mZLQg4Z3kgq1Ao2VhUvOPA50aZeZbt+gLdi2XWHAdYJQtL3/pRlzBmt3SCvYArGpk5Seq+AJV0pkUZgME2aOlFPgAUa0y8mPWkgu+rJZQ8BxQlCKefzaJVhDbI1xQBqiqF81fAhhaL5Zf18Aq2LZSqGARYOqlSP6dSHbBzqMCBd8Ah8qN56+K4imI2GC4YDPg0n2j+QcBp24Zy7fk8BY0rzCGimiK/4ojqGN+dG3MSH48gnLAfQlerIxFFsPjDRSkajgIXX2CrqaKo9xpdLlEA3rjwijXgS43iIJeo/ZO+YmuZ2nXO0PLd+pQLhI17BcEvQf6ZxdqiNUZfominBOBVUHQq8FsV+Sd5hrGkMafn4tDuWHY0SMEeniUq4Qt40Kg14o90drYLDAkhHJvYM9trKWUD/R24M/kfNgzhPec4clfR0C5c7DpAN7jHuYoqCc8Ve7BpqsEU4YAyqVwXEYO0KvrII1ehG1HDIMeGeXiYFsa0TZBz39LRrlR2PeaDHqXqQXlZHQ4AfvyyL4LRlHOAxvtBkGPhnI5sPE9xThIzj9Po6ta2FlKc/YaQzm3BXK6ac4qC6FgjA6Ik/POCbqzggvl9A+UkjlyLFHpTgLopQJpuovL70+WV5Dcjyn4yJ5TvcU+X5FXZTvz9PkD7CnbnxlroYftHdChXBZzxhF8zXC2M81ZTk1BNfvzQ68xTjvTXq1BuUzm9c+F/U3M7yHzg8J3vs/pJvqJbMwB5PlRyL4xitGCIvYt1wSh3F72beFDCz6zJ7zz7yQPOfZOAVrwlGMkSFIx3RxuzSVycYx0v5o1p/Dsfi9a8IJn5gYd5TTqCvGumFLS42lzBt8B5oH2pnWEb8bxz9zGe0S2BxusZbwzjxQluo/7ELbPbubEEe6RvmilM4RjXrW5CgpcNjWEkcPKKDBVGUqzuQVVit/cgkyl1tyCUzSMlaFBnt92RTSpxESamf9s6jeY3gTz8hO+Tr81RZmVH7VvZut/spuT72kJno4trjiHZMUVmffnggX9Z/oLBP43y7vhZSkAAAAASUVORK5CYII=",
    "FaEye": "iVBORw0KGgoAAAANSUhEUgAAAGAAAABgCAMAAADVRocKAAACVVBMVEVMaXH////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////fZCMHAAAAxnRSTlMAJHDmAQYLA+zpMZj5cgUNAtnYyAQSDJLukehIr90jmUaykxcH2iXKL+0KV+ci8PV0+uqsbvQwEV7kKvM0yUcJ8faUMuHHqNFvVcEm78Agphq/kIBmiiiuLQ6z1xPcer3rJwjWO35FStC2NRb+331JsBC1WP3Ezk3yLrT3OJ/iw4Wp0lG4Vvs3GJ48f8JliWqtLLc223m8Wl+i5YMhxsuNsU/Vc1SXu5yCxWNEh9T4lhW6mzneXeOBYqWG03EPUjMUmnsZvj1yndYlAAAACXBIWXMAABcSAAAXEgFnn9JSAAAEaklEQVRo3u1ZV2MTRxA+yZLVLLnIRcWyVcAF9xZ3wDZu2AZjCDW00EvovZcEUqgppFJS6CSBhARIp+R+V7CjlU/ab3fvZB5vHjXfzNzOzk6TJOmkk0466aSTahr2V3vziwNWz+bNHmugPd9b7X/+qnRndFVvWCEDqt3xrT1jotrN6w5bZQ5ZD291TUD9a4UeWUiljv3Jabfs3C2rpKM7Ldo9358la6CH87K16f86U9ZIWXs1qC/aISdBl1pUqs9ek4bky5/UBw2NLlejIVg/uxwh0rao8lPLFRSO3rZ41MY8FL7F88X6zx6k5X7tm0IDU43NNLLjsUD9pKe0UI03F4OnzP2BRn8+jaf/3XYQHiWclwiCzR1h4w0BGl+Xy/ui9AW0xMcFLHSbjUZfF7xRSyEt4+tinBckHoc4KvJoqco2BCwpBf6J+/7c317MqOo4UhwuW6Q8Qz6I6oW0/gHgn6wE/4cIo9UxoLiHybSkrShR/xsAlZYYP6sVvGPjT8NeQ8tm5iTE/0zwLu9Rx/QpuCdmxX5+HQgPmeMkfwKQ5lTKwPI4N9hjl5MiC+LjU5S7+uiLCsWnv9gZnEi+TJGe96H8BvKPOx7y/ZskL6HMN90QC7Ri9AF5tP5pDQmYLwknjDS0kyCvhvUDPJYyKs5INH4EVTj/ZxZUIWaKidLvost0OMoy2ZCOhoox5tvQ+mz6AA4a1UredA9UsmGUNRVX2Hpyf7HvdyDYX1GuEWtZ+/KGezErGJV8FFo9dr9luI0hoeDHWnotUhOjR6iISl57mX6HtrkbGDA3KSUMfpM0k8EhuWQVv1eZQXIZgz9Hms7gkFTyFt/AQZLMGPxBqVJg4Au+gX1CAyIXrVfnon+YLmJdMqnc2/gGdgsu+RdmmE6NSt7iG7gahW3F7PUW4UP7mW/gJP+h3R3lfchNFZYU7ghFEvZxyD4wxquAya6cJLtPeAZG+MluFqceyRujstsr2fpbh6OgQ5BtJF5o5xaceWwDywimE3G7Y12VAT1nK8mjpgfM4c/MK5mVFZxSpTyg1Pg71n8kwi36/YJaomhbhuFIuLSA27Z0xk/0QwByLsbevpjmXo5wG6/lCRuAHPCRNfbxsfl8bUL89Fm4rWNWIzW6gkjOTB/n/7lGkbhtH2ziN78+A13S94NIWKBs303/LruwJ9C8JzT3fZegfa+9CAcQMCAUJjWADP6NoV0+YEE0QoH4S2GOjQVLaXR+OncIBP7J/IqNj7hp/GQ7G28H93v7PncQ99ISp7zprEEcLDUKzYI7u9MBnPpHKlglOMH7rfpMHBXzUTNvDZ+Oa4ZNhzpLURs2oGqdswWuc2w9Rv/KHLM5Z6Xf2APrS9oztWsvw4FkFlI31C6kRmnvEq3qv1mndSmoycSqJu17R8s7Z9Sqn/NdkhvghT+WirWvCLdIyZPL/x53tVx7M5gtTZAy7M7FPjgmhowlple1go8EnSN13QGbZ9cujy3QXTfiXLtJ/9NDJ5100kkn9fQfWBq/u9XcsGYAAAAASUVORK5CYII=",
    "FaFileAlt": "iVBORw0KGgoAAAANSUhEUgAAAGAAAABgCAMAAADVRocKAAAAeFBMVEVMaXH////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////0cwFZAAAAJ3RSTlMAzvyIMTABAi+JggrNzy6AignchAPbfvfdI8hM9sf1g8ZPOsWBJMONCu75AAAACXBIWXMAABcSAAAXEgFnn9JSAAABCklEQVRo3u3ayQ7CMAwE0BIKhbIUCrRl3/H//yFIcIBDEyeOJQ4z58hP6TInJ8l3RnljiJP3caqv58Qnow0x8wFeyX2AnPwBWngATQjgI5gggI5sgMIAmmkDbCEYoKE2wBQEAE+QACxBBHAEGcAQhIBbkAJOQQxQVxtwCBEAuxADsApRAKq0AdpqAw9t4KANUDTAO3xgqQ2stYFKG9hdlAGqi/tSFQj7xgAAAAAgCBhP+/b66g86ImDgbsieCMjcQPbfAOMRpSKg03PcIUsn+NEAoIvQRegiAADQRegidBEAAAAAAACgA5g4800rUMYBVhFWT6wp2pdn9jHmn+aW9ZxbKXwPZlX8zn8CPSxmfwK7sHwAAAAASUVORK5CYII=",
    "FaFileInvoice": "iVBORw0KGgoAAAANSUhEUgAAAGAAAABgCAMAAADVRocKAAAAhFBMVEVMaXH///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////8NoezaAAAAK3RSTlMAGoD5B3f8BnqJggp5gXuKCdyEA9t+iPfdI8hM9hjH9YPGT3g6G8UkdhnDY3zGZQAAAAlwSFlzAAAXEgAAFxIBZ5/SUgAAAU5JREFUaN7t2ttuwjAMANC20BRaCuM2CmxsbOM2////DWk80CJSO7a1i+zXRD6S2zhR2ii6jiKvHGDiezos9m8RJYoVIOMCnCOnADnQAZgQgCoEoAguCIBXNABhAIy0AbQQDEBfG0AKDAAncACUwAIwwvX0QVYfy3ptAJQkIGsOdluBduE2hX/0pnJ9bQAedUt0jgQN9JoPeYgC/ALvNUVUSQSAuTYAa23gUxt40QZArBeRA7/QpnzAP7rUBubaJXr6YAP+XgSL5DDlAZwwwAADDJACOscY38TitEMGUlqfTMnAiQY8kwHZ3fAfA2GHMAMMMMAAWcDatWdSTMu/09703+nHllT52GJHRwMMMOD3A4F3kngAc/kvcDHL2tF+GFAvEeb7iC00A/4m4GTyu7vAWAaYCfx64o37nzOLjUT+7YPn95xyzHwObpbU838BRIh9ESQbtzsAAAAASUVORK5CYII=",
    "FaFileSignature": "iVBORw0KGgoAAAANSUhEUgAAAGAAAABgCAMAAADVRocKAAAB71BMVEVMaXH////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////joNC+AAAApHRSTlMA7+T4AoknqlX+A+V9ASSINWz3OPD0PvFqS/WDnPz2dz8Eivs3nlA2+UQl8zTyn9TMlA0F5pNCCUwtVn9Gjm/6nTtARcuXWUmM2Tne40O+I+pXIqRH7M2Ps7jcWzzhwqNBZeutMXkXfqLGp0qN0JIw1RFUuzobwKEgxbCRU5YV/XwfxGKl6GdI7YsKTZAuM3aaez3igB7nDiEMY78IdJhaoA/Y3Y+VVaAAAAAJcEhZcwAAFxIAABcSAWef0lIAAANcSURBVGje7dlnQ9pAGAfwQ2IhFNQKalVcOGorjjrqqtbW0VarVbv33nvvvffeu70PWpJLkITT3JMnvJJ7YUie3P8HR8DcQUi6zYJWsa3PY9WCyolBT88CG/ntF6l18ypneimV/XAgQgEApcVgYA0MgAsyEAALFArQzFQDQMEGABPsACDBFgAR7AEAwSYgLtgFaHaqAVHBPiAoIAAxAQMICSiAdqcaYJVUAhtTDdQ6CAS59wO8TCla02kDEG4vGmNdrlyGAPcg+TkutU9GFgA4Cs+PC0JAyUJ4vi4IAbSy+AM4XxMospmun9OGoiI4C5A5cw3VdwedBszCa8cBk/DAQSBnD0fIdQ6oCsnzk4UjjgHBUGz6kCTIHVjAreUPF6h5ZuEA9jKtfcYSzxdoz3ge21+Uoe76sR80dwfLW14fH5NEwU+QgLvQnG8Q/AQJlGpfgfl1iUfj70MvQQJFWv72euNx/TUQJFA0ygLe1pkrRsF2/hLe+HAEm/l55az7aBGvKr/EAnm7We/+Um7Zj30FeStY56X8/EPY9yCen8stv8deRfr4r+fnZxIkUDDMej7i50cIEggFWceV/PwnBAmEqmbMX0WQgJ5/zC2YDwRc2qJaIT//OEECFvk9BAm4clif68u45U0ECbiirMtZfv5zggQ2nGI9uiq55WsBCDDS0DBiPqbNlib4+ScDkDnam4rYgq2fCwyVcHssDkAmgWFWquYB/PwbAdAsU5tdZfEAbv5gADSNHdBrTYJAWILNk1/ptTIxoE8CTsTX6rXPsghQLQFn+gNTxSYBwCKfA5SpszlJGyN347dG9wxAgwReq1BHKD9f+Tv58eun2Kb519/pgEsSeDGkST3c1qZuvmhntfzkA4d98NUWNkKtraan9psHrLPOTwJq76gjRGm58bw//5KB+wL5hg6r93l2sVvOtti3i35GC9tMTqibhKWiiA+46ng7/oH/Pk6pPkaDJSduYn7lSnj+U18oP5R9Nkb9sUdXnQG2xI8F1VtONkZ3lYcPHQEen9EO1bAp495m5d87ux3vlhwA6M79PuK70HVO3x/rfFozrj3e2httRgPOLIakgVkIyLj8W8783Dt922EJZOMA65/I28cw+ZtbrD/KvqGwx2YLe0m6pRsh/wF5CfdEGpipWgAAAABJRU5ErkJggg==",
    "FaFilter": "iVBORw0KGgoAAAANSUhEUgAAAGAAAABgCAMAAADVRocKAAAAk1BMVEVMaXH////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////bzZmPAAAAMHRSTlMAxvoKD9h9EsnOBHoRNQbz9QnXe8IIR/vwweVktvn+IJHZdxUkBa9S/Dja8n7DedMBq3OuAAAACXBIWXMAABcSAAAXEgFnn9JSAAABNklEQVRo3u3ax1LDQBCE4ZbECnCO5GSiTdT7Px1VhEKAJc3sTt+mz1v/d5njAsVzyCvK8lAWwPSlIm6xxLyi7hUzLvCGCRd4wiMXCCi5wByjjNnP7oCdPV5/fxdgCp99nvDdZwk/fY5Q7zOE33174W/fWvjftxW29S2F7X07oalvJTT3bYS2voXQ3k8XuvqpQnc/TZD0UwRZP16Q9mMFeT9O0PRjBF1fL2j7WkHf1wkxfY0Q15cLsX2pEN+XCSl9iZDW7xZS+11Cer9dsOi3CTZ9YNAEDGC0JgAOOOCAAw444IADDjjggAMOOKDeERs4YwObNRnAw5gM4HJBBtC7PucCwDCQAUyvLrgAcHBCBnBzTwaA09rB5gwAt1ntSwNlvfLwCyhB2nD10T9esgAU/TCZ9Ueit+/+FNdkhsUKCQAAAABJRU5ErkJggg==",
    "FaFlag": "iVBORw0KGgoAAAANSUhEUgAAAGAAAABgCAMAAADVRocKAAAB11BMVEVMaXH////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////LJsPXAAAAnHRSTlMA8az3AQx//P0JzQ7hwsfbAgf5xQaq7Xrsa0yZnifR9beiIWlPkjX+oXKW4wXOCnbdgNzmDbyn7hDZbBlErhFmSyMEqYooUdd1vTxgQSJlAwjwjiyykxI2eTvgHMFAxkXq7ytul9p4FrucOn3ApumHJWiMC7Dz1LU0uho+gYbJSE0PUpXYFHu+XD3iHsP2W2qdX3P43uhnsS/6cEl9tPyHAAAACXBIWXMAABcSAAAXEgFnn9JSAAACK0lEQVRo3mNggADOiWZG5pMtmfpYGWgBWPqN50DBVHsW6ptvYjsHCTBR3RM6pnNQQIwIlS3QmoMGVKlr/gwOdAs4Kqhqge0cDCBKTfNZmTEtYG6kogWz52ABNlS0wB6bBVZUtEAcmwWTqGjBLGwWTKOiBQ644sC6iU2jxyCBcU5glByvWUSDEgllSJgmX6lQnWQAiG03E0sqsnCYMAVDVIxJuYAo45Ok5SE6uKJB3OmYFvQyzsEBDIoiBQgUPCqm6nDlXIJAkXD2OSQB42Yff1ymC1SnSKAolgSJ1s4hFeRJS+lhmu6U5dKFrpIPJCHSPod0wKyg2YlUrpvEepRgC9dEsKyg8xyyAHtZuWJQW76qBq+2Oi410KCbQzMA9eGoBaMWjFowasGoBaMWjFowasGoBaMWjFowasGoBaMWjFowasFgsICd1hbw0NoCRRpbwK5GYwvEGWhqgXwhAxEWSLQIZ4SIilbq8giRaL6nIwNeC9h1ZZLrZQURikTUlGV4iDZeuwN5pBhDml/LzQLryGiwPi8HEcYzGhajaEOVtRRXwje2a22VQ2AsV4zNDk0PsmymF+FR8HQ/BdwxwueNOWqLpfQgBLhVQt0xDecQ1pfFppoMC0CgSiotV04MGifmXEzZ8dw4VJJpAcwz3a6+rfjnDymzgAgwasGoBaMWjFowIixgx2xNUhcYISxIpYkFNQgL2GhiAWcczHxhTgba2GDID4wHdn42apsPALU5iUN+xEWaAAAAAElFTkSuQmCC",
    "FaGavel": "iVBORw0KGgoAAAANSUhEUgAAAGAAAABgCAMAAADVRocKAAABKVBMVEVMaXH///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////+Ovq5KAAAAYnRSTlMA/MwRL/sWeBfV03tayhDYGVUB+eyWd+t0zhMMDdRyLcIIoojkxBRrntb65Y0rBv18v10pbPfiiijNDnqjQY5QErd5O+DBQGQmB0/038Chh2hJ7irzU/jZFTmg44ZI7Yu0UkOm+YIAAAAJcEhZcwAAFxIAABcSAWef0lIAAAIQSURBVGje7djXVsJAFIXhgyZSVCxBxQaCqNjA3nvvvXfP+z+EKBbAkJlkZt9xbshF1vcvVsiQCZGHWZjdMm8XM4SaQIg/Z7wT5Nd0c376urA+qFDgQwqBQS6c4wTWZx7R7AdLfL7ax/rMe2CfA2C/Huxb95jf/6+/qs8P2fjsq8P6+grlfF2F8r6egpOvo+DsqxdEvmpB7KsVZHyVgpzvvSDrey30B1l6zCYPgXVmaCHOjC1MMLhwzeDC3TS6MMbgQnOL10LjXAJZ2JxnPlxJwAoH1tfBdqebghWNSBes74MeucJTNnfuyzMZEddfRq5Ak/7X3XDuE1f4GQ+F8Qy6sEHgQgOBCzduV9gdl4Ejl351lctAL9gPudvnhuvdXuLGio/y25fbsH4HGW1Yn0QFZV9Q0OA7FrT4DgUpP36SGhpeMkrXh6pq8doalHlBMmXmH/tjDj5R1LPv/1sOHfxzS90vmhJ/7QLr0xnYpxGwb3OiXp8ms1if6FSfb//nXfzsreQzx0QFRZ99hnNB1WdOOu4f1H1Oldk/jH3uRd/6lX22yr7ReJ+4JHWfZxTfzIl8ToP9oVGszwNKflTo11Z8oD8l9P1KftzE+vQA9ikF9skC+zQM9ikJ9snwYX2iGNgnaoXcv+KCPt++oNO3K+j1/xd0+6UF/X5xAeEXFjB+7o7L39PmAKHGSD5aM+lRvegHUhrDoe3SRqQAAAAASUVORK5CYII=",
    "FaGlobe": "iVBORw0KGgoAAAANSUhEUgAAAGAAAABgCAMAAADVRocKAAACdlBMVEVMaXH////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////8rISQAAAA0XRSTlMA+y3O/Qn29wED8PX+ye9xHPoCCivpzAYdNvkfYzDPbTPigKXLcvyiNe34HiRzdnCj6zHbFxsgqyrzL+gYkya7qJVFtAdmGj7nTJk3Ro50eZ3q0PR91LWBSCnyLhmfVifFFdGmERAEWVoPamxHa/GUMnpbf+aEiUssb6eNbpLVoe5OkTR344uQFLmarFiWFijXYMdEyqDTUg5BI55Qv2IFOXxUqeyKrlF1E9wiCLw7fuWDIXjfXuSCJUkLU95nqq9XOHtcpJc9zWUNt8FAur4/OsZo1BYAAAAJcEhZcwAAFxIAABcSAWef0lIAAAVpSURBVGje7Vr5X1NHEH/BkASEkCAQQI4iQgBBRUE5xCIqiqIWsVpBRaCFel9tvWq11VZrq622VuvZ+77v+77v/Y+azOTBe7uzb19I8vn4A/NTdr47M2/ZnWNn0bSbirZcLCu7eClh6uf9nM9CVHWmLq5qp1VEftzRxSKU3Rlh5eXGqv3O3/6sbMefHzvYCPn3Is/nvVZwVwzqnwhWMXYOfz9Yzgzk2I/cxYy5L8wfo/qsQWdI1XkP2upmJmpErZ5fQ7/Lf1o/BvWTX3SCpiUwCqxmHE0IANAGA+ehaDcj5WAj6lmI4x+YQKsQuRVHPcWZ0ejvTdfVPIcnySUamIgf/aw+Pn/AtnrPa35dajNyFjCCdiGWo4/zM2zqnzQ4quRV4NRXUQbceH6HRzkzk+3o900dlWgeAtYRRtIK3P+uUU7RHLX+5bUGFS8Dq85FG5jYD/BDBlZ2hUr/LceMKpKAV8YkdBDgJCMrNU/x/anG2UXILJIZyEH8ASPPe9z+97M16NFMSrjCe5ndNVyupOQfkRu4Dz/LzKydLNP/xYB55hvoFe/JDTwMgcpTwkURyWn1tHLiZ4BdyJjqb3Qbxy2lDZzgpb8E9nQrAzUw5W6eTfr0bgc3y/m6MZrR9Cg6ShrH9jcR8TOHF16NgcNtZcCVLB7UMKWnCAZEb1qGmZhZUie5CYzdL+SXRmHOOeKQC7ROz5wcdX/AGfhWlG0AYK21AUxIvSIw16y/KU0MZZiillobuAc3UNyotCSTgbmi6FRMBUxBG2DaBBFoNepf7xQnbMRqRWWgA6YdFQFnlsHATEJyJSBTVAbehWkrCSSojdM4xY0WkfHyCGDFTElfw8RdFFT1FmBLSLl99tyAsT0WE98BbBWJFQPWojawwmKp0wF7jMQeB6xVbeAonTUN8ehJEmsD7G21gbUwsZPEnoE9dpDYblmUFLIjVRvpFf2NEFRNy1WD3Gy1gdth4nwaDNdg/9LQNJArURuYjSmXBv+Tbg+bB3LNagMl2AWgwRdC0FYawvu1V20gG2/sNFgQgmpoCAsel9pAKpZPNLgtBD1PQ1g5lasNuDHtM6mnJdxAwv9ECd/kYRqqi9sxpaM1y42Lo4U7EZcTGSrC8SBAB7tZcQl2Q/Jw/ZT6emOsr7eT2ICs8A3Rd4AdUhu4ABP3khjeBj8ksa3xSJk1Fkl/jd2kj52vdST2vUXZchKwbWoDi+UtK1dgvCgcp/iR/BK4T3VK0V9mEMigrWtsh8rAfkmzgruJfyO7iLfbu4g/reol9MbWSsjsFhdQaN4FoobPst8MaRCBV/iXmh4yYWva+9YGpsCkAnH9b/IHSTwHp+00pPAacZqpu3Yp6eTtwV5LbUA4IWJLTWvy89vUb7cpeIo/5Y5CytsyeOlNdtuam3h2Ge3PG8muoI3G7DJ7jVktmStSPsKG8Ceq1rL2lZlbNEkWkubUUrWRsjnexDXHLZ5BZpnb75+q2vsHiIRv/URx3CtWhfLyLvJAYXrE8yoeWcxrKKSPl7m9a1ph6nJVaqjIFmqXs7JHolMAHzawuvLUyafe4NIvLRLemQx0GMChpdE9c4VOa5BvidST4cK9gfcyew914cCXb846dN3WgtjIgvNn2M/RDZt1qb9gnEs9lvoA+mXk2bMhmiogc8dOFPsbx3tEAycQ+QNHO3dkRllo+EqdhotIQPCFnBuG115nqW8MtUxfMHy5ysFw08c9YPR8hoEqnO3Lg31jLJeqr7j18lzrMF3lHJ9req50X6mO5V9Brv/jjaz+R0NC8g8jr73y2vUtsVZ9Pt05t4/4eHOb7va5cS0wzy6Al3F3S3/CathLVzMyrv5+c9XV/wMznaVgIelZVAAAAABJRU5ErkJggg==",
    "FaGlobeEurope": "iVBORw0KGgoAAAANSUhEUgAAAGAAAABgCAMAAADVRocKAAACf1BMVEVMaXH///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////8Un7H2AAAA1HRSTlMA8fj+Afr57hEFQhLkCgQj+4Qx39Wu2FQJbwJzNfexsNYD8PUXp/R4IAb8r0Uld6TNLXAyrQcrfRv9KURgbF+F9tfcGA2y2+DBbhbp8xC/H8Tohh6CJmFDeefs0jd/HaOoicwsdDYc5YMhy7uMNBVY3vIUON0keidtVWtms5ThxwjRULdaiEqNDDC22lmcXqFjpoerKtRyU5bZYqou05oZgEZSyCi4mb2eImXrao6TmLw7fqLG6tBPl8XKSc8vtbqbfBpdPuOlyUiLzk2QDzN2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",
    "FaGraduationCap": "iVBORw0KGgoAAAANSUhEUgAAAGAAAABgCAMAAADVRocKAAACHFBMVEVMaXH///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////8EoJZ8AAAAs3RSTlMARJGA+6ICzAMBZvURFrsSVer+8bP2VjwiCOQGd3b83QQTiScHKwz01Zb9hsmq8tifw6SuN+utb5g2rE+2+SCHbfM5H8QzVyPIqewt0nBRmdwYnsKEqPBQ+hdanTsc5YMhRcuxl9qcOoLFY0lyELU0upvev6CBQ+gFSCnOCk0PlTi+XIWKKGsJj7h6vX8d0Q3Wt7x+X8aIStBuczVZwF6h6SULsNTZfOOlJItsr9NxUrk9YVNIuyQAAAAJcEhZcwAAFxIAABcSAWef0lIAAAOLSURBVGje7VhnQ9pQFA0gQ0SWiqjgniAi7r1nW0cdde+96tZatY7uvffee74/WEjSmCAvgUD6KefbTU7OTe7LefcmCMKDBw8ePHjw8DEiBKtvbbZvq4IILtQVavEJgCMgWB7oW/VKeXAxoCCyyhDnK/UEQ5UWuEKUMNTfW/HoUOFHQAONeDveN2WHw74gXWzUS+TBncBN2BfkvIdlF0m1wDPYF0TinniYQBgOWEEjVr9hUlepxUHAC4QEy4/D1QdzlZ3Aa0RKRachZQc+g/OCsC87HNlidSKmLqvJTAKcICmzRoYgMVLAIYZjkBbAKcaRI9wmyOY6gQYZ5zbBIqJK5VI/VYUgEl0OZ6+pTvY/jIZBkrXjQ/Vk8+7hvTvDR+KwzW4IW5Q/BeXeqJcXvIiB7NWtjvNTF+0LMpPXz069P+9hGLQX1KL5K7HA/2nbMU/Vv5/Np51hyhykFNIBfcOen7vifnsNeqZemecgrjn1z31LGrN6mmVfxdzpJTYH9/BtxG+M0jokfHTDrcEr4jJKf+eqYfvPD9xwrf5oYN7d0TEWv+QXHpuN1Mm2JPeq04L43XlcQqHEGc1w/Xii1AnYgWQArLG9FDMquuuJgSaovltBKXBvrNVuX3iCOOLGnmFVCUCDVz2FW5RCDrUV2JQ5V9YoHyBbhT0abEiF16uPSLCNxoEHg9TINO2nRuD0SAhBpmH+s1UI9uRZ5GqbpKJq11fpc5WUlcmHJxCYMMosFuqc35f350Jl1Ctkv7+MObN0NI96D2UI8cjoqrm2dBD7e2JjusYFxUiToBBlTOJRCmRHqEsX/TB8WEyF7CEpjEZQ4lERu930JE2CVmxbwQt9i2Ufg+ufwl58cJRsA88/DuBGqMYpcjTqYtvP4Eb4ijPSD9vAE8CN8AlnWNGohm0CuBGWcIYW7dgGtgmgRjjYe2ZobMCMHViCDoIy6YUNaIzwgKCgVltmm2AZluA5QQlyWE3ENsFd2C8VkrEcVpuLYqfffIbBZoTVXi+x+jiuhVXoAomFWU2R4em/EKAVK6AuIM/VVvzYxG2TJ/Kmol2arbSZfCPElKYvXXdXft08SDsVUf6vfCa1xQnR/UjGuXTqWoWMYezKIfHHoqnn5gRPVjaLXWsXb67MCs64MdfdvFQnxTD8stJlu7he8bOstKndsqCUSpULlvam0rLGir5o/l81Dx48ePDg4TP8BWRM7mGeZYi0AAAAAElFTkSuQmCC",
    "FaHandshake": "iVBORw0KGgoAAAANSUhEUgAAAGAAAABgCAMAAADVRocKAAACSVBMVEVMaXH///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////8+w1HyAAAAwnRSTlMA/v1+yAF3ZswzBuYDAhQJ8fCABedQp/nFmuKZDQRr0h1qS/pA6gw6IPgaKa/32PxhpKlR9pQyNxjC6zHWdBK3eRdaouUmy6wwtjV4Fj/kgknujPMQlr8+44HEQ8ln7Qrydrndvh5W3y9Yrbwc1cqlD0KF7Iooza4Os9e4+3pbf0HHRoknLJg2wYMhxmRFiGkHSu/QsZJU2n0bwF6hY6bphyVobQuwkXJT2RW6mzneXR8kSIvObE3TtFI4XJ89wxO99d1lSUEAAAAJcEhZcwAAFxIAABcSAWef0lIAAAOUSURBVGje7ZjlQ9tAGMbTboXSFija0iLD3X24DJfhDNnYYLjM3d3d3d3d5f3L1rsr7SWDQnvptzyf8r65PL/kcs5xkiRJkiRJkiRWuQXolVaNrsFJz+AvSlr6gCxH/QvkQGliHCdXh4BQ8gLH/FfyXCY9cDLeADNoAbv/WCBO1iwGEIngwjOY8sHJ7j4AkQh8/0x3nCwNBxCJwPc/RZLBcgCRCDz/5N6ZoEwEnpVuEOc0hwAcIrjxhNviTvqpvrW42JlKmFsuqKQH35HjF3FFRZRUYvk37B+1ah7+sBAVdeXn5gD4LsH+dUpwDkCbh/1jngqt9n46uV0MQAr2P1gudEpTcZxPEjtgEvev6CahURvOG9kBu1BUpRX6eOPvUvQwAzLQAKf5LBycg0n7Xsb+D5QoOCsw0Y2Q928WoRWloWAdv0DLPuzvkypGMy1DQSbv/vps7L+iWpR+kICCWvq2IZ7Uv2aDKIB8FAxTd/MXWYaw0gwRALISFCVYbvZ40IOkQYSe/BiPjtPLiKIo7Jx1hUxvJb/YAZ0qFAbeQTO94T7xjdDCVrJAUhnZx6L3pDqC4rPfKshleqIpvY2MsdzpcFaAzE8wPZ1IxvlzS0nYEck6H+he0/aao7LpedqFfNGbD4wAkL14ZvE/cIS6ccuTdDl/RgCAV1s98vLwq+SvWjaRTqGolTMC0Gd09kfK/muUt3+SDwvzYgbMoh0uhHB3wkkAACPuKFzgQ2cBoCuH/AhvmZMA0FBDqmnwr5MAkHidEIaHnAQAeKTBhN9/nAWA6lYyjRptALo2ax0HQNJGUk0jVoB2VE8DQk27R1/HAeAVZh2wMEBt6uYfrQC1ZZ3jIADkxQoeAK/NfC2AchR6swAAUoNowG50pbZWkcn3qpoNAPtpQEMcx/XSP3nAv8m6nP7q6hpiPyCxkABSyMbFv33GZtphfo3WXPsJ5lkuwlY/6J+uSM4zwW6AeXr1GZodEFtPTY8BdvrLjpmf7NbNBtAf5k3wjfYBLlgejK6gAZb9Zl6dcIcbKrfDXx1HPRlHbWNt6VLsvP1jXzp0ZBThNZfx+ejLwrHCLlWpbftHxnCeN2/4F+c4fGwXk2TLP3wP+8Fg4XEbgFAxjh4FxxTtZdbrAHEON1U/KP+iB+YVhEm5KrHOT9NfTZ+N3EPD/vcWHDzfIt4Jrfu75jHltYtPzK883jhQMeWnkE6uJUmSJEmSJPH0D9jz7kS2af8nAAAAAElFTkSuQmCC",
    "FaHeadset": "iVBORw0KGgoAAAANSUhEUgAAAGAAAABgCAMAAADVRocKAAACT1BMVEVMaXH///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////9W/TobAAAAxHRSTlMAAvIH8/z+AQYDBQQbz/1s9G3YGfEcEbWbtrvu1/tGbvklIU6ZNtBzOtkav+31O3cI+oiAV8iuuDfHZdFvMbfBRcDdPevN8OWqr7lExDj2CTB+a3XhHWAsDVAS27zqJrHaeJx9sJG6H0OGyQpNLtO095Ue7LJis6NBolIze6ik52aKKEzSDpR6GL2ef+aEakuT1lWYWp3gX0DGZKfLaUrvKwxPl9/kh8pJjAsvljQVOd4+JEgpkHYUmr6f4lRjXCPMoYHCaMIlowAAAAlwSFlzAAAXEgAAFxIBZ5/SUgAABMFJREFUaN7tWfd/1DYUV27kRsblLrmRu8sld9mLlIyG0QAJJIwkkEAZpdCyoaUUKHuVFrondO+99967/sOa2PdkeciWLPND+kE/SXpffb+W/PT0LCPEVNZN9fddmGgZ6Wpo6BppmbjQ1z+1DrlVys71pOdLhuL/q3f8BnH24G8n10rU0n38bFCI/tO9xySbsuL7Q47pp1aVSiwlvcAR/eQSibkkm7npE02lEk9Z9QUf/zPbJc6SuS/ATt+xTXJQ9j3Cyv/SRePosfTRaG5zIREOJwqbc9Gj6Roj5onlTPSBH+r0W2pfT3NIDws170z7dcC6VoZlCr6qG7W4tZqGrW59QAc+FbYNOknNgNLDbdb4eUmttx23iR4PLtXAbxyyn/LQas2Qzi1W4C0fktj2ZWxOkdtBjhp4m44MDxLAht1eVrfzTpPRtraM6j+/ELCb23g25ucjxNB7ab70JLn6Cc7QsocY3GeOeYPwh70lvMErcD/he3eZIcpXqIA+J+H3JvUBt3cYzaFhlf8VZ+fHIlXhUeMKnFdnuMHpCdWjcuzU2/Z3Y9sV50fsx5hkbUFnOolNqwPOBQK3Y5o9Wss7qv9vFckSEu9hoqc1IfQrvH8XIKEyiff0reSGPoF1d4smUo9jqhcJF22Hzh3CuZo3hqegHlLPYdWPxHPBHCbrx3134wjkRi6LPWkx9LRhzXNuCAzhDQ0Ok4KOw+6k47XA95nSPogzkDZ3BOYB35uKy/xhWDOE7okOxqpmuqoGopbnmrd+QIbFBiMbkfGdKq8ZH9qtAIg/T2Q8FeV0/vIWInvahYEbcMYqf8CMQdoE+c8tPk2eUEGdg7dFA/Q0gi48XyZErlgaDo5KXTYVoQnU64CV9UVDJ/RMzjT+gUZP8fn1/NJCmsCdemRlcQ4PQ8f0TGMCGspHRNxnSGo9NAETqHJWfgntfxEqAdSYEjoeM0nMaQIm0JTyYQqe7wmgP8G0TfFPv5iAPy5bfoZ2AZ2F6kOm741XQFovW96FZiOahmpUtiwRFUgqngjN8+gUVHOypV1UICZbvoXm32gNVN+iOAangE+2HMKZMBqF6n7qGC4BBVyA1mvoV6hudVPgR2i9gPBVR9BNAS++LkE6FpcEgmpz7gusLFbev1YCvcXKB64KhNVmsHd2Dit7g9dKwDbEcwrU2AgIx6KsjYBL0bSMOjYiKvC6zQw2Cp5odXGbGaBdYgLPIjuBco9IVlFTbSuAGkXyIrhDCFnNvp45s4vq+aOIRQA1alepgnr/7a3Qrs9lxCaAOlKs2TWhUNdEXB6G7Bwkvr42O/sKfS9HLO/vg5GFMiybfCquuVqx9UDBMvcFSq4L/P8FAtcF5qTAspR8PTJ79ZH9adwOzS1wm/bXiDSct4R/gm8xGfm/zhh+zFn+07qE/6UyPn/GmD90fWMxAKdXo2wCS81SrE6LjYxPujNs79f83+IB+kU/xvzOJJAyF2iiXtodwVcX1UwCMXOBLGV9Fh1R/0ixvQKPuYBZxnfwuwiZy5wWErAvaxh3Qcwhf1WeUSDlUGATa5wYd8Z/B3ugG3bEz/EzKJ/hpu/exBXfl3fx+k+e8wTJd3Kw+2tPOzikDjRd9dlzzz82euYEGR/+A6GCBqi645IvAAAAAElFTkSuQmCC",
    "FaHeartbeat": "iVBORw0KGgoAAAANSUhEUgAAAGAAAABgCAMAAADVRocKAAACTFBMVEVMaXH////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////ocqefAAAAw3RSTlMAAfvCCf7mA1vXDBAHBvwS+P1s5PH1+u5tCAsXafn0xG6Cjw2/CvIPDgLn9lD3qc118OWItnhejKBidrk4ZEFPnNnHosq0sXfOBExRN+EiZUbMao6Tt53AY0TpKnK1XSTJZwUuvoAe7OtvFrvtmnETPFWrTmukQoVm0nCzlDK4vZ5ghKiJrUssMZg2eeB+HMbLSo0rkjARVNo63xs/IKaHJUnP85E0FbpYm3zjH6XoSItNUpUUGVw94sMvIaewU3uZhtN0pZsWAAAACXBIWXMAABcSAAAXEgFnn9JSAAAEVklEQVRo3u1ZZUMbQRDdEEIoIYRASEoKlOJQpFAKpVhpKXWBuruXuru7u7u7u+8f6+XuIpe7nd2zfmI+bbLv3uRl92ZmdxDqsi4zwVaVvh7S4vPiyjjrwIzpFhhsOZSyu2eJB8f4Ls+sHfmezl7UONmNI+1z33Nk9PzaVAnYZp3rAOl73KrEMrNVf1JGN9XLwdizvweRfvhSD1a2bc/l6GuDCeCshU5l/tntmGhZe6LRjWlkdMVtJf6RHgxZ/8RIcOwLEOy5JOef68aw5RWFwQnLKWDbymj+PjZMs9bQBnFVUMG4n5R/LGawiosCuLCVBb0+kn9EEssjuGdCAJw7iAkccybMf3Ixxswe8vMYwb5FQf5iK2a1XbGxC5jBecEwU4fZrb5eBbhGDD/TsElWwq8ZGopNM/5tSIwzz0FqICr9xCbaOs7BDzMdrOX+od5mOshyoiZsqq1Gv8x1MBR9N9fBYDTMXAetqMpcB6NRGQXxZYfzQp52B2koBgZU5XKviv+qZgfJKBMGDODDVYYOBXCuSV4mFBFLNAdU9BcuVsSU0Uurg2FoCDg/KlhTlml00Ib+QNPtobQ9UDaXOcORm0ENZL3QGmi6LuRguk0eBFjKnVLkAioWX3648pgZNRcXG/jW3g3mT3IhBFQ5ZyNKp41Rc92Fr+MpdQ4HySbOumfzHMJOjQpaVenCtxQJgfJuFfFVmyhwLBdOFB+VBFAklLmUN4hoI3iGr9xOCNi80bIVoEr4zUPGEMLRcTs/3YlX+IUCXEkAKMH7TYBshwpwRxbGGcKoIFyOpIcdABLOixDHS8Uw9JCfXMoNW4r5YU6UAHs6LKGkMPgjmm2KyY4/B7cExs38+EaMVEB8NijB9iYsM4cYhsbxY6sA2ydOpYjE4lITJIxnOva3CeAG/sMDyRbieLN1XyssEgNJufDRKhEQuVu1Wt9gTBkjLJZUgH4JiaGjQ46w5N0kAvRLqAkXma5gcpaurE4Jc6KzP5ecUyR7U5WEwz7aWYJLzvDW5K3yKMlDQwHw2CwhOTdSQ1xBA1nDFcDDlsi7L0BA2nXoX1p9j/zkVKYAl9YEr8NUcp0wkUVA8h3aSo8jV0Dz6QIyD9L30qtk0tOTqAIy17Hs1lIvSf4HigBvKdv78JZU0u+FBSQ9Zn3j1hA8rEiABCTFq4hBbiBNbyZUUbPUBKEpyld43j6uwg7ljeyuUxfmptAvCaUJuENtIO2uyoMtRX2o3qrGwQAtyWAGO/8xbenmBCv/U60JjfHot1B7ytzEwl9rRNkC2HhdWd8ygcZ/X2dlZHkH8z+x6K3tLEcg/knF+rtelk7gpsEAfi48byDxl9uN6dzZbxIuRYcb1Ru0lyvxH5hnXPfR2Sbnr040sr/plN367DSUnzspVEv5H6Ub3aPNl7RWrH7ju8D5k6PbUUabP9RLOpVrTie7UDzwDCoyq1fuOB3gv+syrxufMCH1WY4fdVmX/V/7B9t/SBMLQv95AAAAAElFTkSuQmCC",
    "FaIdBadge": "iVBORw0KGgoAAAANSUhEUgAAAGAAAABgCAMAAADVRocKAAABZVBMVEVMaXH///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////9UqgXGAAAAdnRSTlMA7vNm5CgMAgQNJvl3Gu0OUZSZA61Lk9Z4GyVJC5skKdg9QmsJLbPXGOZl6wh5wUWx9JLAIERollj93uNDCg8Zvlzi5/H2vVs8qCdqb7K3VZjbO+AcX6LlIeoHStW2l/6hY+kGq5HUv12gPgVIr7RS95V2mt1hAe46pgAAAAlwSFlzAAAXEgAAFxIBZ5/SUgAAAitJREFUaN7t2mdTwkAQBuADkpBQpIgURREFe++999577717v99PMgp34XJZZnC8/fxmH0jZCwkIpVdzaXkr5qrW6vdOlK0cTdhUtRfr97+zYJNliet+ftP9MS4o0wFGMEA10fs3Y5AaoAKllC2cb+7MsNvlpMT9VKCcsoWLHA9T4tVUwErZYpgcH6bErVSAtlPd5HiEljcMhMnxRjDAGSZ8h0ijEwwwWgIQgAAEIAAB5DUwFnyQo1G5xNefE0Cyqd85+/I6PNCi/Uxqi9BAm/131B6EBQ4L07P2FkhA0jLDnlpA4JSUtsEBY4WktCqBAW3k+CsYUEKOF4EBlJ+mc2CAQo4rYIBGjnvAAAc5Hvs7B9lHjh+BAf1qji80XEdKdwHOom3CeTQ0DjmulzLHdRXsghNMX3B80EtmlefX/nmCX/TrF1Lnkmqrzclti9S9H1OUWFG3JO7sBJBvwNntaLJGTfQEUk9tLYGehFqTvDnZBQBW6/ZS40H2VrpclV45NTYabB9mga2o/sPuhstjM0DvZtbH9eiz1wTQhRhqnh8IIKZ65AUmI2zAxgsnMIsY64IT8LICz5yAgxWQOYFBVmCQE0DM9W+Be9b+55zAGiswwwlcKWz9tQPeWdQ33ZG9fcdon1j0BSCAfAKsMP0nDL+wNlghw6/cDdYKFeiEASrok3EKon+7zujdqce5/eMGiheY7h/XXz6Kr831HynLukIN+EOc14M15M88vl/jWAEASwCXWwAAAABJRU5ErkJggg==",
    "FaIndustry": "iVBORw0KGgoAAAANSUhEUgAAAGAAAABgCAMAAADVRocKAAAAh1BMVEVMaXH///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////8s+2B1AAAALHRSTlMA7JncCtsE/AnfbxKIywcw83xI2iN/IutV+rxfxhZ9ggFJC/iBpdi5gMPHxWiSLlQAAAAJcEhZcwAAFxIAABcSAWef0lIAAAEESURBVGje7dnJDsIgFIVhtFZaZ611nseqvP/zaUw0RiuXSyEx8Zwtv/0WskMIDCu+yjKQKmcy2EVOvn9WXzdzISyVZpkDINABoQNA6gDpAFDaAQAAAAAAAACKAZtuuE26LUEHdsA6vh+fWnRgA8z3j/MFHfCB1bT0PD/QARvoT17Ox3TABKq1DjPgAe0mO+AAlYbkBwygF9sExkC9bBeYAqOhZWAGDI7WgQnwfvVYgQHwcfVYAQnkXD1WQAF5V48V6IH8q8cJCKBE/JwOhPI8AAAAAAAAAAAAAAAAAADwQ4D0+32pfywtvkTs/AKpiGY+v3+5PXlHWejpf5Bh6uTRHvv7XQGckwCMKWh/TgAAAABJRU5ErkJggg==",
    "FaInfoCircle": "iVBORw0KGgoAAAANSUhEUgAAAGAAAABgCAMAAADVRocKAAABwlBMVEVMaXH///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////+V6nZNAAAAlXRSTlMALAEF/q75+oAOd9WB9y1v+K8usNZFQmx5IwmE6zESp+or85EQCnGxgthtBERUZkyPfx2y9cFArEqSNSovfB+qDzjdJkYkJWFr5+wozfFws9f7N9xB8DbbF+VpB+8RFqYGcjT9xKXokJUzvidDVWBf9jJ1vSLMaq1Ljg2TfqIC9Nq7fcAgyhW/yWcpTfJ2FPweelFSXsr0NvIAAAAJcEhZcwAAFxIAABcSAWef0lIAAAOKSURBVGjezVrnX9NQFG1K23RQVgd7FyhQ9pYpsyKKCIKiiIjKkOHee+/9/l9DW2jTpn333iTK+XzeOfkl7913RwyGQ4O0Ec/2RK7dnjuxHRhJ01a7ylPoqhVZDMRaV2GgShv18pYOC1NE1usnvWrVSxqvsJRw2M+qkM8pnWFciKUZRPniQROD4eVRgvybGoZA9gB2T2ZWMBRuWM9j9OdPMTTcl8HyxkwLI2DGboTpDw0yImpKIPpjw4yM2Xy+fsY5pgLXinn6nWeYKlR4UuvPOZlKODtTxoYLTDVO3Eyun3+MaYCykaShc5hpggVfkvOVzTTCtPKJW2GaoU5J/4hFOwNRIX5fAsS3suZuQXiwVsRnOhKzgjz+quZHYaqtks/9Gq8/0M5d8zHKXuSSn7fGGbj4zx9L7+HSu+T6p/nv3yYL6fzvIL9/2rj8h/In4n8GsywG8fOHbrlBE3fBankMvZS/LYS4a4O/4mqUfVFEG+QATtu3A3Ym4HCiXxFjuQfsKQB7TW7wGLDkw0H+DAkvRbJt6puErDkeYbeAAlhlrME6aEl/hN0BC5GLUf17sBWbkfolCxiEe4Yi72cduMBZEOJ7wGG+qLJJED73TIIXeEMGhUw3+EMGNfoZhONRrX4GjlCtIepnYKmXDK4DiKaTscfM2DcFdRiX6KPgdxlFH9QgIJEbALz0OINnUIPfEnkCwLPG5whQgz8St05Pg6DE/amnQR4s40r4BiiDOsIuAhvsXWo7kHNQbaMZ7AK3aQgUg1/Ag0Y2CABDBdlgDBPsCAahYAcO1wQDB+rCIRiYUVcmwcCPuvQJBl5U2oI3iKQt0MQLb3AblToSDPZTx169DPaTX8Nd3KUfLparuWXXW3gBYlbqPWTDCxBuCZWuZJDOK6HewYtAq5KBFV4EGspXtTcwCYhCnGJwH9NKoHyDDUwzhLCLulDtnIRLXzoHKynPQUUroSGFwa3Elppb08oj7d83BQ2Gap3bmho2ZtuM+raWX/h0bo6P69zeTzkE8Wap1XfO6TxiGf3fQyLpS8/S9RfuQAZpNvKgzrUEnDU2kqKGCB017sUlQuRzbODGvV9w8u24ca+EH2aMflcrYSbe9/Q7TN00PU+c6mcsA+o3cVlQ8V/Ckr6/PYRz71fvkwQo52b/J21+DinY8pvdsrNhcZv93gJtf3CpH99q2Ana7cHdBu9Y/eH5qecvkxK9q8EZgcgAAAAASUVORK5CYII=",
    "FaKey": "iVBORw0KGgoAAAANSUhEUgAAAGAAAABgCAMAAADVRocKAAABlVBMVEVMaXH////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////r32jLAAAAhnRSTlMAfwkn2/wpgPn6Edoe5ctm7BIHDAvi3VcD5/H7wg0coojf7lPtTSFCluQdRlkoDr2eQSLwMbxkAia7P8XPkdm63j7jxM6Q04mGqYKbH5qZqNyBzS2z9hM34YRl68ytk1WYwUCn6iuSlzX+faFjRCUv1Pg0/RqgCg/YPYzmWoPQqkiL1xumxhX2b90AAAAJcEhZcwAAFxIAABcSAWef0lIAAAKqSURBVGje7ZlXV9tAEEZtbAnJcgMXjHGn2MYmtFADCSX0FtJ77733XuZ3wwsnWplodrWap/C9zu69tqWd0ZE9Hmfpra+uTCVLml6a+frty+d1j5tJ1Tb8YEly7XvKJXxwqQr75tPHHhfw4wM6/DMdj1RJfLzQAbZ5lT8rw+8MAZr2Nsf41mngSj7ujL8+BZxZHHLCVxPAnfbD4vxzXSCQ4Q+i/DtCfIC04HfInQfB3B4Xun+Og3AW3woI7oGDrAmcL3AU7hN3CD2/x0Z98XjfxQnLzfqOU7CA4I/O7q2cNJjCXT6+D+lvNwJ/1wYy5ori4xIMIJ8/YF7cwkyiN1yzUbMXDLLLZ801nWcCLSHX17p+xFz9gfPH5uwFj60bIuZqCJ/TNeQWariOfUy5hgo2EEHWuiHLlDdRwQlEEGs4lkzZjz6jgOhPpLL1XkRQxwSjthcZoI4IVjHBhHVHP1vfRgQraM8ssxsmLeUtRIBPGsOmVezmJCIo4m0/02LiZ6zVIiKIcgwW4+be6sGGh26IIgKNa3aNRNRYTI3071PSEIEOknmBCIZlBZcQQUhWMIMIwrKCMCK4Iiv4jQgKsgIvIrgmK8CevnpkBb+wgZCU4/9EJ1pFTlBBBVflBGWM3yzHr6Zo+bBMzNeDtHy4TszXh2j5UCDmJ2K0/AfdtHxYIOaHW2n5c0FaflQl5t+n5YdOEV/fHC1/k/b+SdCeL306S8nX8k8p+1tiOUfY/99XymPuzsf5C1uXiyVd60qGb3nbXuNvPQT5zwOCb14F+dUgLT/9jJavPSHmd5Ly06dfEvGd/nPyEGgF3UAsmKcWnKEWwP8jUKgFBrXASyw40kQpUPxep3wPTf85EBwIXBUoFP2H+08awwXBNkX/MafpD0H/YQ3NfsXt/mPKDjJJi7yNrt73AAAAAElFTkSuQmCC",
    "FaLandmark": "iVBORw0KGgoAAAANSUhEUgAAAGAAAABgCAMAAADVRocKAAAArlBMVEVMaXH///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////+3y3gDAAAAOXRSTlMAROZvu6qZiAblVf0ZzBIB8/czA17ud+goDvva/i8f4ucJifrr9QLJ5IfS2zAg6RqAeid5foEFTXZXUdB/AAAACXBIWXMAABcSAAAXEgFnn9JSAAABKklEQVRo3u2ZW0/CQBBGhwKdBaugQkFRbnLxfr/N//9jEhMTS7rVbmd88TuvTeakPWk2u0tUguR9vVyu04RMOKw1z+STnUb7VX38Iw/kG/34VHP6rN04kW1263tK4w/iruSy36x1Kk9P6pdSwOh+rNPVz6a4Cxw/zHb1E1T8qJfT1U/Z4t6ufkoUT9KRBDHh8a+6HkswPxYf8q1UpB9fF3QVFfKLB3QtUfwuvRJlJnzzNd29VelaVLw324xfPSzEjMXTil7ElGca2AoiEmMggEBB0LKdf05sK7ggx4bv0GJHAOiS+QNVn0IAAQQQQAABBBBAAMH/Ebh5ZLv9M9jMckYQ6Qum3i+odogAAQQQCA42gBrqa9n2umZ8MMsU2QqmON6HADcgf3AD8gEOi4rKZxZXMQAAAABJRU5ErkJggg==",
    "FaLaptop": "iVBORw0KGgoAAAANSUhEUgAAAGAAAABgCAMAAADVRocKAAAAvVBMVEVMaXH///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////87ucsBAAAAPnRSTlMAIfReQgnbZjP78yX+j9yEEvpASo3VEQaQFDhcYewm7cTDri3SUVZBCG+y9ZjQtlk6fUltkS/UEPjZXSQp3e1pjeEAAAAJcEhZcwAAFxIAABcSAWef0lIAAAErSURBVGje7ZlpT8JAEIanFz0AActVCsjlfYG3qPv/f5aYqCGwkEmZRYjv83nnfZrsTJvuEgEAto3nOv5GOK63Jr7ctNXG2MX6qvzugRIhl1/x/EL5M8OzVvCgxKhq99eWE0RnGoGrBDnUCBxJwVQj8CUF/t8K0iAjKVMQZH3RBBBAAAEEEEAAAQQQQAABBBDsnaCU9bCotCv/yRBA8H8FSSBIgkP0XaLbCxldcbJcWGOUhf089Vhtd7WUX6+wCvsUsta9vC8KJryBCIk5OR8LdwLHFWYhFZgLL4t3v2PUuh1wywp0pIzSoZFZwTU1zQqqNDYraBHdmMy/mDXcm8Ft7lhfLX0amcqPvq9yHnNm8l+HP2PZOK/Jxz/dN+ZH34rbviDt2MKHCoDt8Qks7CgoF1naegAAAABJRU5ErkJggg==",
    "FaLayerGroup": "iVBORw0KGgoAAAANSUhEUgAAAGAAAABgCAMAAADVRocKAAAB9VBMVEVMaXH///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////8OI5egAAAApnRSTlMAXqL4gG/+s02RxAEC+TwMBO79BffzAzfl/PLxLBfQESkZZuzf9frLRO0zVxZW5y0Y8A1VBzUVLvZ3mR2E0bdA74cGLxA0m94+Q58eusgoCNvgISuxJcrPC9Sq4hy72Q8jqc0JDvvceuFg5iLHqMyd6m70MNW22ng/IMWrsDkaoIbJCpDYFJrDQjLXE8JB6ycxEjY7xoisOn3k6Sq1fOPozq/TtLndrxdKcAAAAAlwSFlzAAAXEgAAFxIBZ5/SUgAAA/pJREFUaN7NWvdD00AYvTYJbaCVFlpoyyyUDQIiBZS9l8hUUVFREFGRJSAyFQH33nve3ymyzDXrjl6i78ekeWny7r7v3csBQIConD6nsy8nCmiD8KQzcBO+krMa0GezyXAXMfxpyvQ9fDFEYDXGUqTfZwqBYoSaw6iwW8xXoQyOGyKDV5brgAq4zNYGRR/HZkAVxPCH90yfe8gBMWA15lJUVhrkelvMtyAR4on0PsBdgsTYzyZg0uexdrgneLD0LsJTVhqpJwspKkuud2TBMUgBZUlHZZQ9BSlBSm8vWwkpwsO3IPQ1xlRIGRGmaIrKKuq9UAY1Q/zGU2Q6oIYongQfoaboBvna3iBD+xv80PYG8+BpsZb8KTUARMdrOEyf6zHRdCgVW22MarFzBhQ7Hcq1Dg1Hh5apR9MP0ra0/B/GSwfrqJX5tZiH/Ih9X7Ji2fclxL77h8wWSfqDhusQMsbE4BYg5/gUCC9yt0X0ZznXTg1JRxfH00r0HSXIEirT6N46PsauIvRX/tx3F+cNNtxFIPrDm4JzjLFNfN9dNHAncJax6Bu+I6qnW++iaF3qzyWXxgkv7xpl0PPMaBciVmmyFMvDDfm7nDLPz9QjenuXBfPbvuwVnkusZ+QmdyL4Jq9giKkC0Xu8YfsFjocLj1coTZjXQLnupGXZkDin1RPTisQ5tqw0RYJk4FEZ5z4uXCkH8Klcng9+qs5Ue5VXmt5bpV52p8ACRjmwNktkN7HNGFcyG53/bjtGvQkpb0Lpm8pxSmH7o60Z+JKwIFvMR3Cu+Pp3ojd2R2Bc4OIbszeytRu8C8cXdTcij9y2hteKGQbrZ461NpFofvYxLV+Uzy/KdIT7NOh/oZUyoKeFBkt/DSnhEhDXbiLHm47RmtHuQ4CAfquA1RIXOf3YKy+BZ9l0ACSQ6vNqHqaPRFmsRN6W9RbxRdFzWHq755C1jP9Tlk0+vWeMH4THOvkJdWWfiHyRr6RWPuMNGGkD3Dsl+vfsA8kxHpj/9/AOWV9UbeiVrcbcBVlfJMz/xbYH9UVh6SZJx1sQpeiLdixxXZq0L8oTPmLdyCB6fnCkDsMXpWUCMJmC54tmVwQGyrkyi+eLJoqACdsXJcxsz2/XTAK2L5pXiXNQX1Q9HGq1rg9XE/giO4gh80VhYWS+qBJMYfiixb37IhOoc2jpi4rvAfCiVztf1PtGYR4F74sEM7Gw303bF7n70VShk3fS9EVOvlMk2gD3mZYv+sI+kxx21YbvNOinkRqrf14UVD4YmATK5UWV9JJAQO9zDlleRJwPKiaB/+JjKUk+iJcE7vmDtSeID9Y6fHJXKZ00Ng3osO1Bj40bOmw90WHzjExepIrfPkK4GPZhrecAAAAASUVORK5CYII=",
    "FaLeaf": "iVBORw0KGgoAAAANSUhEUgAAAGAAAABgCAMAAADVRocKAAACcFBMVEVMaXH///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////8h3Fl2AAAAz3RSTlMAGQIG9/v9qv4BcvzMVdQ5BfLo1g8DCECDtt+lvvrApPEHOyMOs5kYQUvliKwr0E99Gz/F7k74Gi7d4twhXqFstYDgvwTszUzSMuHC8G+ytzbbxqdpMNp4ICVoKgvzgcRDJMnTFJpcYULXDfUM6Zbe44v2Zl8msB/nqWuPLXW4N3q9Wzx/5iLHqEbriSetLNFQMXQSmHm8fsGiZEVK77H0ktVU+Zc1uzpjRIerSc+RUzTZulibfF0+hmdI7QqvkHFSlTPYOHs9HsNlhUeKVnenAsnpAAAACXBIWXMAABcSAAAXEgFnn9JSAAAD7ElEQVRo3u2aZUMbQRCGscBBCBYIwd3d3QoUCoXiWqBAsVLc6u7u7u7u7t7OXyoa7pLdTXK39435mJs872XvbubduRgYzMVc8Ix2j6fLxeRHAsAxc/H4B93GBWCJaHzj/Ak+hIomoJzkg8RQJP4ZyZQAvBeHv6Bhmg8tovDHts7wgUkVgR9WALORQZ9vncviQyZtvHmdCZsPg150+QF+oBbeNPFeGbHqfPhHEb+qHzRjlNbaW9+zA1TMowC3qTx1xRkwwdjrUcHqPxzd7e9ixo69ruEF8UAKb12XYIs7GYSLYV3wPo19wDeytOMTy6XAP9K14aN2MCAk1pHxDha2ICwsiXxDVxAaxAfB0U4wH3IJ/AMmwvlwHs9/aUqBDxIsfxNDgw9uOP7Ny1T4sBDDb/akw4dgTHGfT4kPVWiB1bT4EI62TlJqAoUovlEZNT54oAQe0+PDL1R3yacoMIQQiKPIh98IgUU0BUYQNVqi0zfP9iRZWbpfC20ipnsaaQpYELlMilNb5HBnrUKVX9JASM9GrNBGXLL0ooeFI8LnvCAI/NFMV6C7wOd3rxZj6tYNgsBfRBvTzOoerSNtVW4TBMa0XwLGvcOX1FcTQgh8OeIL/pyM0uM1mh5yW3Ra5BMnV8vWJLtYctsbQAi4s443xalNAwzzsvyS9WilgQiBFNXaFG/m7u7WngvWt53ZIARmrMrbMNaHb4J6nXk8xi6oi9Y9ZciiWfYuKHSQX51ATkMm5iT9abNP+J6QYL51KMkIJdADTO8yVeUOOiGg0H1F3tddfZ2qzia7K8hzRSEFfs5Mwszbu4RVaheyc38QI7ATmO4jGnd/wc50JWnPV2EruJUxiYTTd6LQK8vx/Pu7KPBto7Ab949Umr0Fjr/UlQrfygbD/5RMhS+xxvADbanwYQOGH2FCh5/tgOZ3UHLtdgFo/jNTOnyT52h+nIQO3/QS5vpSWn/cMLY2lhJfieb7yjlZzul8S1w1pnwWc9L2Jzjw2yRIhzAPANcvZo9baHs5n/tHhuGv4VzgmEkbbWil/2yoBMNXcM62ZdpTBLTqyTfDmuRGTh08rCqs2/VqABXYBjPCXqDXzbMHckJ0H+n4fcd3MHaDdHPkHJLpOHR5eJrQ4aPZmYfU2/NJHX6EVZoNyaGw/c8jTTfprW2yUxZBfqV4nb2rTUCNRu5k4stsirJe22jXg5VehBteD6B2B/E/bn3RYfK9gjU8yiFMsIvMqlT1XJpcmF79TaHb5H72HmW0zfqNfazzZLKdV1ON9XmBIRfvpddUZKl+wHpxBOxLpwXaxHr5W2M5ef5HjMUSMDCKuOCirJz7n8Nc8I7/BM1pGJP8QLEAAAAASUVORK5CYII=",
    "FaLightbulb": "iVBORw0KGgoAAAANSUhEUgAAAGAAAABgCAMAAADVRocKAAAB2lBMVEVMaXH///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////9F490OAAAAnXRSTlMAA/qnDXcC/vkaAfj96OTpyowd20TuzoVHzfvcnuao6/DgK9BPEVSXNQZY3uND7fd2GT3ilINZoedmBDfHzFCdfhyi5Qf0MHgbSU4Q2RW/oAUpD4AeiS6pCJpCI+xrCS3SDlEydRO4VhjhwoQik7xaxurLSu+NbgySttoW3z+CIKaHJasLL3K1U5uBH6VnSIsKr3G0FDjdviitVV6GVPje6wAAAAlwSFlzAAAXEgAAFxIBZ5/SUgAAA1BJREFUaN7tmWdX2zAUhh0y7CQQSAmEsEcIBAh7lg0dQCkFyijde++96d57L//XlvhAZRL73iuZ03IO70eN97Et6Uq6liSkvv389X1/dVauN6s6e7aldZ1kpXb8aPOpy5Q5F3BZZD/TnKUmlW/aivc49F5WDaX0PhC0P243sV9QyhWHgL3n9WUVlO8gt7/joYrSwDs+/303VaSyuQa7creKVsEM3T+2USWotpP8/CR/VS1+TPM/na4S5XtJmj+jiQ7esb6eifJxt1NJTghtIADaEidKx5bFSqfBOzTj/VuX9y0b9vytdRt9pR6sf3mRvmPdh/VstSGgBhs1ri8LzJv01ZsNB9qP8y/N1fXK2KqrNRzkP8rdhQIM6P2X3tvW6O9KTzGdq88x/lM6j8zF599bj1gaCiYovdWN7+L3jzShltt92N82yHZ4ohW6gjJuPdfA+3RAN/+1+emyoyNGDATMsc2HtbIgPiRVgIA8Nj5o6zci4wGjkH8/27pDOxU1UaJqPyEMebX4Vk8K21UAoIVp+1WbVrStYTsAKGTa9sVLGmkbjx0AhJi24XiJnwboBgAlTNuJeEkXDVANAGqZtkfiJRdogCIAwMbiaEIJQl4AwIZS90LBEPF0oQAA9irgNN+/DI5gAIBdtYrTbbZ/JVcJAOhWBZUBAGZFAfcAQFAUcBcAVIkCjgGAIVGAG9oQbov554E7WoUY4AwI2CkGuAMfW9JE/Ntt8MHolAjgBeJk51b4/ZUGvusNWvmo0/W8l9e/bgp3QfDzAqaRN5zJGj7/Ewewl7SAzOMvv8FfM7mWM+EaK42fpPufp1zEpYYCqv9gg0TSZ+JyS2mkpluu0gB76AmjbRT/VxwZr0dfCGeVazw5tRh+NeTwZQXRH6mXM6056UOmBMt5E6fnrLmWmWR+Qxj/MoEsfA4GEBBJjp+F/ceEsu+V4FSVO4UA0lEIcEPwB8KlFVpjS3LlmftnekR/sbSIZQ4QOUhzQKm0pjX9BzKfpisNkFcDQP6ngFQLAKmrHlBs+ofRAsBFkfwWRqa58cMWABwmyYs0hxWxYsQY8MyaaFTIm2nH6pPBRyq6ZVVADScHRKwL2fn8qRWcokmGwR61dNsJP9UNRPvHEWTH36tBWf3/6rnnAAAAAElFTkSuQmCC",
    "FaLink": "iVBORw0KGgoAAAANSUhEUgAAAGAAAABgCAMAAADVRocKAAACcFBMVEVMaXH///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////8h3Fl2AAAAz3RSTlMA8QoJaQMC/gzCAesN7y+s1AT7qGruxA5bmLr8vhdPgK1Zhi7nI2tlkzvlB9XawJEQ+B8F8jOIxVMcsVTXhxmwjNtmlBNW8I6yUNa3+iGNMPk1eAZtCzT9Q5Ueri2EWskbqexMj9Kz9rgYPOEiJwh0EnmdfqJk9JJzEbZ9giCmq07zltl3WF0+JKpsr5BxtFL3dntc4tHkzChVyokrV6SFyIpRMjd6vZ6jS/Uxwcanlxa7nDpeP2NE6WhJKnIVmzka44Gl6Ejt0w+5n2HQy4tL4xJkAAAACXBIWXMAABcSAAAXEgFnn9JSAAAFZ0lEQVRo3rVa90MTSRSeEBNCAgQCGMALLWABRJoiRUBFBA8FBVRERM/zPHvv7e4sZ29nb9d77733Ov/SbWY3knkzO7uT3byfkjff++bNzNs3b2YXIQuyuvi/toxSB67/JaOtfdVqZK+sLJmNaZk9K2wbu7v5DubJzXkuW/gLpmI9mdptnX7zo1gkf5ZZ5J+ZhsVS/5wVetcz2FjyMuPmrziGzcgTSXHye6qwOUmJr4ekR7BZuRjPLHlSsHnJS6j/EclNqP+KOMoS6r8i+Qnmx/gp6/OTk3VqZcgdCudm8R7v4y6L/v9WmPoQ0T/nCxYw05r/w9UUqDqLQYxa4U8u9AKYd1cyBN2zwD+Og9wBUSP28iOvH8BaTPD3f2iaH6HgYwDYasjv3c3jn6gHXwKQ9w07+EnCf0UGO2noXSP+ZyfI+K8IWIUiA/7gj1L+K7KHBl836GCRpP8ILabR5WL+r3sk/UfoCMjZYvRRaX70B40PCMELA5Lzo0iZzAgWSfuvhB1tcUNYY5XDDs5qLVteq2nR2XKzaYt/RB0sh/yTtYZ0khB2OnlGT9ImflEH8wH/O1qpk+5T/7/QwdpU19A2M0Qd3KSxy3ppfowXNDI2I8Cp2wL+VJAlslR12DemqtwObDbsBVEREnSwFDjzKdGGrsXqxtPr4L4IbDJEM1QMCgh1Bd6jtR9RJrNgWBSKOljCy4sLlwGK4tgQZfbkW2zsd2/bqP0sobHziXIcUyC2Cvgb2P03X1EPqb/30eAr3NDFeHfUdhLDjwt09vfp5M9kGvwNTxkJlO9V2xV9TFOlW6d++JK3Yi8S5SV2gzioGueyLV/p1SeHyN+rNPpnFTSXHUI6aficaejSrX8eJ/8/AXuTVsgNMDn8O6J3vgXUabW69ZUakR8Dg+kacC2sA3IqiL5LXPnG1m/riKYjmZ9L0QX+OeAsrZxG8ztjz6fr64huAW3RdyQKPgM6GOKkliL6bOA6R7W+TJR5mJvtlMe5j5dx6i7HqKq20gMA9wPZRNkNp+JCFA5CKfkA0f4ecw730PwnAVObem0Aj0WBtfziB79CtMO6/JsdwGCvmoV3wiEEBrS0XErr1euV83r8iL3/WU70vWx2matarKG1O4jykB5/AfuQF3N38Ihc4hU0z8dmWobfxd6P1Wib7RR2CPtIwzZezC/l86NmNrvsibY9YNraeZuRmlvQUOR4z/CjO4L67Qe4/HgO0U+jlYc1+Mb9b7rZ+0+GPzumdRVonKBmsCZau1+09ZbADkqo5tfpxjfUhxYMrFnUAby/baIH6WyIbdz0LbeoTBfdP8M7yQ8AYHvlWKMvzD2I9SSZr344lUzjw6zq0zxdDKL3XdEMtdPYy0EW0jFK89c1YIkKtw1cePIwW0lU5kS3tPdhWEwRdZDBC3NGclvW52/Rfg9A/mteUQelEs6oGYdJH8IZQk/T4BVG/JNgcYp7xDagRAhK+4//QjIjqJXmrzktNinlFz+m+UFmMYyiz2T5N3kMOgB58YxofTn8AcOwA09y2qCU//iqYVjDhH9Fxn981G3YQSMw8QUl/M/vN3GrCPcDv9c0f4rHBD979HyVd8URPz8KM3vyLjgG92EL/Aj9y5hm0bfeB5qwFX40j/Na4MTYFcHgiRwe//hU8+/2eG8+HWtOvb0hs7X372Od2Jr/irwk//JEwv+IHEwwPzpdL8cvNT9662yj/7zznr3+RyKpKKH+R95jmX2PWOVB8UmF3xT/OWfc77ozzazDeVdCvwZwnLT4vUGZ+HuGrlobvsg4rktf+atN35SMcukbCtzILrk30gLvbwtvIXul9f7d4XXlnTjtxnX/jNshK1T/A5cWLGcTiJ0aAAAAAElFTkSuQmCC",
    "FaLock": "iVBORw0KGgoAAAANSUhEUgAAAGAAAABgCAMAAADVRocKAAABhlBMVEVMaXH///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////8dk3SyAAAAgXRSTlMA+O3XmfYBgPPuJzHWfoP0JA/YPd2b8ELs8RNBhEbrVZjBIWQCRSb52sBEJQtTlkMz/L4bCbKxnFhXwwQ65yPIqSgt0g6zMjfcveHMS9ENEja8O+AcQOrLB+9Pc7v+5KbpygarSYwqTnI02f18vz5iZ0jOCpAu0/eVmjgZjo91CHQN2JbzAAAACXBIWXMAABcSAAAXEgFnn9JSAAACFklEQVRo3u2Y51cTQRTFE1LYDekFCL1EukAaJARFpagg3QJ2kKpUBQvW95/7Ed/sLsxsXs7xw7tf5977y57ZmZN9DoekSuO/Rsfiene3Hh9L/BkvOUiVK+Q1QNJGir/J6rM/a8FEtT+yJPUNg06wkPNrQ9n13s56uEItx97y+iPv4BrlF8vpn5mFa7VaY7//SRIkpDXb7X/kBCk5t+z13w+BpDwrdvpv6yCt+g71/rU5UNDGpjLgDJS0rdo/nFYDpBVf1tQzUNRESglQNOuYDDb6BgZ8jU2fzVZvKt1A88aCl8OX6/vLxvWPKrfSqSG+vosdz403uMqBviOGe3yi5W6P6HmjcId6hOytx0bT99fieY5IAwpidMnMdfFFsLVKA9xCstrc1iTYBqUB73Hw7YK5beop9rlk+z+9wsEDK2MA+z7kJAE1wqPvWBnvCcZzSUArjj2wdsbtHeYTHBuxdk5jZ6ck4AWOtUm/bnuSAGHvgtbOIHYGJAHVUqdAzckABjCAAf8N4EYiGoKyFXJ96zOt72oHMvX2G/uPdCCUHjP8ftJ+gKqwAMgAsdqF/QVyDSHAKD3AjQBRegD+I6nRAzQEgAqIAQxgAAMYwAAGMIABDGAAAxhgD1CBj8BkpT9j/QiQoAfgoX0fPaAODyseUvf3CtOWjpYKj3McsSrS/phx5NV/SNefCZtO7YbcfoLzoPnd/+7vX/GvahIMuju5AAAAAElFTkSuQmCC",
    "FaMagic": "iVBORw0KGgoAAAANSUhEUgAAAGAAAABgCAMAAADVRocKAAAAtFBMVEVMaXH///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////9vsrUyAAAAO3RSTlMAu8A8Eb9A0BTMRMQnP9ka5izpCdq5F9UbChO9d+j11gsEzRgDwWMqYSjx9sIiCLf67wz5IM/zzvK+w+Y0XgMAAAAJcEhZcwAAFxIAABcSAWef0lIAAAKKSURBVGje7ZjZViMhEECxFRI1GpfoODGJcd91dDZn+P//mqymgaqC6i4e5pzmLafIvXRDUwVKOc0YhTUipNT5zdOvx5+DsYo1rauEVP/azlqvE+Ebaw0/pDY27aId7EcewFrND634MYOZdjHcUJkfMehpD80MuXzSYOY9gHGur81D21tRPmXQ8w4aw0OKkI8bzLKDQfGBAuKjBr2M62QBzIcNxHtAQxgfMJDvAZtknO8bdoqwR7ETW6btlvOHXe38HN6iowefIvzQvPH3Tg7XwGfA8L7C3yoCvlKwIVXgbXYAHzOkvSJvuwb5qCFlkt2E489vdxn4eurOdAdbRwGeXD/dVcgz9MbgUqXxFD8wDOJbBY/vGy7imx2TPzH8KYVHKdu1UvH1U2qXZ6X4W1rCYYxfdXed1ZKaMivy7V1qZq/IL66Sa5NK/FafUV2x53fCb7PqQ2/835+j42+zKlyPb+3rC5+fvv6n7ejQ6XF8Jsz3DMdfxPmOIQu/ZMjE/zRk4y8MGfkzQ1b+xPCjFj/4vsL2Lev4/dbwG37Db/gN/7/gG5N5/Frn5dPls8D7pw8A9fn0EUZg/ZCHMAE+eYyUWP/UQViCTx3lRb5f4jJCgs+9v2XyK9zfcvjQrdlKUZsfuZSry49dK/ZbNccP3fiV7s7Or+uvH0wxu5q7EclfxL3lX6H8iN5bfojlX+RDGwnmd3Cr+C1YP4Cb3UCyPoG26/eeYP0DJpz7oWB9BabMzp5c/QYn/RRDan0Ily1xQ3L9iRReMQOjvkVKx9uhVP2MFb/UMzDrc8U1yPBxgxQfM8jxYYMkHzLI8kODNN83yPNdQw5+2ZCHPzEsMlDRV5naeHAxeijurmRo/wDDMIQ04lTHkwAAAABJRU5ErkJggg==",
    "FaMap": "iVBORw0KGgoAAAANSUhEUgAAAGAAAABgCAMAAADVRocKAAAA51BMVEVMaXH///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////+mKIV5AAAATHRSTlMA2zx/1HcBqtUqb5HE3SJEFwIr+ZkzgPsSwaeIB/3352az8NZVZCYRIAbu44tN8nsWQgkO9hPcQQMsUDFaHKKDT/S7XunzgRRX/DjDOQoFDAAAAAlwSFlzAAAXEgAAFxIBZ5/SUgAAAblJREFUaN7t2cdywjAQAFA74AIhgE3vIUASOum990T//z0xsgm2LMltfAize4GZ1ertWIaDVhAgICA2OeQs/khF2OLq+oOZW0iigr+gZCKcoU1rCM2O6M3nlAJCFmBECKM56CFcSsnlJRHn1kBgozrvWHUtV65UbFs5B2DEzvaWv+0PG3vrKmfqYFddp0jAp3FzX7GX0JtnAZ6GnOsS6+nNcwCeMXxJuhbbXxu/AMM4SY8pK83X5py2CQ9wG46TdQDNhzZ9Bw/ACLWsr46vXmGsETSFVe4NLA28aMJeINyiSIB5iAon/xM3MIobQAAAAAAAAAAAAAAAAAAAAAAAAAAAAPwfYBQ3UIsb+ORsMMG1aiSgrrGASr1k3VjqZTUscPo+ZDyCt+K+/dKVZ7CBceOO8Yxfy1XXvbEuZQIBX8/9S4F+iGL6mH71nU1n/AKilP8rI1KZvsa5vacYbqCg5GRbjT3V6n57zh9IgwTszRPAk/NkfRoOgGgeR8vKdebVAFOgVCLpBlzN235IvUEz6CDLMlYArXkcFzPjD2OqhZrFLQ0TEKUFc9XZox5hmmhO17IyjGYhIDY6fgEW4jwMGPFerQAAAABJRU5ErkJggg==",
    "FaMapMarkerAlt": "iVBORw0KGgoAAAANSUhEUgAAAGAAAABgCAMAAADVRocKAAACN1BMVEVMaXH////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////ZmOKXAAAAvHRSTlMAAQMFFwf5/fz+9xuu9ghzBvAE+x1QAu8MFvjZFe3xiUKp7AlRdTfCEuU65CDuEFfdez2AHsPSinCzNc1MLddW3HpbqNGTdLzBRYisK5J4fT9EJatyd1g5fMSGJM4K0w+Vdjhdb7uwY6QjyGaPlDITuL3hYOaEIsdlRusnzPXW+jZaO+B+X8Zkp+rLSo3QsU/0MNURVJfawF6CxaaHSYzPTi/Ult6/46VD6Its8nG0M9gUuZq+XOJhurIcNIPCva0AAAAJcEhZcwAAFxIAABcSAWef0lIAAAPFSURBVGjetVllc9tAEJVjxxCyXYcccpiZuYE2TcNMDWOZmZmZmZkZ78d1pp1OM/bu6UB+n9++N5JOewuKwgaf8rmS2eLcLKMxK7d41jX3y0fRED9Ly2KIG+w/pv21UU+YmjEREObvX4Kk5cOn0wgFJwrGpOT1SdFEBdFWvbh+/ihhwPtOQfmgcQthgqXVV0Q/cyVhRsUyfv3OGMIBeyyvfqIf4YIxh09/lYlwwrSdRz/ETLhhWsiuHxtBBGDcwarfFUOEYF/Bpu9bQQSxm+1/GCfCaGXKDxZxA0uHuv6CRUQCo+qZL4lI4Y1q/o+XM4gPVzHoJZKwqtyPL/BQc2PPpFOvd05aUyg/+mb6LTqFBgbejvpPc7oCUWId1WAGC4tzKyGCuzHmepq+P/LslgKdO1XXiyRcE62aKeVJlDmIwyuKQRkc0gaz82B2AOUvtsPvH6HrmkF6Fl5VlsPnB32pwfC9+g41OAbyh/BHdoEB31B+Cfh/ReEGBvDUNaH8sxD9Gu1cF/F9ZTBT99AMGqCIeyj9EERPphmkQxGPUTqYqqNoBgYoYidKN0J0G7VAgCIiULpZGwMzSg/U5hXtQelhEL2fZnAQbHpQeg33HTgBRUSi9G0QPYVmcBOKGETp9eAnc1I+AXgl3EL5BWDucuEGd8GAAbytgdN1MMYvhNP1fdTgA1yWduuQCycAvmBT8Uc+A1+CtTC7HWYPUw7FYeTSBzu8PuTSz6MY7ML6r1rPsmUA6xOpPXMKVk41u33pwgCM+Zla2V1GK0K/EsO88/8Sb6ND6O1HLqVPLWpIN9hshvSkQUoXfVVlFLZJtnyvU+ugIuX096rO8kLkDCrVu8CPMvpLGNrYKoe4/r4Wlka5SdxgC9sooVpUP5Jxytkh2Oub8lnHLZ/EDB4xz4vGMkT007LZJ1KVIgbneGZq9fz6y7mGgtlfefUzUvnGjks5T5LlDu/g9AqfwWLuye8IV1qtCec2UPo5ZqeOZJHp+BC7wQ2x7cRJVv1hm9gCoYpxAOzXogjCymZQKrxiCd3AdIuFii9x/KPV9cMKZdZQB9QN+uT2aMfV9J9JLuqyN2qb4zzxnLpsMcUq0lhNMzitwbLUtxjXv5igxTr2ghHTj3igzcK3DTOY0GhfHfoE1i/SbCueCa6l7Jna7dyPQAbXtdzqx3FNkAVw6alHHbdOUwOPMsZyStEYa2TLFNUF/MP5+tUjmhsoa+f90I79ihfQzjbvkPihG//pH/XxioGy9fVf/fNdipeQ+KeadIQoXkPy27CwsnKukN8qqy88IW8DWgAAAABJRU5ErkJggg==",
    "FaMobileAlt": "iVBORw0KGgoAAAANSUhEUgAAAGAAAABgCAMAAADVRocKAAAApVBMVEVMaXH////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////Mid1sAAAANnRSTlMA+O3WmTDzLwIMNu7jJC7YOD1C15MxwUUm9JL5wAFEJQuRm0PO8pD3vs0ON73rIZzkzzm/PpqWWsoUAAAACXBIWXMAABcSAAAXEgFnn9JSAAABOElEQVRo3u3aSXOCQBQE4EF2RdS4xX1fkrhknf//03JIApdM2fPIHFLVfW7ehyhgwSj1e4br1kjDGXU/5som9wNtnekCn//e0II0OvD+i+ZrXWuCwFELM8DmD7U4Gwi4yoEAAlpyoAsBkRyIIKCoP+QzpD+Lt8UmdkCM/q5DIZCgQCIE8FOfAAECBAgQIECAAAECBAgQIECAAAECBAj8CeD8yS/87LovBLYh9BmS/osQEIQAgVvAYV9fLp9Wnisgy78KceYGOOQ/jdhzAuzLytgJUC8rbSdAWlZS18Dj/zxEq7LSqwTsDBt7xQ0onBgqO1XphXX2LYTPpoYPAWvjeeqN22n62ptUfOU+l1/r7rB7+Ek6fwr+SbicHS/cUJ2aaD689ESpxZv9/GNT2WQT+BbLHyI/MH2/n3mtaBdgTd8QAAAAAElFTkSuQmCC",
    "FaMoneyBillWave": "iVBORw0KGgoAAAANSUhEUgAAAGAAAABgCAMAAADVRocKAAACc1BMVEVMaXH///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////9Ln5MhAAAA0HRSTlMAD8hYRBAKM/4E98dbqFU2iAVN/Pp4zoABqY9W+5lL20Dv9CXujG2wL/jtGQl6FlnpkUPoX2xHiihMcA5RlOF/HeaEZUbMrVAxF+AcomTqywxzETWc3xuhygZO1FO6/T7jpclni69xHmYTCMHQktpePyDF2ZugqprdQn3kpOcj7M2uLbPXddyePMLriSdqjtFv9ZPWdBK3vFqdO4MhxqcmjSswVPmXuzqCY6Zoz/NytTR3Od6/XYEfYoZI8pDTtFKV2HYUnz3i0jIY8Jh55QJFQLzVNAAAAAlwSFlzAAAXEgAAFxIBZ5/SUgAAA0ZJREFUaN7tmeVX21AYxtM2LS0VSqEt7u7u7jZ8wwdsuDuDubu7u7u7u+dPGhRYaXKbJunNOXzI86W5J7fPL3pfCYJw4sSJE6eFoZq4iGP7MpSvnMOVwWWPTwRCtD48kOm90dUJM9T+W8mCJWZ7W1dEFw1jxrWlxY+5OS++ZcgSM6k6dT4Dc6HqXLYNRlHuSfTMO3MdNEsxOvLPo2weWN+7l565TvZTfxUJTpahIhLvrqtt8gKqjraGw3QEcUuf3rA8et+tnPioxOV+OF5tS+eQtd8jon/2/R+KkfMe+p0pRyq1HYq3KJqlEPg8ufymz4n+NRnVHZlnwvbZMSLH4Gql3czJW8wBdkAGYBM4gBg2YMjOEIBBl5ptgE0UCDCSswIaQewtIgIsplaZgahf4WPDMBBNwWsAgBmNQDqNl0YAXZGQ7wgeEE06O+C03F4aYA5g8AfJvctxkUzPkbisFzMFdPYYn/kQ1S+CWYuZATbUGZ8o5RnENT4DQLuDP4l/Ii4s82kCRrf2kF1YL0IaUuhOD2BCocQ4J4MJKAFF0mUQAY4gQCg8QKQEBJCEQQM0g7MNOTTAIjBgOTTAPTDAChqgHQwYhAYwkvaJ2AaUQgM8BQNOQQNUsH2TWX9Mb4AB9mwvFcXwFjsBCCCDuJpmgwAamAFHRvTvhxrRvArx/mnu1Go1qtk1n4cL+qTxrDjDz4pu+s4Xzvcvryab+47HpD4oiNH7v0ghy947mBYgrSrd+2Ct8iWdFmxGhbO2oaroionkV/qH5RLqL8IugFAnz/5+9rxYZQkD8Ht8sn8iBFijIWlnoJ3HF99KYH2QAL0sx1c4k2wD7Loh+4vxRaAFZEAYoU7WwAU0zwEeZYbMALS0PS5tak1O3a2oR1G0VvEs1fnuvArx9i59O0f8tTfqWxDFrH+2pygdPxTvSQzYwsbNDbqXqqkRQVbjrhnVnojrKh+rGpJOYlDEnWuOQVMbYwz6Nev21H6k3A/No7dA2JR4q4T02rlJXlTNnfg7XRIZdKND1K9Nm3t0P48pZd5O94t1JXkZD366eUFi9veAQMcDGg986vHeN/b6NiECT/kPZG1qZbizg1LrI3A7a819PuLEiROnBaJ/46TpR0kNT0sAAAAASUVORK5CYII=",
    "FaNetworkWired": "iVBORw0KGgoAAAANSUhEUgAAAGAAAABgCAMAAADVRocKAAAAq1BMVEVMaXH///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////9Hbg9cAAAAOHRSTlMA+O4EhGExZjOZYws3GEFlQGQ/3lzvlUkCQsjCx61L1hfGSgw1Fv7fxTk+Bd1iSI9gjl+NXgFdkIkc+SMAAAAJcEhZcwAAFxIAABcSAWef0lIAAAGbSURBVGje7Zlbc4IwEIUjitysiIrQeq32fr83//+X9QGYEQlkk7hlOt3zxkD2m8kQ9uyBMRKJ1Lruh7bTIHvUNat/es0lmoxM6t92uFSdgQHgjAN0rl//hIO00Qb0YYC+NsCFAVwC/Aag51bUOyrAqd50CECA/wZAP2j0sWsLgN4yYU3/6lLfViwggBsDXzSAGK8LE2s3nMjqP4zNzGl3vGd+51nN+Z75vTM0v2V5GcBD8/IEwAGkLzMPqnUGWIMXzN5S9s1R9cpiXIDFOLLwAehb9IwLeGfpx+Frmt9aegpa5osOX9OvT8HJ0Om0RfcGPUyAlgB7uVDhd/PLcBvULdqtwvyhwgnX5kiNuZDli+uHMThHkuVCQkKokCPJciFLsEu7GJ4jyS1iUgWsFHIkuckV7FGkYIrlNr1xRpPbegIoAMrj6hQGmErH27qBW2sQdwhAgL8OQD9o9LFrC4DeMuVN/6kK2KrkSAsN2xJYCjmSLBd6FBkvXyVHasyFLFtsHX1LIUcq5UIlRUmt+Q2SqO4n6nFzJBKJ1KwfOo+4ZbZ7+McAAAAASUVORK5CYII=",
    "FaPlug": "iVBORw0KGgoAAAANSUhEUgAAAGAAAABgCAMAAADVRocKAAAA9lBMVEVMaXH///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////8A4nJTAAAAUXRSTlMAweXQz8j6B4hu9BsG8xjHbQvJbPzsDlb75vBFxbC/BaoPGT3nBM1wmTzCo0FGidEXWjvgHECsEfmXeLtZwD/kh6vuELoaPqVD7a+VFHuf4sN/qzSzAAAACXBIWXMAABcSAAAXEgFnn9JSAAABjUlEQVRo3u2Z507DQBCEzyZxSXHiEFJIT+i9dwi9N7//y4AQIhHiYMtZSLDze3c+S9add9ZKaeX6yVQqmXeV4tXolB6J3pRI82q0z/be+9rtcmq08qMPjXJqtCoMmnOcGq2yg+Ysp0araEicGgEIQAACEIAAfhcgEtGUWapGbNleRgvwIiPytICqGcAE6PRyJID/DLDN+PfjPmjH+qvCs+O9KvhvAv01FsAfBTh0/x0QYJwOGAMBnuiAdRCgTgdYIMA8HfAIAjzQAXsgwDUdcAkClIpUfyeEDXiHVMA+cIKcpgKgq8eNCs2/cgEdghs0QBs8ZddogE34HG9R/J8RQWGBAuhhokgb73+Fyjpd9J19uoJLU2UsYBWb13Zx/nfoQNg6x/g3CT9ZOstw/0SJEmq3C1D/qXtabD4DnrdmhxrMJ0Hz/EyLkf0Pbn6cVMq87UJ363v/k4C9wJjN6e3ra0Z2JEeNL7/SxdueMqXAn/s0sTpWPlBGFS4O+9dCZV6xLzEFIAABCEAAAqADXgBWRybatlcQQAAAAABJRU5ErkJggg==",
    "FaPlus": "iVBORw0KGgoAAAANSUhEUgAAAGAAAABgCAMAAADVRocKAAAAM1BMVEVMaXH////////////////////////////////////////////////////////////////x7/yuAAAAEHRSTlMABsjQG2zHzwfzbvRtyfHGRnLvpAAAAAlwSFlzAAAXEgAAFxIBZ5/SUgAAAMxJREFUaN7t2WsOgjAQBOAttPRB1d7/tBr/iIupiZ0lJswcoB8EWDJZEYY5LqlMa3tmnWrCn+/ntkn08Ot/O/8hoO+hNJUABi4aWMBA1kAGA20XAgQIECBAgAABAmcFXv1iPJ8aip8bNLqhJPD5u4ZSGjyh3y/Gs/T7xXjyl7dwPATOBhi8ptdjP7RqPSpShA87p8Z1tB3XIi4ssCedb8Hxp0+AAAECBAgQIEDACMj9fiH/v2qs1stS3VCiE/TCOtourLcN5Zd+wTADuQPgZ6Q3twbD0wAAAABJRU5ErkJggg==",
    "FaProjectDiagram": "iVBORw0KGgoAAAANSUhEUgAAAGAAAABgCAMAAADVRocKAAABI1BMVEVMaXH///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////+GeTxlAAAAYHRSTlMA/QRl7gsBZsxjZP5epMey9frkykkQJCPNCVEyN1snaksSF1rg5UVK7wyx9BvFYu3igGHIqUeU19wYf2DmIgiODZMx1hzGpyaNK5Iw1fkWWSWrsNQ0v8TJBapICg+VMzhQHGfvAAAACXBIWXMAABcSAAAXEgFnn9JSAAABtklEQVRo3u2YV1PCQBSFNyR0RLooFgRpKqjYe++9d83//xUKD4kj2ZvNZC8zOve85OEm55vN3OSeXcZIJFI35LuL+DtV7HPikSlaWEQ+fK3a1aNuqfBcWtQ+fW1toUczjC2v6TzVPWL+T1GuRanA3nW+VsUADcBimyWA6o0YYBywmGYKUE2KAZKAhcKAoh4QAwQgDwK4A8QCdhpwB7CXnwAE+PcA9A+NfnbdASjuAfDITLgHwEN/AahmxQBNwOKEFUr86vmhEMBT5zokLr5j5Syf0CsYHZ85z++3A64v9yv8Zs07yoLpdP3SKvzucO4eDBuAXQ9KoO8xlzCCApicMABToyiEGXMJmyiA/jEDEB9GIQw5blWn+7YHkzCPQvjRqmfUqvatekStymlVY38dXzpGIXy2W1U5KG9gHZO0WvU2h3gOsxfT77dQT3pWFt9wj5JeT9uXWl7tHIhqxSsLo1atR7oWlOTPTx1SCLUqH6DJeEt5KNqmJABCEEDGO/Lb7tEIQAACEIAABPgLAPSRWYEALxIAXg05trAgH9CQE+2CnDVoTVnh1JsKdYbfUEpa+CWRSC71BQrVGLMe4KdwAAAAAElFTkSuQmCC",
    "FaPuzzlePiece": "iVBORw0KGgoAAAANSUhEUgAAAGAAAABgCAMAAADVRocKAAACSVBMVEVMaXH///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////8+w1HyAAAAwnRSTlMA/AnmqwRO7gIBVQUP9ygIDfUhB+/5hjj7W/3+IELxDpl6PKMiRid0eV+xFqGw8/gVztHq9NobmCPI7GtMjy2z13W4VtzhwvBvULfbnTvG0G5zEZc1WTp95KZEh2iMkS/UclMaXaDjQ+gkZ6rtKfKQUjMU3Rlc4oAew/q2tWJxYK1kba/YZqSFqUeK0nD2lDJ/hMdlA6jricwssjFa4Byi5YOIaawMMFR4wD+CY+klykkQljTZWN4+xMmLdnu+YR+awfq18k0AAAAJcEhZcwAAFxIAABcSAWef0lIAAAOgSURBVGje7ZrlXxtBEIYDkYsQJAUiuLu7leLS4lDcpaVAgUIp7tTd3d3d9f6yNj/C5S63e02X3U9lPt6+vA8Xdmdmh4hEG/G/x9nqxYxuv5bCYTkR++hGMW2K9xWb8Pvv7adZ0bYZt/9DmhvnMRNktGVsdcDp77iTB6C/4ATM8f3pLTgBmQAAPY4RsAwC+GIEFIAAMowAMQhQjREQAwKEYQTsB/h/wrmL9ABAGk5AZST/JOPNd4OW/m4BeHNR524LwCDubJrF9Y/BXhBUsWx/dSD+gqM1sAD7SJTMJLO/kxcJgL8NA/Aj01S0MYB2MoBiBtBBBnCEAeQIC90Xkfw1Bqs2kd3CL7d5FH/FPKscw1q7PXNL6j/rKwj+tufYB80HJLGvyDdtNHeozenXC8OgJKD6nsBNdY3crkjeWi4xdzbL8N/TuKzLzT4+W+Q7EL36OUjHPdLuOvOydWrVIY3U+Gaakqefs7nrScIAc1AU9cyJhsckRXmCnndZC0ANd8KAmyLCAAlpwFfSABlpgD9hQLycMCBFRBhwgzTgFGnALtKAIdKAHaQBWtKA7aQBH0gDOkgDgriWgV1Bt7NqojEC6tltTO3I6kPd5UBsAEq6Zu9QG8+6UsxgS9drRzkgl/PYoMcFoOyN9qP1vDz+A1dXEdmccyY8ClisaaKhS/w7gIpaD2FWENDfIks0Ttd8g9+iNzTwtWNB5nZXUVSHBkiAA3rsOcfz1QTa/oICUnstx5BlKIBJKGCGl7Xe2CAA4mCABAU/L4YiAGJhAFfQLBsB8BMGeAm6kcF9PCckH+NAC0oYoBUA8ILZ2wTfM97aAqZ5K2IVDJD8LwA9M1C3HIauQHPRYwDgEkQbapaETXFWphVQQAgA8M6KtkhRzrpDTtnC6wHoUhQBljpzm/fr4WrTEWiSC1Q0dSXPP9kFLP3Gm13on4R7n2yQChf9ZmtfgN4mPJGAAVwGLIQ1MOUFNAAda8udCBXAhLcQAdxXV6VAda6oAPoOSxYMlyUhA1xKGNWowIRkCRlA5zF/Bh8BVSk6gC5cU10UEI2tAxCxpnIWELnZoQOKTSKpAfmCIwxgduCYkKoJGRDFDDmqhGR+qACdOWcnigV01wQB+Z6wnxvxYMmev4CclbyyR6GOwvPdXuUBSYzlP2wiM69KubL07h5WFa4rlWQcfaDtVFg95w3Tpoe0u3r39R30PuyjvKICDqnvewwplQ0nNHYbX6fYCFb8BopBch5NEncuAAAAAElFTkSuQmCC",
    "FaQuestionCircle": "iVBORw0KGgoAAAANSUhEUgAAAGAAAABgCAMAAADVRocKAAACQFBMVEVMaXH///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////9tN2DHAAAAv3RSTlMArvf5sAUBd/76K237+ELWRGwJIm/V6CQmBHkthKesscoQCt2CJ2soDuF/A/USNkACadoq8zT9qthFQ2ElVCPIZs3xMtc3icyt8DHbO8Fk6gduDOSHBnwfLoAeRmBf5+yPcPbcGL2eQceO0ZgXfqLliO+N9DARNRYbwCCmCy9TlhXjxKXJZ6/yMxT8OBm+UVWFitKzlHUTVpkdo+Zl60ssDbKT4BwhxstzeLvffelo7ozU2bq/XYEpznEPlXbDUgJV/QEAAAAJcEhZcwAAFxIAABcSAWef0lIAAATaSURBVGjezVr3X1NJEE9CIHkJpAGaUAQUQlUEkY6oiKIgCKLoKfbu6XmevSsW7OXOetVerp93p1f0/WuGJCbZfbO7s+/BR+fH7Mx3P9md/c53Z5/J9NFYsr/hapPXZvM2XfX4k0cXe2JDjrVeURNMqbfmeCaODrqjbGGSClrawR3TjKJPnnlK5Vqm7VMD8MtyU1WhKbnZOuEX1NlVnL24qAP+UJ4qYWcHZHMyo0aVsnXLF8ngT5mjSpt7Oho+JSNJ1WGpthQc/r91qk7Lm4zBHz9J1W3rZ4jxsz9TDdi4BSL8yvmqIatp4ONvtqgGzVLJw2+bLYzvKf31Atdj/h7O/o7jBC7etLbxUnKkMlz57TXbcVU6kzrZ+fPt9XYqydN3LGY5dxYwzheTfb5+1gf4B5p+ZPib4RP3D8Pd/l8f4y8vKkuDQ3pB/mHww0kXr9xlwrUO4KUlDH7bdoCb10fMcKHTqoIsGH9jieBklt8B4/6n/T75AvTL7xNz7yYosOIQ5WYF8f84gSlOb6DQCaRTO4h/rhnF749WQ8HkPt8HJ/gLWaG6wcNA6BNQP9zDlsCps4DoWkeCRy74B7TVI+Co9LQBivE2FD4UH9+vgBNsoGA81opw7T3YRGd5OcSSyvnYeAZ8Br4iQG6djY9Ud1EzbIfivbHheQwWqithlWrlGDnBMfAQxQiFydLVR9/7DJykhBZJ+o/B8OLoaBmngHVHXH64QY88IPOoFopeGh1dyKuAO4+MxGtLRSu5Rp9DsZej95c0/g0g21R+XPvzE8QEFl94rEFQ52vnQcXxbzJPYaUfKSU5etTJJJLGN8BehRwi5ds2Sjl8xyjO4cF6aXj7cR910K4xJH2YzhVZ/FldmqJTDXsmBUKDRZLwPd1TNWTnYjmPFJQVUvDKGkBVpWxhuXtCoy1Sq+OA6sE3TP+3odGVEvhrlkD4w2xJPhgadqLhazNg8cWRzF62IALsOYjf9TMnJEtmglxQFfVyb6RZEktUcwm6UbzkB3klNjlHC+9bLrpxDUqkqbY11CjuB1zBH7ROGv4AhiQ9eKqwUvjtPZioZjzZbSfxz6/GBIXJDknXheQEv+A6MBIFZys5wVZUkFmiZD4kWLSgVOJve3C7XLrbFrPdOPxo0RfIFiN9Cx9CeBmxjWLpaMzeS8dpYzVBsUi+a+RKa7/T2X8a26/NF11ANDdTf8Td34Hz94quUFTX8kz8GFSlYgTICdElkLQziQetChEwJLzG8m7uZvGGEQpH2Am3+8kJiuw4Hoo1i0TuW+h6c1oUMR3TDIlbPz3BXLklZbVzYhakJxCIkYoBZENKm9NR6+X739S21NzcgF1yS/Q78NC2jyvRWmn3p9wzBj7qrOWmaRHpXMxNUyfcfONmUodE1nXIt5ZDVpXoGuR5Vhfoao6nJswQ5JEduzkuau+bo/tQzF2f2Xt43TcXVwDY83c5nXPzuftr2TzGTywrPvQjkck0Y71+/M50TBe0RPdDnXUvstE6U9dTo4J9ahzhJbc8fuaw3HPvOTn4Crnn3pEKZJbBn3BXx5v40Z0XkILvyyk6X/WzDyMUmXL4ewPfJewd288eItr7pz8ZBGW5vPTV6Hwc4mssNLuJs5HkNhe6fKP7gUugubFlZdBmCw62uMYHPp6Pet4BBs3mDK+2IlwAAAAASUVORK5CYII=",
    "FaQuoteLeft": "iVBORw0KGgoAAAANSUhEUgAAAGAAAABgCAMAAADVRocKAAAA6lBMVEVMaXH////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////XnmiNAAAATXRSTlMAmdbb8wzx7fxExiUkmj1IQqkJ1wMIDZPB+RvA5O5DCuKAyEH6F3OhtcRmBCgOVtx6vcfrmCEC0LH0kra7Wd+CY0kLm7+BztjdGb77e2TTpygAAAAJcEhZcwAAFxIAABcSAWef0lIAAAHHSURBVGje7ZjHcgIxEAVZ8pIxGQw2GZxzzjnN//+O04Eg7erN2D64SnPuVlNbGyQCATt2/mgi2a1h56lZo/jm5dKvW8+DcY8ms4MVYKs7itPs3APLw5abIGVCxuVhq54nzcQNy+PWeZS0478+bJXSRPwAbkXOSBDArf4hCQK4tf5AggBulaokCDCsNEkCuLVBkgBu1aO+pMeDxrDy/j/F41WBW64/SW39+we3Ev5kYkEbwK2u75UMtfXrM6yRHlquZNduvJ9h3Oq/6sDFlVXfVyjDGujIo7LhG8Cwxhry5MD0DcOtSE8lT4vG/QNuZTVXsmz8CDOsNxW9Mm8iGNZQvdOK5gDD6ihoBdgGMawLBb0GAgwrpqBJIMCwagoaAQIMi7vNYls2YAM2YAOiwNwcF9LuTyxCppHkBiYWhFIzIwh8WxhKsZQg8GWBKFUlgU8LRaklCXxYMOqIAg6OFkSBAo6GRYEwjpIoQDZgAzZgAzZgAzZgA/8lEEbJ6HSAYT2iaHA6wLBuUXRvOsCwXBTNcf7bn7a2MbIxe4ZiWPu7giMUy8rEADIzf5DlWMk7E/mSUg/YLKvlBD3v7HDQyenP8DLLjh32vANRCBkHH2jAEwAAAABJRU5ErkJggg==",
    "FaRandom": "iVBORw0KGgoAAAANSUhEUgAAAGAAAABgCAMAAADVRocKAAABVlBMVEVMaXH///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////+0VpgTAAAAcXRSTlMA/QnhzNj+zhMRAQr6AgXCJxLpzx71Fmc90A6Yp4jutNHwMX7GB9q6WMkPMvH2A61L1ttFhwb4Yi4zV/yfmdx/HWDmhMeoCHkXHMHltvl4OsBJC9R3fIHEkFIUe4CFZlEYb1/q9GOmJZYVOT6l6Cl2udN+2w8AAAAJcEhZcwAAFxIAABcSAWef0lIAAAMJSURBVGje7VlrW9MwGO26tYOxDobjMmATGCKCqJsoiKIIoiIo4AUR8X6/+/7/Lz6lraVN8iZpE/TDzscmzzlP0+a8b04Mo402/h+YT6zZycYlbfxdT8HFrU5dAjfAQ75Dk4DlC0BO0zsUQbMChNCzSqBbAXQrgG6FqACc6NAsoF4BWAozozZQsDLP5/ywAhj8/VBlDNsDPP4BG0BAIcMazizi/FMZABEF9nDxAsZ/8TLw4X4HZLhvkM0/2AciOGNiAjA2zuIfHwMx7KIC0NtD5+/pFeSH67gAZJdp/JWsKD/c5wiA1U/y91vC/NDiCUB+Os4/nRfnhwOuAHR3xap7twQ/5y/ycPrUUf7SSQl+17X5+xFGnJDfWZPgP9zJVYGJw4WAv7Aky2/MVG3+1NVAoCbNT0GN7XsPKDxzjHrArpqF4bhzv/RHblJe9lqBLoBVZWckOveX//zZVZL//Dl6RcOrfvRXfIEYdPgLy3UV5r1w7mf/2df3JP/EQ0ZfVOaVw6FcMPe7vwkW6iT/aCejsyvzC/p+05v72z9RfPpB8k/eZfSmZZE2xLPkbZNt0NlKuu7aLSq5IbZBx4pRgvPBQj1Yg/CLhDh7JTq7IbU+HrbW2QZNNAT+pIkEp4/SJqWl2SGm3b7zqG49L8nzx/f1oTlNqWtqCWcSaislQDPoxwr5EYNWAsygVQA1aAXYKZL8myV1YQitgw66pPWt9GEIZtD7zbhVyIchtA466FRdc0prdphBm9vp7fpLk+R/9zNasluvk4chmEE7H7klk6twtOgHmP3mD+4JFH2OAq2D/nvmnEvftqAG/Uas8cKiCtSg39qCrSOigBn0K9Kc9pKFIVHUAvFV+fZd5ACyVMCWL2kYEmLNwX+ApGGIgkMgJww5nmNsyoP4ru4ogRuGVGg7ZVlZGJI+zmnpDqQO/mWkpiAURMMQXge9mEkXhigLZhlhyIZAtDy/kSwMURSOq75C0M1v6OY/5isWmcNyomsuDfyGbn79V43aL0sbWtcnXRgiiMRhSBtt8PAHjOXjUDrV3ugAAAAASUVORK5CYII=",
    "FaRecycle": "iVBORw0KGgoAAAANSUhEUgAAAGAAAABgCAMAAADVRocKAAACeVBMVEVMaXH///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////+CY/WAAAAA0nRSTlMAAvcStrIBB/4G/ee1/An2+vk6+NCbBc4O9dbp7QoEQcyeIgjwMbc21RBdJLTI++u8ZMvkFR+NpNIy3Fs85itP9BEWnKEl7lh8xNM4PcNHN9sXneUbaDnqSsmFii2zE5l6GGDHZYlqjpN0O35FiGm7fV5jREkqkVM03hqqkC4PlTMUmmbRolSXWYzPvlyjQqnsKGuPlHW94cKERietS28NVZh5HMHGDLEwczXaeN/APyDFh8oLsHegPuOl6IZniynYdld7GYAdIW7zukhsUrndn2FelKmJAAAACXBIWXMAABcSAAAXEgFnn9JSAAAE4ElEQVRo3rVZ90MTSRQeEg1LuAgJJQFCU4pIFRBBREEOpQooRTzLnb33s/ezlzu9qlf0eu+9995v/qJLCJudt87M7uzOvp8yb998X2ZnXptFyLT88uvV9Gq/f/TS2T9nIemSeaUAEzJ2IU8u/rJ/sE4GFkuEz52NKfLJFFn4Q0FMldrDcvDrszFD3hiUgT8liJlSME0CwQnMkU328fcpPAJln+0DdAhz5YmXbRJ8jQ3kHnv4K1KNCKqK7eAnVAIw//7XQ6GeR/1AWZlgg2ADgKp7NaY9uRSoN9ggGCCBXAdV9UEXqS/60TrBWyRQp6b/ASzhIeKd7v3+PhGCKz4NZnqSpk+aDhiWqfpAI8ZZ80UYSi/FXaqX1M8F7ncod1xZ9pw3OgoLxY+811pjIBlQnwGWsD6qWpQ8MVohtg+zVkYn5bih1p1DEqQeQ1NL4qMFolt9oA5n3ZG/msESXnozK/57joWEXLzqTuVOlm/fLSnPlacwCNbJyqSdDIL7ZRHonEGVhQmyCHTOoEqFuTB6bMfsgg9mGFhl0AjuNQT/MH9XSVHMeDfXsOwsjeAd7pzDzx6trdKMG3m2F5Jp+Km5zAlTd6/Wb5tSz7YuoZ+hSnpqz++cVEQzZ0XGzOdZXtBAsfYsb2V55Xo6fqCbmaMPUMz3sFN6FxV/nY89g5ZvVrLN0xIp9rdT2BOoke4IpyrZQrEv5thTI10+Z8XfUexDLrb9A9R3+kwLc8IrNPuH2fXqPPqpOH7Cy5hQ7aHZN23vcwlGuqZaBsN1VrI+eetIjkiky3ssi0qwhxctnlrTnq2YjnRDT9MICo1iY2jvTzvVc+sr5zdLX1AIvjWVewKbP4q8L9dMvtmX1B6m3nSS7m/jG8yj9wLzZSXAhC5HWhhNvmKc0/ck4Q8uZMW7aXIIPgeoY4RnzxXGClSsbdbrngT42Un96fHBx4LwZecj/07ZpnOBOeBofhrx7M2t1mrNiSp+AN7FLAcLaI/V7RWx0SMi8FotcBO4AOhMR9Ve4Fy0nSwqF8Af1gKan6hRPX1gAVpBF9r6eOE3Ih0XCdOnRfqZAL/GY/korgFA8ci0IAza4Q7rZ/0FkIbCai8Fb+R22fGmrQDq6kTO17mArZ4BXr/lj+8wuPBQHrQXEHpBgquNbuc5wLnEbshZAuD+imguk4prp+wSwAb6d4R+9tJdgC+JG8Phhkzqoxkknm8IDVtxgcS7otYbqc88NbDS+5cM/R0i+Lia/rQjDezpGWKUIYSPMeN5OwEZRGSj1SOGzyIIEJBLEVlKHBfDZxG0kYEHkdGjzYyvTsZGBKcIyN/AK9oi9P+ZBL3kTSUiv5ycFsPHJi4AxsCoakTg/TAJSsltPQ+TTUsp//9PwsYEPeAC4D80CBq0tNUjbtP4FAJ34HQauNmOlM417I8D6U1cfBPyd/QLFud5uk18/Ec0OHG+0Xht4heMB88RL9vCHr6vPzb5qAmC7Rbw8UX1buayMUGyBfwX4/XDqm5DghZx/DPEgXevNSJ4Xxi/qwyU75sUuZus7Nc3Rou7ZR7Txm2ULrP5hiLH0ZS3FzHuRN7dURisSyFCRbFoqEj5LFg4bPmDvC5YM/OBjdJ2ssME5hKORAbkNANymgE5zYCcZkBOMzhFoPp0NXKYoQE5yNCQPHox04zl/0V7LYFy91m9AAAAAElFTkSuQmCC",
    "FaRegCircle": "iVBORw0KGgoAAAANSUhEUgAAAGAAAABgCAMAAADVRocKAAACFlBMVEVMaXH///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////+SsMzEAAAAsXRSTlMA+kKu/gUC+fgB1hIKBPF/wRGwRHfw7R7slNyeMdvLKzWrr7GCRWxUayMJMhMdhK1vp4hprE/VfQYLFR/E3Z/D2EYkUl5hZs0tcNc3GPWTNhedB24Woabz3nyB6MlN8iYnbVVDeeeFyLNR+6NlA6jriWpLskDG6vSSGyDFY8ruKk6RchA0/Rq/oIZnqiku9xR7voAleqRM0vYsWn6ijTC7WcDpjC/U2ZtdpdNxdpo4X2AOgsFVAAAACXBIWXMAABcSAAAXEgFnn9JSAAAEqUlEQVRo3s1a918TMRS/tlBoC5YOWyhLKpQhsqfIUMCBiCKKCgIKIuAAcQ9wM9x7771H/0MrkpfcNW1zdzk/vt+SvHxfL3kr71UQ/huK7svqb/Y5nb7mfkdfNF/suKwCe4POT5CuwV6QFccHPd2aEOOnUmzC03S16JZVy/1hqcK5UQX89dG9/oikK76uED7lapSfjV7+UAB/arFfBtlr5Opk0Va/LHp0Z5Mc/A0tftnUsogZ3nSBqpcLd2Qkt5WUtCVnTC6krcesNTGq5mGKxlv1rSRPq55mHXYLC/7dbVL0jHNUm40rL4uVsNaujIy/5I14T2qpLTSzbdwj5u5eEgn/yALRhi9N8eH5jT2rRRu2rgjPrxd99NgEw7UZ9yeKDlQf9nzI3x/1MZ9NK262kya/ICU0Z243wZh4iN1uesiPOLYzpH6S+rN8qRzLdJ8mtnpCqIWJ1P8tMmNW9D5is5d+dRcIlvNGuc7RdJDYXkTjWEdY5i0l7n074TUofqnKjNfblAWQtxjhQ7BvvYVX7xuVCTDVYQxnUHxZj/VHcU4SfRRAjrslazh+JS4VFJPbBTAGyQ1j+y1Xk4UcwjYtvudhmP+uLs9JxsGBnH4Hgl031QnIH4OjWENMj4LcCbWp2hWAKsaTmyG/emhSK8AI9qRbBpNrQeor9bnsJwDrgDkISp549QIOpCK0NMifQWYpj3Q8B+DOzs9YIdzZeAiwgNesnJ9JQBNX+bwoDAjv9nx+A4H+HB8BI3AieXPjLPDinN5EcXBGfzOMAjR8zetNV48QM8WO1MpLQJfYpTaIP4gDORCieS5KwPt0Iy8B03Crfwx3J+T//B7WkCD+ybZXoMEOfgLSEKYjMMhGgwx+AgoRZmNg0IwGyfwEbEGYTYHBLjQ4wU9AEsLMCQzuocEefgIuI8zd/0KA5kek+SWDmpZppKaaG5rmrgI7u1Ze+CdFzk5zd61BwMkUBxwImQkahUwc9Kv54FdLgj5OW8r5CGhEeDfyJIkXp4jglSRevFPH/KDUESe/4zwE+ADuWlD6Xss1fZ+kPECeqBfQBGC78BMKvMV7o1p8YwvlCSUUg9T9agWUAlQnMZs+Bc/YZSpVKJH6jCUqCe3qBOB6iFdcrOZUSngQqpQg2GHB5VaOX3MGYAqlS7icc1pxOadqCJdzgmo2u3ExqVNpQWoWY3wLW1JLUibgOdE7opzCBqIouF1lUTD2Eo3hK7+yZgedxU6w7JN501Wfic2GEDUbSy3BdFSWttYMkS2Hi6HYcgfJlk02O36PiyyO9wmM5f12Rr+0uY61vC9tULiuMNy1sfQMe4Mi0GIRd7A8qyK1WPrNslosgQ5mt6Sx5AvTu7L4JC2rwZTIn5zrkTbgDCPUjKy60SttpaXmstyZzRvc0q3vckyL8mdHV70uiM1wkdEoi6itxt60srqkkpKk2bK0XmqDvIO9JrrILL9ZWnFJkNXu7ZUHv15euzdAboMc/EIlPYEXz6bY0KOG1ykMgWuKdQx/Guh8rCJLsGn7t4c5Ols5E0sHvzFTeY3PiyJvINNgFtlGjNmQqc8TuFL83YHsXzlOZ87P7IGV8cJ/Q78BO9UXgVts/fYAAAAASUVORK5CYII=",
    "FaRobot": "iVBORw0KGgoAAAANSUhEUgAAAGAAAABgCAMAAADVRocKAAABelBMVEVMaXH///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////9xrg87AAAAfXRSTlMACATFzP4BmWYzpvcNDJxyx8vQyhD4xGIFe+fsro+zVtx6GOGjqOvR1nRVF1qdIY0rc1klaAbU/TkfpU0ZXD0jqSgt0g5RMr2eYEZqrSxvslA223m85cYCp+qIaQexu8BeY0SrjM8LsE7zv4YkZ6psCvKQLrRS2Jo43b6fYesmeMgAAAAJcEhZcwAAFxIAABcSAWef0lIAAAI5SURBVGje7dnnU9swGAZwQSyH4OxABmWPsveebaGDTkbL7N57Dzp4/3cIjX1OYsmqXnRwVz2f4lPu+V0c21IUQnR0dE5b3vX1nVNYX/MUAOj5GlX92TgcJVCpCMhAIVcVAfdt4KUiIGUDKUWAYQOGBjSggWMAKmajyToKwqF1yehshXD91qMNkMjOLcGnbG8AJJPuFekfbQPptI3691+cAER++M92HYBKtV//GsUBdMoHaARkVvj9kVosUBvmApcAnVUu0I0HhrnAIB54xgWCeCCoAQ0UZ/39dLa1hzFBBHpas9Njzb5AaC9tHCbnAczcO3qH2eLVf7swvVZ7ALl8Yzp4Of86NukaKAauOHNgQ3n/oD1W+aYcKPRMxgi5RoEF0CHnU4bPlPafjTiDi0wA6BNyB5jAc9d57C8Foq7BTiYAD8hrNuCenYZLgZhrcJsNfCYgBnTzgA42ADwg6epYljxFXIDeFPuSv4IkAL+dy/RL+WXa4FymD6UBmPm73jRzXjdai8m+0UQBaB4bGJ9rins/Kjab5sYH7v4CDCD7sNOABjSggVMGWBygK2SnPn84bx/N54/qncEuDmCRBAcwild8IWepVryCMzhAijxWC+wTs10l0H4465mZhKUGsBIZU/8A+Q+BXTzwlgt8xAPXucAQHrjA3zCawPYv+ex4fccCr3wA8wau/4PvruOLOKY/8Elg3/qP/JZjv9j+9dQ3qa1Bq/Gn8PZ4eGGk6h8zshDR/yXr6JxUDgDIcxHYFsJJ/AAAAABJRU5ErkJggg==",
    "FaRocket": "iVBORw0KGgoAAAANSUhEUgAAAGAAAABgCAMAAADVRocKAAABwlBMVEVMaXH///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////+V6nZNAAAAlXRSTlMAFAnd8hPz/Idw6u79hgWL/udm8Znwjvr0BpEf256F+/ir3uutmBfvjTog1IH39YIQHYQDJ8ZprE/VEfkB6Xel6CRnHtzYgJJCDQ9kMMiKKK5MUZQYW2VGicxLLDESt+AcokBFiErQlxa7Wd/kaIwqzwuwL3y/PmIKTdOan+LDxbTWV2OvuahvNlqnDJwVONc8alOQlQEEpl8AAAAJcEhZcwAAFxIAABcSAWef0lIAAALiSURBVGjetZj1XxsxGIcD17WFutACdVuBUmGCDNjQwWDu7u7uPpj7lv93rHxo73K9kjfN+/31Dc9zL7lckxDSeNbMfpi5uyNpy9Kl2G0e5fbrSHH+4xciI98+fU3S2nErn7c3SJ97txij9VJoyLBhn4mulvfi+B6/ha6epCh+UyZGeeLeJsY/MEY5813o8ROUO/MC/O4X/HwahfOv5gF8GoHiR85TUE4C+eufwfj0IpC/BcgHLoQRP5RPPSDBTTCf+iD8R3A+HYa8/3FDTHYg0JWw1arE+Pl9HiN864Tz/wDnwVF9zcLNbz5ixM+FVsZMPtUV7dyCjOHzh6qDJlvZapCXf85tJPCqh3lF36IT1434bR3qcU52psc4BcYrOK0dyH7JFT7+XuMXvUs7cidTfsWQzkaUOGwlBbQAF1NOaPcJD+BLdbdWMMCU/6iLgwWBbwEzyVmm/EtVPBSjIulXCybY6s9qrTQqxKepniojxC4091y16KCCGaoYQjm29kO1RTaJCmiqvzwPHTd0Hwr6V0ID5ZlOBwLpthqF3zIaqJN4SU4DhllAboC+rQisKPxUM3IDUeQGgn3IDVxCbiB/eoXfZEYROJAbOHUHuYELyA1kkBfx0GPcBny3cBuwP6n+DhQxGmhR/UxPI/B71fuAuHy+drM3hfgNLadX9vxuZk8uG6XyTYP6s9HWsEUaXznGe6JZJ8R/eZmXf1gEb7E6UZ//2kPuE2VUhH+G/3ruvgDevYf/SH/cLvL/B9xJPKfIAgVb4MMWhLEFfmzBrmFkgf7KRLaAtGALBAxAAdwAFYANYAHUABfoDC7ZAsbgItIFGoOLIAhUhiU+hoB4g8t/Xd5MYQjI+EzOFo7cI2gCdeoKCtiCaWyBH1vQjSy4QnAF5nFcQfAoQRV4ZgmmwOQoEURBe+daQvAEMvE1BHLxOkF7534iN5qzqNnaRGRnChevuhPAwVfuBLDwy3cC+TdFIP4fSyY4BZEayEIAAAAASUVORK5CYII=",
    "FaSearch": "iVBORw0KGgoAAAANSUhEUgAAAGAAAABgCAMAAADVRocKAAAB/lBMVEVMaXH////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////unmXlAAAAqXRSTlMAA5NA8fT5ATvyOuz6BscLJArn6wcMLo6JiPz75icC727+kVNGQg5vDRJU6e74pe3TkgREDyPIcFEiqMyt8NHG0DCXjHK/44aQgDEryrmkhYrN0tdWN4RL4KLlJmkR2rs/5IerSc8v1DR3WN6gBUhNcbSV2BQ9jbpryalHKK4yuL2ePOFQt5g220/VczV4333AIKYlKm2wTvOW2YvOr1IzON0ew5lqRUNsBTqrZAAAAAlwSFlzAAAXEgAAFxIBZ5/SUgAAA/1JREFUaN7tWWVXG0EUnaSQQCMkeIA0EAIBgru7W6FYcStQitNidXd3d5t/2dLTnjK2O7Ob/dBz+j7fd29mZ96b+yYASEZ+3rZ+K3sh3WxOX8juyPmSlw/8F7me6PYYiEXMjuOrzi/0bY5QyIiszxfVshufFULJsK4Y1dD3HoKyUZ2mmH94C3LF/TZF9OEOM+QMQ1O4OP+DQigQ3TZR/mMBUCgqBHdiPgEKRkKLyOmpgwoi5wAv/4FBqCgWeUsiCiqMdT7+e1BxzHOdH+X8MCFRnv8R63zGZT7x2EaXlkZtnviHrP7nKpGt39fUxDD3OxR3OyeMCmy3yAg4aFmN8XYSafdO0rD90vyPKf3H7I6gg+0tGZS+dFqywhrIjJ5INn72OIkflxLoJfHTEVIJEeVkRrHEAsj75YVM/etqiZQZdkHfIMC18ufaTSQNMLHEFTDN0b90xFeqYkFbceRYCk/pR/TgeZGcTc4Qyde8avCz/Yrhr/D6P8jb36OxxCG6I/Pg9WvnFQgZwlKTuH5HPP8VuMq19nasv9n5BUJMmN+j+mfM3zaLmAQ9mhubTMHkYcs8KiJwjWMTvmP3i5CpNWIncER+lZliRqpU/vt2KD5Du+FFs69QIHMoxCMmcATNrqRADqOQEjEBG5p9lwLJQiGpYgInsCKiQAwoxCImEI5mZ2gtYND6E6XLb3KBqk1upECyUcgpVcc02++FVoRmr1EgOapaxUk0W0+BbKOQULFm50Oz6ymYb2ra9XsseYLjwnGLCKxjTzFUv7njvyuzmms08PIL3MJSo6ko/M6cDOHlT8Ed1Qaf8XLwCuBzexjjKWwTt46zfPxnceuoZ3lM3MReT+Yyv8SYU8OCWnFkOcezn+4p8czGxK4Qs8RVYc8lOUMZuwnwhVzhEWpOYmpJI0e6M5JDSPIbMkPyOeE8iX/Zyoa3jpH4BskuOUwZrWPdKaxB3EBBy4xFTdR3kFVKUYfUx9GwU3LuoIz6xmHS9yErN/bpTVRgmawbKahgvNT4SouWnakWS6pzuajUxwAFOuXPdZrwgyO+2k4ZR3IHqg1XkLTCpmqFwH3SN/iiaoVg6TUYu7ReA/AmaK2QWCHIJ/qVwMcqEf4yp/gaLP0GXnrzlAWI7zQAznE+/obdv1igEgVQPCNPX5n4q0uZlCkYB2S2wvr2dxPshMoUAIisi2N2Hf3f5m9zKVYAuqQmayzxV2N11wbiOYKClSvsXr5JI83nKj9kGWLTL91c09dPkFeQSgWO2B+gqKb/K/xzCoyz5CrRWmEQaKzgAxorPAcaK3wCGitcBtoqRAM/B1YPUQBoqqABP6KgCf8eBY34f+60S5v93fMUmWnKsP4ZOn8AS2DSTnTon44AAAAASUVORK5CYII=",
    "FaServer": "iVBORw0KGgoAAAANSUhEUgAAAGAAAABgCAMAAADVRocKAAAAVFBMVEVMaXH///////////////////////////////////////////////////////////////////////////////////////////////////////////9FDDDEAAAAG3RSTlMA80xL8M8Hx47JShtJjGwGyI/vjdBu9Mpt8cbmxI83AAAACXBIWXMAABcSAAAXEgFnn9JSAAAA3UlEQVRo3u2WSxODIBCD16qAj4J9aG35//+zWMeDjNPT5tBOcsopcUB2PxGKUpAZhykCNA2zWfK7e4Sp7tL3A/NTg5ExQuWlwRZYKbAFhUSwNAqsc/bIKxVc++Vnb6vcqxX063Mtc69VYLeB0Oy9WsF5C73tvX6B2/vfOaLYrpmn3KsVVOUns8q94kNr3OVx5NUKvo8K8LB74sf1jF44poauzJCWfo1d+iLBW8hNFy8fCIUU6Zp0TbomXZOuSdeka9I16Zp0TZGuSdeka9I16Zp0TbomXZOuSdfUP+oNfuugtYcRKQ8AAAAASUVORK5CYII=",
    "FaShieldAlt": "iVBORw0KGgoAAAANSUhEUgAAAGAAAABgCAMAAADVRocKAAACFlBMVEVMaXH///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////+SsMzEAAAAsXRSTlMAgLuqBAUbAQIDVer9+0Rm9vD+8g8H5Cn61RTznswIJRpFd93XmeyKE1bcPCKoJw2TEjY75Su23+nPkbXOCjNX/Fyf28Yml15TYpj4yLN1uDd6HcJgo+ZGUPUcIQwwc/k/xcqwED7jxCSv9zLRQMva9Eepa67x0g5RW+HrS7J0F53gwV+n79CxkhE1eFmcoYKmaAbuC0402RW6WDm/oIEf6Mnti9O0lbmaPR4Jia3WQ01IyOJhAAAACXBIWXMAABcSAAAXEgFnn9JSAAAD+UlEQVRo3rWa51sTQRDGL0C4SCehBwldEAGlKYKISK+iYsEOFkTsvffee++9e/+hPI9LMnvZ3du7HefjzMv+knvZbJnTNPnw5HUX+v2FRXs92v+IlMRMg0RBfhb68DuK9hggeiLLMEdPjYkwwmJ64gKk4aNdsQYzkufPUh/dkxAZb3AjLiJmGpaxvFAx3GQsLxwazjSWF/YN5xrLC1uGi41VNlzCWBXDm84kiQdZv1hUHW+5pmqs5QPkGy5nrISSbXheX5zUc/6nzunoFRresNc0fMUmWSNDbmUL3dpUAcff6DNsAzRtzvqZAqFvS0gZSDacACxmjC99SqYvMxwCxIa36kSz0s5k0sP+PXI63nK0CUTxUg0wGY/YhreQ8mJlwKThpxiGPybFJEPBA5HhSaQUjwJgGN5M8j4sgPnXbJxk5+EBNK0SzgSSm4tgcjBygfYdyRViAtKB9hjJNWIClgDtIZLbhglIANptJLcdE7AQaD+S3GxMQD/QtpPcdUwA/LSDJLcIE3AZaBeR3E5MQD3QTq1pY3YAFuczHS6Oa6ZWjGY8wBq4t3BrDqayBWAVkG4NZsvxPHABaUMwO4gHgJP2djB7Au8RpQHpqmB2GA3QBqXDofxxLEAiUMaC/AEswAWgrAP5/UgAfQVQvuIsQyoAuNoY1TzzFQAzgPB+FKwM4QDg7m4fVTkvDRAdVddBYRdVKsnAAMAn5I2ma3UIgHtwBzdhKi5H8KAU6vJNxc096gC4LY1baq6eVH5EFV4geyrczzgDZBuM9T4U7jRFwNINQOVfK16LnACo7dtN1r47QwkQDQ9i3lqWpFwJcAuK+piSMhVAE3WVspEteqEAOAc1I5xv+VUGEMX80xuUplNmKtr6BlnUOfISd7Y/dwqIpCR5/F+TOmePqJNSNIiOiBlOvsG3Yup4Xyu7qMoC3K2UoF24r1mbZv8RXaHqsU/kD4lygC66/szqruGzTcBK2rZRy7vl3GJbgNqL9J1stPXleKkdQFsNXV0uc/0+JA8ImP4n7sjd75+VBQRMn//DA7kGxbrVcoCqN3TJny7bAjkqBThi/hyl8k2W93zA1NlX7zff1rts9IimDVgBro6aC9m6DYB2eJcYUBV2zzfgttfnysoUAL7/CLtqbay03SKdywXsDL+Lbiyx3woc28oGtP1kNIbmOGk2vt7NAnQz1qSRzc6apafr5XaTESmOu71SrZEWt+Y4Ur9YDu916ZpK3N0gHv9gjGrPvXqFaPz6gHpXvySSP353DsprCTGcVbTmIdZ7FbnlXoa7nyo1vKhqNY/fW4376omnqwAOv7sT/+UZz5/gT/hER9R/eT9Hzyv65fdltv+2M7X+AojcI+FjhuxaAAAAAElFTkSuQmCC",
    "FaShippingFast": "iVBORw0KGgoAAAANSUhEUgAAAGAAAABgCAMAAADVRocKAAABYlBMVEVMaXH///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////8CSbcJAAAAdXRSTlMA+wMFBHyqZswzAkYG9fQpR/ZfSNjXmY3nawm9b7x+RcABTtm/peiaXkJqaf6RkA/3/D2kyA4Tt8HGJtpZnDp9P221ut4+xGxNuTitwijN8ZQy3DzhQYTrJ44S+jsh6svQsZIRxav9XR+GZwoU3XsZHmEl7pU4H/tHAAAACXBIWXMAABcSAAAXEgFnn9JSAAACc0lEQVRo3u2Y6V/TQBCGJ201tIW0lLYCgiBni4ByKCqXgjen931yKCCXuv+/oQlJ+stmj2TST/t+y+TNPpm9ZhMAJSWlRquzWtGFVKl2hmj+1lgLEZYx/kG2/bUHREqH7yTfX7J9Ql5sSQF+EGm9lxpfQx5gyIx0lYTQawlAJQzgpwRADwPQFUABGgFYIQgyvu0dBwFGCY6M0wAAQdMFOkCLm3D0tUlQX8LmIKrrb+MmNPF76Qo24E26pos4OZQCF1qKR3gpchTaZKxkDuHXbtiZ6WwVbEKaRAZAYpAx0hgAJuFkAgHA7KVPnPFdvCGyXacizNYpoXoQgVAWKzjMkcYAeHO4JtN+Ylu0ZLqEjzkrsiiyU78Sr8kuYQqj4FCKvkPYSCKUTNqpwhnpzwhFn3psSdk3n9eudkbRAee9khaaQAffQ9PFALAfN0CPGzCCBfidpmpyHQugE2wpgAIoQAMAK+V/lUycgDNdCgC0lEqtMu0E+v9SAd09Zr1OZi+LNs/yUwCac2TqEDpysP0UwIgbuy8CYPv9gAHv1Ormt8/x+wE93gfyfADbr/kARtL7QDP3ZzbHv+oDFOuXXz8PwPRrq3/cW3NWrL/+gVkrWvAfBmT9cM9eMvUp28vnSXQ/3LXTynqDbXYwF90PmXbrXq83aC/OoUx0P8C8Te9wQ112qIzhd9LTHjv+CVo3hPXD1TvO4sw3m+PV1nt+/QjFDzDpTuDWYtHz7b2A4wcY5xaNaH5ILNH8yyksv/nf76nfP3MTz29+v4358k1g+k09u+21Tz8EZL85+/LDQ5a7fTgL+P7aLpDrKxT6chmIya+kpISs/9qZVdWcgBHPAAAAAElFTkSuQmCC",
    "FaShoppingCart": "iVBORw0KGgoAAAANSUhEUgAAAGAAAABgCAMAAADVRocKAAABqlBMVEVMaXH///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////8tn/2VAAAAjXRSTlMASO1NnQcGqvv+Nm6CAZ8VCNZ0EqJFyww1Ot/PThD4/R5BQPrje7Vk7+cjCfEOdTe9fx0DrdENvOBpT/Q/yjS6WN4aXQryFL7D13rm8OUhCwKewCqG2IUEqUfsKK5ws1H2mdzHqCfMSyz1mBdaXybQkjC2+VmcfcVjJWjukZa/PgWLry4P95W5V+KA2R+11hGZAAAACXBIWXMAABcSAAAXEgFnn9JSAAACoElEQVRo3u2Z11sUMRTFs7hlFhCWqiC9SlcBBUFAASki2BvV3nvHXrHN/6wPfpvZSTLJZM99Uc7zmfv7ZmdvcnLD2Kb+A20diyb8ujXeC6v/vsWV6fI2FGC3K9dXFKBJAZiMgQBJBcDdDgKo6ruD1IAoCDCpAvSDAB9VAGcCA7j6SkWoAL3CnkSnHJBHsnAMu0B19r0UAHtrkQQ3555AKIIC3D4BUIoFzAmAPCxgpwCowAJOCICORijglPhHvQ8FbBEBT5H1k4UiYB0JOC3p5REk4LgEEK8BAlply1E5EHBSBngDbLNjMsALHKBFumJ/wgGapYDeIRigXr7pPEfVdxT58zUKcE2xbY6jAA8UgEUU4InqxDAHAtSposVbTP2jyuzyAQM4pwR8wwCWlYB2DCBXfS68gKi/FlcHyA2aQMF1hihQpLVKFCh4pCcKFFxHaAIF1wpNoOD6QRQo0rpOFCjSuk0UKLi+0wQKriWaQME1SxQo0irLsv4X3UCrJJkdQD9KuZRV/fLPWsAV7p5OhNTFu2f1M4Vi6qFFAweUkgA8Q4simoEtH1rUdpEAPEOLBhKA7dBi6NFDM4D90OKm2ch1wrEmTJu9Qr/9ft9hBIjaA1JGgEFrwIzZTzRCuJRmObRoNewE66HFDbP6uQO2gB1GjXDnseeRJv0m4L1KqRnT149k9FmjdjmqzxgmOmU6f3Wb78S4EOxf8MXltq6wl1Ozwf5mv/+Z5vZuIGTz/PT7K4P9knPmryD/vOifDwScFx94F/LaIfgzV4U7csn8VYGAnjDHags/KxF2G6ca6Wdsyv/ALqyf7Q/1CcL7GduX6R9F+1lhpdd/8DDaz9ihA571tw7vZyzVnfP3Mqw4RuH/o1gkv6AgPxJnRP5N/SP6DZi5eq1D6RYpAAAAAElFTkSuQmCC",
    "FaSitemap": "iVBORw0KGgoAAAANSUhEUgAAAGAAAABgCAMAAADVRocKAAAA5FBMVEVMaXH///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////8fIpK+AAAAS3RSTlMAkl7k7PZlZpnM/Vzr/i2UwhYkI8gyqEYxF+XGKxFjAZEQZ6pI7eJL1vqYvGTvDPl9h+4LBPFREwNqjnSISo0laElyU5YVYobyLhQM2LS3AAAACXBIWXMAABcSAAAXEgFnn9JSAAABb0lEQVRo3u2XZ2/CMBCGmbELCat708kqs3vSlm7///9TlXxx5LMVekeFyj3fsN+7R4qCfUkkGIaZDTo1Txh496dU/b22AhnsEPVXVkgMnbZdMKB4SjXl4IlAcOISUDwj4RIIFrCABSz4JwLPJfAIBCOX4IJAUO7a+3dfKe7Mvl3wTnLpn70EcPugeEU0t3z2mvKHeti4Pv7R7BXIBzAZCuTUJjwWsGAOBMO71oITPxT47lSrX4H73zYUEY1nqP+1IuTR7P/1QSk4v5zsO2Ny3gzBDa3gwfYKUiFtgpJwUIqfkb/6E6EyLJgDQX487Mi8q5gqw/w52/tJGaEq9igzh0fmWZhZjdZiMoUsdNpmt/RaVOYAPs9X9GJUJgNv+HoxKmO5kQK9GJWx3XkxiuNlWMACFrBgFgQ5eD2nF6MyaXjjWC9GZYrwRlIvRmXWN6D1dFkvxmUqS+b65mJ0YsBldpdT0c+s1JoxO1FlGIaZCt+dP10KJj6xaAAAAABJRU5ErkJggg==",
    "FaSpinner": "iVBORw0KGgoAAAANSUhEUgAAAGAAAABgCAMAAADVRocKAAABSlBMVEVMaXH////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////k2/KfAAAAbXRSTlMAJD30me7Y8+3WJiXXm0LBRcBEQ771TwHxZTEK4gT7veYSNuUhjX0LTr8+mmYo9jLrUJgwEVSc5Ix3/Xylex6PURNW4R0ijvrgostKDHhZzy9T2R8FFAk3eqPMavB5fhxkad8gaJEVWM5N8lJXAlspLgAAAAlwSFlzAAAXEgAAFxIBZ5/SUgAAArlJREFUaN7tmVd32kAQhTcGI2PcbYSRsEMxJQmYOIYAxsYlcYtb3NJ7b/r/rzkHEtAWiZW4evA5mkdp5l6QlmG/WUIcxNR4KBAYuzdDvInRSeNfTEe80L87bHRjeMiDz2/SN4yRW3CDOYOKSbT+lMHELNjgIWsQBBuEWIMxsIHCGiheG8x7/YjCrqVS33P1E1Vnro6jXvLu+mK7fnnzmro+wxpMuNM/mu8qJNPUndu0/rQ7/ZU1k0Zg1XwrgmgVu/RaOcibbw6NmPRdNrt15jlvUHcjd/5fn3PZ6lKLjEEiRifMBsOKEg5OuF2gp+xKMZ5if0w5zuAca5DlDOJYgxPOoIA1UDkD8J97fhnW0CxikzH4gDa4zlD6TxbgW4d0wKSfeezB7ko76H3+957sD/WNROf9xheIRxFbPY8XRokffvhx4wLLp5walk95NSyf8mpYPhWoYfmUV8PyqUANy6cCNSyfCtSwfCpQw/KpQA3Ip2La5fj0wSAGAjUUn1rTLoZPbWh36T6yVYjUEHxqT7uD82lf2h2UT7G064cffgAgUKsW1ZLzupJarGqxvmnH72rtPhL64ghjrz5etMtqubx9YjnRbYU/WvL6rbPeGK5sl/g8amrmSelRQitpKoumrRO/Jam/o1eST+nqjCrLbFtmfmX+UItyBq+Zsj3pgZTkJu+CKYvqFokr3EhNamRRkh71veUyVRmDLa6sQaBjzThXlrXIbHKZv2UMqlzZCwIdLWtc2R/icjgupt1YjSnbrxBX431L2mVn3ofW+z56e/9JlzuNPX5Elf26JG6OWOxot2xuYWu2a/voTa+lpOVPY7Xed1C27JfE55edNx3Nppycxm7vdb7E/uFl31VX2WnWGz+fOT2o0wuNenOn4p/G3ojTWGk+xQaWdiX5FBtY2pXlU2xgaRfGp38BfMPAsn4U6ecAAAAASUVORK5CYII=",
    "FaStar": "iVBORw0KGgoAAAANSUhEUgAAAGAAAABgCAMAAADVRocKAAACDVBMVEVMaXH///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////9xxQJjAAAArnRSTlMAKPuZ/YgEAQP+f8zWVQL581N3qk0jZvH2E71gS2/6WoNFywdUu0T42V3tzg/8mjhhpOfsa0yUWzxB5iLHRusn9TGY2xdAIWTq7ytu9NVzeBZZnN99Xj/FBu7PbU4vv4EfQwVI8vcz3b7igEJHCS1wUdd1VjfcGKOEifDRUJMSNp07waLlxmkMT5IR2job5CDpyquMKgvUEDQVm3w+4+gkyWdsLtNxUth2V1yfPcOLba28AAAACXBIWXMAABcSAAAXEgFnn9JSAAADMElEQVRo3u2aZ1caQRSGRZooKN0ooqjBRkQ0xl6iJhp7iybRNI0lmt57TO+9997L/MZ4OBxgl93l7uydT+H9unfe58xhdm5ZkpISSoijLPeI/9Nzdv4rzWRZmWdZ+ddpSFCa9Wz8dR0kpA4dE0CAhOVi4Z9SEgEcY7GFUhKlXww28D0acDsFHZBHOBrG9k//wgUY05EBRwhPJlx//Rk+4BnuFiwkRt9QN5AWC5jQIwJuEAFtRQQ8FAL48LawhggqBw3wQRiQhrWFXUREe5EAf8UATTj++URURSiAS+KATRj+9URCGNl5uxRgt3L/O0RSb5TVWRWlnn5pQL+ntCJLvnNhj8mq9jYQoLak2bTOulUQ582NRW61d95AaLTMWTTNXRfJ6Sr7QHv2Ap0zR4aF7PYBu4pbFQyNlBFkrfa/jPg73hEGaugM5ysfYaICffzLRpn2hACprACvQwAtK8DJEKCLFSAQAtzqY+NvDr/bw2wAY5EX4TAL/43Rr/I1A7a94Sv3LvqoQfbv5d92XagEQ7VAC5OM518eELqxf6IRysdEJgW5SP61YvksZxbDP+OgRBW6Qrl/sWTfMG5W6r82Xzrr/5lR6D8Vr664mKnEv68+fuUypyD/n66A1EZtLbT+3Uuwum7bBjr/shfQylF1hcb/wX54bdp6T75/VY+c6ndSPmBSVnl9Tj7gJl5bg9HsUJyjFjn+x2lO0ToZgCIawJQMwAEagFsG4BUNwC8D0EQDuCpjgkZ1Zc/Apzv36S67nUonXPEEn4AN0gGsYMAJOsBlMGCCDuADT/Ipq6NZ6MxfRZuTW6Ffs2gBFiBAqm8uKZF4mAoEPBJ1eOvUpZvmRR/bgIBDYgMU677gTMlZJRJgBI67hBsRszbc9j61dguGJBfC6i7Bol/NSVjntYJH+TMIUBu7MFd9lx81pC2OjcsDAXbwl2kqHUJxjsqYxnQRBHjMW+VtE4t8YuP176dAgF7ueHpcKvZoNie4GgTozIiaijXHi24uiGrPgCkn/CMYLYAsqLcYBccfUks8wYZ51AW8HVNco8GTnArP+u8Hf9fYp+Hx0/aaCz8aE3+q+F/1DyGomHjC4DS6AAAAAElFTkSuQmCC",
    "FaStore": "iVBORw0KGgoAAAANSUhEUgAAAGAAAABgCAMAAADVRocKAAABj1BMVEVMaXH////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////EwBIFAAAAhHRSTlMA+Pzt6BT+3RHgAcAkbUP5E/C4+71/J/ohB6zVFkQGq+5I01rqCPep7Chw154DzG8SVZgcwUpuDE+cICUq1Lo+Kc602JpdVBWU61JCI0ez9jLcGFtg0dZ0t7ydfkAm7yv0MHPauxteoT+CpulOU9l8Gr+ggapskHGVXIAeYQmJy5fnhslRHc6HAAAACXBIWXMAABcSAAAXEgFnn9JSAAACBUlEQVRo3u2YV1PCQBRGF5GgAgIq2BALiL333gV777333nvbHy5YCCSLIt4744x7Hr+dfGfysHs3IYTD4fiTWFmu/h1DGl1i4H5DCwWgsDhQf1QMBWEyKoBgnQLRwe4v3oESZDuZgnwKxjSrv0mAEwg98n7rOAVkP1MmuKagmKT9cw5YgaJZsgXiKDDl/pvhloIz6ndG2OEF2RYfwQZF4Fjs3xQwBDTae0jnofTTlc/NsEWROHnvNyuwBPr6N4GGoqHx9HdTRKrdghZMQSEhgxSVJWLCFRyQVVxBDInEFajJAK5gkehxBQ5ixxU8kQdcwQt5xBW4SCeuoJvcGTH7je6RcI8peHYfdk4lXn+XwXNedzDX4sOYceyPYtv7SHOx+veuWM/Ed46w4tgSGys+/5jJBTWyJaX7RlBhDDpOISR9XhaXFngvjheSsdy/5olrGyVPtJZ9Ed9IPsAcRb6XR3OW+IrClPdCk64Rb0xCXMo3cd+RGIdt10mu14eVp21dCmVG6lmCb6zNTUpO88S6YOLeoqSGNIU9L0s3zP8J/EEuJ2brVFZLc7WpKvSSqplds8Vq0DYtt8vWxOGvCl2g8paEcwEXcAEX/HeBNiJktEEJQOACLuACLuACHMEYrKBNJiiFFeTLBDmgf7D1C/J5WgJo0NtYEzunvUENQnJNPf8k5XB8eQVALxMbhs7tbQAAAABJRU5ErkJggg==",
    "FaStream": "iVBORw0KGgoAAAANSUhEUgAAAGAAAABgCAMAAADVRocKAAAAHlBMVEVMaXH////////////////////////////////////Z6AnKAAAACXRSTlMAiOUG5ufkh4nWCmnpAAAACXBIWXMAABcSAAAXEgFnn9JSAAAAY0lEQVRo3u3XuRHAIAxFQQT4oP+GHWsceSyy3QJ4GfBbg9/GPVepHiMFYpWLFJj1gSMF1gYCAgICXwO9/vxz93V95Qcn+t4HBwBjxBjxCRMQEDBGjBEAjA6jQ0BAwOgwOuDlAfYcnhHmZ3cuAAAAAElFTkSuQmCC",
    "FaSyncAlt": "iVBORw0KGgoAAAANSUhEUgAAAGAAAABgCAMAAADVRocKAAACQFBMVEVMaXH///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////9tN2DHAAAAv3RSTlMAAfsEIN/4gHeI/P4F+vHh/fce0LVxzAis9JI1BvMf8i5BsVb1MdYHMN02thAaySIlVXp/HUYnkxcRFpx9yquRL5bjJDMUe7M4GdzY+UJuzQmujy3ScA5RMpk3o8cDia1L8LImy0rVG0TpSYwqzwuwTnI0WJt8qtO3c9mBD28jRWznqeyKKPaUuBi9nsKEqI4sUHSY21o74H6iQAKnaY0rDFSX2nhZoYJtujnePsSlQ+iLKc4KkLSVduJHZmJlYF/KtNIAAAAJcEhZcwAAFxIAABcSAWef0lIAAARKSURBVGje7ZrnQxMxGMbT0g1FS1sotLLKkiVTQIZMFRAcDLcCTnDvCbj33nvvvbfmX5MWwV4ud3lzvQ9+6PMxlzy/u2vufZM3Rei/UUTed++PfpNp78/8wbwwdb07iwt36C04QNZdU5+s+qiO+6GC6LGYKuur3e+DdZ9UsALLasbGnCDs5y8wYKas19oU2rufYZjCP18WDU7fXBPnlZ0KE13hGK79RcLR8dG+VnujpL3nqgFzKaUiJmC4Tj/cmquR8J+2HXMrLUPkj3ES1V7zUIsVyODUkP7YSfPPTsYKlVVG+GMj7ddtwoqVphP60wCbenEQ2ueOGo9lAcXjcFCKrcWygBYDVlUkIFPu/g1pji2zzebJ8xzN0xUC1hyQfHLXrbkBH/6i6qfbIvkB2RLzR5u8IUY8G2LObdPyATQ99CBwe6VUQIlaHckDmELrYakplc0XFyxgQCata3Q9K6gnzgACPJT4Fj7Zw04bYUYLCFAovhrZB8tMfQkAQJ54auvvQHNf4hg2oEp07YoOnFyF8YcKcIvy44sisL8gfkoAjpNXuteo4T8KcIumf7Eq/qOABeSFUyq8/wBAKRmkcyNUuf9RABkkLBkq+Y8A2onmOKh/vB4U7CYSreMagP7p0bBwXUG0LoY+wCNgPlhGLJePQQFfYIC7RGJaAp6iR5gA/8puFdHYBQZ4mYBMXzejsC3BAwaErWX4T/V3cyl9Qwg15sr77/H3eids3cyzC9Jc32iWUof7b6ey9baAb2D2DRRSSCGFFJK8FpXE/guctvVlI8UVp0lKHUkaHsAlYeg/OLy5zJLNGpU8RblZwrEuf2MWI+/Z4UXSdGIjX+h/P8zM7QUDNhAj3/gadzIB4JUkIipN1nWURQVFJVD/lcQKayGCAe5BAWeIgSYYoBK6WShNIUbWgQD6eOgD3CdGrkUQgB68m623UN8QAwD395CbBVsDADBmAniKLibHliM2AH7/aLnokCGDDeDwbxOVpnYgJoDDXycqh4xNZQI43j9lN/4asQGJ4HLON3ENdIANwAnLYf5dlILUFAQADJXU0gHna0bKIZU9AgTA+EM18yDJTqu0piIgAFvmTZKzz9lCPWM7j8CAoV9idZRkqaLERh3i0vAAhh64ag6ltDxzjkPihLC5E/EBfBXOqtbUgNwfkdmanCLVt3cp4gf4M+yDrPI4s/lT+dFdcqcM3alIIQCmwy+RLKA2Njj/SPHJmVEYfxJrg/F/OxfJAnzxU7dCuf/FPCQLGI7P2Q6l/rP20L4Wpyj+a/oVnUVpd9IX4i2U/FKdxu9v3yRVkcml5K+YrZyzyVYgvTzL8cfDSmJ9VXST48A6/HSdbGXMG1eTL76B5z1QhANcKSZ09rEV8HJO1CPlanQuk7dvNzWgILW04qTErNUu/FqnzuZ0ZtLuZOEBqbbJsTVpnbpb4LCBlvzfe02m1l/5gxNU/vNMMPoD2rJisiIZAJAAAAAASUVORK5CYII=",
    "FaTachometerAlt": "iVBORw0KGgoAAAANSUhEUgAAAGAAAABgCAMAAADVRocKAAACQ1BMVEVMaXH///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////+2X7EAAAAAwHRSTlMA5ue3AuXWVf77FAGBHRX9ufy6BQQtuAMsZOROZWFQgwulyA5R3A367wyx6PdXo7yCpoCKKI/CJ/CyWsGsT/Mv1BC1Uz7EqtMPHuy79fQpoNsgQmvxs/aUE5k3x+uJ0ZMXiMuNkjBztpcWWRtEJcqrjCrPbfjZORrjYmyQLlIz2CahVD+RaqJnGUCkhWapzQky13VbPGAiRsxLbxI2eX5f6mnVEdo6wMVjBu4fyYvOCq+VPeLDTe1WK76GI44c3kNZ+cR6AAAACXBIWXMAABcSAAAXEgFnn9JSAAAETUlEQVRo3u1Z518TQRDdQCJNEmNUlGYQIx2kiSIIKmAX7F3BLmKXasHee++9997L/mnmLneXvdu5uz3MftG8T9n83rzJvZ2d3b0gFEYYYfxX8MSX7G6prGzZXRLvCbV22ZPi7NYMrCCjO7u4c0+IxNd35jswiMRDKe1/q346ak4MNoCr9sSwv5DfuXcaNsWY+v49lF9VF42ZED1nYU/kp2ALGLDTory3KQZbQsyBdCv6zTZsGY48ZvlhabhHqJjIpr+lGvcQrVUs+oNy9RUWr7x4e5lB6c5NMdcvMviFJ8WJbC8woNwykU9Ya1SLCQFS1lYD0lu7YcPcZLSeymTaVaOJqDNotXbD8lka7H8u8Ysr2/Ka7tC8/DjdBDsMayRb4W0Wxyv7+T86z9LEFXr6N42LcLTMi+sQhssDZg900cxzsP54k96WKzfmh+IwWRoNAWarEtK/lGm2jgZLzFKx+cRKo1EA89kMWv/pRtOF6i4POPReGPSVAxsg6hCPxQmWHr1wf1ZjUq3KoYEwlZroNnICovPfeb01aUZzojjUWwxISvfWFJJ8l2YP6kcuTrfUeJvd+gkUh3qJjogfy0n+qQRVgmtkcJH87WD9BGqHVgP8WaT+kdlkS8xSNoZcRoeeBwbpJD9zHZGgFFxQCI2w4pCW/zKo45wKtwSwxvUd0vDHzFd0jus0NTSBzaFJMF9emGjsSHW1lyibJ5tDi+TRdDUrU96ju7ThUoUlFLA5VO3V4R+UaFSTqBNTj61gqyGMHzWKfGq3WgQ+mLi5L0tJmTeZcZUJ87ngMMgPnPC/WT6eDJL0+5sRh4u0X1b1tQ7po0Fgfbf8AH20DulD2BeS+TmEsXAQK7TskFNKMNyce8i/P0VwdAg7EPrJ0yGMj6DfPB3C+Cg6yNMhjA+jA1wdwj9QJLdVJmIcWs7VIfwCdXN1CPdCiTxrCOMONJurQ3gkUh+ubOPWGK1st2+f8YmRinej7eT7mST/5cSe6tD79XmvleMAWENA/HZEbPiJ8YHgeDCD7zJ5VIMcguIzUU5wkCRHDwWiH89XnWVfARQo3obuBTufcnmz0/MwIUt9WnbSjwDGV6PgW5vIYPgAbfDUDdrzfhX1OgaM74v2MiVYQl+JSpkS1BPlYLPrWqQsLwIztPcTML4IfQ4yUpX7pu4NlkStlgXF70NfiFmSymySTe9+qUIxNctA/HTkIY7ujqH+p7Tvol933YUSbKPriIqf5n8b8JWkRERG5gAVXg4lOA8QtfEFjBvHLChBF0OgcN07ysBbCyVYwBC4X1j1n8x5DwD9OJ953GThhQzKZ/glF+gENQxh9SIzlYE5IorCB4awj4HOaPnwyIocD/kCiwPkDlbm4qPvUlrwGz4JKpSCuMEnQVuw5Kbw0O9Dbk8cZiH6DLlq0kKf4JhqWV73hVq/4b564ccWZoRSPiMtluot62ZGhQwzneF/ZsMI45/CH5SATvw7+r+2AAAAAElFTkSuQmCC",
    "FaTimes": "iVBORw0KGgoAAAANSUhEUgAAAGAAAABgCAMAAADVRocKAAAAk1BMVEVMaXH////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////bzZmPAAAAMHRSTlMAHhIg4SEfI97iE8rPEM3zBAkDwNNg8PX0v8m+0nC9wswIdHNj1Apxw8jxZMvuYmG/HkQ0AAAACXBIWXMAABcSAAAXEgFnn9JSAAAB0ElEQVRo3u2Y6XLCMAyESQgqvUgI930fvfX+T9dOCwO0sSVH1nSY8f6W90tix5a3UgkKCgoKulpl633am6+avOrtstFLZ+uM7x8P8FvtKqe6uvmpHsRc/8d7POj2hq6G+rH64c7Vn0M4+XMJ5/404dyfR4gbeKGkZquuJZfVDXIeshYin/DbH3GUE4AXRD7hrz9ihwDskU8o8scWAUgLxhhm+nJ+j+oSgH7RoMJ3KHx+xJQAvCKTYPDHIQFYIY9g8scJAWi2WQSj/3hBbl0Jg2D0TxjbY1Q3DD6tJaBLRAShP0kQ+xMED/5Wghd/2zJJENk/e6l38PL87gR3fzdCGX8XQjl/PqGsP5dQ3p9HkPhzCDJ/miD1pwhyfzvBh7+N4MffvD97AoDyJwLlSQblZQrKPxoobxWgvNmB8nYNygcOKB+ZoHzog3Lbot14WfrDyEfraO0/PRCI/lZMIPtnIYHRn4sI2pfA7UZ2jW1TOd9SehHfEYC5dpTQ0w5DnvmLo3gtvRGAmXYg9a4dqWUjWSg4zf871nQNZiPnYNY1Wo6co2XXcDxyDse/5uEw02NevH9I4aYx/0jLP1rd/tNkwatu7ob97qCTV4KCgoKCrlafyu7GQ7M9mRoAAAAASUVORK5CYII=",
    "FaTimesCircle": "iVBORw0KGgoAAAANSUhEUgAAAGAAAABgCAMAAADVRocKAAAB6VBMVEVMaXH///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////9MyKhJAAAAonRSTlMA2BjaFtX++hcd+Br819zbG/lU+9TdVXcc2dNSsNauBUI1IAEesxIhKxVYGbEjCS3SMuGEb8Gn6ll9vx8KLhSCRUZDJAREYWtmE3/14MsH0BHp8xDomoAmJ215583xcDciqOvMrdGyNlpA7270tt+mBipT/ZvexKopr/K0OCV6bFFeDvaZW+bwLDF+ouUCMLvKq88vcrWWNDlxD/eVM3Z7X2Bhk2lZAAAACXBIWXMAABcSAAAXEgFnn9JSAAAEaElEQVRo3s1aZ0MTQRDdJCRHQAgpJAghdBSRJh2U3hSR3lREFLAi9q7Ye++97C/1klBSZnb3SoT5yOy9CXs7b97MHiEbxrLGW9xdSx7PUpe7czxLX+ydLZVeg0RDTDJ4K1t26oO+dU+1nYKWUP1sq1b0wp7LlGlmT7wG+F0zTso1qWaXSvinVyxUzF7fVAGfu50qMG+j0jNZNUQV2f2xfUrwjyVRxZaUIwyfV2WnKsw59lIM/9tbqtLmC0Xwb12iqi21nY//c5BqsJEKHv6JKarJhlrZ+NkJVKMlZDO54STVbNcZu1QyQnWw2gmUOi9RXSx5AMmva1QnK8gDA3ygulkVhJ9j1y+A9CUa/2wm1dE2R3Prb6qreaLqSy+y0uL98Qg9kA9uYEXvdH5EAKx+SW5Cdh+EfelyvXdLyIO+cPxpZNmmbX7vcB9YYYx+XwYWIbz+FLDwCRk4Gu0zFAd9WARvKP57CxOfkKZyDB+NYCkLCTAD73/G2ooiF7Q/zAg1ayviJPbvD6wJi2BKDPXBEZxrmu8QHz88gikl3AdHOLXqh46hxR2ZKsUGDJ8QB/QS+1b1M8iI0WySaAq6bGnRPjCPzi0790DOUYAPU2wBurcCrk8QRv2ysxpyHjECMNZUWZrkA47dByCMhuX+BS70ScVghB1bgD8Pv4EFQGnA24LkugtqLRpzod//DoFoC7gr0e5FsHlp6scQmplEKtNBogh+UTkKEKRUA6MwCUSIK2d0cIFeQ2KVvhQufjrjcXuHvGKCWfpMnAjxLubjfrXdyi6upjQWfqKB/XSnn0Y45ftMLo5v5EmRO/KiLp5A2LsFw7faeM/elVdNUrURrMncR7vlZbf5ImfHYQg/n49P9/+PADHfIu5Lxo+R2EvmHVNWIqTZBI4pJ9HYZJFm4ifahAZ8OdXMXKpgkh2fsOPNPLJj0bVIQWDRnZlTcMCyrISwfeySmR4nVjLxXQqWzE7EO1cETdoalexSNku29DeJyxakLC/LFlh4DUL7k+8XXpCwM9YyhJe4dAySTzKU2hdZ0hEUvws4MdiA5PsIYVxd8R6FvHVR0teEqxk3U74jDcjDiKOSiTU4hJwH2WCS10Idxw97+P9wntdCkRp4EBrS5BiTcN3qgNlsMfSO4AJl/w9xLrgNZ+CHtbFkHmnE63C6WYmA4EfMEo5RVoQikM7mAqnucAqNEoiXMQwxInTZLyd7HYY/G9m5oOOcGxePYGQ5OLqAuU5HEcp+fQdSv2I8UjNnxXgoCF7qfNUvwKn1GcySwlR98G3D6HD8gC7D8fHYjvenKtbzgkK+Ynkc2ysWQspexPaSSH7TyRrOT4mIErx3RS2+t1BMa5IedVeNnjzh28wcFcxn/qzsuve6MvheZde9fhnqU4I/a1VxJ/7q+QUxdEvBtMpb/bIaSYD7F59o+C5hILafPQTsXH0DQlDfG+qv6vNxSGlbsy8zLDfsmb7m7FJ9P3DpuNXm+Nvt8XT/cbS1d2ycj3r+ARanBS8ydAnSAAAAAElFTkSuQmCC",
    "FaTools": "iVBORw0KGgoAAAANSUhEUgAAAGAAAABgCAMAAADVRocKAAACAVBMVEVMaXH///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////9f68wEAAAAqnRSTlMABAGjAxB4SvjK+U7IC4ACYJ6n9NWqdYefEnv9RyvOD/fvfNcfHvV6mPHwnE2dLeaojnSixo3+SEvyYZD7Zo/2x3kgSSpyyR24Xd3AboprVpm9VTshDOTu1LV3GsTohosKdrm3HAVsmk+Sk18524UjrkyzUTJBImVGzCxvstblZOrLsXNUuzrffRuhROmMz7DzkZY0FbpYm96BZ1KVM9g4GVzDcNy2lD2/0xKLXtQAAAAJcEhZcwAAFxIAABcSAWef0lIAAARrSURBVGjetVpnQxNBEF0CEkCUnoixEAlBIgFRUIoiRToIKEWaYAEV7L333nvv3fmVXu6ScHfZMnuH85GZeW9nZndmbwkhs7Jw0QIiJ2kl6b86tpemFA3k99ZWLxVYLwTI3iyD/nt/KRikRoQPkJWOhZ+sPghm8fSJ8BXBpckx8wYocliMj0tTQSsNHjIw+Jg0fS2i4mcmoPAVWctP0/A6Kn5OHBafn6a4INhcf0g2bWSWNwNsr5+bpq45wgfIL6FZb50zfHqa+gKy+GQRsCX20HXI1peQBdkcBnOauqXXr8jSLA6D8dBN3pZef0jSgSf63dRmYf2iMhjSlGwNn18G3W46YRFfUAZFwmY7YjXl3PyP3MWVIUKQH6tJ5U28ZwB7cWXQbF5QmmiRg42vNq1zqDJoNvNpqhZmU1yp6qtRZdBMVtBU67nrB6hFnQbNYhVNdZ2PD72o06AZ3KSpgnx8uIg6DZpBKk1Vx8l/SAZQp0HTP6KpaNtIhw9VqKbEqQFlG6Xph2oeqilp2oR4mm4Jb/0Al1BNaV6YIZOia+KtH+AvqiklcxhWcfHhD242cBiu9Jy/P8LEh2nkbFiWy6sDBM6uHS4+5TDlPyQz2NnAy1JEBkdj7wQl6NkgiIEhpQn4EW2JoUYw6pan2mQ4IBqm85Il62C6MDiE49peDG2IC4EdhnbUlcNGljAB2IlhCPtJTI+hrOZy462PJ7cW5jyg43v3EFsMugZ94ENVLP66VxKvBgIGQnLf+s0Ei6VePYQMpGEXt41zHg2wDI5Ej87kXhwS/058ApaBdCdFDdwOJL5ym2MxTFQeejyavPt1dUPU2hlmGHxP8Pgshv3e6HZ52WhkOOaTwmcwGORCwSzD9mIiiY9hGPsc9llyI00eH8OQ5ySyslPvvzqO0pcMkjLfDv5/YHCZ/eeYwRXrH2GYhjlgWEzz71JVE172AEhyWl+/Ik9VXSVvxCBjoOPDFtbXuCxDIcNZi/8If0wisuRi+Z5R1d8Eg1gYQyHT1cn+FJSIgY0Px1WDIbDFwMEHl2rRA3YYXDy/IOZhhl+Heq7bVdXG5wHLMazhe3m+qFa7wGoM9SKv06rZetS9lxKDEB/eaYbXrDEkCl0SI//+GLPCsEboMHsVfO4H6TrI4CuTPU+WQQ5fngGff/MdTlSHbap1kzQ+mqG0X7Et9sjjoxmUBtZSZAUfy1DOeilDfKqgKu0lvg1RsoqyOvT6sTFkkc4IlfqG2+jF46MYskmr8Y3YKYGPYagg7dFi0J5fEwUTXMSQ5CPhvvIw4lIhgy9kUO5S/kgoYSmTwhcwuBWDdtO/Mprl8LkMIfxokSs186Oy+BwGFT+6TQPqt1txQBqfyaDh6w5ac1ltM1jAZzC4I9qgVH9Ad43x2V8YVEn1B2QMbp1yn8fm+ikxjBuU300MniYiL4YY3CblPkOWApWE2GNwx/7SIxjdSxv27iHWZFv4pzQpUzStr3P3D7//4M+pAmJZ+l11T5I+1TQY/vgPU6XtzIAFW7EAAAAASUVORK5CYII=",
    "FaTrophy": "iVBORw0KGgoAAAANSUhEUgAAAGAAAABgCAMAAADVRocKAAABwlBMVEVMaXH///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////+V6nZNAAAAlXRSTlMAd/z68v6wqv0Dp6tmKI8iAgELNN3c8a2NDERtqSDmWkDlPuikQgRHs/YdYIknjpM2O+/02t+haBW/4wVsr5B2gBP4Ec+9ZVUhZLqaDs2udZl6nsKoRuvMCLJQMXQSneCiReoHrEorbpIwc5d4OsDFpuklBkmMEBofhsmLcZUZHsP70TWcYedf9zPSHIO04opb4bzQWEDkncMAAAAJcEhZcwAAFxIAABcSAWef0lIAAALRSURBVGje7Zllc9tAEIbPEFlmiJ0YYg62gQYbZmxSZmZmZoaUWf+38tjSSYpsa0+nmc5Y+8ne2X0f+4Nu31shZEbNxpZuByii9TD9nXUcMOqugH4/WJ8nQP6DjyOIKADQRAJoAgAcJACHCahdgJexcwaFfaKL1+/jDIy+LsRwhkYjshsLaECcwWECagswikf4IIlUGPePqgIkZ8lzEoBP9SxTBzwmAUyrA16LHxlcYCEBPMP9h8WkDeXFzxFc4CEBbMb9YTH5Ab3BFW1iwXkSQFxs78XJefQOfzkYFCoukADuCN2BvTi5it5LSnYLJTECffas0H1Mkl1HM7OSmiNCTR4O+C70LrE4+XFOPn/775WKtsEB30qtbZJfzLn4xFup0bVailU/4ICvxc5Fm9QQfyqk3LK61bVC7iccsFToW9ulYlf98qmZ/cvnluGAL3xbOiWfl8niv1pRlP7OIfQLqs8/ppmFco/eLWX1nxU3FOBa/KxMnRKfvEOvjJgEN+7is2Nmnr7+9jmpf3xwk7b+k+NyhxqboOzoQkoPHLRY6clbpwMqNjsXpqUfyakb+VC3jYa8LRoqe/HwtuvXb/dWvNvUM6wedZapft9f7kyRyqc6/ZpuaIFMzwBcfaAnEwBcA5PpyZGsVu3syGQ6SbI08mu6/jT4yddSW7UAPDr2XkENs3khqGez9vJANf0zm/Tt7lqqnFDsfr3bwWuVAVP694+XKulfpbDgdDaW13c7aaxQx8oaAPcYnSVt6x51/ROt1PbAHpVBIbhMOnH6tlI/H6e7yx5WAoZpb8s3XLBNgAkwASagOGqOXvQNCVuyDbv2Ugz5EpfJxk7MclKrsxucCsH1e0HXkch1qH68A2Z9O1pg+o/6wS8jYAZsB9y+n4PovyC45rAPabtqPS77PgngKQDQTAJo/p9e1O0jAbgAgAQJYBxyzCVcDmC4xp3IjBqNf0VBdwuDaRgIAAAAAElFTkSuQmCC",
    "FaTruck": "iVBORw0KGgoAAAANSUhEUgAAAGAAAABgCAMAAADVRocKAAABO1BMVEVMaXH///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////9uRaD0AAAAaHRSTlMABQJGfqUEZjMDBvspR/b1X/RI2GnZqo3na9e9b7xFwAFOv+hsmpnEkD33yAkOt8HlxibaWZw6PwuRtbp8Pk0Pufw4Ldxe7sKkQuzxMjxBhCeOEvoh0HMRVMVy+F2G3Rlhnq1aJf2LlXX8UmQAAAAJcEhZcwAAFxIAABcSAWef0lIAAAIvSURBVGje7ZhnW8IwEMcPUBALCuLeE/fee4t7773N9/8EFgkdtk1De/HxRf7vuPx7vydNLxcCICUl9Udav44UqM2dAf70b1XEg5I9p3zpoy/Eox6PuQCfxLM2rnhe/4d3ADnhAKR95CdJjpXe8AMgh+6APl+Ae3eAr/wkIgES8P8B+Yq+eHgSCsgyngUDCFkQDSBFogH2BCKaQEQT/CU8KvlRNWOlUQotWuw8B5xKZhCQtgpnAtZe5EhA2+ycCHi7qQMBcbu2J2D2A1sCasOxI+B2tLBe05ViWqZOOIuJ6ck6YVtQ09cIewFBpwptpc9FHVvCdHBf2LmIDpYI6GgmSYAESIAESACGXtNfm0GRgKxqHADl9fVlheRx9G/Blk20qVnt14FQA296lh9uLCElf2SCOoUnPdsPl5ZQpd5cO3kAbD/MjPyer7F9t3O8H7YfoPfOHGk2PhB3B7D9iho8MM0hGTA+UFruegnC9mey0Znb7l0tkjIfcVzvzZl+JfOuD03kYlXmB8Zz0YT1dFWoH6ZpyZinTMtnyb8fxui0QsZgBQ3G/PshWJsbazEGaXF2BP37ASYpvU4PNdJQGsOvTU/p1Pyrdq/Bqx9aR/NfV3u8VF2vipb8724Uv1rT+gdclkoZ7v2ncPwAPU5NA8kP0Tk7/3wYyw+wvGL1Dw3i+dX/b7OW+UYx/arWho32/kVA9qtfX7yrI+eu7QoBvv9nF4i1JRJtsSAI8ktJSSHrG1KzFEYLpFEnAAAAAElFTkSuQmCC",
    "FaUniversity": "iVBORw0KGgoAAAANSUhEUgAAAGAAAABgCAMAAADVRocKAAAA8FBMVEVMaXH///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////+f7GOoAAAAT3RSTlMA+SNm/vzciND9CfGzMpk8zG9VArtedx+qTYAoDhMDiQjR2xyiDPTVgkSR48QK95rihfbmhMcnMdb64MbvK9oWfQHpBhD4fIEkBe0uM9gZeEHrYQAAAAlwSFlzAAAXEgAAFxIBZ5/SUgAAAWRJREFUaN7tmWdTwlAQRYmkGkKTGlQEUYoi2Hvvdf//v1FmAmM0wZfHro7Mno/5cM9M5u2dvE0sxjC/S6Xa71crVOnz+ZwKA9KZBfz0uVRyFUZslMqLqPHLBR2+sGbMYKUvZRMaBGCtx1cQ4l86TxCK3WhfTpT+lqnBD7y6Rdn063LpEASwcvlZifhT4wqEOUmmoqXfxNMQkWbhXPjItxtbIIGSyD4KxO+5RyDNdudZsA3kGdMj/jaQJ6RHAtpAnm89EtYG8vh6pNizgQC7N5zxAyCi6QmAjKkR1KnyzzzBw4XxAW72IPHOPwwkL+czFuqcBQhMTIEZIHBMFSteNff5K3giwk5b1OcsYAELWMACFrCABSxgAQtYwAJpgVNDvCc7f3HTJ99VkG9b/r+AGUe3pSOvrgE0/X70Z6e7S7PVPB4aWlR7U9cT6FSCuifQqATa1CzHFar8W0+wSWRQdrhiGBTeAdMo/LVf0hY3AAAAAElFTkSuQmCC",
    "FaUser": "iVBORw0KGgoAAAANSUhEUgAAAGAAAABgCAMAAADVRocKAAABpFBMVEVMaXH///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////9bMWvtAAAAi3RSTlMAEwWP9BXt/vzzjPgk8fvrb/r5JfL128tZbRiRuUIEa/bXmceJ1hcHFhsBU87YVxnKhSPsza5RdVY3f2AiJ1C3mOWN0LG2wD/pBu4LTv1840NnLrT3FD2AWm6QpMhmiiiz3L3ho2qODbJ0ElWdwSFkAogmDHNUeCBEaM+W2Xebvx/oyYtsld2+n+IclCRQlQAAAAlwSFlzAAAXEgAAFxIBZ5/SUgAAAq1JREFUaN7tmGdPG0EQhs/9zsZgG1dscMHGNBd6SeiEFjqEAOm9915Jb/Ong5QvrH0+b5mViOTn653eR9q7md1ZRaGlaTcaCy55vUvBWDTcpCBj2gsBQTFiQoz/sx+AMqyWDqR4s0cDXQJnsxj5YzNQkbdPxfO7NDDAPyiaf6IeDFHfiOVvQFUiQuujVheo4/z5dW6gQLvEm784A1R0mzkFL4GSO3z5HVZagVbgEuwDNdNc/S1AL7DydL49YCDOIQixCL6w5xccLAJHG7NgF5jIMwuibAILsyDGJjjFLPjBJggyC+xsggZmgY9N4Dp+AulLdEX2R5b+m0ovtDCbYJhZ0Ca72SlFFkErx34QYRHc4NkyrZK3TMVCL3BKPrb4OSeqD7SC57yjTTddfr9Z8uF3jP94PUhzfF8WGRAuVK/hR5JHqA3RIW3cb5TvXlaEaTT4l/rrMAblFo+/0iC+iDTrF27qFLV1GvNGxPTtPLE/OFrjd7EvXNrzlomc3eez5yYs+XalRo1jgO2ecyvZlzj85z+nGys0kfTC4eNEX/Ly1xc2xnRP6GhdOeaGNktf2Rx6QLxSjNA7Tk66yjvD1EBPeKTXls3aekfCPQNTOiPCu1dU8c+c9cCJN326en7zGgjwe75KfOeoCkKoq52Grf8+CHPd4BBjvg0IzLZUXJ9fgMK5xxUEo4DEqn7+vIolWFnXy7+1Bmhc06sHJyDyRKc/1GMKvJ/KBClA5XvZmcGFK9g+UyLwADJxkUtMjsnZ5sAWrJBrNAzodMkrgn9cJARb+IJZQvAeX5AkBA34ggVCsI0vSBACkEBNUBP8bwINP99NCHL4ggwh+IgvmCQEr/EFzeSeeRU7f65k03+I3LDtZVdVO3bU/B2dqfcAL/+n/lXbeiqDUA9aJnX0+/4F9ltf91/nTXoAAAAASUVORK5CYII=",
    "FaUserFriends": "iVBORw0KGgoAAAANSUhEUgAAAGAAAABgCAMAAADVRocKAAACZFBMVEVMaXH////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////DzP8zAAAAy3RSTlMACAL+ATMDmcz6+fdf5w0m+DIt+9ESwfTu/cnYuRaF7GsJ8VtBIusx741uT17kBgvzkS9TfONipUMKkJUU/JrDYeoHaEKKrg721xNWwseoJ/D14EDlgyGniMsMMHOXNRuWNLqgH8ToJO2Lr9MPn+KAI8q1Odbf2bGkyGYEqUfSUZS4N9wYvTyjhGVGiWosk3S3VZg223m8OxzGZEVKktURVLtZoT+CIMVjpkTphyWrSYxtsBAVWN4aXT6BhkgpzmxNLnF2V1w9Hp54Tk3u7Z8AAAAJcEhZcwAAFxIAABcSAWef0lIAAAO2SURBVGje7ZhVVxtRFEaTkEBjhRAIAYK7u0OxYsXdKVK0QIvX3d3d3d3d7f6pZhKgk8mVyaykT/M9Zu05e/TccyMQ8OHDhw8fG6blTnjTkbjrkc5iW1TPU2xXg9nktkusXV6UXwdMstnZqvUbygEz6jf21qtfGwog0S20Vv09cgCNB9rQciksRaNJCUuIYVE/JBYgokPcJdWT+VOSF7aS6i88BpBZATtg2207OuO+0wsvSEfXB+p4cz6+hkklYV/q+1kYAThkxqe6mVNuizGCRQAbFQN31MKogX5k/ZhYvKDTFPfzhGOTQpQgAV8f3DXtS19RXBtKsIkgAPvptKsWhQ1cRAh6SYLvdDoczfkiBEqSwORTOIHmuhECOUmQToOHpWhOWgQXqEmCDBrcigMn4IJQkqCJBp/DgaVwwXKSgP7wFDgwGS5YRxLQu0AEDrwCFzwm1Jel0eAOHBkIF6y0wwt0dNgLA9uhmnY2XhBgAjejQQ9Ur9iIrT8abQInosnjKIHYAyfIMYVvyFCg0zJkw5ZgPs9sMdvlLwyz5BxE1vf5yWQzfeBk7hLcqnkAUb94hzn7A3qTpIHYZd8evmz6p8LgfMirarebMLmI2yDPoQQxnT4vZpLVpeThS8V8l6QFyJmti9G/XFRsxkdxRBnt2v07G3C3NGfqH/pCIWI7ogb5Vrhos6qVmrgIP9LpdHhn98lkyrJdXTbZrPCxdpa+l1RCfhZKJEGwHYjwlLOjybQaEiTEbD68bxqm5TEFo1bloKHrRD1Y40p/QVPbXahpPGrklSQzOvqZJLInitoibO05DJOIvn2g9Vvdl0BjLVF/wiPajOs0MnfCeR+T0I1XW+9qtpNYbUZV1biUKM1mMele43dxUoZdnIIZi8c+J8A6G4wXkUzAGmkTgngQWJLgLYajUghY7/C84C2wLKGfqaOe9hEwzdyzjgSWZtSwZt0iYQ6z05CbxQIwZHjDpknYH4NgHHDIJ+rIMyRqKk9PTXCpDy5QjT8tloTl66kZTgLwkLqEVSRqUj9iSrkJKihBABFT4cd8TPzXUq2RiP0W/OIoAEepSwgmUTOCMa4CwybmNImaFuRyFdRTgkISVSVw4io4TwmukiipgGt9Yx9wIFFKWwvGbS0IsLGgWSSQchVksBDU6fdSSVwFieT/4C6/1CONXAXXKIH3AmS6414bVoOzHOt7sv43e4hTfff1rIfFonIu9cMtGEeFBXKL7887yybeTMU9B/bJSKy153cJfPjw4fN/8xc/fknGoMq9KgAAAABJRU5ErkJggg==",
    "FaUserGraduate": "iVBORw0KGgoAAAANSUhEUgAAAGAAAABgCAMAAADVRocKAAACH1BMVEVMaXH///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////9T2eBEAAAAtHRSTlMA9Pnt/gJmgECImfx36xMI3x9VBVaiJPGPnqPWX7Hp8/jeTb7VRPag8qrn++YDb/UHK24wAQZtkdT32BTdGcyyFcQl6DMipELI7M0JGMdlahK3+jYX5SGNlxZ9XqFOcr+GzgpXnz0eBAy7C7ANs8qbmoVrUde43Hq9Wzx/HUGEJ/COUJN0mNt5foPGJu+Sc7ZUNVk6G8CCY4eMLxBT2XxdQ4tskC7TtFKVdntc4g7h48liYODXCRRsAAAACXBIWXMAABcSAAAXEgFnn9JSAAAD40lEQVRo3u2ZV1saQRSGVxDWRRCiIood7L3EHrsxajSJib1HY40tvffee++99/mBCcICg8vOGZjkyYXf3Q6H731mp+ycMxwHVsTPxe86hUI3v/CjjmOttWrjL+Sm+jltg4qRd3a+Ol2HJJQ2rT1Y5Kd5WIN27jKSUUCtkLzeR/OpZH66DAGUWSVYHlGa11iE2s2IQsFJmtAKoHllaElVIPJFazT9q0iTpV+ThPxSVro6f7WktypfbfyMmCgjPWfZHFadj0NM1fkB78cIYq5DGCCOPaATA9SzB2RhgKPsATw+yE8Zv6T6c56TVfWMmXlGevFDia02yPbbK6PBP3PDwL1LXlbxEiDIr63CqO6S2SZEgE1fkoVaBe1mV0PYh9wBNs2mwv1TZwG7qCdg8isc8C3MB0A4zSsK9wFQTQPYTg+YSqEBpMxQAyx0U9RCDbjrz84jJb3ZFmfW+zQECMUSAY32QLP4HE0HSCL5F4qRhY6GDH92fwkViJHFjoYEOkACCZAjRhY4GhSMAc5X9NrRoKQD6IiDbLQHDojPe+kAZ4mAvNilJZ8nPu+hA1wjL7QoW1yU87GEDhBJBiQufZScjzfpAOTt9LA90Hk87qD6bAZ2EAEx9sgYZ8M8DaCN/Ia22SNdg1BMA9gIHAKEEp0NIQFw/4AQov+EGOvKUSLhgAccdAjcB+E2uAtpreAhcB8ETgsFXAV8kZ0nRoNblrwV5t8COLRMuMLdEsUIE8Q//higA+HSa3JXMODUOAhJjq+7/vDYvV1NXsMbQNn36LD4ByteelAT+hAM85fRyTQ5f9Mg57fqZOZSSwSLgtFkjpdO5PJFHBtVnpBY1AHdrRw7hbxswr4PgW07N7Eu2/WORbaXKxUKZXl75Fgvt6IV/QfSHxlKtZb2kQP7Sq2pQ0F6SvdP1bnA3It3rOvYA3nwZTvSDM7ueFdk8w7Y4i76aILnjzwWenwf4Ou8pZQiQ+U9Y9/cIPmrJVInAez/p1KulbVX9dBk2bxk8HuZi4vsRpo8nvcSbM72Cjjl9dsowP0RWvTmr6WpRfAywae9zB/ZKwkB7o8yb0n5zxgo6im8fGziqATgPukMJ4D9EbogcfmRAK8JEf1R2fK7lh7AOVqA+iN0xdO/IgVc14L4o9wQb+UcYh945FPe2QStzQH90UWPmznfyuFyeQN+bxfK/gbkHQboZg94gQHesgfghc4n7AG7MYCOPSAOA5SxB6RgAPQX9G8B8ez9TRignD0gGgMssAecwQDJ7AHr8N1uP2v/Ox7bdRfjpaZcVmQYVzL1H5coegyz87dKF0mea6IZrIf4aI37+P4GzjnN2oOAN6AAAAAASUVORK5CYII=",
    "FaUserTie": "iVBORw0KGgoAAAANSUhEUgAAAGAAAABgCAMAAADVRocKAAACB1BMVEVMaXH///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////+QMwV4AAAArHRSTlMA8xMh/vQi4e2gYfu2YgX8j/UCYPHrEvqS+SRtpZnWoqHy21oYkbnu+LTgBGvSUW8XjRYbBnfO2FcZvuL3hy08agh5B3jfPSBCZijN9scnUHS30NUlSU5T/RrJZymVFIAey25ZypBpn4UjyOwJTJTXVty9W6MD8I6yVTY7QIPGZCYrDLFzEZfAP2NE6WiMz7DUcpY02RWb3r8fQ4aLbArTcZo43XtcSxzCw7MueENYIAAAAAlwSFlzAAAXEgAAFxIBZ5/SUgAABAJJREFUaN7tWWVX40AUTVtqlELRosV1cXeXXXSBhcV2F2fd3d3d3X3nRy4lTSYTmcmkyTl84H7j5s27PbmZee8NDKMWWZnZORm5Xm9uRk72QhajM8zzcwDBYpJZx/S/liKBBFZPhU7pw0tMQBaRuzr0yD84AhTx8HLw+bebAAbpxcHmj3YBLCLuB5d/GhCRFNT7iSALRAxpz58YA1TA1Ks1f+UIUIXqcI0CnUAltmrLX2FVK2Cq0ySwBFSjTNP5FqlewKrl5JsHFCjUIDBHIzBFn78uhEYgpJRaIBNQoZZaIJtOwEMtkEMnEEYt8I1OIINawEEnEEstkEAnkLf+BAx/RUeMNtnwz9TwjbZAJ3CYWqDU6MOOWaQRGNBQD5JoBFq1lEyrwSWT8agXeGBw25KucaKqUitwXetoU60uf1+4wc3voPb2ulhN+z4ezICwhbyHLxk8Qk0HO6QNpePyx4wzQWMb5lvqS9RjUN5Ukq40iFfqNOvXnZLZ1NYyPW9EzB/2IvUhZKAwTu8Ll/Jaz3CNIyHBUTPsqS1nNrCBdYD9558460dt3J9hir2ibbTe+fl2G132th+n2THczTEzXjS/d4Z74mb39aNj6i/wzO/2cXlSeTIeFYjnH6TyHfxudZu78rugUDbwdBwyMDhgrgbInmxXUZ3PPkZ+qU2+04P3aDYkvOgZsUaKRic3PLU/QdYHf+kdNN5VhU0flSL+WKAJzBhkxxipBRx2RGFKS7Lk1LcLHofJjTN2yZLjJxQFDsqULWgC0xv4VF1XILdHZskXqlbRLQiIF3+i3C5Qd4d3TrZDSRNETLLUpIBKk1vj+iqX/2qXbGUXmtDCUi1YC/x49VFGoEmhNbHAkHaW6WQUdgHEBWl+W55CbCiMyWeZfMiEKixqfCMRuKnUXDnhLgmcFrlReAvEH8IaXvQrhUITmjmqmWCBvynbKRIoUe4PeRNaOWaZZIHMReeUcihvAr/Pk0kWrOKiqKvCzPVOyfuwEy1Y7f3Qd3QG06JzNWEWUrMBqgCzDP3vThluCLBIblIzWcaCW/UeEcjHhQZMeCo4kjEHEYdriMBzXGjAhCJB4WIZJ27VIUQgFhfKenpjAjITcUQLwD1EoBE76fX4Qw4ImZ9EC0A/IoAfJaMlh2HT2gSKX0UhsOIP+SNkfhMtoBLwm9CBvMW8NpIFVAJ+E7pRpoVkAZ1ANF9sOHRiDyJqgRXJVvxLtIBKwM4XGw6rRadARwHQ80/MNJMsQAVMJBNaxcwyyYIYRKCGEO2UNJXJJAt8iMBLQnSXpPjauwhLUhCBu0B3bEZr5lG9898SFX1LrL75HZKrqm6Hrvm7Za7OXuuXv17+qu1tis8UfHKTL0Xo739dOXjxXrtwPwAAAABJRU5ErkJggg==",
    "FaUsers": "iVBORw0KGgoAAAANSUhEUgAAAGAAAABgCAMAAADVRocKAAACTFBMVEVMaXH////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////ocqefAAAAw3RSTlMA/PTk+gTMmTMByYH3FBgXGyX9Bb1yvAnq+bDe8xCUdZ7mHOUCJivaOiBE2ejtKQrdfb7nyGbxLbP2E38I0bL1ErdVmJ1+y6wRXkn4PiS0u+NChakoa64O+zfcoyLH644sUF9FShb+ymjuL9R3fKDEYpXYuXsZPWFw4iNH7IqP0lHXuFZbwmBBZagnaksNMTZaO+CiQIMhxqdpB++NDNVztpc1eN8/gqbpBquMC5G1NFg5XR+lZ6pIzmyvTfKQ03ZcgB6jNS8xAAAACXBIWXMAABcSAAAXEgFnn9JSAAAEMklEQVRo3u1ZZUMbQRBNSI6SBJpAAxGcosWLVoECpUDd3d3d3d3d3d3d9o+1e9mTlctJj3669419M+/dsbszw2GzWbBgwYKF/4tHhTWjMzJG17zq0RXqrrmPYwBCzJgOl9n6c3cADAO2mCo/4SKg8HaCefpTCwADq8+bpT8sGzDRc7M5+gmpQAGTG005Ps+BIgaaYdAOouDkv+v3z45mMPKyDqmy++6qkMOf7MVW34GoOIgFd0/u7QhVuT1jGfKNWXbhnvbjZOuDoxu4ZaE9PMJdt7cMIfXz5Ec9JV1cLwIq+C2Gpvtly8WJuP68u1jWCPEdNqgZzBaf34+tpyZgBsOJtEqBeKZmsF2I9BBEWK5/lkxL8yo4U6gX9jeGZIbKDB5QeRsRs0LNYB8KvEExUyT93CsUm4KoajWDahTYm2IWST3DS+dtQlSKmsEIFDiSpuaJBjNpshNRy9UMslBgJ03NFA1yaLIKUUvVDDJRYAlN+aQrGE+R0xBVqGYwCQW6KaYPp3wNAOiHGC4YXb+EU7gGeC1/SZIT8wTqW3SDGULcWDtJDZIZ5H4kyG5SP54YTd9ZJAa2ENRgbLLJ/4r3c6na2ZqjGWRIcUOKMaa2DK92c2plZIW8FDaWKOtny6tynrx31/Yi63V+rEiO74sxvexK+vYD+HQQln4/ZYzmvqcentbSM+tJZqGSQQ0ZOXTKIng+Bw5SmCw5X52Xtf4piSWfNI4V663zcfo7/5dSWn/rUTOH05w2Un+Mz+TxetJ4p6QearvdBX8hJLRnLQiGQkF3y+u+enPhPCM9YHxqeEO5XomnK8OpUs10YvOVNM/IJud2ffr76QFQnK/S/cxD/kSP/lWmRGS+4thdsb5Ij4Gvjd1POWYhh9swW7a5iXGLc2nN3MVxibLhamU8S6eSNc/8heOFMDkuXcPfsYrM7sSxyKyA66W7f7xBKyd6MoT+zlfJLP1hkZz1TfOlwuBu7chBt66j1S0Vj6SmU6iashw2MuYZEIxU63Xv6QdyLFniSKOWD0Ve7/B0xj4z5hn7qEiVfgg0oySSMsfOmK+c1NpaPvinE+iAcy+fVEkTNmrpO/8lYpYd6EIS3+K5ARRBGcwvhJFxnUAn4o/zHUfdYA2MCxQA3SjuDzP9qgZ8sz4HDOAXzPxAGRB7OZm/uzFGDNJgEXedJjcZn2bANmjQCgxhGcydga/tIk8W3GJXhTGDAmhwjzzzgQXYG8GdqgMGAYehcux8uwM227FVDbIPWfAhxhk1uASzr0s/N2Rc4+9fID9OAF/lbhk1WAWzE0Wx/IBC92gyanBTY3u6YNRghUaDWKMGsZaBWQbVRg2maTQIGzX4rNEg06hBstYxrY8x/YZ1WgfBZcYM7mj/2LvciH6Wjv8puI5M1ysfXKhv3i/f6WnuphnNnlnlNgsWLFiw0HX4A1BZX6xH3uRzAAAAAElFTkSuQmCC",
    "FaWarehouse": "iVBORw0KGgoAAAANSUhEUgAAAGAAAABgCAMAAADVRocKAAABPlBMVEVMaXH////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////lgy2UAAAAaXRSTlMARP3dzKpVZjOZAgTLuyn8AyP7Iuvwb/X6iAf4v8QKL5S4ZbdaIU8ByhB3Yui0lRS5vjKO25/554oozXDcer3CiScssjYXO8GsjSv0czXaFpzA5OloSU7zkbUV3nzJSK9NcQ+AW6Kbi1zIR25nAAAACXBIWXMAABcSAAAXEgFnn9JSAAABtklEQVRo3u2Zd1PCMBxAA6WUUkSGA0VFEUEFHIC4xb333nvl+38BreedorTEXxKP3OX9Ry/JO+6lpS0ISSSSGiOwmEhsebktHx47wO9kZnt4rO51qPX4E2edO8B4+fsmBZcx1T/IcHmP3ox/EVWX2aze6m7HFri0DvqwLQ3YhsbTMKuwVlAE7/sZ1goFFLxiWCuiqoNZWBbBg/ZhLYPrflZhbYIbVcNeYCr2H85Zhf1zcEBYKyYqBAeGJQxOE5Yk+KuCuaA8fixvPGFujJiXqTzmyAtCOSdPgTOGnjFXdKTwFRRQka+giDBnpEAKpKCC4GTt2gPmKL9RTbBN+fQSLNgLIn7a2/H1kK1ApX+eWLUV3NELNm0FmjnkMALbMKFLc7ZWXZCFbkkfoaAbKnDViqBLeMGQ8IIbqGCFUODv1EEM5AgFdEgBC8E8NLKHULAA3aYlQsEM7xNtWHhBm/CCvZr9BqQ/+pPQ25ZpebETRmDEYI8fMYNQsAt+/0QouOJ9ovUKL/AJL3CJJciUfbw1h5xBBVlzdrLs0Cgaj2tfHKfMIUtpDUR6x5ydSn47FJ+T/zBLJP/AG/GOFRTcerCsAAAAAElFTkSuQmCC",
    "FaWrench": "iVBORw0KGgoAAAANSUhEUgAAAGAAAABgCAMAAADVRocKAAABp1BMVEVMaXH///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////+JEg7+AAAAjHRSTlMAqP6iAwSjdHt4x/Ad+0zMjqH0AkvGnHJ3yd1J8fMeY7YBkefNIvVVIUrvVKaMGiSvV1yAmSC1m9j8a9sc/WzX3OiVX9OP0rNR9pS4ep5/wg3W+ph5vJ1A5YOIyweNK9X5Ndq7Wd99XiXKzy/UU/jev6rtTZAudji+R09t2QVI8ppdYjA/lsMt5pPAOZuzv/UAAAAJcEhZcwAAFxIAABcSAWef0lIAAAJ5SURBVGje7ZdXd9pAEIUnApwYYXqxAYNxiWuKe4kd1/Tee++99+b0xJkfHWGwjgWL4jmayZPu20o6371nmJ1dAJzI+zX//ceR2aPRtquvW/MvgFexb3OzaNFeTvyHLzNYKW0LG74wfwMVYjPI7VPhcZSLP5BV8uu8TPxGTclf5+PB+5Iom38ORfPDcWF+vyZbn809svlhqzC/V5jvnZHlQ6cwH/zC/CuS/VnUmWp+wpZ/u4Nm0F5tELJriZOoHSNtMsUm3hirza8rfnCPYLBb1UNdNfme5fd5gkG9yiBoz8dWgsEBlUGjXX0MPSYYHFIZeOz5OEYwCKkMQnb1MXSaYHBXZaBqo1V8bCYY7FSO0g6b+hh66LRE+MyWj22OD5vKNvJGLK8XCAbPcQ1tZM2PuEQwSCkNnkZeBrsKtfj4m2BwCmvq8qUng50XY5X1MTRPMDi/B+0VPZeoevaTMk5PIFm/SOdRhG5AmXWwns7HP06n9T80XZDNj5+F+WmfaH0QB2Tz44gwXw/I1kfrl82P+4X5Sdn64JBPNL82KJtf75Wt/y7Z/nz0QHR/td/inD99iQbrHeXsK9b5Xw9wOH4zeW06nb7v7zt4fRh4z68Njv7wuXyX7/Jdvst3+S7///NTa7k/ONDHcdn8gaxsfhiVzQ858/bnD4f9/Plhe5nTkymuunVuPrSUQZnSMs7Nh23l+qysm5j50FwibVpZX2DmVxmEmfl2JWLhmz9yd2kZ5OabbarHl/k6Nx8C5kZrCu+4w57f0KLEfFityazAfLDozbhk/qLeTVnxUylgVsYy4j69BXZNDpm91OAZBgnlJlreR6NjIxMBZvBfAECEiimVZxgAAAAASUVORK5CYII=",
}


# ------------------------------------------------------------ components ----
_KC: dict = {}


def _kcomp(name):
    def deco(fn):
        _KC[name] = fn
        return fn
    return deco


def _P(text, size, **kw):
    """paragraph helper; returns None for empty text so callers can inline optional lines."""
    if text is None or _ks(text) == "":
        if not kw.get("runs"):
            return None
    d = {"text": text, "size": size}
    d.update(kw)
    return d


def _B(x, y, w, h):
    return {"x": x, "y": y, "w": w, "h": h}


def _kdk(ctx):
    return "FFFFFF" if ctx.dark else ctx.t["dark_bg"]


def _list(v):
    return v if isinstance(v, list) else []


@_kcomp("kpis")
def _kc_kpis(ctx, b, box):
    t, c, dark = ctx.t, ctx.c, ctx.dark
    items = _list(b.get("items"))
    n, gap = len(items), 0.2
    if not n:
        return
    tw = (box["w"] - gap * (n - 1)) / n
    compact = box["h"] < 1.05
    for i, it in enumerate(items):
        x = box["x"] + i * (tw + gap)
        hl = b.get("highlight") == i
        fill = t["primary"] if hl else (c["surface"] if dark else t["light_surface"])
        tc = t["on_primary"] if hl else c["text"]
        mc = _kmix(t["on_primary"], t["primary"], 0.25) if hl else c["muted"]
        _krr(ctx, _B(x, box["y"], tw, box["h"]), fill, name=f"KPI {i + 1}")
        pad = 0.15
        value = _ks(it.get("value", ""))
        vcol = t["on_primary"] if hl else ("FFFFFF" if dark else t["dark_bg"])
        if not compact:
            d = 0.38 if it.get("icon") else 0
            if it.get("icon"):
                _k_icon_dot(ctx, x + pad, box["y"] + pad, d, t["accent"] if hl else (t["secondary"] if i % 2 else t["primary"]), it["icon"])
            vx = x + pad + (d + 0.12 if d else 0)
            vw = x + tw - pad - vx
            vs = _k_single_line_size(value, vw, min(30, box["h"] * 19))
            row_h = max(d, 0.46)
            _ktext(ctx, [_P(value, vs, bold=True, color=vcol)], _B(vx, box["y"] + pad - 0.04, vw, row_h), valign="middle", quiet=True)
            trend = it.get("trend")
            if isinstance(trend, list) and len(trend) >= 2:
                val_w = (_k_str_em(value, True) * vs) / 72
                sx0, sx1 = vx + val_w + 0.15, x + tw - pad
                sh = min(0.34, row_h - 0.06)
                sy0 = box["y"] + pad - 0.04 + (row_h - sh) / 2
                if sx1 - sx0 >= 0.45:
                    vals = [_knum(v) for v in trend]
                    mn, mx = min(vals), max(vals)
                    rg = (mx - mn) or 1
                    P = [(sx0 + (k / (len(vals) - 1)) * (sx1 - sx0), sy0 + sh - ((v - mn) / rg) * sh) for k, v in enumerate(vals)]
                    sc = t["on_primary"] if hl else t["secondary"]
                    for k in range(1, len(P)):
                        _kline(ctx, P[k - 1][0], P[k - 1][1], P[k][0], P[k][1], sc, 1.75)
                    _kell(ctx, _B(P[-1][0] - 0.045, P[-1][1] - 0.045, 0.09, 0.09), t["accent"])
            ly = box["y"] + pad + row_h + 0.08
            _ktext(ctx, [_P(it.get("label"), 11, color=mc, spaceAfter=2),
                         _P(it.get("delta"), 10, bold=True, color=t["on_primary"] if hl else t["secondary"])],
                   _B(x + pad, ly, tw - 2 * pad, box["y"] + box["h"] - pad - ly), field="kpi label")
        else:
            vw = tw * 0.42
            _ktext(ctx, [_P(value, _k_single_line_size(value, vw - pad, min(26, box["h"] * 26)), bold=True, color=vcol)],
                   _B(x + pad, box["y"], vw - pad, box["h"]), valign="middle", quiet=True)
            _ktext(ctx, [_P(it.get("label"), 10.5, color=tc), _P(it.get("delta"), 9.5, bold=True, color=t["secondary"])],
                   _B(x + vw, box["y"] + 0.08, tw - vw - pad, box["h"] - 0.16), valign="middle", field="kpi label")


@_kcomp("timeline")
def _kc_timeline(ctx, b, box):
    t, c, dark = ctx.t, ctx.c, ctx.dark
    items = _list(b.get("items"))
    n, gap = len(items), 0.22
    if not n:
        return
    tw = (box["w"] - gap * (n - 1)) / n
    line_y, card_y = box["y"] + 0.5, box["y"] + 0.78
    card_h = box["y"] + box["h"] - card_y
    _kline(ctx, box["x"] + tw / 2, line_y, box["x"] + box["w"] - tw / 2, line_y, c["line"] if dark else t["light_line"], 2)
    for i, it in enumerate(items):
        x = box["x"] + i * (tw + gap)
        cx = x + tw / 2
        hl = bool(it.get("highlight"))
        dcol = (t["accent"] if dark else t["primary"]) if hl else ("FFFFFF" if dark else t["dark_bg"])
        _ktext(ctx, [_P(it.get("date"), 13, bold=True, color=dcol, align="center")], _B(x, box["y"], tw, 0.32), valign="middle", quiet=True)
        _kell(ctx, _B(cx - 0.12, line_y - 0.12, 0.24, 0.24), t["accent"] if hl else t["secondary"], line={"color": c["bg"], "width": 2})
        paras = [_P(it.get("title"), 11.5, bold=True, color="FFFFFF" if dark else t["dark_bg"], spaceAfter=2),
                 _P(it.get("text"), 10, color=c["muted"])]
        if card_h >= 0.45:
            _krr(ctx, _B(x, card_y, tw, card_h), c["surface"] if dark else t["light_surface"],
                 line={"color": t["accent"], "width": 1.25} if hl else None)
            _ktext(ctx, paras, _B(x + 0.12, card_y + 0.1, tw - 0.24, card_h - 0.18), field="timeline item")
        else:
            for p in paras:
                if p:
                    p["align"] = "center"
            _ktext(ctx, paras, _B(x, line_y + 0.2, tw, box["y"] + box["h"] - line_y - 0.2), field="timeline item")


@_kcomp("cards")
def _kc_cards(ctx, b, box):
    t, c, dark = ctx.t, ctx.c, ctx.dark
    items = _list(b.get("items"))
    n = len(items)
    if not n:
        return
    cols = int(b.get("columns") or (n if n <= 4 else ((5 if box["w"] > 9 else 3) if n == 5 else (3 if n <= 6 else 4))))
    cols = max(1, cols)
    rows = (n + cols - 1) // cols
    gap = 0.2
    cw = (box["w"] - gap * (cols - 1)) / cols
    ch = (box["h"] - gap * (rows - 1)) / rows
    hls = set(x for x in _list(b.get("highlight")) if isinstance(x, (int, float)))
    for i, it in enumerate(items):
        r, col = i // cols, i % cols
        x, y = box["x"] + col * (cw + gap), box["y"] + r * (ch + gap)
        style = "solid" if i in hls else (b.get("style") or "outline")
        lo, shadow = None, False
        if style == "solid":
            fill, mc, title_c, dot = t["primary"], _kmix(t["on_primary"], t["primary"], 0.15), t["on_primary"], t["accent"]
        elif style == "tinted":
            fill = c["surface"] if dark else t["primary_tint"]
            mc = c["muted"] if dark else t["text"]
            title_c = "FFFFFF" if dark else t["dark_bg"]
            dot = t["primary"]
        else:
            fill = c["surface"] if dark else "FFFFFF"
            if not dark:
                lo, shadow = {"color": t["light_line"], "width": 0.75}, True
            mc = c["muted"] if dark else t["text"]
            title_c = "FFFFFF" if dark else t["dark_bg"]
            dot = t["secondary"] if (i % 2 and not dark) else t["primary"]
        _krr(ctx, _B(x, y, cw, ch), fill, line=lo, shadow=shadow, name=f"Card {it.get('title', '')}")
        pad = 0.15
        tag_c = t["accent"] if style == "solid" else t["secondary"]
        stat_c = t["on_primary"] if style == "solid" else (t["accent"] if dark else t["primary"])

        def body_paras():
            out = []
            if it.get("tag"):
                out.append(_P(it["tag"], 9.5, bold=True, color=tag_c, spaceAfter=3))
            if it.get("text"):
                out.append(_P(it["text"], 10.5, color=mc, spaceAfter=2))
            for bl in _list(it.get("bullets")):
                out.append(_P(bl, 10.5, color=mc, bullet=True, spaceAfter=1))
            return out

        body0 = body_paras()
        if (it.get("icon") or it.get("stat")) and body0 and (ch < 1.05 or cw / max(ch, 0.01) > 3.2):
            d = min(0.42, ch - 2 * pad)
            tx2 = x + pad + d + 0.14
            if it.get("stat"):
                sw = min(cw * 0.3, 1.5)
                ss = _k_single_line_size(it["stat"], sw, 24)
                _ktext(ctx, [_P(it["stat"], ss, bold=True, color=stat_c)], _B(x + pad, y, sw, ch), valign="middle", quiet=True)
                tx2 = x + pad + sw + 0.12
            else:
                _k_icon_dot(ctx, x + pad, y + (ch - d) / 2, d, dot, it["icon"])
            _ktext(ctx, [_P(it.get("title"), 12, bold=True, color=title_c, spaceAfter=2)] + body0,
                   _B(tx2, y + 0.06, x + cw - pad - tx2, ch - 0.12), valign="middle", field="card body")
            continue
        head_h, tx = 0.0, x + pad
        if it.get("icon") or it.get("stat"):
            head_h = 0.4
            if it.get("stat"):
                sw = min(cw * 0.5, 1.6)
                ss = _k_single_line_size(it["stat"], sw, 24)
                _ktext(ctx, [_P(it["stat"], ss, bold=True, color=stat_c)], _B(x + pad, y + pad - 0.03, sw, head_h + 0.04), valign="middle", quiet=True)
                tx = x + pad + min(sw, (_k_str_em(_ks(it["stat"]), True) * ss) / 72) + 0.12
            else:
                _k_icon_dot(ctx, x + pad, y + pad, head_h, dot, it["icon"])
                tx = x + pad + head_h + 0.12
            _ktext(ctx, [_P(it.get("title"), 12, bold=True, color=title_c)], _B(tx, y + pad - 0.04, x + cw - pad - tx, head_h + 0.08),
                   valign="middle", field="card title")
        body = []
        if not head_h:
            body.append(_P(it.get("title"), 12.5, bold=True, color=title_c, spaceAfter=3))
        body += body_paras()
        by = y + pad + (head_h + 0.1 if head_h else 0)
        _ktext(ctx, body, _B(x + pad, by, cw - 2 * pad, y + ch - pad + 0.04 - by), field="card body",
               valign="middle" if (not head_h and ch < 0.9) else "top")


@_kcomp("process")
def _kc_process(ctx, b, box):
    t, c, dark = ctx.t, ctx.c, ctx.dark
    st = _list(b.get("steps"))
    n, gap = len(st), 0.18
    if not n:
        return
    sw = (box["w"] - gap * (n - 1)) / n
    d = min(0.62, sw * 0.32, box["h"] * 0.3)
    _kline(ctx, box["x"] + sw / 2, box["y"] + d / 2, box["x"] + box["w"] - sw / 2, box["y"] + d / 2,
           c["line"] if dark else _kmix(t["primary"], "FFFFFF", 0.6), 2, "dash")
    card_y = box["y"] + d + 0.12
    card_h = box["y"] + box["h"] - card_y
    for i, s in enumerate(st):
        x = box["x"] + i * (sw + gap)
        cx = x + sw / 2
        hl = b.get("highlight") == i
        if card_h >= 0.5:
            fill = (c["surface2"] if dark else t["primary_tint"]) if hl else (c["surface"] if dark else t["light_surface"])
            lc = _kmix(t["primary"], "FFFFFF", 0.55) if hl else (c["surface"] if dark else t["light_line"])
            _krr(ctx, _B(x, card_y, sw, card_h), fill, line={"color": lc, "width": 0.75})
        _k_icon_dot(ctx, cx - d / 2, box["y"], d, t["secondary"] if hl else t["primary"], s.get("icon") or "check")
        bd = d * 0.4
        _kell(ctx, _B(cx + d * 0.22, box["y"] - bd * 0.15, bd, bd), t["accent"])
        _ktext(ctx, [_P(str(i + 1), max(8, bd * 36), bold=True, color=t["on_accent"], align="center")],
               _B(cx + d * 0.22, box["y"] - bd * 0.15, bd, bd), valign="middle", quiet=True)
        paras = [_P(s.get("title"), 12, bold=True, color="FFFFFF" if dark else t["dark_bg"], spaceAfter=2), _P(s.get("text"), 10, color=c["text"])]
        if card_h >= 0.5:
            _ktext(ctx, paras, _B(x + 0.1, card_y + 0.09, sw - 0.2, card_h - 0.16), field="process step")
        else:
            for p in paras:
                if p:
                    p["align"] = "center"
            _ktext(ctx, paras, _B(x, card_y, sw, card_h), field="process step")


@_kcomp("chevrons")
def _kc_chevrons(ctx, b, box):
    t = ctx.t
    st = _list(b.get("stages"))
    n = len(st)
    if not n:
        return
    h = min(box["h"], 1.0)
    y = box["y"] + (box["h"] - h) / 2
    aw = box["w"] / n
    notch = 0.5 * min(h, aw)
    for i, s in enumerate(st):
        x = box["x"] + i * aw
        fill = _kmix(t["primary"], "FFFFFF", 0.62 * (1 - i / (n - 1)) if n > 1 else 0)
        _kshp(ctx, "homePlate" if i == 0 else "chevron", _B(x, y, aw + (notch * 0.25 if i < n - 1 else 0), h), fill, name=f"Stage {s.get('title', '')}")
        tc = _kread(fill, t["dark_bg"])
        _ktext(ctx, [_P(s.get("title"), 11, bold=True, color=tc, align="center"), _P(s.get("text"), 9.5, color=tc, align="center")],
               _B(x + (notch if i else 0.12), y + 0.04, aw - notch - (notch * 0.5 if i else 0.12), h - 0.08), valign="middle", field="stage")


@_kcomp("layers")
def _kc_layers(ctx, b, box):
    t, c, dark = ctx.t, ctx.c, ctx.dark
    tiers = _list(b.get("tiers"))
    m = len(tiers)
    if not m:
        return
    side = b.get("side") if isinstance(b.get("side"), dict) else None
    side_w = min(3.0, box["w"] * 0.25) if side else 0
    side_gap = 0.2 if side else 0
    label_w, lg, tg = min(1.6, box["w"] * 0.14), 0.12, 0.12
    band_x = box["x"] + label_w + lg
    band_w = box["w"] - label_w - lg - side_w - side_gap
    defaults = ["dark", "secondary", "primary", "neutral", "info", "accent"]

    def item_paras(it, role_c, hl):
        tc = _kread(role_c) if hl else (t["dark_bg"] if role_c == t["accent"] else role_c)
        p = [_P(it.get("title"), 11, bold=True, color=tc if hl else ("FFFFFF" if dark else tc), spaceAfter=2)]
        if it.get("text"):
            p.append(_P(it["text"], 10, color=tc if hl else c["text"], spaceAfter=1))
        for x in _list(it.get("bullets")):
            p.append(_P(x, 10, color=tc if hl else c["text"], bullet=True, spaceAfter=1))
        return [q for q in p if q]

    needs = []
    for tr in tiers:
        its = _list(tr.get("items")) or [{"title": ""}]
        k = len(its)
        iw = (band_w - 0.2 - 0.1 * (k - 1)) / k
        meas = [_k_measure(item_paras(it, t["primary"], False) or [_P(" ", 10)], iw - 0.16) for it in its]
        need = max(meas) + 0.32 + (0.42 if tr.get("base") else 0)
        needs.append(max(1.0 if tr.get("base") else 0.62, need))
    sum_need = sum(needs)
    avail = box["h"] - tg * (m - 1)
    if sum_need > avail * 1.12:
        ctx.warn(f"{ctx.where}: layers content needs ~{sum_need:.1f}in but has {avail:.1f}in — trim item text or give the row more height")
    hs = [v * avail / sum_need for v in needs]
    y = box["y"]
    for i, tr in enumerate(tiers):
        h = hs[i]
        role = tr.get("color") or defaults[i % len(defaults)]
        role_c = _k_role(t, role, t["primary"])
        if dark and role == "dark":
            role_c = c["surface2"]
        _krr(ctx, _B(box["x"], y, label_w, h), role_c, name=f"Tier {tr.get('label', '')}")
        lc = _kread(role_c)
        _ktext(ctx, [_P(tr.get("label"), 11, bold=True, color=lc, spaceAfter=2), _P(tr.get("sublabel"), 9.5, color=_kmix(lc, role_c, 0.18))],
               _B(box["x"] + 0.1, y + 0.06, label_w - 0.2, h - 0.12), valign="middle", field="tier label")
        band_fill = c["surface"] if dark else (t["light_surface"] if role in ("dark", "neutral") else _kmix(role_c, "FFFFFF", 0.88))
        _krr(ctx, _B(band_x, y, band_w, h), band_fill)
        its = _list(tr.get("items"))
        k = max(1, len(its))
        iw = (band_w - 0.2 - 0.1 * (k - 1)) / k
        base_h = 0.32 if tr.get("base") else 0
        ih = h - 0.2 - (base_h + 0.1 if tr.get("base") else 0)
        chip = t["primary"] if role in ("dark", "neutral") else role_c
        for j, it in enumerate(its):
            x = band_x + 0.1 + j * (iw + 0.1)
            hl = bool(it.get("highlight"))
            _krr(ctx, _B(x, y + 0.1, iw, ih), chip if hl else (c["surface2"] if dark else "FFFFFF"),
                 line=None if (hl or dark) else {"color": _kmix(chip, "FFFFFF", 0.7), "width": 0.75})
            _ktext(ctx, item_paras(it, chip, hl), _B(x + 0.08, y + 0.16, iw - 0.16, ih - 0.12), valign="middle", field="layer item")
        if tr.get("base"):
            _krr(ctx, _B(band_x + 0.1, y + h - 0.1 - base_h, band_w - 0.2, base_h), t["primary"] if dark else t["dark_bg"])
            _ktext(ctx, [_P(tr["base"], 10, color="FFFFFF", align="center")], _B(band_x + 0.2, y + h - 0.1 - base_h, band_w - 0.4, base_h),
                   valign="middle", field="tier base")
        y += h + tg
    if side:
        sx = box["x"] + box["w"] - side_w
        sf = c["surface"] if dark else t["dark_bg"]
        inner = c["surface2"] if dark else _kmix(t["dark_bg"], "FFFFFF", 0.1)
        _krr(ctx, _B(sx, box["y"], side_w, box["h"]), sf, name="Side panel")
        _k_icon_dot(ctx, sx + 0.15, box["y"] + 0.15, 0.42, t["accent"], side.get("icon") or "shield", t["on_accent"])
        _ktext(ctx, [_P(side.get("title"), 12, bold=True, color="FFFFFF"), _P(side.get("subtitle"), 9.5, color=t["dark_muted"])],
               _B(sx + 0.66, box["y"] + 0.12, side_w - 0.76, 0.5), valign="middle", field="side title")
        sitems = _list(side.get("items"))
        n = len(sitems)
        if n:
            top, ig = box["y"] + 0.72, 0.08
            ih = (box["y"] + box["h"] - 0.12 - top - ig * (n - 1)) / n
            for j, it in enumerate(sitems):
                iy = top + j * (ih + ig)
                _krr(ctx, _B(sx + 0.12, iy, side_w - 0.24, ih), inner)
                _ktext(ctx, [_P(it.get("title"), 10.5, bold=True, color=t["accent"], spaceAfter=1),
                             _P(it.get("text"), 9.5, color=_kmix("FFFFFF", t["dark_bg"], 0.1))],
                       _B(sx + 0.2, iy + 0.05, side_w - 0.4, ih - 0.1), valign="middle", field="side item")


# ----- charts (native, editable) -----
def _k_numfmt(prefix, unit, d):
    f = (f'"{prefix}"' if prefix else "") + ("#,##0." + "0" * int(d) if d else "#,##0") + (f'"{unit}"' if unit else "")
    return f


def _k_font(font, size, color, bold=None, name=None):
    font.size = Pt(size)
    font.color.rgb = RGBColor.from_string(color)
    if bold is not None:
        font.bold = bold
    if name:
        font.name = name


def _k_chart_data(b, series):
    from pptx.chart.data import CategoryChartData as _CCD  # type: ignore
    cd = _CCD()
    cd.categories = [_ks(x) for x in _list(b.get("categories"))]
    for s in series:
        cd.add_series(_ks(s.get("name", "")), [_knum(v) for v in _list(s.get("values"))])
    return cd


def _k_series_fill(ser, color, alpha=None):
    fmt = ser.format
    fmt.fill.solid()
    fmt.fill.fore_color.rgb = RGBColor.from_string(color)
    if alpha is not None:
        sf = fmt.fill._xPr.find(_kqn("a:solidFill"))
        if sf is not None and len(sf):
            a = _K_ET.SubElement(sf[0], _kqn("a:alpha")); a.set("val", str(int(alpha * 1000)))


def _k_gridlines(axis, color, width=0.75):
    axis.has_major_gridlines = True
    gl = axis.major_gridlines.format.line
    gl.color.rgb = RGBColor.from_string(color)
    gl.width = Pt(width)


def _k_axis_line_off(axis):
    axis.format.line.fill.background()


@_kcomp("chart")
def _kc_chart(ctx, b, box):
    if ctx.dry:
        return
    t, c, dark = ctx.t, ctx.c, ctx.dark
    ctype = b.get("chart_type")
    if ctype == "combo":
        return _kc_combo(ctx, b, box)
    T = {"radar": ("radar",), "bar": ("bar", "bar", "clustered"), "column": ("bar", "col", "clustered"),
         "stacked_bar": ("bar", "bar", "stacked"), "stacked_column": ("bar", "col", "stacked"), "line": ("line",),
         "area": ("area",), "pie": ("pie",), "doughnut": ("doughnut",)}.get(ctype)
    if not T:
        ctx.warn(f"{ctx.where}: unknown chart_type '{ctype}'")
        return
    kind = T[0]
    rnd = kind in ("pie", "doughnut")
    series = _list(b.get("series"))
    cats = _list(b.get("categories"))
    d = int(b.get("decimals") or 0)
    fmt = _k_numfmt(b.get("prefix"), b.get("unit"), d)
    label_c = "E8ECF4" if dark else t["text"]
    axis_c = t["dark_muted"] if dark else t["muted"]
    show_val = b.get("show_values") is not False
    legend = b.get("legend") if b.get("legend") is not None else (len(series) > 1 or rnd)
    XL = XL_CHART_TYPE
    if kind == "bar":
        xt = {("bar", "clustered"): XL.BAR_CLUSTERED, ("col", "clustered"): XL.COLUMN_CLUSTERED,
              ("bar", "stacked"): XL.BAR_STACKED, ("col", "stacked"): XL.COLUMN_STACKED}[(T[1], T[2])]
    else:
        xt = {"line": XL.LINE_MARKERS, "area": XL.AREA, "pie": XL.PIE, "doughnut": XL.DOUGHNUT, "radar": XL.RADAR_MARKERS}[kind]
    gf = ctx.slide.shapes.add_chart(xt, _ke(box["x"]), _ke(box["y"]), _ke(box["w"]), _ke(box["h"]), _k_chart_data(b, series))
    ch = gf.chart
    ch.has_title = False
    ch.font.size = Pt(10)
    ch.font.name = t["fonts"]["body"]
    ch.font.color.rgb = RGBColor.from_string(label_c)
    ch.has_legend = bool(legend)
    if legend:
        ch.legend.position = XL_LEGEND_POSITION.RIGHT if rnd else XL_LEGEND_POSITION.BOTTOM
        ch.legend.include_in_layout = False
        _k_font(ch.legend.font, 10, label_c, name=t["fonts"]["body"])
    plot = ch.plots[0]
    colors = list(t["series"])
    if kind != "radar" and not rnd:
        try:
            plot.vary_by_categories = False
        except Exception:
            pass
    if rnd:
        pcols = [t["series"][i] if i < 6 else _kmix(t["series"][i % 6], "FFFFFF", 0.45) for i in range(len(cats))]
        ser = plot.series[0] if plot.series else None
        if ser is not None:
            for i in range(len(cats)):
                pt = ser.points[i]
                pt.format.fill.solid()
                pt.format.fill.fore_color.rgb = RGBColor.from_string(pcols[i])
                pt.format.line.color.rgb = RGBColor.from_string(c["bg"])
                pt.format.line.width = Pt(1)
        if kind == "doughnut":
            dn = plot._element
            hs = dn.find(_kqn("c:holeSize"))
            if hs is None:
                hs = _K_ET.SubElement(dn, _kqn("c:holeSize"))
            hs.set("val", "58")
    else:
        single = len(series) == 1 and not legend and kind == "bar" and T[2] != "stacked"
        for i, ser in enumerate(plot.series):
            col = t["primary"] if single else colors[i % len(colors)]
            if kind in ("line", "radar"):
                ser.format.line.color.rgb = RGBColor.from_string(col)
                ser.format.line.width = Pt(2.5 if kind == "line" else 2)
                ser.smooth = False
                try:
                    from pptx.enum.chart import XL_MARKER_STYLE  # type: ignore
                    ser.marker.style = XL_MARKER_STYLE.CIRCLE
                    ser.marker.size = 7 if kind == "line" else 6
                    ser.marker.format.fill.solid()
                    ser.marker.format.fill.fore_color.rgb = RGBColor.from_string(col)
                    ser.marker.format.line.color.rgb = RGBColor.from_string(col)
                except Exception:
                    pass
            else:
                _k_series_fill(ser, col, 70 if kind == "area" else None)
    if kind == "bar":
        plot.gap_width = 55
        if T[2] == "stacked":
            plot.overlap = 100
    want_labels = show_val and kind != "radar"
    if want_labels:
        plot.has_data_labels = True
        dl = plot.data_labels
        dl.number_format = fmt
        dl.number_format_is_linked = False
        dl.show_value = True
        try:
            dl.show_percentage = False
            dl.show_category_name = False
        except Exception:
            pass
        dcol = label_c
        if kind == "bar" and T[2] == "stacked":
            dcol = "FFFFFF"
        if kind == "pie":
            dcol = "FFFFFF"
        _k_font(dl.font, 10.5, dcol, bold=True, name=t["fonts"]["body"])
        try:
            if kind == "bar":
                dl.position = XL_LABEL_POSITION.CENTER if T[2] == "stacked" else XL_LABEL_POSITION.OUTSIDE_END
            elif kind == "line":
                dl.position = XL_LABEL_POSITION.ABOVE
            elif kind == "pie":
                dl.position = XL_LABEL_POSITION.BEST_FIT
        except Exception:
            pass
    if not rnd:
        try:
            ca, va = ch.category_axis, ch.value_axis
            _k_font(ca.tick_labels.font, 10, label_c, name=t["fonts"]["body"])
            _k_axis_line_off(ca)
            ca.has_major_gridlines = False
            _k_font(va.tick_labels.font, 9.5 if kind != "radar" else 8.5, axis_c, name=t["fonts"]["body"])
            va.tick_labels.number_format = fmt
            va.tick_labels.number_format_is_linked = False
            if kind == "bar":
                va.has_major_gridlines = False
                va.visible = not show_val
                _k_axis_line_off(va)
                if T[1] == "bar":
                    ca.reverse_order = True
            else:
                _k_gridlines(va, t["dark_line"] if dark else "E3E7EE")
                _k_axis_line_off(va)
        except Exception:
            pass


def _kc_combo(ctx, b, box):
    t, dark = ctx.t, ctx.dark
    label_c = "E8ECF4" if dark else t["text"]
    axis_c = t["dark_muted"] if dark else t["muted"]
    series = _list(b.get("series"))
    bars = [s for s in series if s.get("type") != "line"]
    lines = [s for s in series if s.get("type") == "line"]
    ordered = bars + lines
    gf = ctx.slide.shapes.add_chart(XL_CHART_TYPE.COLUMN_CLUSTERED, _ke(box["x"]), _ke(box["y"]), _ke(box["w"]), _ke(box["h"]),
                                    _k_chart_data(b, ordered))
    ch = gf.chart
    ch.has_title = False
    ch.font.size = Pt(10)
    ch.font.name = t["fonts"]["body"]
    ch.font.color.rgb = RGBColor.from_string(label_c)
    ch.has_legend = True
    ch.legend.position = XL_LEGEND_POSITION.BOTTOM
    ch.legend.include_in_layout = False
    _k_font(ch.legend.font, 10, label_c)
    plot = ch.plots[0]
    plot.gap_width = 55
    try:
        plot.vary_by_categories = False
    except Exception:
        pass
    bar_fmt = _k_numfmt(b.get("prefix"), b.get("unit"), int(b.get("decimals") or 0))
    line_fmt = _k_numfmt(b.get("line_prefix"), b.get("line_unit"), int(b.get("line_decimals") or 0))
    show = b.get("show_values") is not False
    for i, ser in enumerate(plot.series):
        if i < len(bars):
            _k_series_fill(ser, t["series"][i % 6])
    if show:
        plot.has_data_labels = True
        dl = plot.data_labels
        dl.number_format, dl.number_format_is_linked = bar_fmt, False
        _k_font(dl.font, 10, "FFFFFF", bold=True)
        try:
            dl.position = XL_LABEL_POSITION.INSIDE_BASE
        except Exception:
            pass
    ca, va = ch.category_axis, ch.value_axis
    _k_font(ca.tick_labels.font, 10, label_c)
    _k_font(va.tick_labels.font, 9, axis_c)
    va.tick_labels.number_format, va.tick_labels.number_format_is_linked = bar_fmt, False
    _k_gridlines(va, t["dark_line"] if dark else "E3E7EE", 0.5)
    _k_axis_line_off(va)
    # move the line series into a <c:lineChart> on secondary axes (keeps workbook references -> editable)
    bar_el = plot._element
    plot_area = bar_el.getparent()
    sers = bar_el.findall(_kqn("c:ser"))[len(bars):]
    if not sers:
        return
    C = lambda tag: _kqn("c:" + tag)  # noqa: E731
    lc = _K_ET.Element(C("lineChart"))
    _K_ET.SubElement(lc, C("grouping")).set("val", "standard")
    _K_ET.SubElement(lc, C("varyColors")).set("val", "0")
    line_cols = [t["accent"], t["danger"], t["info"]]
    for k, s in enumerate(sers):
        bar_el.remove(s)
        for tag in ("invertIfNegative",):
            e = s.find(C(tag))
            if e is not None:
                s.remove(e)
        sp = s.find(C("spPr"))
        if sp is not None:
            s.remove(sp)
        col = line_cols[k % 3]
        sp = _K_ET.fromstring(
            f'<c:spPr xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" '
            f'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:ln w="31750" cap="rnd">'
            f'<a:solidFill><a:srgbClr val="{col}"/></a:solidFill><a:round/></a:ln></c:spPr>')
        mk = _K_ET.fromstring(
            f'<c:marker xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" '
            f'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><c:symbol val="circle"/><c:size val="7"/>'
            f'<c:spPr><a:solidFill><a:srgbClr val="{col}"/></a:solidFill><a:ln><a:solidFill><a:srgbClr val="{col}"/></a:solidFill></a:ln></c:spPr></c:marker>')
        tx = s.find(C("tx"))
        idx_at = list(s).index(tx) + 1 if tx is not None else 2
        s.insert(idx_at, sp)
        s.insert(idx_at + 1, mk)
        if show:
            lbl_col = t["accent"] if dark else t["dark_bg"]
            dl = _K_ET.fromstring(
                f'<c:dLbls xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" '
                f'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><c:numFmt formatCode=\'{line_fmt}\' sourceLinked="0"/>'
                f'<c:spPr><a:noFill/><a:ln><a:noFill/></a:ln></c:spPr><c:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="1000" b="1">'
                f'<a:solidFill><a:srgbClr val="{lbl_col}"/></a:solidFill></a:defRPr></a:pPr><a:endParaRPr lang="en-US"/></a:p></c:txPr>'
                f'<c:dLblPos val="t"/><c:showLegendKey val="0"/><c:showVal val="1"/><c:showCatName val="0"/><c:showSerName val="0"/>'
                f'<c:showPercent val="0"/><c:showBubbleSize val="0"/></c:dLbls>')
            cat = s.find(C("cat"))
            s.insert(list(s).index(cat) if cat is not None else len(s), dl)
        sm = s.find(C("smooth"))
        if sm is None:
            _K_ET.SubElement(s, C("smooth")).set("val", "0")
        lc.append(s)
    _K_ET.SubElement(lc, C("marker")).set("val", "1")
    _K_ET.SubElement(lc, C("axId")).set("val", "60000001")
    _K_ET.SubElement(lc, C("axId")).set("val", "60000002")
    plot_area.insert(list(plot_area).index(bar_el) + 1, lc)
    # secondary axes: hidden category axis + right value axis
    axes_xml = (
        '<root xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">'
        '<c:catAx><c:axId val="60000001"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="1"/>'
        '<c:axPos val="b"/><c:majorTickMark val="none"/><c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/>'
        '<c:crossAx val="60000002"/><c:crosses val="autoZero"/><c:auto val="1"/><c:lblAlgn val="ctr"/><c:lblOffset val="100"/><c:noMultiLvlLbl val="0"/></c:catAx>'
        f'<c:valAx><c:axId val="60000002"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/>'
        f'<c:axPos val="r"/><c:numFmt formatCode=\'{line_fmt}\' sourceLinked="0"/><c:majorTickMark val="none"/><c:minorTickMark val="none"/>'
        f'<c:tickLblPos val="nextTo"/><c:spPr><a:ln><a:noFill/></a:ln></c:spPr><c:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="900">'
        f'<a:solidFill><a:srgbClr val="{axis_c}"/></a:solidFill></a:defRPr></a:pPr><a:endParaRPr lang="en-US"/></a:p></c:txPr>'
        '<c:crossAx val="60000001"/><c:crosses val="max"/><c:crossBetween val="between"/></c:valAx></root>')
    root = _K_ET.fromstring(axes_xml)
    last_ax = plot_area.findall(C("valAx"))[-1]
    pos = list(plot_area).index(last_ax) + 1
    for k, ax in enumerate(list(root)):
        plot_area.insert(pos + k, ax)


@_kcomp("waterfall")
def _kc_waterfall(ctx, b, box):
    t, c, dark = ctx.t, ctx.c, ctx.dark
    st = _list(b.get("steps"))
    n = len(st)
    if not n:
        return
    d = b.get("decimals") if b.get("decimals") is not None else 1
    cum, bars = 0.0, []
    for s in st:
        if s.get("total"):
            bars.append({"lo": min(0, cum), "hi": max(0, cum), "v": cum, "total": True, "label": s.get("label", "")})
            continue
        v = _knum(s.get("value"))
        lo, hi = min(cum, cum + v), max(cum, cum + v)
        cum += v
        bars.append({"lo": lo, "hi": hi, "v": v, "end": cum, "label": s.get("label", "")})
    vmin = min([0] + [x["lo"] for x in bars])
    vmax = max(x["hi"] for x in bars)
    top, bot = box["y"] + 0.32, box["y"] + box["h"] - 0.5
    ph = bot - top
    slot = box["w"] / n
    bw = slot * 0.6
    Y = lambda v: bot - ((v - vmin) / ((vmax - vmin) or 1)) * ph  # noqa: E731

    def f(v, sign):
        s = "+" if (sign and v > 0) else ("−" if v < 0 else "")
        return s + (b.get("prefix") or "") + _kfmt(abs(v), d) + (b.get("unit") or "")

    _kline(ctx, box["x"], Y(0), box["x"] + box["w"], Y(0), t["dark_line"] if dark else t["light_line"], 1)
    for i, br in enumerate(bars):
        x = box["x"] + i * slot + (slot - bw) / 2
        fill = t["primary"] if br.get("total") else (t["secondary"] if br["v"] >= 0 else t["danger"])
        _krect(ctx, _B(x, Y(br["hi"]), bw, max(0.02, Y(br["lo"]) - Y(br["hi"]))), fill, name=f"Bar {br['label']}")
        _ktext(ctx, [_P(f(br["v"], not br.get("total")), 10.5, bold=True, color="FFFFFF" if dark else t["dark_bg"], align="center")],
               _B(box["x"] + i * slot, Y(br["hi"]) - 0.27, slot, 0.24), valign="bottom", quiet=True)
        _ktext(ctx, [_P(br["label"], 10, color=c["text"], align="center", bold=bool(br.get("total")))],
               _B(box["x"] + i * slot + 0.03, bot + 0.06, slot - 0.06, 0.44), field="waterfall label")
        if i < n - 1:
            lvl = br["v"] if br.get("total") else br["end"]
            nx = box["x"] + (i + 1) * slot + (slot - bw) / 2
            _kline(ctx, x + bw, Y(lvl), nx, Y(lvl), t["dark_muted"] if dark else t["muted"], 0.75, "dash")


_K_RATING = [
    (re.compile(r"^(high|critical|red|no|✗|x|fail|poor)$", re.I), "danger"),
    (re.compile(r"^(medium|med|amber|partial|moderate|◐|in progress)$", re.I), "accent"),
    (re.compile(r"^(low|green|yes|✓|done|pass|good|strong)$", re.I), "success"),
]
_K_NUMCELL = re.compile(r"^[-+]?[€$£]?[\d,.]+[%xMBK]*$")


def _k_cell_border(cell, color, w_pt=1):
    tcPr = cell._tc.get_or_add_tcPr()
    for tag in ("a:lnL", "a:lnR", "a:lnT", "a:lnB"):
        old = tcPr.find(_kqn(tag))
        if old is not None:
            tcPr.remove(old)
    # borders must precede the fill in tcPr
    for k, tag in enumerate(("a:lnL", "a:lnR", "a:lnT", "a:lnB")):
        ln = _K_ET.Element(_kqn(tag))
        ln.set("w", str(int(w_pt * 12700)))
        sf = _K_ET.SubElement(ln, _kqn("a:solidFill"))
        _K_ET.SubElement(sf, _kqn("a:srgbClr")).set("val", color)
        tcPr.insert(k, ln)


@_kcomp("table")
def _kc_table(ctx, b, box):
    t, c, dark = ctx.t, ctx.c, ctx.dark
    cols = [_ks(x) for x in _list(b.get("columns"))]
    rows = [[_ks(x) for x in _list(r)] for r in _list(b.get("rows"))]
    k = len(cols)
    if not k:
        return
    rows = [(r + [""] * k)[:k] for r in rows]
    wts = b.get("col_widths") if isinstance(b.get("col_widths"), list) and len(b["col_widths"]) == k else [1] * k
    wts = [max(0.05, _knum(v, 1)) for v in wts]
    sw = sum(wts)
    col_w = [v / sw * box["w"] for v in wts]
    hl = set(x for x in _list(b.get("highlight_rows")) if isinstance(x, (int, float)))
    allr = [cols] + rows

    def row_heights(size):
        return [max(_k_measure([{"text": cell, "size": size, "bold": ri == 0 or ci == 0}], col_w[ci] - 0.14) for ci, cell in enumerate(r)) + 0.1
                for ri, r in enumerate(allr)]

    size = 10.5
    hs = row_heights(size)
    while sum(hs) > box["h"] and size > 8.5:
        size -= 0.5
        hs = row_heights(size)
    if sum(hs) > box["h"] + 0.05:
        ctx.warn(f"{ctx.where}: table is taller than its box — fewer rows or shorter cells")
    extra = max(0, box["h"] - sum(hs))
    if extra > 0:
        hs = [min(h + extra / len(hs), h * 1.6) for h in hs]
    if ctx.dry:
        return
    gf = ctx.slide.shapes.add_table(len(allr), k, _ke(box["x"]), _ke(box["y"]), _ke(box["w"]), _ke(sum(hs)))
    tbl = gf.table
    tbl.first_row = False
    tbl.horz_banding = False
    for ci, w in enumerate(col_w):
        tbl.columns[ci].width = _ke(w)
    for ri, h in enumerate(hs):
        tbl.rows[ri].height = _ke(h)
    border = t["dark_bg"] if dark else "FFFFFF"
    for ri, r in enumerate(allr):
        for ci, cell_txt in enumerate(r):
            cell = tbl.cell(ri, ci)
            align, bold, color = None, False, c["text"]
            if ri == 0:
                fill, color, bold = (t["primary"] if dark else t["dark_bg"]), "FFFFFF", True
            else:
                rr_ = ri - 1
                z = (c["surface2"] if dark else "FFFFFF") if rr_ % 2 else (c["surface"] if dark else t["light_surface"])
                fill = (t["primary"] if dark else t["primary_tint"]) if rr_ in hl else z
                bold = ci == 0 or rr_ in hl
                s = cell_txt.strip()
                for rx, role in _K_RATING:
                    if rx.match(s):
                        fill = _kmix(t[role], "FFFFFF", 0.15 if dark else 0.84)
                        color = _kread(_kmix(t[role], "FFFFFF", 0.15)) if dark else _kmix(t[role], "000000", 0.5 if role == "accent" else 0.15)
                        bold, align = True, "center"
                        break
                if align is None and _K_NUMCELL.match(s):
                    align = "right"
            cell.fill.solid()
            cell.fill.fore_color.rgb = RGBColor.from_string(fill)
            cell.margin_left = cell.margin_right = _ke(0.07)
            cell.margin_top = cell.margin_bottom = _ke(0.03)
            cell.vertical_anchor = MSO_ANCHOR.MIDDLE
            tf = cell.text_frame
            tf.word_wrap = True
            p = tf.paragraphs[0]
            p.alignment = {"center": PP_ALIGN.CENTER, "right": PP_ALIGN.RIGHT}.get(align, PP_ALIGN.LEFT)
            run = p.add_run()
            run.text = cell_txt
            _k_font(run.font, size, color, bold=bold, name=t["fonts"]["body"])
            _k_cell_border(cell, border)


@_kcomp("scorecard")
def _kc_scorecard(ctx, b, box):
    t, c, dark = ctx.t, ctx.c, ctx.dark
    opts = [_ks(o) for o in _list(b.get("options"))]
    crit = _list(b.get("criteria"))
    k = len(opts)
    if not k or not crit:
        return
    rows_n = len(crit) + (1 if b.get("show_total") else 0)
    legend_h, head_h = 0.26, 0.42
    crit_w = max(2.0, box["w"] * 0.36)
    ow = (box["w"] - crit_w) / k
    rh = min(0.56, (box["h"] - head_h - legend_h - 0.06) / rows_n)
    _krr(ctx, _B(box["x"], box["y"], box["w"], head_h), t["primary"] if dark else t["dark_bg"])
    _ktext(ctx, [_P("Criterion", 10.5, bold=True, color="FFFFFF")], _B(box["x"] + 0.12, box["y"], crit_w - 0.2, head_h), valign="middle", quiet=True)
    for j, o in enumerate(opts):
        _ktext(ctx, [_P(o, 10.5, bold=True, color="FFFFFF", align="center")], _B(box["x"] + crit_w + j * ow + 0.04, box["y"], ow - 0.08, head_h),
               valign="middle", field="scorecard option")
    ink = "FFFFFF" if dark else t["dark_bg"]

    def ball(cx, cy, dd, s):
        _kell(ctx, _B(cx - dd / 2, cy - dd / 2, dd, dd), c["surface2"] if dark else "FFFFFF", line={"color": ink, "width": 1.25})
        sc = max(0, min(4, int(round(_knum(s)))))
        if sc == 4:
            _kell(ctx, _B(cx - dd / 2, cy - dd / 2, dd, dd), ink)
        elif sc > 0:
            _kshp(ctx, "pie", _B(cx - dd / 2, cy - dd / 2, dd, dd), ink, adj={"adj1": 270 * 60000, "adj2": ((270 + 90 * sc) % 360) * 60000})

    wsum = sum(_knum(cr.get("weight"), 1) if cr.get("weight") is not None else 1 for cr in crit) or 1
    for i, cr in enumerate(crit):
        y = box["y"] + head_h + i * rh
        _krect(ctx, _B(box["x"], y, box["w"], rh), (c["bg"] if dark else "FFFFFF") if i % 2 else (c["surface"] if dark else t["light_surface"]))
        runs = [{"text": _ks(cr.get("name", "")), "bold": True}]
        if cr.get("weight") is not None:
            runs.append({"text": f"  ({round(_knum(cr['weight']) / wsum * 100)}%)", "bold": False, "color": c["muted"]})
        _ktext(ctx, [{"runs": runs, "size": 10.5, "color": c["text"]}], _B(box["x"] + 0.12, y, crit_w - 0.2, rh), valign="middle", field="criterion")
        for j, s in enumerate(_list(cr.get("scores"))[:k]):
            ball(box["x"] + crit_w + j * ow + ow / 2, y + rh / 2, min(0.3, rh * 0.62), s)
    if b.get("show_total"):
        y = box["y"] + head_h + len(crit) * rh
        _krect(ctx, _B(box["x"], y, box["w"], rh), c["surface2"] if dark else t["primary_tint"])
        _ktext(ctx, [_P("Weighted score (0–4)", 10.5, bold=True, color="FFFFFF" if dark else t["dark_bg"])], _B(box["x"] + 0.12, y, crit_w - 0.2, rh),
               valign="middle", quiet=True)
        totals = []
        for j in range(k):
            tot = 0.0
            for cr in crit:
                sc = _list(cr.get("scores"))
                w = _knum(cr.get("weight"), 1) if cr.get("weight") is not None else 1
                tot += (_knum(sc[j]) if j < len(sc) else 0) * w
            totals.append(tot / wsum)
        best = max(totals)
        for j, v in enumerate(totals):
            _ktext(ctx, [_P(f"{v:.1f}", 13, bold=True, color=(t["accent"] if dark else t["primary"]) if v == best else c["text"], align="center")],
                   _B(box["x"] + crit_w + j * ow, y, ow, rh), valign="middle", quiet=True)
    _ktext(ctx, [_P("Harvey balls: empty = 0 (no fit)  ·  quarter = 1  ·  half = 2  ·  three-quarter = 3  ·  full = 4 (strong fit)", 9, color=c["muted"])],
           _B(box["x"], box["y"] + box["h"] - legend_h + 0.04, box["w"], legend_h - 0.04), valign="middle", quiet=True)


@_kcomp("gantt")
def _kc_gantt(ctx, b, box):
    t, c, dark = ctx.t, ctx.c, ctx.dark
    P = b.get("periods") if isinstance(b.get("periods"), dict) else {}
    ph = [dict(p) for p in _list(b.get("phases")) if isinstance(p, dict)]
    n = len(ph)
    if not n:
        return
    label_w = max(1.4, min(2.3, box["w"] * 0.2))
    tx, tw = box["x"] + label_w, box["w"] - label_w
    head_h = 0.3
    ms = [dict(m) for m in _list(b.get("milestones")) if isinstance(m, dict)]
    ms_h = 0.46 if ms else 0
    ticks = []
    if _list(P.get("labels")):
        labels = _list(P["labels"])
        N = len(labels)
        cwid = tw / N
        X = lambda v: tx + v * cwid  # noqa: E731
        ticks = [{"x": X(i), "lx": X(i), "lw": cwid, "label": _ks(l)} for i, l in enumerate(labels)]
        for p in ph:
            p["_e"] = _knum(p.get("end")) + 1
        for m in ms:
            m["_at"] = _knum(m.get("at")) + 0.5
    else:
        s0 = _knum(P.get("start"), 0)
        s1 = _knum(P.get("end"), s0 + 1)
        step = _knum(P.get("step"), 1) or 1
        span = (s1 - s0) or 1
        X = lambda v: tx + ((v - s0) / span) * tw  # noqa: E731
        v = s0
        guard = 0
        while v <= s1 + 1e-9 and guard < 200:
            ticks.append({"x": X(v), "lx": X(v) - 0.3, "lw": 0.6, "label": (P.get("prefix") or "") + _kfmt(v, 1 if step % 1 else 0)})
            v += step
            guard += 1
        for p in ph:
            p["_e"] = _knum(p.get("end"))
        for m in ms:
            m["_at"] = _knum(m.get("at"))
    rows_top = box["y"] + head_h
    rh = (box["h"] - head_h - ms_h) / n
    for i in range(n):
        if i % 2 == 0:
            _krect(ctx, _B(box["x"], rows_top + i * rh, box["w"], rh), c["surface"] if dark else t["light_surface"])
    for tk in ticks:
        _kline(ctx, tk["x"], rows_top - 0.02, tk["x"], rows_top + n * rh, t["dark_line"] if dark else "E1E5EC", 0.75, "dash")
        _ktext(ctx, [_P(tk["label"], 9.5, color=c["muted"], align="center")], _B(tk["lx"], box["y"], tk["lw"], head_h - 0.04), valign="middle", quiet=True)
    for m in ms:
        x = X(m["_at"])
        y = rows_top + n * rh + 0.04
        _kline(ctx, x, rows_top, x, y + 0.02, t["accent"], 1, "sysDash")
        _kshp(ctx, "diamond", _B(x - 0.09, y, 0.18, 0.18), t["accent"])
        _ktext(ctx, [_P(m.get("label"), 9.5, bold=True, color="FFFFFF" if dark else t["dark_bg"], align="center")], _B(x - 0.8, y + 0.19, 1.6, ms_h - 0.22), quiet=True)
    cyc = ["neutral", "secondary", "info", "primary", "dark", "accent"]
    for i, p in enumerate(ph):
        y = rows_top + i * rh
        role = p.get("color") or cyc[i % len(cyc)]
        col = _k_role(t, role, t["primary"])
        bc = "FFFFFF" if (dark and role == "dark") else col
        _ktext(ctx, [_P(p.get("label"), 11, bold=True, color="FFFFFF" if dark else t["dark_bg"])], _B(box["x"] + 0.08, y, label_w - 0.14, rh),
               valign="middle", field="phase label")
        with_text = bool(p.get("text")) and rh >= 0.48
        bh = min(0.22, rh * 0.4)
        by = y + 0.07 if with_text else y + (rh - bh) / 2
        x1, x2 = X(_knum(p.get("start"))), X(p["_e"])
        if x2 <= x1:
            ctx.warn(f"{ctx.where}: phase '{p.get('label')}' end must be after start")
        bw_ = max(0.1, x2 - x1 - 0.06)
        _kshp(ctx, "roundRect", _B(x1 + 0.03, by, bw_, bh), bc, radius=bh / 2, name=f"Phase {p.get('label', '')}")
        p.update({"_x1": x1 + 0.03, "_x2": x1 + 0.03 + bw_, "_ym": by + bh / 2, "_top": by, "_bot": by + bh})
        if with_text:
            _ktext(ctx, [_P(p["text"], 9.5, color=c["text"])], _B(tx + 0.03, by + bh + 0.03, tw - 0.06, y + rh - (by + bh + 0.03) - 0.02),
                   field="phase text", min_size=8.5)
    col = t["dark_muted"] if dark else t["muted"]
    for i, p in enumerate(ph):
        dep = p.get("depends_on")
        deps = [] if dep is None else (dep if isinstance(dep, list) else [dep])
        for di in deps:
            try:
                di = int(di)
            except Exception:
                di = -1
            d = ph[di] if 0 <= di < n else None
            if not d or di == i or d.get("_x2") is None:
                ctx.warn(f"{ctx.where}: phase '{p.get('label')}' depends_on {di} is not a valid phase index")
                continue
            down = p["_ym"] > d["_ym"]
            if p["_x1"] - d["_x2"] >= 0.2:
                ex = d["_x2"] + 0.1
                _kline(ctx, d["_x2"], d["_ym"], ex, d["_ym"], col, 1)
                _kline(ctx, ex, d["_ym"], ex, p["_ym"], col, 1)
                _kline(ctx, ex, p["_ym"], p["_x1"] - 0.01, p["_ym"], col, 1, arrow=True)
            else:
                ex = max(p["_x1"] + 0.12, min(d["_x2"] - 0.12, p["_x1"] + 0.25))
                _kline(ctx, ex, d["_bot"] if down else d["_top"], ex, (p["_top"] - 0.01) if down else (p["_bot"] + 0.01), col, 1, arrow=True)


def _k_ramp(t, i, n):
    return _kmix(t["primary"], t["secondary"], i / (n - 1) if n > 1 else 0)


@_kcomp("pyramid")
def _kc_pyramid(ctx, b, box):
    t = ctx.t
    L = _list(b.get("levels"))
    n = len(L)
    if not n:
        return
    gap = 0.08
    lh = (box["h"] - gap * (n - 1)) / n
    w0 = box["w"] * 0.52
    for i, lv in enumerate(L):
        w = w0 + (box["w"] - w0) * (i / (n - 1)) if n > 1 else box["w"]
        x, y = box["x"] + (box["w"] - w) / 2, box["y"] + i * (lh + gap)
        fill = _k_ramp(t, i, n)
        tc = _kread(fill)
        _krr(ctx, _B(x, y, w, lh), fill, name=f"Level {lv.get('title', '')}")
        _ktext(ctx, [_P(lv.get("title"), 12, bold=True, color=tc, align="center", spaceAfter=1), _P(lv.get("text"), 10, color=tc, align="center")],
               _B(x + 0.12, y + 0.04, w - 0.24, lh - 0.08), valign="middle", field="pyramid level")


@_kcomp("funnel")
def _kc_funnel(ctx, b, box):
    t = ctx.t
    L = _list(b.get("stages"))
    n = len(L)
    if not n:
        return
    gap = 0.08
    lh = (box["h"] - gap * (n - 1)) / n
    wmin = box["w"] * 0.46
    for i, lv in enumerate(L):
        w = box["w"] - (box["w"] - wmin) * (i / (n - 1)) if n > 1 else box["w"]
        x, y = box["x"] + (box["w"] - w) / 2, box["y"] + i * (lh + gap)
        fill = _k_ramp(t, i, n)
        tc = _kread(fill)
        _krr(ctx, _B(x, y, w, lh), fill, name=f"Stage {lv.get('title', '')}")
        runs = ([{"text": _ks(lv["value"]) + "  ", "bold": True}] if lv.get("value") else []) + [{"text": _ks(lv.get("title", "")), "bold": True}]
        _ktext(ctx, [{"runs": runs, "size": 12, "color": tc, "align": "center", "spaceAfter": 1}, _P(lv.get("text"), 10, color=tc, align="center")],
               _B(x + 0.12, y + 0.04, w - 0.24, lh - 0.08), valign="middle", field="funnel stage")


@_kcomp("comparison")
def _kc_comparison(ctx, b, box):
    t, c, dark = ctx.t, ctx.c, ctx.dark
    cols = _list(b.get("columns"))
    k = len(cols)
    if not k:
        return
    vs = bool(b.get("vs")) and k == 2
    gap = 0.6 if vs else 0.25
    cw = (box["w"] - gap * (k - 1)) / k
    head_h = min(0.72, box["h"] * 0.22)
    foot_h = min(0.55, box["h"] * 0.16) if any(x.get("footer") for x in cols) else 0
    for i, col in enumerate(cols):
        x = box["x"] + i * (cw + gap)
        hl = bool(col.get("highlight"))
        hf = t["primary"] if hl else (c["surface2"] if dark else t["dark_bg"])
        _krr(ctx, _B(x, box["y"], cw, head_h), hf, name=f"Column {col.get('title', '')}")
        tx = x + 0.15
        if col.get("icon"):
            _k_icon_dot(ctx, x + 0.14, box["y"] + (head_h - 0.38) / 2, 0.38, t["accent"] if hl else t["primary"], col["icon"])
            tx = x + 0.62
        _ktext(ctx, [_P(col.get("title"), 13, bold=True, color="FFFFFF"), _P(col.get("subtitle"), 10, color=_kmix("FFFFFF", hf, 0.25))],
               _B(tx, box["y"] + 0.04, x + cw - 0.12 - tx, head_h - 0.08), valign="middle", field="column title")
        by = box["y"] + head_h + 0.08
        bh = box["h"] - head_h - 0.08 - (foot_h + 0.08 if foot_h else 0)
        _krr(ctx, _B(x, by, cw, bh), (c["surface2"] if dark else t["primary_tint"]) if hl else (c["surface"] if dark else t["light_surface"]))
        _ktext(ctx, [_P(bl, 11.5, color=c["text"], bullet=True, spaceAfter=5) for bl in _list(col.get("bullets"))],
               _B(x + 0.16, by + 0.14, cw - 0.3, bh - 0.24), field="comparison bullets")
        if foot_h and col.get("footer"):
            _krr(ctx, _B(x, by + bh + 0.08, cw, foot_h), t["primary"] if hl else (c["surface2"] if dark else t["light_line"]))
            _ktext(ctx, [_P(col["footer"], 11, bold=True, color=t["on_primary"] if hl else ("FFFFFF" if dark else t["dark_bg"]))],
                   _B(x + 0.15, by + bh + 0.08, cw - 0.3, foot_h), valign="middle", field="column footer")
    if vs:
        cx, cy = box["x"] + cw + gap / 2, box["y"] + head_h / 2
        _kell(ctx, _B(cx - 0.26, cy - 0.26, 0.52, 0.52), t["accent"], line={"color": c["bg"], "width": 2})
        _ktext(ctx, [_P("VS", 12, bold=True, color=t["on_accent"], align="center")], _B(cx - 0.26, cy - 0.26, 0.52, 0.52), valign="middle", quiet=True)


@_kcomp("matrix")
def _kc_matrix(ctx, b, box):
    t, c, dark = ctx.t, ctx.c, ctx.dark
    gx, gy, gw, gh, g = box["x"] + 0.42, box["y"], box["w"] - 0.42, box["h"] - 0.4, 0.08
    qw, qh = (gw - g) / 2, (gh - g) / 2
    for i, q in enumerate(_list(b.get("quadrants"))[:4]):
        x, y = gx + (i % 2) * (qw + g), gy + (i // 2) * (qh + g)
        hl = b.get("highlight") == i
        fill = t["primary"] if hl else ((c["surface2"] if i % 3 == 0 else c["surface"]) if dark else (t["primary_tint"] if i % 3 == 0 else t["light_surface"]))
        _krr(ctx, _B(x, y, qw, qh), fill, name=f"Quadrant {q.get('title', '')}")
        tc = t["on_primary"] if hl else ("FFFFFF" if dark else t["dark_bg"])
        _ktext(ctx, [_P(q.get("title"), 12.5, bold=True, color=tc, spaceAfter=2), _P(q.get("text"), 10, color=t["on_primary"] if hl else c["text"])],
               _B(x + 0.14, y + 0.1, qw * 0.62, qh - 0.2), field="quadrant")
    for p in _list(b.get("points")):
        px = gx + max(0, min(1, _knum(p.get("x")))) * gw
        py = gy + (1 - max(0, min(1, _knum(p.get("y"))))) * gh
        _kell(ctx, _B(px - 0.09, py - 0.09, 0.18, 0.18), t["accent"], line={"color": "FFFFFF", "width": 1.25})
        qi = (0 if _knum(p.get("y")) >= 0.5 else 2) + (1 if _knum(p.get("x")) >= 0.5 else 0)
        q_right = gx + ((qi % 2) + 1) * qw + (qi % 2) * g
        right = px + 1.7 < q_right
        lc = t["on_primary"] if b.get("highlight") == qi else ("FFFFFF" if dark else t["dark_bg"])
        _ktext(ctx, [_P(p.get("label"), 9.5, bold=True, color=lc, align="left" if right else "right")],
               _B(px + 0.13 if right else px - 1.73, py - 0.12, 1.6, 0.24), valign="middle", quiet=True)

    def ax(a):
        a = a if isinstance(a, dict) else {}
        return (f"{a['low']}  ←  " if a.get("low") else "") + _ks(a.get("label", "")) + (f"  →  {a['high']}" if a.get("high") else "")

    _ktext(ctx, [_P(ax(b.get("x_axis")), 10.5, bold=True, color=c["muted"], align="center")], _B(gx, gy + gh + 0.08, gw, 0.3), valign="middle", quiet=True)
    _ktext(ctx, [_P(ax(b.get("y_axis")), 10.5, bold=True, color=c["muted"], align="center")],
           _B(box["x"] + 0.17 - gh / 2, gy + gh / 2 - 0.15, gh, 0.3), valign="middle", rotate=270, quiet=True)


@_kcomp("hub")
def _kc_hub(ctx, b, box):
    t, c, dark = ctx.t, ctx.c, ctx.dark
    nodes = _list(b.get("nodes"))
    n = len(nodes)
    center = b.get("center") if isinstance(b.get("center"), dict) else {"title": ""}
    nw, nh = min(2.5, box["w"] * 0.25), min(0.9, box["h"] * 0.24)
    cw, chh = min(2.3, box["w"] * 0.24), min(1.3, box["h"] * 0.34)
    cx, cy = box["x"] + box["w"] / 2, box["y"] + box["h"] / 2
    rx, ry = (box["w"] - nw) / 2, (box["h"] - nh) / 2
    pos = []
    for i in range(n):
        a = _kmath.radians(-90 + 180 / n + (360 / n) * i)
        pos.append((cx + rx * _kmath.cos(a), cy + ry * _kmath.sin(a)))
    for p in pos:
        _kline(ctx, cx, cy, p[0], p[1], t["dark_line"] if dark else _kmix(t["primary"], "FFFFFF", 0.55), 1.5, "dash")
    _kell(ctx, _B(cx - cw / 2, cy - chh / 2, cw, chh), t["primary"], shadow=True, name="Hub center")
    _ktext(ctx, [_P(center.get("title"), 14, bold=True, color=t["on_primary"], align="center", spaceAfter=1),
                 _P(center.get("text"), 10, color=t["on_primary"], align="center")],
           _B(cx - cw * 0.36, cy - chh * 0.36, cw * 0.72, chh * 0.72), valign="middle", field="hub center")
    for i, nd in enumerate(nodes):
        x, y = pos[i][0] - nw / 2, pos[i][1] - nh / 2
        _krr(ctx, _B(x, y, nw, nh), c["surface"] if dark else "FFFFFF",
             line=None if dark else {"color": t["light_line"], "width": 0.75}, shadow=not dark)
        tx = x + 0.12
        if nd.get("icon"):
            _k_icon_dot(ctx, x + 0.12, y + (nh - 0.36) / 2, 0.36, t["secondary"] if i % 2 else t["primary"], nd["icon"])
            tx = x + 0.58
        _ktext(ctx, [_P(nd.get("title"), 11, bold=True, color="FFFFFF" if dark else t["dark_bg"], spaceAfter=1), _P(nd.get("text"), 9.5, color=c["text"])],
               _B(tx, y + 0.06, x + nw - 0.1 - tx, nh - 0.12), valign="middle", field="hub node")


@_kcomp("tree")
def _kc_tree(ctx, b, box):
    t, c, dark = ctx.t, ctx.c, ctx.dark
    br = _list(b.get("branches"))
    k = len(br)
    if not k:
        return
    root = b.get("root") if isinstance(b.get("root"), dict) else {"title": ""}
    gap = 0.2
    rw, rh = min(5, box["w"] * 0.5), 0.62
    rx = box["x"] + (box["w"] - rw) / 2
    _krr(ctx, _B(rx, box["y"], rw, rh), t["primary"] if dark else t["dark_bg"], name="Tree root")
    _ktext(ctx, [_P(root.get("title"), 13, bold=True, color="FFFFFF", align="center"), _P(root.get("text"), 10, color="E6E9F0", align="center")],
           _B(rx + 0.15, box["y"] + 0.04, rw - 0.3, rh - 0.08), valign="middle", field="tree root")
    bw = (box["w"] - gap * (k - 1)) / k
    bus_y = box["y"] + rh + 0.22
    hy, hh = bus_y + 0.22, 0.5
    lc = t["dark_line"] if dark else _kmix(t["dark_bg"], "FFFFFF", 0.6)
    _kline(ctx, box["x"] + box["w"] / 2, box["y"] + rh, box["x"] + box["w"] / 2, bus_y, lc, 1.25)
    _kline(ctx, box["x"] + bw / 2, bus_y, box["x"] + box["w"] - bw / 2, bus_y, lc, 1.25)
    max_leaves = max([1] + [len(_list(x.get("items"))) for x in br])
    leaf_top = hy + hh + 0.12
    lg = 0.08
    lh = min(0.55, (box["y"] + box["h"] - leaf_top - lg * (max_leaves - 1)) / max_leaves)
    for i, x0 in enumerate(br):
        x = box["x"] + i * (bw + gap)
        fill = t["secondary"] if i % 2 else t["primary"]
        _kline(ctx, x + bw / 2, bus_y, x + bw / 2, hy, lc, 1.25)
        _krr(ctx, _B(x, hy, bw, hh), fill, name=f"Branch {x0.get('title', '')}")
        _ktext(ctx, [_P(x0.get("title"), 11.5, bold=True, color=_kread(fill), align="center")], _B(x + 0.08, hy, bw - 0.16, hh), valign="middle", field="branch")
        for j, leaf in enumerate(_list(x0.get("items"))):
            y = leaf_top + j * (lh + lg)
            _krr(ctx, _B(x + 0.08, y, bw - 0.16, lh), c["surface"] if dark else "FFFFFF",
                 line=None if dark else {"color": _kmix(fill, "FFFFFF", 0.6), "width": 0.75})
            _ktext(ctx, [_P(leaf, 10, color=c["text"])], _B(x + 0.16, y + 0.03, bw - 0.32, lh - 0.06), valign="middle", field="leaf")


@_kcomp("bullets")
def _kc_bullets(ctx, b, box):
    t, c, dark = ctx.t, ctx.c, ctx.dark
    inner = box
    if b.get("panel"):
        _krr(ctx, box, c["surface"] if dark else t["light_surface"])
        inner = _B(box["x"] + 0.2, box["y"] + 0.16, box["w"] - 0.4, box["h"] - 0.3)
    paras = []
    num = bool(b.get("numbered"))
    for it in _list(b.get("items")):
        if isinstance(it, dict):
            subs = _list(it.get("sub"))
            runs = ([{"text": _ks(it["lead"]) + " ", "bold": True, "color": "FFFFFF" if dark else t["dark_bg"]}] if it.get("lead") else []) + [{"text": _ks(it.get("text", ""))}]
            paras.append({"runs": runs, "size": 14, "color": c["text"], "bullet": True, "numbered": num, "spaceAfter": 3 if subs else 8})
            for j, s in enumerate(subs):
                paras.append(_P(s, 12, color=c["muted"], bullet=True, indent=1, spaceAfter=8 if j == len(subs) - 1 else 2))
        else:
            paras.append(_P(it, 14, color=c["text"], bullet=True, numbered=num, spaceAfter=8))
    _ktext(ctx, paras, inner, field="bullets", min_size=10)


@_kcomp("callout")
def _kc_callout(ctx, b, box):
    t, c, dark = ctx.t, ctx.c, ctx.dark
    tone = b.get("tone") or "primary"
    fill = (c["surface"] if dark else t["dark_bg"]) if tone == "dark" else t["accent"] if tone == "accent" else \
        (c["surface2"] if dark else t["primary_tint"]) if tone == "tint" else t["primary"]
    tc = ("FFFFFF" if dark else t["dark_bg"]) if tone == "tint" else _kread(fill, t["dark_bg"])
    sub = _kmix(tc, fill, 0.22)
    _krr(ctx, box, fill, name="Callout")
    pad = min(0.3, box["h"] * 0.14)
    kind = b.get("kind") or "statement"
    if kind == "stat":
        value = _ks(b.get("value", ""))
        narrow = box["w"] < box["h"] * 1.6
        vw = box["w"] - 2 * pad if narrow else min(box["w"] * 0.42, 3.4)
        vs = _k_single_line_size(value, vw, min(60, (box["h"] * 0.42 if narrow else box["h"] * 0.7) * 72 * 0.85))
        if narrow:
            _ktext(ctx, [_P(value, vs, bold=True, color=tc)], _B(box["x"] + pad, box["y"] + pad, vw, vs * 1.15 / 72), valign="middle", quiet=True)
            ty = box["y"] + pad + vs * 1.15 / 72 + 0.05
            _ktext(ctx, [_P(b.get("text"), 13, color=tc, spaceAfter=4), _P(b.get("attribution"), 9.5, italic=True, color=sub)],
                   _B(box["x"] + pad, ty, box["w"] - 2 * pad, box["y"] + box["h"] - pad - ty), field="callout text")
        else:
            _ktext(ctx, [_P(value, vs, bold=True, color=tc)], _B(box["x"] + pad, box["y"], vw, box["h"]), valign="middle", quiet=True)
            _ktext(ctx, [_P(b.get("text"), 15, color=tc, spaceAfter=4), _P(b.get("attribution"), 10, italic=True, color=sub)],
                   _B(box["x"] + pad + vw + 0.2, box["y"] + pad, box["w"] - vw - 2 * pad - 0.2, box["h"] - 2 * pad), valign="middle", field="callout text")
        return
    tx = box["x"] + pad
    if kind == "quote" or b.get("icon"):
        d = min(0.5, box["h"] * 0.4)
        _k_icon(ctx, "FaQuoteLeft" if kind == "quote" else b["icon"], t["dark_bg"] if tone == "accent" else t["accent"], box["x"] + pad, box["y"] + pad, d, d)
        tx = box["x"] + pad + d + 0.2
    attr = b.get("attribution")
    _ktext(ctx, [_P(b.get("text"), 17 if kind == "quote" else 19, italic=kind == "quote", bold=kind != "quote", color=tc, spaceAfter=6),
                 _P((("— " if kind == "quote" else "") + _ks(attr)) if attr else None, 10.5, color=sub)],
           _B(tx, box["y"] + pad, box["x"] + box["w"] - pad - tx, box["h"] - 2 * pad), valign="middle", field="callout", min_size=11)


@_kcomp("agenda")
def _kc_agenda(ctx, b, box):
    t, c, dark = ctx.t, ctx.c, ctx.dark
    it = _list(b.get("items"))
    n = len(it)
    if not n:
        return
    cols = 2 if n > 5 else 1
    per = (n + cols - 1) // cols
    gap, cg = 0.14, 0.4
    cw = (box["w"] - cg * (cols - 1)) / cols
    ih = min(1.0, (box["h"] - gap * (per - 1)) / per)
    for i, x in enumerate(it):
        col, r = i // per, i % per
        X, Y = box["x"] + col * (cw + cg), box["y"] + r * (ih + gap)
        hl = b.get("highlight") == i
        _krr(ctx, _B(X, Y, cw, ih), t["primary"] if hl else (c["surface"] if dark else t["light_surface"]))
        nw = min(0.9, ih * 1.1)
        _ktext(ctx, [_P(f"{i + 1:02d}", min(26, ih * 34), bold=True, color=t["accent"] if (hl or dark) else t["primary"])], _B(X + 0.15, Y, nw, ih),
               valign="middle", quiet=True)
        _ktext(ctx, [_P(x.get("title"), 14, bold=True, color=t["on_primary"] if hl else ("FFFFFF" if dark else t["dark_bg"]), spaceAfter=1),
                     _P(x.get("text"), 10.5, color=t["on_primary"] if hl else c["muted"])],
               _B(X + 0.15 + nw, Y + 0.05, cw - nw - 0.3, ih - 0.1), valign="middle", field="agenda item")


@_kcomp("progress")
def _kc_progress(ctx, b, box):
    t, c, dark = ctx.t, ctx.c, ctx.dark
    it = _list(b.get("items"))
    n = len(it)
    if not n:
        return
    rh = min(0.62, box["h"] / n)
    unit = _ks(b.get("unit")) if b.get("unit") is not None else "%"
    lw = max(1.2, min(2.8, box["w"] * 0.3))
    vw = 0.75
    tx = box["x"] + lw + 0.12
    tw = box["w"] - lw - vw - 0.24
    y0 = box["y"] + (box["h"] - rh * n) / 2
    for i, x in enumerate(it):
        y = y0 + i * rh
        mx = _knum(x.get("max"), 100) or 100
        val = _knum(x.get("value"))
        frac = max(0, min(1, val / mx))
        bh = min(0.2, rh * 0.36)
        by = y + (rh - bh) / 2
        _ktext(ctx, [_P(x.get("label"), 11, bold=True, color="FFFFFF" if dark else t["dark_bg"]), _P(x.get("text") if rh > 0.45 else None, 9, color=c["muted"])],
               _B(box["x"], y, lw, rh), valign="middle", field="progress label")
        _kshp(ctx, "roundRect", _B(tx, by, tw, bh), c["surface2"] if dark else t["light_line"], radius=bh / 2)
        if frac > 0:
            _kshp(ctx, "roundRect", _B(tx, by, max(bh, tw * frac), bh), t["secondary"] if i % 2 else t["primary"], radius=bh / 2, name=f"Bar {x.get('label', '')}")
        if x.get("target") is not None:
            mxp = tx + tw * max(0, min(1, _knum(x["target"]) / mx))
            _krect(ctx, _B(mxp - 0.025, by - 0.07, 0.05, bh + 0.14), t["accent"])
        _ktext(ctx, [_P(_kfmt(val, 1 if val % 1 else 0) + unit, 12, bold=True, color="FFFFFF" if dark else t["dark_bg"], align="right")],
               _B(box["x"] + box["w"] - vw, y, vw, rh), valign="middle", quiet=True)
    if any(x.get("target") is not None for x in it):
        _ktext(ctx, [_P("▮ target", 9, color=t["accent"], align="right", bold=True)], _B(box["x"] + box["w"] - 1.5, y0 + n * rh, 1.5, 0.2), quiet=True)


# ----- images (URL, data URI, or file in the images folder) -----
_K_IMG_CACHE: dict = {}


def _k_load_image(ctx, src):
    """Return (png_bytes, w, h) or None."""
    if not src or not _HAS_PIL:
        return None
    src = str(src)
    if src in _K_IMG_CACHE:
        return _K_IMG_CACHE[src]
    out = None
    try:
        raw = None
        if src.startswith("data:image/"):
            raw = base64.b64decode(src.split(",", 1)[1])
        elif re.match(r"^https?://", src, re.I):
            if _HAS_HTTPX:
                r = httpx.get(src, timeout=15, follow_redirects=True, headers={"User-Agent": "generate-slides/2.0"})
                r.raise_for_status()
                raw = r.content
            else:
                import urllib.request as _ur
                with _ur.urlopen(_ur.Request(src, headers={"User-Agent": "generate-slides/2.0"}), timeout=15) as fh:
                    raw = fh.read()
        else:
            base = Path(getattr(ctx.valves, "kit_images_dir", "") or "/app/backend/data/deck-images").resolve()
            p = (base / src.lstrip("/\\")).resolve()
            if not str(p).startswith(str(base)):
                raise ValueError("path outside image folder")
            raw = p.read_bytes()
        im = Image.open(BytesIO(raw))
        try:
            from PIL import ImageOps  # type: ignore
            im = ImageOps.exif_transpose(im)
        except Exception:
            pass
        im = im.convert("RGBA")
        maxpx = int(getattr(ctx.valves, "max_image_px", 1600) or 1600)
        if max(im.size) > maxpx:
            im.thumbnail((maxpx, maxpx))
        bio = BytesIO()
        im.save(bio, format="PNG")
        out = (bio.getvalue(), im.size[0], im.size[1])
    except Exception as e:  # noqa: BLE001
        ctx.warn(f"{ctx.where}: image '{src[:80]}' could not be loaded ({e}) — placeholder drawn")
        out = None
    _K_IMG_CACHE[src] = out
    return out


def _k_place_image(ctx, img, box, fit="cover", ellipse=False):
    if ctx.dry or not img:
        return
    png, iw, ih = img
    if fit == "contain":
        r = min(box["w"] / iw, box["h"] / ih)
        w, h = iw * r, ih * r
        _k_image(ctx, png, box["x"] + (box["w"] - w) / 2, box["y"] + (box["h"] - h) / 2, w, h, ellipse)
        return
    pic = _k_image(ctx, png, box["x"], box["y"], box["w"], box["h"], ellipse)
    if pic is None:
        return
    tr, ir = box["w"] / box["h"], iw / ih
    if ir > tr:  # wider: crop left/right
        cut = (1 - tr / ir) / 2
        pic.crop_left = pic.crop_right = cut
    elif ir < tr:
        cut = (1 - ir / tr) / 2
        pic.crop_top = pic.crop_bottom = cut


def _k_initials(name):
    ws = [w for w in str(name or "?").split() if re.match(r"^[A-Za-z0-9]", w)][:2]
    return "".join(w[0].upper() for w in ws) or "?"


# ---------- strategy ----------
@_kcomp("swot")
def _kc_swot(ctx, b, box):
    t, c, dark = ctx.t, ctx.c, ctx.dark
    g = 0.14
    qw, qh = (box["w"] - g) / 2, (box["h"] - g) / 2
    titles = _list(b.get("titles"))
    Q = [("S", "Strengths", "strengths", t["success"]), ("W", "Weaknesses", "weaknesses", t["danger"]),
         ("O", "Opportunities", "opportunities", t["info"]), ("T", "Threats", "threats", t["accent"])]
    for i, (L, dflt, key, col) in enumerate(Q):
        title = _ks(titles[i]) if i < len(titles) and titles[i] else dflt
        x, y = box["x"] + (i % 2) * (qw + g), box["y"] + (i // 2) * (qh + g)
        hh = min(0.48, qh * 0.28)
        _krr(ctx, _B(x, y, qw, qh), c["surface"] if dark else _kmix(col, "FFFFFF", 0.9), name=f"SWOT {title}")
        _krr(ctx, _B(x, y, qw, hh), col)
        tc = _kread(col)
        _ktext(ctx, [{"runs": [{"text": L + "  ", "bold": True}, {"text": title, "bold": True}], "size": 13, "color": tc}], _B(x + 0.14, y, qw - 0.28, hh),
               valign="middle", quiet=True)
        _ktext(ctx, [_P(s, 11, color=c["text"], bullet=True, spaceAfter=3) for s in _list(b.get(key))],
               _B(x + 0.16, y + hh + 0.1, qw - 0.3, qh - hh - 0.18), field="swot items")


@_kcomp("venn")
def _kc_venn(ctx, b, box):
    t, c = ctx.t, ctx.c
    sets = _list(b.get("sets"))[:3]
    n = len(sets)
    if n < 2:
        return
    cols = [t["primary"], t["secondary"], t["accent"]]
    D = min(box["h"] * 0.98, box["w"] * 0.6) if n == 2 else min(box["h"] * 0.66, box["w"] * 0.42)
    cx, cy = box["x"] + box["w"] / 2, box["y"] + box["h"] / 2
    off = D * 0.3
    centers = [(cx - off, cy), (cx + off, cy)] if n == 2 else [(cx - off, cy - off * 0.6), (cx + off, cy - off * 0.6), (cx, cy + off * 0.75)]
    for i, p in enumerate(centers):
        _kell(ctx, _B(p[0] - D / 2, p[1] - D / 2, D, D), cols[i], transparency=45, line={"color": cols[i], "width": 1.5}, name=f"Set {sets[i].get('title', '')}")
    lw = D * 0.42
    for i, p in enumerate(centers):
        dx = (1 if i else -1) if n == 2 else (-1 if i == 0 else 1 if i == 1 else 0)
        dy = (1 if i == 2 else -0.6) if n == 3 else 0
        lx = p[0] + dx * D * 0.2 - lw / 2
        ly = p[1] + dy * D * 0.2 - D * 0.18
        _ktext(ctx, [_P(sets[i].get("title"), 13, bold=True, color=_kdk(ctx), align="center", spaceAfter=2), _P(sets[i].get("text"), 10, color=c["text"], align="center")],
               _B(lx, ly, lw, D * 0.36), valign="middle", field="venn set")
    ov = b.get("overlap")
    if isinstance(ov, dict):
        ow, oh = min(2.2, D * 0.45), 0.62
        oy = cy - oh / 2 if n == 2 else cy - oh / 2 + off * 0.05
        _krr(ctx, _B(cx - ow / 2, oy, ow, oh), c["surface2"] if ctx.dark else "FFFFFF", shadow=True)
        _ktext(ctx, [_P(ov.get("title"), 11.5, bold=True, color=t["accent"] if ctx.dark else t["primary"], align="center"),
                     _P(ov.get("text"), 9, color=c["text"], align="center")],
               _B(cx - ow / 2 + 0.06, oy + 0.03, ow - 0.12, oh - 0.06), valign="middle", field="venn overlap")


@_kcomp("cycle")
def _kc_cycle(ctx, b, box):
    t, c = ctx.t, ctx.c
    st = _list(b.get("steps"))
    n = len(st)
    if not n:
        return
    d = min(0.62, box["h"] * 0.16)
    tw = min(2.5, box["w"] * 0.24)
    cx, cy = box["x"] + box["w"] / 2, box["y"] + box["h"] / 2
    rx = max(0.8, box["w"] / 2 - tw - d / 2 - 0.15)
    ry = max(0.6, box["h"] / 2 - d / 2 - 0.62)
    P = []
    for i in range(n):
        a = _kmath.radians(-90 + (360 / n) * i)
        P.append((a, cx + rx * _kmath.cos(a), cy + ry * _kmath.sin(a)))
    ac = t["dark_line"] if ctx.dark else _kmix(t["primary"], "FFFFFF", 0.45)
    for i in range(n):
        p, q = P[i], P[(i + 1) % n]
        dx, dy = q[1] - p[1], q[2] - p[2]
        L = _kmath.hypot(dx, dy) or 1
        k = (d / 2 + 0.08) / L
        _kline(ctx, p[1] + dx * k, p[2] + dy * k, q[1] - dx * k, q[2] - dy * k, ac, 2, arrow=True)
    center = b.get("center")
    if isinstance(center, dict):
        cw, chh = min(2.4, rx * 1.3), min(1.2, ry * 1.05)
        _kell(ctx, _B(cx - cw / 2, cy - chh / 2, cw, chh), c["surface2"] if ctx.dark else t["primary_tint"])
        _ktext(ctx, [_P(center.get("title"), 13, bold=True, color=_kdk(ctx), align="center"), _P(center.get("text"), 9.5, color=c["text"], align="center")],
               _B(cx - cw * 0.36, cy - chh * 0.36, cw * 0.72, chh * 0.72), valign="middle", field="cycle center")
    for i, s in enumerate(st):
        a, px, py = P[i]
        _k_icon_dot(ctx, px - d / 2, py - d / 2, d, t["secondary"] if i % 2 else t["primary"], s.get("icon") or "sync")
        bd = d * 0.42
        _kell(ctx, _B(px + d * 0.2, py - d / 2 - bd * 0.25, bd, bd), t["accent"])
        _ktext(ctx, [_P(str(i + 1), max(7.5, bd * 34), bold=True, color=t["on_accent"], align="center")], _B(px + d * 0.2, py - d / 2 - bd * 0.25, bd, bd),
               valign="middle", quiet=True)
        cs, sn, th = _kmath.cos(a), _kmath.sin(a), 0.75
        if cs > 0.35:
            tb, al = _B(px + d / 2 + 0.1, py - th / 2, tw, th), "left"
        elif cs < -0.35:
            tb, al = _B(px - d / 2 - 0.1 - tw, py - th / 2, tw, th), "right"
        else:
            tb, al = _B(px - tw / 2, (py - d / 2 - th - 0.12) if sn < 0 else (py + d / 2 + 0.06), tw, th), "center"
        _ktext(ctx, [_P(s.get("title"), 11.5, bold=True, color=_kdk(ctx), align=al, spaceAfter=1), _P(s.get("text"), 9.5, color=c["text"], align=al)], tb,
               valign="middle" if abs(cs) > 0.35 else ("bottom" if sn < 0 else "top"), field="cycle step")


@_kcomp("value_chain")
def _kc_value_chain(ctx, b, box):
    t, c, dark = ctx.t, ctx.c, ctx.dark
    sup, pri = _list(b.get("support")), _list(b.get("primary"))
    if not pri:
        return
    mw = 0 if b.get("margin") is False else 0.95
    W = box["w"] - mw - (0.1 if mw else 0)
    n = len(pri)
    aw = W / n
    ph = min(1.15, aw * 0.75, box["h"] * (0.5 if sup else 1))
    sg = 0.06
    sh = min(0.72, (box["h"] - ph - 0.12 - sg * (len(sup) - 1)) / len(sup)) if sup else 0
    sup_h = sh * len(sup) + sg * (len(sup) - 1) if sup else 0
    top = box["y"] + (box["h"] - (sup_h + (0.12 if sup else 0) + ph)) / 2
    for i, s in enumerate(sup):
        y = top + i * (sh + sg)
        _krr(ctx, _B(box["x"], y, W, sh), c["surface2"] if dark else _kmix(t["secondary"], "FFFFFF", 0.82))
        runs = [{"text": _ks(s.get("title", "")), "bold": True, "color": _kdk(ctx)}] + ([{"text": "   " + _ks(s["text"]), "bold": False, "color": c["text"]}] if s.get("text") else [])
        _ktext(ctx, [{"runs": runs, "size": 10.5}], _B(box["x"] + 0.14, y, W - 0.28, sh), valign="middle", field="support activity")
    py = top + sup_h + (0.12 if sup else 0)
    depth = 0.5 * min(ph, aw + 0.2)
    for i, p in enumerate(pri):
        x = box["x"] + i * aw
        fill = _kmix(t["primary"], t["info"], (i / (n - 1)) * 0.5 if n > 1 else 0)
        _kshp(ctx, "homePlate" if i == 0 else "chevron", _B(x, py, aw + (depth * 0.55 if i < n - 1 else 0), ph), fill, name=f"Primary {p.get('title', '')}")
        tc = _kread(fill)
        tx = x + (depth if i else 0.1)
        tw = aw - (depth if i else 0.1) + (depth * 0.15 if i < n - 1 else -depth * 0.45)
        _ktext(ctx, [_P(p.get("title"), 11.5, bold=True, color=tc, align="center", spaceAfter=2), _P(p.get("text"), 9.5, color=tc, align="center")],
               _B(tx, py + 0.06, tw, ph - 0.12), valign="middle", field="primary activity")
    if mw:
        mh = sup_h + (0.12 if sup else 0) + ph
        _kshp(ctx, "homePlate", _B(box["x"] + W + 0.1, top, mw, mh), t["accent"])
        _ktext(ctx, [_P(b.get("margin_label") or "Margin", 12, bold=True, color=t["on_accent"], align="center")], _B(box["x"] + W + 0.1, top, mw * 0.78, mh),
               valign="middle", quiet=True)


@_kcomp("concentric")
def _kc_concentric(ctx, b, box):
    t, c = ctx.t, ctx.c
    R = _list(b.get("rings"))
    n = len(R)
    if not n:
        return
    D = min(box["h"], box["w"] * 0.52)
    x0, yb = box["x"], box["y"] + box["h"]
    cols = [_kmix(t["primary"], t["secondary"], i / (n - 1) if n > 1 else 0) for i in range(n)]
    for i in range(n - 1, -1, -1):
        s = D * (0.34 + 0.66 * (i / max(1, n - 1)))
        _kell(ctx, _B(x0 + (D - s) / 2, yb - s, s, s), cols[i], line={"color": c["bg"] if ctx.dark else "FFFFFF", "width": 2}, name=f"Ring {R[i].get('title', '')}")
        prev = D * (0.34 + 0.66 * ((i - 1) / max(1, n - 1))) if i > 0 else 0
        band = s - prev if i > 0 else s
        ty = yb - s + 0.04 if i > 0 else yb - s + s * 0.3
        th = max(0.24, band / 2 - 0.06) if i > 0 else s * 0.4
        _ktext(ctx, [_P(R[i].get("title"), 11, bold=True, color=_kread(cols[i]), align="center")], _B(x0 + (D - s) / 2 + s * 0.15, ty, s * 0.7, th),
               valign="middle", field="ring title")
    lx = x0 + D + 0.35
    lw = box["x"] + box["w"] - lx
    if lw > 1.2:
        ih = min(0.9, box["h"] / n)
        for j in range(n):
            i = n - 1 - j
            y = box["y"] + j * ih
            _kell(ctx, _B(lx, y + 0.08, 0.22, 0.22), cols[i])
            _ktext(ctx, [_P(R[i].get("title"), 11.5, bold=True, color=_kdk(ctx), spaceAfter=1), _P(R[i].get("text"), 10, color=c["text"])],
                   _B(lx + 0.34, y + 0.02, lw - 0.34, ih - 0.08), field="ring legend")


@_kcomp("staircase")
def _kc_staircase(ctx, b, box):
    t = ctx.t
    st = _list(b.get("steps"))
    n = len(st)
    if not n:
        return
    g = 0.1
    sw = (box["w"] - g * (n - 1)) / n
    min_h = box["h"] * 0.32
    for i, s in enumerate(st):
        f = i / (n - 1) if n > 1 else 1
        h = min_h + (box["h"] - min_h) * f
        x, y = box["x"] + i * (sw + g), box["y"] + box["h"] - h
        fill = _kmix(t["primary"], "FFFFFF", 0.55 * (1 - f))
        _krr(ctx, _B(x, y, sw, h), fill, name=f"Step {s.get('title', '')}")
        tc = _kread(fill, t["dark_bg"])
        _ktext(ctx, [_P(s.get("label") or f"Step {i + 1}", 9.5, bold=True, color=tc), _P(s.get("title"), 12.5, bold=True, color=tc, spaceAfter=2),
                     _P(s.get("text"), 10, color=tc)], _B(x + 0.12, y + 0.1, sw - 0.24, h - 0.2), field="stair step")


@_kcomp("flow")
def _kc_flow(ctx, b, box):
    t, c, dark = ctx.t, ctx.c, ctx.dark
    ins, outs = _list(b.get("inputs")), _list(b.get("outputs"))
    core = b.get("core") if isinstance(b.get("core"), dict) else {"title": ""}
    aw = 0.42
    col_w = (box["w"] - 2 * aw - 0.24) / 3

    def column(items, x, label):
        lh = 0.3 if label else 0
        if label:
            _ktext(ctx, [_P(label, 11, bold=True, color=c["heading"])], _B(x, box["y"], col_w, lh), valign="middle", quiet=True)
        n = len(items)
        if not n:
            return
        g = 0.1
        ih = (box["h"] - lh - g * (n - 1)) / n
        for i, it in enumerate(items):
            y = box["y"] + lh + i * (ih + g)
            _krr(ctx, _B(x, y, col_w, ih), c["surface"] if dark else "FFFFFF", line=None if dark else {"color": t["light_line"], "width": 0.75})
            tx = x + 0.12
            if it.get("icon") and ih >= 0.45:
                d = min(0.36, ih - 0.1)
                _k_icon_dot(ctx, x + 0.1, y + (ih - d) / 2, d, t["primary"], it["icon"])
                tx = x + 0.2 + d
            _ktext(ctx, [_P(it.get("title"), 11, bold=True, color=_kdk(ctx), spaceAfter=1), _P(it.get("text"), 9.5, color=c["text"])],
                   _B(tx, y + 0.04, x + col_w - 0.1 - tx, ih - 0.08), valign="middle", field="flow item")

    column(ins, box["x"], b.get("inputs_label") or "Inputs")
    cx = box["x"] + col_w + aw + 0.12
    ox = cx + col_w + aw + 0.12
    mid = box["y"] + 0.3 + (box["h"] - 0.3) / 2
    for ax in (box["x"] + col_w + 0.06, cx + col_w + 0.06):
        _kshp(ctx, "rightArrow", _B(ax, mid - 0.22, aw, 0.44), t["accent"])
    _krr(ctx, _B(cx, box["y"] + 0.3, col_w, box["h"] - 0.3), t["primary"], shadow=True, name="Flow core")
    cy = box["y"] + 0.5
    if core.get("icon"):
        _k_icon_dot(ctx, cx + col_w / 2 - 0.3, cy, 0.6, t["accent"], core["icon"], t["on_accent"])
        cy += 0.72
    paras = [_P(core.get("title"), 14, bold=True, color=t["on_primary"], align="center", spaceAfter=4),
             _P(core.get("text"), 10.5, color=t["on_primary"], align="center", spaceAfter=4)]
    paras += [_P(x, 10.5, color=t["on_primary"], bullet=True, spaceAfter=2) for x in _list(core.get("bullets"))]
    _ktext(ctx, paras, _B(cx + 0.18, cy, col_w - 0.36, box["y"] + box["h"] - 0.18 - cy), field="flow core", valign="top")
    _ktext(ctx, [_P(b.get("core_label") or "", 11, bold=True, color=c["heading"])], _B(cx, box["y"], col_w, 0.3), valign="middle", quiet=True)
    column(outs, ox, b.get("outputs_label") or "Outputs")


# ---------- data & assessment ----------
@_kcomp("heatmap")
def _kc_heatmap(ctx, b, box):
    t, c, dark = ctx.t, ctx.c, ctx.dark
    rows, cols, V = _list(b.get("rows")), _list(b.get("columns")), _list(b.get("values"))
    if not rows or not cols:
        return
    mx = _knum(b.get("max"), 4) or 4
    lw = max(1.6, min(3, box["w"] * 0.27))
    hh, leg_h = 0.44, 0.3
    cw = (box["w"] - lw) / len(cols)
    rh = min(0.62, (box["h"] - hh - leg_h) / len(rows))
    g = 0.03
    for j, cl in enumerate(cols):
        _ktext(ctx, [_P(cl, 10, bold=True, color=_kdk(ctx), align="center")], _B(box["x"] + lw + j * cw, box["y"], cw, hh - 0.04), valign="bottom", field="heat col")

    def color_of(v):
        if isinstance(v, str):
            s = v.strip().lower()
            if s in ("red", "high", "critical"):
                return t["danger"]
            if s in ("amber", "medium", "med"):
                return t["accent"]
            if s in ("green", "low"):
                return t["success"]
            return c["surface"] if dark else t["light_surface"]
        f = max(0, min(1, _knum(v) / mx))
        return _kmix(t["primary"], t["dark_bg"] if dark else "FFFFFF", 0.9 - f * 0.9)

    labels = _list(b.get("scale_labels"))
    for i, r in enumerate(rows):
        y = box["y"] + hh + i * rh
        _ktext(ctx, [_P(r, 10.5, bold=True, color=_kdk(ctx))], _B(box["x"], y, lw - 0.1, rh), valign="middle", field="heat row")
        rv = _list(V[i]) if i < len(V) else []
        for j in range(len(cols)):
            v = rv[j] if j < len(rv) else None
            fill = color_of(0 if v is None else v)
            _krect(ctx, _B(box["x"] + lw + j * cw + g, y + g, cw - 2 * g, rh - 2 * g), fill)
            if v is not None and b.get("show_values") is not False:
                if isinstance(v, (int, float)) and not isinstance(v, bool):
                    ri = int(round(v))
                    lbl = _ks(labels[ri]) if 0 <= ri < len(labels) and labels[ri] else _kfmt(v, 1 if v % 1 else 0)
                else:
                    lbl = _ks(v)
                _ktext(ctx, [_P(lbl, 10, bold=True, color=_kread(fill, t["dark_bg"]), align="center")], _B(box["x"] + lw + j * cw, y, cw, rh), valign="middle", quiet=True)
    ly = box["y"] + hh + len(rows) * rh + 0.06
    if not any(isinstance(v, str) for r in V for v in _list(r)):
        steps, sw = 5, 0.36
        _ktext(ctx, [_P(b.get("legend") or "Low", 9, color=c["muted"], align="right")], _B(box["x"] + box["w"] - 0.5 - steps * sw - 1.6, ly, 1.5, 0.22), valign="middle", quiet=True)
        for k in range(steps):
            _krect(ctx, _B(box["x"] + box["w"] - 0.5 - (steps - k) * sw, ly + 0.03, sw - 0.03, 0.16), color_of((k / (steps - 1)) * mx))
        _ktext(ctx, [_P("High", 9, color=c["muted"])], _B(box["x"] + box["w"] - 0.45, ly, 0.45, 0.22), valign="middle", quiet=True)


@_kcomp("maturity")
def _kc_maturity(ctx, b, box):
    t, c, dark = ctx.t, ctx.c, ctx.dark
    L, D = _list(b.get("levels")), _list(b.get("dimensions"))
    n = len(L)
    if not n or not D:
        return
    lw = max(1.6, min(2.8, box["w"] * 0.25))
    hh, leg_h = 0.5, 0.3
    cw = (box["w"] - lw) / n
    rh = min(0.62, (box["h"] - hh - leg_h) / len(D))
    for j, l in enumerate(L):
        fill = _kmix(t["primary"], "FFFFFF", 0.8 - 0.8 * (j / max(1, n - 1)))
        _krect(ctx, _B(box["x"] + lw + j * cw + 0.02, box["y"], cw - 0.04, hh - 0.06), fill)
        _ktext(ctx, [_P(f"{j + 1}  {_ks(l)}", 10, bold=True, color=_kread(fill, t["dark_bg"]), align="center")],
               _B(box["x"] + lw + j * cw + 0.05, box["y"], cw - 0.1, hh - 0.06), valign="middle", field="maturity level")
    for i, d in enumerate(D):
        y = box["y"] + hh + i * rh
        my = y + rh / 2
        if i % 2 == 0:
            _krect(ctx, _B(box["x"], y, box["w"], rh), c["surface"] if dark else t["light_surface"])
        _ktext(ctx, [_P(d.get("name"), 10.5, bold=True, color=_kdk(ctx)), _P(d.get("note") if rh > 0.5 else None, 8.5, color=c["muted"])],
               _B(box["x"] + 0.1, y, lw - 0.16, rh), valign="middle", field="dimension")
        X = lambda lv: box["x"] + lw + (max(1, min(n, lv)) - 0.5) * cw  # noqa: E731
        cd = min(0.26, rh * 0.5)
        _kline(ctx, box["x"] + lw + cw * 0.5, my, box["x"] + lw + cw * (n - 0.5), my, t["dark_line"] if dark else t["light_line"], 1)
        cur = _knum(d.get("current"), 1)
        tgt = d.get("target")
        if tgt is not None:
            tgt = _knum(tgt)
            if tgt != cur:
                _kline(ctx, X(cur), my, X(tgt) + (-cd / 2 if tgt > cur else cd / 2), my, t["accent"], 2.25, arrow=True)
            _kell(ctx, _B(X(tgt) - cd / 2, my - cd / 2, cd, cd), c["bg"] if dark else "FFFFFF", line={"color": t["accent"], "width": 2.25})
        _kell(ctx, _B(X(cur) - cd / 2, my - cd / 2, cd, cd), t["primary"], line={"color": "FFFFFF" if dark else t["primary"], "width": 1})
    ly = box["y"] + hh + len(D) * rh + 0.06
    _kell(ctx, _B(box["x"] + lw, ly + 0.04, 0.16, 0.16), t["primary"])
    _ktext(ctx, [_P("Current", 9, color=c["muted"])], _B(box["x"] + lw + 0.22, ly, 0.8, 0.24), valign="middle", quiet=True)
    _kell(ctx, _B(box["x"] + lw + 1.05, ly + 0.04, 0.16, 0.16), "FFFFFF", line={"color": t["accent"], "width": 2})
    _ktext(ctx, [_P("Target", 9, color=c["muted"])], _B(box["x"] + lw + 1.27, ly, 0.8, 0.24), valign="middle", quiet=True)


@_kcomp("risk_matrix")
def _kc_risk_matrix(ctx, b, box):
    t, c, dark = ctx.t, ctx.c, ctx.dark
    N = int(_knum(b.get("scale"), 5)) or 5
    risks = _list(b.get("risks"))
    list_w = 0 if (b.get("show_list") is False or box["w"] < 6) else min(4.2, box["w"] * 0.42)
    gw0 = box["w"] - list_w - (0.3 if list_w else 0)
    axis = 0.45
    side = min(gw0 - axis, box["h"] - axis)
    cs = side / N
    gx, gy = box["x"] + axis, box["y"]

    def sev(li, im):
        sc = (li * im) / (N * N)
        return t["danger"] if sc >= 0.48 else t["accent"] if sc >= 0.2 else t["success"]

    for li in range(1, N + 1):
        for im in range(1, N + 1):
            _krect(ctx, _B(gx + (im - 1) * cs + 0.02, gy + (N - li) * cs + 0.02, cs - 0.04, cs - 0.04), _kmix(sev(li, im), "FFFFFF", 0.2 if dark else 0.55))
    for k in range(1, N + 1):
        _ktext(ctx, [_P(str(k), 9, color=c["muted"], align="center")], _B(gx + (k - 1) * cs, gy + side + 0.02, cs, 0.2), quiet=True)
        _ktext(ctx, [_P(str(k), 9, color=c["muted"], align="right")], _B(gx - 0.24, gy + (N - k) * cs, 0.18, cs), valign="middle", quiet=True)
    _ktext(ctx, [_P((b.get("x_label") or "Impact") + "  →", 10, bold=True, color=c["muted"], align="center")], _B(gx, gy + side + 0.2, side, 0.24), quiet=True)
    _ktext(ctx, [_P((b.get("y_label") or "Likelihood") + "  →", 10, bold=True, color=c["muted"], align="center")],
           _B(box["x"] - side / 2 + 0.1, gy + side / 2 - 0.12, side, 0.24), rotate=270, quiet=True)
    cell_count = {}
    md = min(0.3, cs * 0.36)
    for i, r in enumerate(risks):
        li = max(1, min(N, int(round(_knum(r.get("likelihood"), 1)))))
        im = max(1, min(N, int(round(_knum(r.get("impact"), 1)))))
        key = f"{li}-{im}"
        cell_count[key] = cell_count.get(key, 0) + 1
        k = cell_count[key] - 1
        per = max(1, int((cs - 0.06) // (md + 0.04)))
        ox, oy = (k % per) * (md + 0.04), (k // per) * (md + 0.04)
        x, y = gx + (im - 1) * cs + 0.06 + ox, gy + (N - li) * cs + 0.06 + oy
        _kell(ctx, _B(x, y, md, md), "FFFFFF" if dark else t["dark_bg"])
        _ktext(ctx, [_P(_ks(r.get("id") or i + 1), max(7, md * 30), bold=True, color=t["dark_bg"] if dark else "FFFFFF", align="center")], _B(x, y, md, md),
               valign="middle", quiet=True)
    if list_w and risks:
        lx = min(box["x"] + box["w"] - list_w, gx + side + 0.45)
        n = len(risks)
        ih = min(0.42, box["h"] / n)
        for i, r in enumerate(risks):
            y = box["y"] + i * ih
            col = sev(_knum(r.get("likelihood"), 1), _knum(r.get("impact"), 1))
            _kell(ctx, _B(lx, y + (ih - 0.24) / 2, 0.24, 0.24), col)
            _ktext(ctx, [_P(_ks(r.get("id") or i + 1), 8.5, bold=True, color=_kread(col), align="center")], _B(lx, y + (ih - 0.24) / 2, 0.24, 0.24), valign="middle", quiet=True)
            runs = [{"text": _ks(r.get("label", "")), "bold": True}] + ([{"text": "  · " + _ks(r["owner"]), "bold": False, "color": c["muted"]}] if r.get("owner") else [])
            _ktext(ctx, [{"runs": runs, "size": 10, "color": _kdk(ctx)}], _B(lx + 0.32, y, list_w - 0.32, ih), valign="middle", field="risk label")


@_kcomp("bubble")
def _kc_bubble(ctx, b, box):
    t, c, dark = ctx.t, ctx.c, ctx.dark
    B = _list(b.get("bubbles"))
    if not B:
        return
    xa = b.get("x_axis") if isinstance(b.get("x_axis"), dict) else {}
    ya = b.get("y_axis") if isinstance(b.get("y_axis"), dict) else {}
    xs = [_knum(p.get("x")) for p in B]
    ys = [_knum(p.get("y")) for p in B]
    ss = [max(0, _knum(p.get("size"), 1)) for p in B]

    def pad(a, lo, hi):
        r = (hi - lo) or 1
        return (_knum(a["min"]) if a.get("min") is not None else lo - r * 0.12, _knum(a["max"]) if a.get("max") is not None else hi + r * 0.12)

    x0, x1 = pad(xa, min(xs), max(xs))
    y0, y1 = pad(ya, min(ys), max(ys))
    L, R, T, Bt = box["x"] + 0.6, box["x"] + box["w"] - 0.1, box["y"] + 0.05, box["y"] + box["h"] - 0.5
    X = lambda v: L + ((v - x0) / ((x1 - x0) or 1)) * (R - L)  # noqa: E731
    Y = lambda v: Bt - ((v - y0) / ((y1 - y0) or 1)) * (Bt - T)  # noqa: E731
    gc = t["dark_line"] if dark else "E3E7EE"
    for k in range(5):
        yv = y0 + (y1 - y0) * k / 4
        xv = x0 + (x1 - x0) * k / 4
        _kline(ctx, L, Y(yv), R, Y(yv), gc, 0.75, "dash" if k else None)
        _ktext(ctx, [_P((ya.get("prefix") or "") + _kfmt(yv, 1 if abs(y1 - y0) < 10 else 0) + (ya.get("unit") or ""), 8.5, color=c["muted"], align="right")],
               _B(box["x"], Y(yv) - 0.1, 0.52, 0.2), valign="middle", quiet=True)
        _ktext(ctx, [_P((xa.get("prefix") or "") + _kfmt(xv, 1 if abs(x1 - x0) < 10 else 0) + (xa.get("unit") or ""), 8.5, color=c["muted"], align="center")],
               _B(X(xv) - 0.4, Bt + 0.04, 0.8, 0.2), quiet=True)
    _ktext(ctx, [_P(xa.get("label"), 10, bold=True, color=c["muted"], align="center")], _B(L, Bt + 0.24, R - L, 0.24), quiet=True)
    _ktext(ctx, [_P(ya.get("label"), 10, bold=True, color=c["muted"], align="center")], _B(box["x"] - (Bt - T) / 2 + 0.05, T + (Bt - T) / 2 - 0.12, Bt - T, 0.24),
           rotate=270, quiet=True)
    smax = max(ss) or 1
    dmax = min(1.1, (R - L) * 0.16, (Bt - T) * 0.3)
    for i in sorted(range(len(B)), key=lambda k: -ss[k]):
        p = B[i]
        d = max(0.16, _kmath.sqrt(ss[i] / smax) * dmax)
        col = _k_role(t, p.get("color"), t["series"][i % len(t["series"])])
        px, py = X(xs[i]), Y(ys[i])
        _kell(ctx, _B(px - d / 2, py - d / 2, d, d), col, transparency=25, line={"color": col, "width": 1}, name=f"Bubble {p.get('label', '')}")
        lab = _ks(p.get("label", ""))
        inside = d >= 0.75 and _k_str_em(lab, True) * 9.5 / 72 < d * 0.9
        _ktext(ctx, [_P(lab, 9.5, bold=True, color=_kread(col) if inside else _kdk(ctx), align="center" if inside else "left")],
               _B(px - d / 2, py - 0.12, d, 0.24) if inside else _B(px + d / 2 + 0.04, py - 0.12, 1.8, 0.24), valign="middle", quiet=True)
    if b.get("size_label"):
        _ktext(ctx, [_P("Bubble size: " + _ks(b["size_label"]), 8.5, italic=True, color=c["muted"], align="right")], _B(R - 3, T, 3, 0.2), quiet=True)


@_kcomp("mekko")
def _kc_mekko(ctx, b, box):
    t, c = ctx.t, ctx.c
    cols = _list(b.get("columns"))
    if not cols:
        return
    names = _list(b.get("part_names"))
    tot = sum(max(0, _knum(x.get("width"))) for x in cols) or 1
    leg_h = 0.3 if names else 0
    T, Bt = box["y"] + 0.26, box["y"] + box["h"] - 0.5 - leg_h
    H, g = Bt - T, 0.04
    x = box["x"]
    for col in cols:
        cwid = _knum(col.get("width"))
        w = cwid / tot * box["w"]
        ps = _list(col.get("parts"))
        sm = sum(max(0, _knum(p.get("value"))) for p in ps) or 1
        y = Bt
        for k, p in enumerate(ps):
            pv = _knum(p.get("value"))
            h = pv / sm * H
            fill = t["series"][k % len(t["series"])]
            _krect(ctx, _B(x + g / 2, y - h + g / 2, w - g, max(0.01, h - g)), fill, name=f"{col.get('label', '')} {names[k] if k < len(names) else k}")
            pct = round(pv / sm * 100)
            if h > 0.26 and w > 0.42:
                _ktext(ctx, [_P((_ks(p["label"]) + " " if p.get("label") else "") + f"{pct}%", 9, bold=True, color=_kread(fill), align="center")],
                       _B(x + 0.03, y - h, w - 0.06, h), valign="middle", quiet=True)
            y -= h
        _ktext(ctx, [_P((b.get("prefix") or "") + _kfmt(cwid, 1 if cwid % 1 else 0) + (b.get("unit") or ""), 9.5, bold=True, color=_kdk(ctx), align="center")],
               _B(x, box["y"], w, 0.24), valign="middle", quiet=True)
        _ktext(ctx, [_P(col.get("label"), 9.5, color=c["text"], align="center", bold=True)], _B(x + 0.02, Bt + 0.05, w - 0.04, 0.44), field="mekko label")
        x += w
    if names:
        lx = box["x"]
        ly = box["y"] + box["h"] - leg_h + 0.04
        for k, nm in enumerate(names):
            _krect(ctx, _B(lx, ly + 0.05, 0.16, 0.16), t["series"][k % len(t["series"])])
            w = min(2.4, _k_str_em(_ks(nm)) * 9.5 / 72 + 0.15)
            _ktext(ctx, [_P(nm, 9.5, color=c["text"])], _B(lx + 0.22, ly, w, 0.26), valign="middle", quiet=True)
            lx += w + 0.42


@_kcomp("rings")
def _kc_rings(ctx, b, box):
    t, c = ctx.t, ctx.c
    it = _list(b.get("items"))
    n = len(it)
    if not n:
        return
    g = 0.25
    cw = (box["w"] - g * (n - 1)) / n
    for i, r in enumerate(it):
        x = box["x"] + i * (cw + g)
        mx = _knum(r.get("max"), 100) or 100
        val = _knum(r.get("value"))
        f = max(0, min(1, val / mx))
        text_h = min(0.85, box["h"] * 0.4)
        d = min(cw * 0.72, box["h"] - text_h - 0.08, 1.8)
        rx = x + (cw - d) / 2
        ry = box["y"] + max(0, (box["h"] - d - 0.08 - text_h) / 2)
        col = t["secondary"] if i % 2 else t["primary"]
        _kshp(ctx, "donut", _B(rx, ry, d, d), c["surface2"] if ctx.dark else t["light_line"])
        if f >= 0.999:
            _kshp(ctx, "donut", _B(rx, ry, d, d), col)
        elif f > 0:
            _kshp(ctx, "blockArc", _B(rx, ry, d, d), col, adj={"adj1": 270 * 60000, "adj2": ((270 + 360 * f) % 360) * 60000, "adj3": 25000})
        label = (r.get("prefix") or "") + _kfmt(val, 1 if val % 1 else 0) + (_ks(r["unit"]) if r.get("unit") is not None else "%")
        vs = _k_single_line_size(label, d * 0.42, d * 15)
        _ktext(ctx, [_P(label, vs, bold=True, color=_kdk(ctx), align="center")], _B(rx, ry, d, d), valign="middle", quiet=True)
        _ktext(ctx, [_P(r.get("label"), 11.5, bold=True, color=_kdk(ctx), align="center", spaceAfter=1), _P(r.get("text"), 9.5, color=c["muted"], align="center")],
               _B(x, ry + d + 0.08, cw, min(text_h, box["y"] + box["h"] - ry - d - 0.08)), field="ring label")


# ---------- people & organisation ----------
def _k_avatar(ctx, person, x, y, d, fill):
    img = _k_load_image(ctx, person.get("photo")) if person.get("photo") else None
    if img:
        _k_place_image(ctx, img, _B(x, y, d, d), "cover", True)
    else:
        _kell(ctx, _B(x, y, d, d), fill)
        _ktext(ctx, [_P(_k_initials(person.get("name")), max(8, d * 26), bold=True, color=_kread(fill), align="center")], _B(x, y, d, d), valign="middle", quiet=True)


@_kcomp("org_chart")
def _kc_org_chart(ctx, b, box):
    t, c, dark = ctx.t, ctx.c, ctx.dark
    kids = [k for k in _list(b.get("children")) if isinstance(k, dict)]
    k = len(kids)
    root = b.get("root") if isinstance(b.get("root"), dict) else {"name": ""}
    g = 0.18

    def node(p, x, y, w, h, fill, strong):
        _krr(ctx, _B(x, y, w, h), fill if strong else (c["surface"] if dark else "FFFFFF"),
             line=None if (strong or dark) else {"color": t["light_line"], "width": 0.75}, shadow=not (strong or dark))
        d = min(0.42, h - 0.14)
        _k_avatar(ctx, p, x + 0.1, y + (h - d) / 2, d, t["accent"] if strong else (t["primary"] if fill == t["primary_tint"] else fill))
        _ktext(ctx, [_P(p.get("name"), 11, bold=True, color=_kread(fill) if strong else _kdk(ctx)), _P(p.get("role"), 9, color=_kread(fill) if strong else c["muted"])],
               _B(x + 0.2 + d, y + 0.03, w - 0.28 - d, h - 0.06), valign="middle", field="org node")

    rw, rh = min(3.2, box["w"] * 0.3), 0.66
    rx = box["x"] + (box["w"] - rw) / 2
    node(root, rx, box["y"], rw, rh, t["primary"] if dark else t["dark_bg"], True)
    if not k:
        return
    cw = (box["w"] - g * (k - 1)) / k
    bus_y = box["y"] + rh + 0.2
    cy, ch = bus_y + 0.2, 0.62
    lc = t["dark_line"] if dark else _kmix(t["dark_bg"], "FFFFFF", 0.6)
    _kline(ctx, box["x"] + box["w"] / 2, box["y"] + rh, box["x"] + box["w"] / 2, bus_y, lc, 1.25)
    if k > 1:
        _kline(ctx, box["x"] + cw / 2, bus_y, box["x"] + box["w"] - cw / 2, bus_y, lc, 1.25)
    max_g = max([0] + [len(_list(x.get("children"))) for x in kids])
    g_top = cy + ch + 0.16
    gh = min(0.56, (box["y"] + box["h"] - g_top - 0.08 * (max_g - 1)) / max_g) if max_g else 0
    for i, kid in enumerate(kids):
        x = box["x"] + i * (cw + g)
        _kline(ctx, x + cw / 2, bus_y, x + cw / 2, cy, lc, 1.25)
        node(kid, x, cy, cw, ch, t["secondary"] if i % 2 else t["primary"], True)
        gc = [q for q in _list(kid.get("children")) if isinstance(q, dict)]
        if gc:
            _kline(ctx, x + 0.18, cy + ch, x + 0.18, g_top + (len(gc) - 1) * (gh + 0.08) + gh / 2, lc, 1)
        for j, gq in enumerate(gc):
            y = g_top + j * (gh + 0.08)
            _kline(ctx, x + 0.18, y + gh / 2, x + 0.32, y + gh / 2, lc, 1)
            node(gq, x + 0.32, y, cw - 0.32, gh, t["primary_tint"], False)


@_kcomp("profiles")
def _kc_profiles(ctx, b, box):
    t, c, dark = ctx.t, ctx.c, ctx.dark
    P = [p for p in _list(b.get("people")) if isinstance(p, dict)]
    n = len(P)
    if not n:
        return
    cols = int(b.get("columns") or (n if n <= 4 else (n + 1) // 2))
    rows = (n + cols - 1) // cols
    g = 0.2
    cw = (box["w"] - g * (cols - 1)) / cols
    ch = (box["h"] - g * (rows - 1)) / rows
    for i, p in enumerate(P):
        x, y = box["x"] + (i % cols) * (cw + g), box["y"] + (i // cols) * (ch + g)
        _krr(ctx, _B(x, y, cw, ch), c["surface"] if dark else "FFFFFF", line=None if dark else {"color": t["light_line"], "width": 0.75}, shadow=not dark)
        horizontal = ch < 1.9
        d = min(0.8, ch - 0.3) if horizontal else min(1.1, cw * 0.4, ch * 0.38)
        fill = t["secondary"] if i % 2 else t["primary"]
        if horizontal:
            _k_avatar(ctx, p, x + 0.15, y + 0.15, d, fill)
            _ktext(ctx, [_P(p.get("name"), 12.5, bold=True, color=_kdk(ctx)), _P(p.get("role"), 10, bold=True, color=t["secondary"], spaceAfter=3), _P(p.get("text"), 9.5, color=c["text"])],
                   _B(x + 0.3 + d, y + 0.12, cw - 0.45 - d, ch - 0.24), field="profile")
        else:
            _k_avatar(ctx, p, x + (cw - d) / 2, y + 0.18, d, fill)
            _ktext(ctx, [_P(p.get("name"), 12.5, bold=True, color=_kdk(ctx), align="center"), _P(p.get("role"), 10, bold=True, color=t["secondary"], align="center", spaceAfter=4),
                         _P(p.get("text"), 9.5, color=c["text"], align="center")], _B(x + 0.12, y + d + 0.3, cw - 0.24, ch - d - 0.42), field="profile")


@_kcomp("raci")
def _kc_raci(ctx, b, box):
    t, c, dark = ctx.t, ctx.c, ctx.dark
    A, R, M = _list(b.get("activities")), _list(b.get("roles")), _list(b.get("matrix"))
    if not A or not R:
        return
    lw = max(1.8, min(3.4, box["w"] * 0.32))
    hh, leg_h = 0.48, 0.28
    cw = (box["w"] - lw) / len(R)
    rh = min(0.5, (box["h"] - hh - leg_h) / len(A))
    _krr(ctx, _B(box["x"], box["y"], box["w"], hh), t["primary"] if dark else t["dark_bg"])
    _ktext(ctx, [_P(b.get("activity_label") or "Activity", 10.5, bold=True, color="FFFFFF")], _B(box["x"] + 0.12, box["y"], lw - 0.2, hh), valign="middle", quiet=True)
    for j, r in enumerate(R):
        _ktext(ctx, [_P(r, 10, bold=True, color="FFFFFF", align="center")], _B(box["x"] + lw + j * cw + 0.03, box["y"], cw - 0.06, hh), valign="middle", field="raci role")
    col_of = {"R": t["primary"], "A": t["dark_bg"], "C": t["secondary"], "I": t["neutral"], "S": t["info"]}
    for i, a in enumerate(A):
        y = box["y"] + hh + i * rh
        _krect(ctx, _B(box["x"], y, box["w"], rh), (c["bg"] if dark else "FFFFFF") if i % 2 else (c["surface"] if dark else t["light_surface"]))
        _ktext(ctx, [_P(a, 10.5, bold=True, color=_kdk(ctx))], _B(box["x"] + 0.12, y, lw - 0.2, rh), valign="middle", field="raci activity")
        row = _list(M[i]) if i < len(M) else []
        for j in range(len(R)):
            v = re.sub(r"[^RACIS/]", "", _ks(row[j] if j < len(row) else "").upper())
            if not v:
                continue
            d = min(0.34, rh * 0.72)
            w = d * 1.6 if len(v) > 1 else d
            cx = box["x"] + lw + j * cw + cw / 2
            fill = col_of.get(v[0], t["primary"])
            fd = t["accent"] if (dark and v[0] == "A") else fill
            _kshp(ctx, "roundRect", _B(cx - w / 2, y + (rh - d) / 2, w, d), fd, radius=d / 2)
            _ktext(ctx, [_P(v, max(8, d * 27), bold=True, color=_kread(fd), align="center")], _B(cx - w / 2, y + (rh - d) / 2, w, d), valign="middle", quiet=True)
    _ktext(ctx, [_P("R = Responsible  ·  A = Accountable  ·  C = Consulted  ·  I = Informed", 9, color=c["muted"])],
           _B(box["x"], box["y"] + hh + len(A) * rh + 0.04, box["w"], leg_h - 0.04), valign="middle", quiet=True)


@_kcomp("swimlane")
def _kc_swimlane(ctx, b, box):
    t, c, dark = ctx.t, ctx.c, ctx.dark
    L = [_ks(x) for x in _list(b.get("lanes"))]
    S = [s for s in _list(b.get("steps")) if isinstance(s, dict)]
    nl = len(L)
    if not nl or not S:
        return
    phases = _list(b.get("phases"))
    ncol = max([len(phases)] + [int(_knum(s.get("col"), 0)) + 1 for s in S])
    lw = min(1.5, box["w"] * 0.13)
    hh = 0.34 if phases else 0
    lh = (box["h"] - hh) / nl
    cw = (box["w"] - lw) / ncol
    for j, p in enumerate(phases):
        _ktext(ctx, [_P(p, 10.5, bold=True, color=c["heading"], align="center")], _B(box["x"] + lw + j * cw, box["y"], cw, hh - 0.04), valign="middle", field="phase")
    for i, ln in enumerate(L):
        y = box["y"] + hh + i * lh
        _krect(ctx, _B(box["x"] + lw, y, box["w"] - lw, lh), (c["bg"] if dark else "FFFFFF") if i % 2 else (c["surface"] if dark else t["light_surface"]))
        fill = _kmix(t["primary"], t["secondary"], i / (nl - 1) if nl > 1 else 0)
        _krect(ctx, _B(box["x"], y + 0.02, lw - 0.04, lh - 0.04), fill)
        _ktext(ctx, [_P(ln, 10.5, bold=True, color=_kread(fill), align="center")], _B(box["x"] + 0.06, y, lw - 0.16, lh), valign="middle", field="lane")
    for j in range(1, ncol):
        _kline(ctx, box["x"] + lw + j * cw, box["y"] + hh, box["x"] + lw + j * cw, box["y"] + box["h"], t["dark_line"] if dark else "E1E5EC", 0.75, "dash")

    def lane_idx(v):
        if isinstance(v, (int, float)) and not isinstance(v, bool):
            return max(0, min(nl - 1, int(v)))
        return L.index(_ks(v)) if _ks(v) in L else 0

    bw = cw * 0.78
    bh = min(lh * 0.66, 0.85)
    pos = []
    for s in S:
        li = lane_idx(s.get("lane"))
        cx = box["x"] + lw + int(_knum(s.get("col"), 0)) * cw + cw / 2
        cy = box["y"] + hh + li * lh + lh / 2
        pos.append({"li": li, "cx": cx, "cy": cy, "l": cx - bw / 2, "r": cx + bw / 2, "t": cy - bh / 2, "b": cy + bh / 2})
    lc = t["dark_muted"] if dark else t["muted"]
    links = b.get("links") if isinstance(b.get("links"), list) else [[i, i + 1] for i in range(len(S) - 1)]
    for lk in links:
        try:
            a, z = int(lk[0]), int(lk[1])
            p, q = pos[a], pos[z]
        except Exception:
            ctx.warn(f"{ctx.where}: link {lk} refers to a missing step")
            continue
        if p["li"] == q["li"]:
            _kline(ctx, p["r"] if q["cx"] > p["cx"] else p["l"], p["cy"], q["l"] if q["cx"] > p["cx"] else q["r"], q["cy"], lc, 1.25, arrow=True)
        elif abs(q["cx"] - p["cx"]) < 0.01:
            _kline(ctx, p["cx"], p["b"] if q["cy"] > p["cy"] else p["t"], q["cx"], q["t"] if q["cy"] > p["cy"] else q["b"], lc, 1.25, arrow=True)
        else:
            _kline(ctx, p["r"], p["cy"], q["cx"], p["cy"], lc, 1.25)
            _kline(ctx, q["cx"], p["cy"], q["cx"], q["t"] if q["cy"] > p["cy"] else q["b"], lc, 1.25, arrow=True)
    for i, s in enumerate(S):
        p = pos[i]
        dec = s.get("shape") == "decision"
        hl = bool(s.get("highlight"))
        if dec:
            _kshp(ctx, "diamond", _B(p["cx"] - bh * 0.75, p["t"], bh * 1.5, bh), t["accent"])
        else:
            _krr(ctx, _B(p["l"], p["t"], bw, bh), t["primary"] if hl else (c["surface2"] if dark else "FFFFFF"),
                 line=None if (hl or dark) else {"color": _kmix(t["primary"], "FFFFFF", 0.5), "width": 1})
        tc = t["on_accent"] if dec else (t["on_primary"] if hl else _kdk(ctx))
        _ktext(ctx, [_P(s.get("title"), 10, bold=True, color=tc, align="center"), _P(s.get("text") if not dec else None, 8.5, color=t["on_primary"] if hl else c["muted"], align="center")],
               _B(p["cx"] - bh * 0.55, p["t"] + 0.05, bh * 1.1, bh - 0.1) if dec else _B(p["l"] + 0.05, p["t"] + 0.03, bw - 0.1, bh - 0.06), valign="middle", field="swim step")


# ---------- PMO ----------
_K_STATUS = {
    "done": ("success", "FaCheck"), "complete": ("success", "FaCheck"), "green": ("success", "FaCheck"), "on_track": ("success", "FaCheck"),
    "in_progress": ("accent", "FaSpinner"), "amber": ("accent", "FaExclamation"), "at_risk": ("accent", "FaExclamation"),
    "todo": ("neutral", "FaRegCircle"), "not_started": ("neutral", "FaRegCircle"), "blocked": ("danger", "FaTimes"),
    "red": ("danger", "FaTimes"), "off_track": ("danger", "FaTimes"),
}


def _k_st_key(s):
    return re.sub(r"[\s-]+", "_", _ks(s or "todo").lower())


@_kcomp("checklist")
def _kc_checklist(ctx, b, box):
    t, c, dark = ctx.t, ctx.c, ctx.dark
    it = [x for x in _list(b.get("items")) if isinstance(x, dict)]
    n = len(it)
    if not n:
        return
    rh = min(0.56, box["h"] / n)
    has_meta = any(x.get("owner") or x.get("due") for x in it)
    mw = min(2.6, box["w"] * 0.3) if has_meta else 0
    for i, x in enumerate(it):
        y = box["y"] + i * rh
        role, ic = _K_STATUS.get(_k_st_key(x.get("status")), _K_STATUS["todo"])
        col = t[role]
        d = min(0.32, rh * 0.62)
        if i % 2 == 0:
            _krect(ctx, _B(box["x"], y, box["w"], rh), c["surface"] if dark else t["light_surface"])
        _k_icon_dot(ctx, box["x"] + 0.1, y + (rh - d) / 2, d, col, ic)
        _ktext(ctx, [{"runs": [{"text": _ks(x.get("text", "")), "bold": bool(x.get("bold"))}], "size": 11, "color": _kdk(ctx)}],
               _B(box["x"] + 0.22 + d, y, box["w"] - 0.32 - d - mw, rh), valign="middle", field="checklist item")
        if has_meta:
            runs = [{"text": _ks(x.get("owner") or ""), "bold": True}] + ([{"text": ("  ·  " if x.get("owner") else "") + _ks(x["due"]), "bold": False, "color": c["muted"]}] if x.get("due") else [])
            _ktext(ctx, [{"runs": runs, "size": 10, "color": _kdk(ctx), "align": "right"}], _B(box["x"] + box["w"] - mw - 0.1, y, mw, rh), valign="middle", field="checklist meta")


@_kcomp("status")
def _kc_status(ctx, b, box):
    t, c, dark = ctx.t, ctx.c, ctx.dark
    it = [x for x in _list(b.get("items")) if isinstance(x, dict)]
    n = len(it)
    if not n:
        return
    hh = 0.4
    W = {"name": 0.24, "status": 0.12, "trend": 0.08, "owner": 0.14 if any(x.get("owner") for x in it) else 0}
    W["comment"] = 1 - W["name"] - W["status"] - W["trend"] - W["owner"]
    xs, acc = {}, box["x"]
    for k in ("name", "status", "trend", "owner", "comment"):
        xs[k] = acc
        acc += W[k] * box["w"]
    _krr(ctx, _B(box["x"], box["y"], box["w"], hh), t["primary"] if dark else t["dark_bg"])
    H = {"name": b.get("name_label") or "Workstream", "status": "Status", "trend": "Trend", "owner": "Owner", "comment": "Commentary"}
    for k, lab in H.items():
        if W[k]:
            _ktext(ctx, [_P(lab, 10, bold=True, color="FFFFFF", align="center" if k in ("status", "trend") else "left")],
                   _B(xs[k] + 0.1, box["y"], W[k] * box["w"] - 0.16, hh), valign="middle", quiet=True)
    rh = min(0.62, (box["h"] - hh) / n)
    for i, x in enumerate(it):
        y = box["y"] + hh + i * rh
        role = _K_STATUS.get(_k_st_key(x.get("status")), _K_STATUS["amber"])[0]
        col = t[role]
        _krect(ctx, _B(box["x"], y, box["w"], rh), (c["bg"] if dark else "FFFFFF") if i % 2 else (c["surface"] if dark else t["light_surface"]))
        _ktext(ctx, [_P(x.get("name"), 10.5, bold=True, color=_kdk(ctx))], _B(xs["name"] + 0.1, y, W["name"] * box["w"] - 0.16, rh), valign="middle", field="status name")
        pw = min(W["status"] * box["w"] - 0.2, 1.1)
        ph = min(0.28, rh * 0.6)
        px = xs["status"] + (W["status"] * box["w"] - pw) / 2
        _kshp(ctx, "roundRect", _B(px, y + (rh - ph) / 2, pw, ph), col, radius=ph / 2)
        lab = x.get("label") or {"success": "On track", "accent": "At risk", "danger": "Off track", "neutral": "Not started"}[role]
        _ktext(ctx, [_P(lab, 9, bold=True, color=_kread(col), align="center")], _B(px, y + (rh - ph) / 2, pw, ph), valign="middle", quiet=True)
        tr = _ks(x.get("trend") or "flat").lower()
        tic = "FaArrowUp" if tr == "up" else "FaArrowDown" if tr == "down" else "FaArrowRight"
        tcol = t["success"] if tr == "up" else t["danger"] if tr == "down" else t["neutral"]
        td = min(0.24, rh * 0.5)
        _k_icon(ctx, tic, tcol, xs["trend"] + (W["trend"] * box["w"] - td) / 2, y + (rh - td) / 2, td, td)
        if W["owner"]:
            _ktext(ctx, [_P(x.get("owner") or "", 10, color=c["text"])], _B(xs["owner"] + 0.1, y, W["owner"] * box["w"] - 0.16, rh), valign="middle", field="status owner")
        _ktext(ctx, [_P(x.get("comment") or "", 10, color=c["text"])], _B(xs["comment"] + 0.1, y, W["comment"] * box["w"] - 0.16, rh), valign="middle", field="status comment")


@_kcomp("board")
def _kc_board(ctx, b, box):
    t, c, dark = ctx.t, ctx.c, ctx.dark
    cols = [x for x in _list(b.get("columns")) if isinstance(x, dict)]
    k = len(cols)
    if not k:
        return
    g, hh = 0.18, 0.42
    cw = (box["w"] - g * (k - 1)) / k
    max_cards = max([1] + [len(_list(x.get("cards"))) for x in cols])
    cg = 0.08
    chh = min(0.95, (box["h"] - hh - 0.12 - cg * (max_cards - 1) - 0.1) / max_cards)
    for i, col in enumerate(cols):
        x = box["x"] + i * (cw + g)
        fill = _k_role(t, col.get("color"), [t["neutral"], t["info"], t["primary"], t["success"], t["accent"]][i % 5])
        cards = [cd for cd in _list(col.get("cards")) if isinstance(cd, dict)]
        _krr(ctx, _B(x, box["y"], cw, box["h"]), c["surface"] if dark else t["light_surface"])
        _krr(ctx, _B(x, box["y"], cw, hh), fill)
        _ktext(ctx, [{"runs": [{"text": _ks(col.get("title", "")), "bold": True}, {"text": f"  {len(cards)}", "bold": False}], "size": 11, "color": _kread(fill)}],
               _B(x + 0.12, box["y"], cw - 0.24, hh), valign="middle", quiet=True)
        for j, cd in enumerate(cards):
            y = box["y"] + hh + 0.12 + j * (chh + cg)
            _krr(ctx, _B(x + 0.1, y, cw - 0.2, chh), c["surface2"] if dark else "FFFFFF", line=None if dark else {"color": t["light_line"], "width": 0.75}, shadow=not dark)
            _ktext(ctx, [_P(cd.get("tag"), 8.5, bold=True, color=t["secondary"]), _P(cd.get("title"), 10.5, bold=True, color=_kdk(ctx), spaceAfter=1), _P(cd.get("text"), 9, color=c["muted"])],
                   _B(x + 0.2, y + 0.05, cw - 0.4, chh - 0.1), valign="middle", field="board card")


# ---------- customer ----------
@_kcomp("persona")
def _kc_persona(ctx, b, box):
    t, c, dark = ctx.t, ctx.c, ctx.dark
    lw = min(3.6, box["w"] * 0.32)
    _krr(ctx, _B(box["x"], box["y"], lw, box["h"]), c["surface"] if dark else t["dark_bg"], name="Persona")
    d = min(1.05, box["h"] * 0.26)
    _k_avatar(ctx, b, box["x"] + (lw - d) / 2, box["y"] + 0.2, d, t["accent"])
    y = box["y"] + 0.3 + d
    _ktext(ctx, [_P(b.get("name"), 15, bold=True, color="FFFFFF", align="center"), _P(b.get("role"), 10.5, color=t["accent"], bold=True, align="center")],
           _B(box["x"] + 0.15, y, lw - 0.3, 0.62), valign="middle", field="persona name")
    y += 0.7
    attrs = [a for a in _list(b.get("attributes")) if isinstance(a, dict)]
    if b.get("quote"):
        qh = min(1.1, (box["y"] + box["h"] - y) * (0.45 if attrs else 0.9))
        _ktext(ctx, [_P("“" + _ks(b["quote"]) + "”", 11.5, italic=True, color="E8ECF4", align="center")], _B(box["x"] + 0.2, y, lw - 0.4, qh), valign="middle", field="persona quote")
        y += qh + 0.08
    if attrs:
        ah = min(0.42, (box["y"] + box["h"] - 0.12 - y) / len(attrs))
        for i, a in enumerate(attrs):
            _ktext(ctx, [{"runs": [{"text": _ks(a.get("label", "")) + ": ", "bold": True, "color": t["dark_muted"]}, {"text": _ks(a.get("value", "")), "color": "FFFFFF"}], "size": 10}],
                   _B(box["x"] + 0.25, y + i * ah, lw - 0.5, ah), valign="middle", field="persona attr")
    sects = [s for s in (("Goals", b.get("goals"), t["success"], "target"), ("Pain points", b.get("pains"), t["danger"], "warning"),
                         ("Needs", b.get("needs"), t["primary"], "idea"), ("Behaviours", b.get("behaviours"), t["secondary"], "users")) if _list(s[1])]
    k = len(sects)
    if not k:
        return
    gx, gw = box["x"] + lw + 0.22, box["w"] - lw - 0.22
    cols = k if k <= 3 else 2
    rows = (k + cols - 1) // cols
    g = 0.18
    sw = (gw - g * (cols - 1)) / cols
    sh = (box["h"] - g * (rows - 1)) / rows
    for i, (title, items, col, ic) in enumerate(sects):
        x, y2 = gx + (i % cols) * (sw + g), box["y"] + (i // cols) * (sh + g)
        _krr(ctx, _B(x, y2, sw, sh), c["surface"] if dark else _kmix(col, "FFFFFF", 0.9))
        _k_icon_dot(ctx, x + 0.14, y2 + 0.14, 0.38, col, ic)
        _ktext(ctx, [_P(title, 13, bold=True, color=_kdk(ctx))], _B(x + 0.62, y2 + 0.12, sw - 0.75, 0.42), valign="middle", quiet=True)
        _ktext(ctx, [_P(s, 11, color=c["text"], bullet=True, spaceAfter=4) for s in _list(items)], _B(x + 0.18, y2 + 0.66, sw - 0.32, sh - 0.78), field="persona list")


@_kcomp("journey")
def _kc_journey(ctx, b, box):
    t, c, dark = ctx.t, ctx.c, ctx.dark
    S = [s for s in _list(b.get("stages")) if isinstance(s, dict)]
    n = len(S)
    if not n:
        return
    ROWS = [(k, lab) for k, lab in (("actions", "Actions"), ("touchpoints", "Touchpoints"), ("pains", "Pain points"), ("opportunities", "Opportunities"))
            if any(s.get(k) for s in S)]
    has_emo = any(s.get("emotion") is not None for s in S)
    lw = min(1.45, box["w"] * 0.12)
    hh = 0.46
    emo_h = min(1.0, box["h"] * 0.22) if has_emo else 0
    cw = (box["w"] - lw) / n
    rh = (box["h"] - hh - emo_h - (0.1 if has_emo else 0)) / max(1, len(ROWS))
    for j, s in enumerate(S):
        x = box["x"] + lw + j * cw
        _kshp(ctx, "homePlate" if j == 0 else "chevron", _B(x, box["y"], cw + (0.12 if j < n - 1 else 0), hh - 0.06), _kmix(t["primary"], t["secondary"], j / (n - 1) if n > 1 else 0))
        _ktext(ctx, [_P(s.get("title"), 11, bold=True, color="FFFFFF", align="center")], _B(x + (0.24 if j else 0.08), box["y"], cw - 0.36, hh - 0.06), valign="middle", field="journey stage")
    y = box["y"] + hh
    if has_emo:
        _ktext(ctx, [_P(b.get("emotion_label") or "Experience", 10, bold=True, color=_kdk(ctx))], _B(box["x"], y, lw - 0.1, emo_h), valign="middle", quiet=True)
        _krect(ctx, _B(box["x"] + lw, y, box["w"] - lw, emo_h), c["surface"] if dark else t["light_surface"])
        mid = y + emo_h / 2
        _kline(ctx, box["x"] + lw, mid, box["x"] + box["w"], mid, t["dark_line"] if dark else t["light_line"], 0.75, "dash")
        P = []
        for j, s in enumerate(S):
            e = _knum(s.get("emotion"), 0)
            P.append((box["x"] + lw + j * cw + cw / 2, mid - (max(-2, min(2, e)) / 2) * (emo_h / 2 - 0.12), e))
        for j in range(1, len(P)):
            _kline(ctx, P[j - 1][0], P[j - 1][1], P[j][0], P[j][1], t["primary"], 2.25)
        for px, py, e in P:
            _kell(ctx, _B(px - 0.1, py - 0.1, 0.2, 0.2), t["success"] if e > 0 else t["danger"] if e < 0 else t["accent"], line={"color": "FFFFFF", "width": 1.25})
        y += emo_h + 0.1
    for r, (k, label) in enumerate(ROWS):
        ry = y + r * rh
        col = t["danger"] if k == "pains" else t["success"] if k == "opportunities" else t["primary"]
        _ktext(ctx, [_P(label, 10, bold=True, color=col if k in ("pains", "opportunities") else _kdk(ctx))], _B(box["x"], ry, lw - 0.1, rh), valign="middle", quiet=True)
        for j, s in enumerate(S):
            x = box["x"] + lw + j * cw
            v = s.get(k)
            items = v if isinstance(v, list) else ([v] if v else [])
            fill = c["surface"] if dark else (t["danger_tint"] if k == "pains" else t["success_tint"] if k == "opportunities" else "FFFFFF")
            _krr(ctx, _B(x + 0.04, ry + 0.04, cw - 0.08, rh - 0.08), fill, line=None if (dark or k not in ("actions", "touchpoints")) else {"color": t["light_line"], "width": 0.75})
            _ktext(ctx, [_P(x2, 9.5, color=c["text"], bullet=len(items) > 1, spaceAfter=2) for x2 in items], _B(x + 0.12, ry + 0.08, cw - 0.24, rh - 0.16),
                   field="journey cell", valign="middle")


# ---------- media ----------
@_kcomp("image")
def _kc_image(ctx, b, box):
    t, c, dark = ctx.t, ctx.c, ctx.dark
    cap_h = 0.34 if b.get("caption") else 0
    ib = _B(box["x"], box["y"], box["w"], box["h"] - cap_h)
    img = None if ctx.dry else _k_load_image(ctx, b.get("src"))
    if img:
        _k_place_image(ctx, img, ib, b.get("fit") or "cover")
    else:
        _krr(ctx, ib, c["surface"] if dark else t["light_surface"], line={"color": t["light_line"], "width": 1, "dash": "dash"})
        _ktext(ctx, [_P("Image: " + _ks(b.get("alt") or b.get("src") or "placeholder"), 11, color=c["muted"], align="center")],
               _B(ib["x"] + 0.2, ib["y"], ib["w"] - 0.4, ib["h"]), valign="middle", quiet=True)
    if b.get("caption"):
        _ktext(ctx, [_P(b["caption"], 9.5, italic=True, color=c["muted"])], _B(box["x"], box["y"] + box["h"] - cap_h + 0.04, box["w"], cap_h - 0.04),
               valign="middle", field="caption")


_K_MIN_H = {"layers": 2.2, "gantt": 1.6, "scorecard": 1.4, "table": 0.8, "hub": 2.4, "tree": 2.0, "matrix": 2.2, "chart": 1.6, "waterfall": 1.8,
            "pyramid": 1.4, "funnel": 1.4, "comparison": 1.6, "process": 1.2, "timeline": 0.9, "cards": 0.6,
            "swot": 2.2, "venn": 2.0, "cycle": 2.4, "value_chain": 1.6, "concentric": 2.0, "staircase": 1.5, "flow": 1.8, "heatmap": 1.2,
            "maturity": 1.3, "risk_matrix": 2.2, "bubble": 2.0, "mekko": 2.0, "rings": 1.4, "org_chart": 2.0, "profiles": 1.0, "raci": 1.0,
            "swimlane": 1.8, "checklist": 0.6, "status": 0.9, "board": 1.8, "persona": 2.6, "journey": 2.4, "image": 0.8}
_K_DEFAULT_ROW_H = {"kpis": 0.8, "chevrons": 0.45, "timeline": 1.0, "process": 1.25, "callout": 0.85, "progress": 1.1, "bullets": 1.4, "agenda": 1.6,
                    "rings": 1.1, "checklist": 1.2, "status": 1.3}


# ------------------------------------------------------------ schema ----
def _kS(desc, **x):
    return {"type": "string", "desc": desc, **x}


def _kN(desc, **x):
    return {"type": "number", "desc": desc, **x}


def _kBo(desc):
    return {"type": "boolean", "desc": desc}


def _kE(values, desc, **x):
    return {"type": "enum", "values": values, "desc": desc, **x}


def _kA(of, desc, **x):
    return {"type": "array", "of": of, "desc": desc, **x}


def _kO(fields, desc, **x):
    return {"type": "object", "fields": fields, "desc": desc, **x}


def _kR(spec):
    return {**spec, "required": True}


_K_ICON = _kS("icon name: alias (ai, cloud, money, shield, users, chart, rocket, ...) or a Font Awesome name (FaRobot)")
_K_COLOR = _kE(["primary", "secondary", "accent", "info", "success", "danger", "neutral", "dark"], "theme color role")
_K_NUM = {"type": "number"}
_K_STR = {"type": "string"}
_K_ANY = {"type": "any"}

_K_COMPONENTS = {
    "kpis": {"use": "Headline numbers. 2-6 stat tiles. Values must be real, sourced figures.", "fields": {
        "items": _kR(_kA(_kO({"value": _kR(_kS("the number, short: '€22.9B', '+27%', '200+'", maxLen=9)), "label": _kR(_kS("what it measures", maxLen=90)),
                              "icon": _K_ICON, "delta": _kS("optional change note, e.g. '+4 pts YoY'", maxLen=22),
                              "trend": _kA(_K_NUM, "optional sparkline values, oldest first (4-12 points)", max=24)}, ""), "tiles", min=2, max=6)),
        "highlight": _kN("index of the tile to emphasize (0-based)")}},
    "timeline": {"use": "Chronology, history, milestones, release waves. 2-7 points.", "fields": {
        "items": _kR(_kA(_kO({"date": _kR(_kS("'2023', 'Q1 2026', 'May 2026'", maxLen=14)), "title": _kR(_kS("milestone name", maxLen=40)),
                              "text": _kS("1-2 sentence detail", maxLen=160), "highlight": _kBo("emphasize (e.g., the latest/current point)")}, ""), "milestones", min=2, max=7))}},
    "cards": {"use": "Parallel ideas: implications, benefits, capabilities, recommendations, case studies. 1-12 cards in a grid.", "fields": {
        "items": _kR(_kA(_kO({"title": _kR(_kS("card heading", maxLen=45)), "text": _kS("body, 1-3 sentences", maxLen=260), "icon": _K_ICON,
                              "stat": _kS("optional big figure shown instead of the icon, e.g. '-90%'", maxLen=8), "tag": _kS("small label under the title, e.g. 'Finance · AR'", maxLen=40),
                              "bullets": _kA(_kS("bullet"), "optional bullets instead of text", max=5)}, ""), "cards", min=1, max=12)),
        "columns": _kN("grid columns (default: auto)", min=1, max=6), "style": _kE(["outline", "tinted", "solid"], "visual style (default outline)"),
        "highlight": _kA(_K_NUM, "indexes of cards to emphasize")}},
    "process": {"use": "Ordered steps / how it works / method. 3-7 numbered steps with icons and connector.", "fields": {
        "steps": _kR(_kA(_kO({"title": _kR(_kS("step name, 1-3 words", maxLen=24)), "text": _kS("what happens", maxLen=150), "icon": _K_ICON}, ""), "steps", min=2, max=7)),
        "highlight": _kN("index of step to emphasize")}},
    "chevrons": {"use": "Stages/maturity levels/progression in one strip (low height). 2-6 stages.", "fields": {
        "stages": _kR(_kA(_kO({"title": _kR(_kS("stage", maxLen=22)), "text": _kS("short descriptor", maxLen=40)}, ""), "stages", min=2, max=6))}},
    "layers": {"use": "Layered architecture / operating model / stack diagram. 2-6 tiers, optional cross-cutting side panel.", "fields": {
        "tiers": _kR(_kA(_kO({"label": _kR(_kS("tier name", maxLen=26)), "sublabel": _kS("tier description", maxLen=60), "color": _K_COLOR,
                              "items": _kR(_kA(_kO({"title": _kR(_kS("component", maxLen=32)), "text": _kS("detail", maxLen=90),
                                                    "bullets": _kA(_kS("bullet", maxLen=32), "bullets", max=5), "highlight": _kBo("emphasize")}, ""), "boxes in the tier", min=1, max=6)),
                              "base": _kS("optional full-width foundation bar text at bottom of the tier", maxLen=110)}, ""), "top (user-facing) to bottom (foundation)", min=2, max=6)),
        "side": _kO({"title": _kR(_kS("e.g. 'Governance'", maxLen=24)), "subtitle": _kS("e.g. 'Spans every tier'", maxLen=30), "icon": _K_ICON,
                     "items": _kR(_kA(_kO({"title": _kR(_kS("item", maxLen=30)), "text": _kS("detail", maxLen=90)}, ""), "items", min=1, max=7))}, "optional cross-cutting column on the right")}},
    "chart": {"use": "Native, editable PowerPoint chart: bar, column, stacked, line, area, pie, doughnut, radar (multi-criteria profile), combo (volume columns + rate line on 2nd axis). Every number must come from a source.", "fields": {
        "chart_type": _kR(_kE(["bar", "column", "stacked_bar", "stacked_column", "line", "area", "pie", "doughnut", "radar", "combo"], "chart type; combo = columns + line(s) on a second axis")),
        "categories": _kR(_kA(_kS("category label", maxLen=50), "x-axis categories / slices", min=1, max=24)),
        "series": _kR(_kA(_kO({"name": _kR(_kS("series name")), "values": _kR(_kA(_K_NUM, "one value per category")),
                               "type": _kE(["bar", "line"], "combo only: line puts this series on the secondary axis")}, ""), "data series", min=1, max=6)),
        "line_unit": _kS("combo only: suffix for the line axis", maxLen=6), "line_prefix": _kS("combo only: prefix for the line axis", maxLen=3),
        "line_decimals": _kN("combo only: decimals on the line axis", min=0, max=3),
        "unit": _kS("value suffix like '%', 'M', 'x' (shown in labels)", maxLen=6), "prefix": _kS("value prefix like '€' or '$'", maxLen=3),
        "decimals": _kN("decimals in labels (default 0)", min=0, max=3), "show_values": _kBo("show data labels (default true)"),
        "legend": _kBo("force legend on/off (default: on when >1 series or pie)")}},
    "waterfall": {"use": "Business case bridge / value build-up / variance bridge drawn with native shapes.", "fields": {
        "steps": _kR(_kA(_kO({"label": _kR(_kS("step label", maxLen=28)), "value": _kR(_kN("delta (negative allowed); for total rows the value is ignored and computed")),
                              "total": _kBo("true = show running total bar")}, ""), "bars in order", min=2, max=12)),
        "unit": _kS("suffix e.g. 'M'", maxLen=6), "prefix": _kS("prefix e.g. '€'", maxLen=3), "decimals": _kN("decimals", min=0, max=2)}},
    "table": {"use": "Structured comparison, risk register, RACI, data table. Cells matching High/Medium/Low, Red/Amber/Green, Yes/No/Partial are auto-colored.", "fields": {
        "columns": _kR(_kA(_kS("header", maxLen=30), "column headers", min=2, max=7)), "rows": _kR(_kA(_kA(_K_STR, "cells"), "rows of cells (strings)", min=1, max=12)),
        "col_widths": _kA(_K_NUM, "relative widths, e.g. [2,1,4]"), "highlight_rows": _kA(_K_NUM, "row indexes to emphasize")}},
    "scorecard": {"use": "Option evaluation / vendor or capability assessment with Harvey balls (0-4).", "fields": {
        "options": _kR(_kA(_kS("option/column name", maxLen=22), "options compared", min=1, max=6)),
        "criteria": _kR(_kA(_kO({"name": _kR(_kS("criterion", maxLen=40)), "scores": _kR(_kA(_K_NUM, "0-4 per option")), "weight": _kN("optional weight")}, ""), "rows", min=1, max=10)),
        "show_total": _kBo("show weighted average row")}},
    "gantt": {"use": "Roadmap / plan with phases over time and optional milestones.", "fields": {
        "periods": _kR(_kO({"start": _kN("first tick, default 0"), "end": _kN("last tick (required unless labels are given)"), "step": _kN("tick step, default 1"),
                            "prefix": _kS("tick label prefix, e.g. 'M' or 'W'"),
                            "labels": _kA(_kS("label"), "explicit tick labels instead (e.g. ['Q1','Q2','Q3','Q4']); then phase start/end index these")}, "time axis")),
        "phases": _kR(_kA(_kO({"label": _kR(_kS("phase/workstream", maxLen=26)), "start": _kR(_kN("start tick")), "end": _kR(_kN("end tick")),
                               "text": _kS("scope line shown under the bar", maxLen=110), "color": _K_COLOR,
                               "depends_on": {"type": "any", "desc": "index (or array of indexes) of phases this one depends on — draws dependency arrows"}}, ""), "bars", min=1, max=9)),
        "milestones": _kA(_kO({"at": _kR(_kN("tick")), "label": _kR(_kS("milestone", maxLen=26))}, ""), "diamond milestones", max=6)}},
    "pyramid": {"use": "Hierarchy of needs, licensing tiers, value pyramid. 2-6 levels top to bottom.", "fields": {
        "levels": _kR(_kA(_kO({"title": _kR(_kS("level", maxLen=34)), "text": _kS("detail", maxLen=110)}, ""), "top (narrowest) to bottom", min=2, max=6))}},
    "funnel": {"use": "Conversion / pipeline / filtering stages with values.", "fields": {
        "stages": _kR(_kA(_kO({"title": _kR(_kS("stage", maxLen=30)), "value": _kS("e.g. '1,200' or '38%'", maxLen=10), "text": _kS("detail", maxLen=80)}, ""), "stages top to bottom", min=2, max=6))}},
    "comparison": {"use": "Side-by-side options, before/after, as-is vs to-be. 2-4 columns.", "fields": {
        "columns": _kR(_kA(_kO({"title": _kR(_kS("column title", maxLen=30)), "subtitle": _kS("subtitle", maxLen=50), "icon": _K_ICON,
                                "bullets": _kR(_kA(_kS("point", maxLen=110), "points", min=1, max=7)), "footer": _kS("verdict / bottom line", maxLen=90),
                                "highlight": _kBo("recommended / to-be column")}, ""), "columns", min=2, max=4)),
        "vs": _kBo("show a 'VS' badge between two columns")}},
    "matrix": {"use": "2x2 prioritization / positioning matrix with optional plotted items.", "fields": {
        "x_axis": _kR(_kO({"label": _kR(_kS("axis title")), "low": _kS("low end label"), "high": _kS("high end label")}, "horizontal axis")),
        "y_axis": _kR(_kO({"label": _kR(_kS("axis title")), "low": _kS("low end label"), "high": _kS("high end label")}, "vertical axis")),
        "quadrants": _kR(_kA(_kO({"title": _kR(_kS("quadrant name", maxLen=26)), "text": _kS("meaning / action", maxLen=90)}, ""), "exactly 4, order: top-left, top-right, bottom-left, bottom-right", min=4, max=4)),
        "points": _kA(_kO({"label": _kR(_kS("item", maxLen=22)), "x": _kR(_kN("0..1")), "y": _kR(_kN("0..1"))}, ""), "plotted items", max=12),
        "highlight": _kN("quadrant index to emphasize")}},
    "hub": {"use": "Ecosystem / partner map / capability around a core. Center + 3-8 nodes.", "fields": {
        "center": _kR(_kO({"title": _kR(_kS("core", maxLen=26)), "text": _kS("detail", maxLen=60)}, "center")),
        "nodes": _kR(_kA(_kO({"title": _kR(_kS("node", maxLen=26)), "text": _kS("detail", maxLen=70), "icon": _K_ICON}, ""), "satellites", min=3, max=8))}},
    "tree": {"use": "Issue tree, value-driver tree, capability map, org-style breakdown. Root + 2-6 branches with items.", "fields": {
        "root": _kR(_kO({"title": _kR(_kS("root question/goal", maxLen=60)), "text": _kS("detail", maxLen=80)}, "root")),
        "branches": _kR(_kA(_kO({"title": _kR(_kS("branch", maxLen=30)), "items": _kA(_kS("leaf", maxLen=60), "leaves", max=6)}, ""), "branches", min=2, max=6))}},
    "bullets": {"use": "Narrative points when no diagram fits. Use 'lead' for bold run-in heads. Max 8 items.", "fields": {
        "items": _kR(_kA({"type": "string|object", "fields": {"lead": _kS("bold lead-in"), "text": _kS("text"), "sub": _kA(_kS("sub-point"), "sub bullets", max=4)}},
                         "string or {lead, text, sub}", min=1, max=8)),
        "numbered": _kBo("numbered list"), "panel": _kBo("draw on a tinted panel")}},
    "callout": {"use": "Key message, big stat, or quote. Put the 'so what' here.", "fields": {
        "kind": _kE(["statement", "stat", "quote"], "style (default statement)"), "text": _kR(_kS("the message / quote / stat label", maxLen=220)),
        "value": _kS("for kind=stat: the big number", maxLen=10), "attribution": _kS("for quote: who said it, or source", maxLen=80),
        "icon": _K_ICON, "tone": _kE(["primary", "dark", "accent", "tint"], "fill (default primary)")}},
    "agenda": {"use": "Agenda / table of contents / executive summary numbered list. 2-10 items.", "fields": {
        "items": _kR(_kA(_kO({"title": _kR(_kS("item", maxLen=50)), "text": _kS("detail", maxLen=120)}, ""), "items", min=2, max=10)), "highlight": _kN("current item index")}},
    "progress": {"use": "Horizontal bars for percentages/scores (adoption, maturity, survey results) with optional targets.", "fields": {
        "items": _kR(_kA(_kO({"label": _kR(_kS("label", maxLen=40)), "value": _kR(_kN("value")), "max": _kN("scale max, default 100"),
                              "target": _kN("optional target marker"), "text": _kS("note", maxLen=60)}, ""), "bars", min=1, max=10)),
        "unit": _kS("suffix, default '%'", maxLen=6)}},
    "swot": {"use": "SWOT analysis: four colored quadrants with bullets.", "fields": {
        "strengths": _kR(_kA(_kS("point", maxLen=90), "strengths", min=1, max=6)), "weaknesses": _kR(_kA(_kS("point", maxLen=90), "weaknesses", min=1, max=6)),
        "opportunities": _kR(_kA(_kS("point", maxLen=90), "opportunities", min=1, max=6)), "threats": _kR(_kA(_kS("point", maxLen=90), "threats", min=1, max=6)),
        "titles": _kA(_kS("title"), "optional 4 custom quadrant titles", max=4)}},
    "venn": {"use": "Overlap of 2-3 domains/capabilities with the intersection highlighted (e.g., data + process + AI = value).", "fields": {
        "sets": _kR(_kA(_kO({"title": _kR(_kS("set", maxLen=22)), "text": _kS("detail", maxLen=60)}, ""), "sets", min=2, max=3)),
        "overlap": _kO({"title": _kR(_kS("intersection", maxLen=26)), "text": _kS("detail", maxLen=50)}, "label for the centre overlap")}},
    "cycle": {"use": "Closed loop / flywheel / continuous-improvement cycle. 3-8 steps around a circle with arrows.", "fields": {
        "steps": _kR(_kA(_kO({"title": _kR(_kS("step", maxLen=24)), "text": _kS("detail", maxLen=70), "icon": _K_ICON}, ""), "steps clockwise from top", min=3, max=8)),
        "center": _kO({"title": _kR(_kS("centre label", maxLen=24)), "text": _kS("detail", maxLen=40)}, "optional centre")}},
    "value_chain": {"use": "Porter-style value chain: support activities as bars, primary activities as chevrons, margin arrow.", "fields": {
        "primary": _kR(_kA(_kO({"title": _kR(_kS("activity", maxLen=22)), "text": _kS("detail", maxLen=60)}, ""), "primary activities left to right", min=3, max=7)),
        "support": _kA(_kO({"title": _kR(_kS("activity", maxLen=28)), "text": _kS("detail", maxLen=90)}, ""), "support activities", max=5),
        "margin": _kBo("show margin arrow (default true)"), "margin_label": _kS("e.g. 'Margin' or 'Value'", maxLen=12)}},
    "concentric": {"use": "Onion / layered-core diagram (e.g., clean core → extensions → ecosystem). 2-5 rings inner to outer, with legend.", "fields": {
        "rings": _kR(_kA(_kO({"title": _kR(_kS("ring", maxLen=20)), "text": _kS("detail for the legend", maxLen=90)}, ""), "inner first", min=2, max=5))}},
    "staircase": {"use": "Ascending steps: maturity journey, phased ambition, capability build-up. 2-6 steps.", "fields": {
        "steps": _kR(_kA(_kO({"title": _kR(_kS("step", maxLen=26)), "text": _kS("detail", maxLen=90), "label": _kS("small label above, e.g. 'Phase 1' or '2026'", maxLen=14)}, ""), "low to high", min=2, max=6))}},
    "flow": {"use": "Input → core → output (system context, operating model, data flow).", "fields": {
        "inputs": _kR(_kA(_kO({"title": _kR(_kS("input", maxLen=28)), "text": _kS("detail", maxLen=60), "icon": _K_ICON}, ""), "inputs", min=1, max=6)),
        "core": _kR(_kO({"title": _kR(_kS("core", maxLen=30)), "text": _kS("detail", maxLen=120), "icon": _K_ICON, "bullets": _kA(_kS("point", maxLen=50), "points", max=5)}, "centre")),
        "outputs": _kR(_kA(_kO({"title": _kR(_kS("output", maxLen=28)), "text": _kS("detail", maxLen=60), "icon": _K_ICON}, ""), "outputs", min=1, max=6)),
        "inputs_label": _kS("column label (default 'Inputs')"), "outputs_label": _kS("column label (default 'Outputs')"), "core_label": _kS("label above core")}},
    "heatmap": {"use": "Capability / coverage / readiness heatmap: rows × columns with scores 0-max (or RAG strings).", "fields": {
        "rows": _kR(_kA(_kS("row", maxLen=36), "rows", min=1, max=12)), "columns": _kR(_kA(_kS("column", maxLen=18), "columns", min=1, max=10)),
        "values": _kR(_kA(_kA(_K_ANY, "cells"), "matrix [row][col] of numbers 0..max or 'Red'/'Amber'/'Green'")),
        "max": _kN("top of scale (default 4)"), "scale_labels": _kA(_kS("label"), "optional labels per integer score, e.g. ['None','Low','Medium','High','Leading']"),
        "show_values": _kBo("print values in cells (default true)"), "legend": _kS("legend left label")}},
    "maturity": {"use": "Maturity assessment: dimensions × levels with current (dot) and target (ring) positions.", "fields": {
        "levels": _kR(_kA(_kS("level name", maxLen=16), "levels low→high", min=3, max=6)),
        "dimensions": _kR(_kA(_kO({"name": _kR(_kS("dimension", maxLen=34)), "current": _kR(_kN("1..levels")), "target": _kN("1..levels"), "note": _kS("small note", maxLen=50)}, ""), "rows", min=1, max=9))}},
    "risk_matrix": {"use": "Likelihood × impact heat grid (5x5 or 3x3) with numbered risks plotted and a legend list.", "fields": {
        "risks": _kR(_kA(_kO({"label": _kR(_kS("risk", maxLen=50)), "likelihood": _kR(_kN("1..scale")), "impact": _kR(_kN("1..scale")),
                              "id": _kS("marker label, default number", maxLen=3), "owner": _kS("owner", maxLen=20)}, ""), "risks", min=1, max=15)),
        "scale": _kE([3, 5], "grid size (default 5)"), "show_list": _kBo("show legend list (default true when wide)"),
        "x_label": _kS("default 'Impact'"), "y_label": _kS("default 'Likelihood'")}},
    "bubble": {"use": "Bubble chart (shape-drawn, labelled): portfolio, attractiveness vs position, value vs effort with size = spend.", "fields": {
        "x_axis": _kR(_kO({"label": _kR(_kS("axis title")), "min": _kN("min"), "max": _kN("max"), "unit": _kS("suffix"), "prefix": _kS("prefix")}, "x axis")),
        "y_axis": _kR(_kO({"label": _kR(_kS("axis title")), "min": _kN("min"), "max": _kN("max"), "unit": _kS("suffix"), "prefix": _kS("prefix")}, "y axis")),
        "bubbles": _kR(_kA(_kO({"label": _kR(_kS("item", maxLen=22)), "x": _kR(_kN("x")), "y": _kR(_kN("y")), "size": _kR(_kN("relative size")), "color": _K_COLOR}, ""), "bubbles", min=1, max=15)),
        "size_label": _kS("what bubble size means, e.g. 'revenue €M'")}},
    "mekko": {"use": "Marimekko / market map: column width = segment size, stacked shares within each column.", "fields": {
        "columns": _kR(_kA(_kO({"label": _kR(_kS("segment", maxLen=22)), "width": _kR(_kN("segment size (e.g., market €B)")),
                                "parts": _kR(_kA(_kO({"value": _kR(_kN("value or share")), "label": _kS("short label", maxLen=12)}, ""), "stack bottom→top", min=1, max=6))}, ""), "columns", min=2, max=8)),
        "part_names": _kA(_kS("legend name"), "names of stack layers in order", max=6), "unit": _kS("width suffix e.g. 'B'"), "prefix": _kS("width prefix e.g. '€'")}},
    "rings": {"use": "Circular progress gauges for 1-6 percentages (adoption, completion, coverage).", "fields": {
        "items": _kR(_kA(_kO({"value": _kR(_kN("value")), "max": _kN("default 100"), "label": _kR(_kS("label", maxLen=30)), "text": _kS("note", maxLen=60),
                              "unit": _kS("default '%'", maxLen=4), "prefix": _kS("prefix", maxLen=3)}, ""), "gauges", min=1, max=6))}},
    "org_chart": {"use": "Organisation / governance structure: root, up to 6 direct reports, up to 6 each below. Initials avatars or photos.", "fields": {
        "root": _kR(_kO({"name": _kR(_kS("name or role", maxLen=30)), "role": _kS("title", maxLen=40), "photo": _kS("optional image URL/file")}, "top")),
        "children": _kR(_kA(_kO({"name": _kR(_kS("name", maxLen=30)), "role": _kS("title", maxLen=40), "photo": _kS("image"),
                                 "children": _kA(_kO({"name": _kR(_kS("name", maxLen=30)), "role": _kS("title", maxLen=40), "photo": _kS("image")}, ""), "reports", max=6)}, ""), "direct reports", min=1, max=6))}},
    "profiles": {"use": "Team / leadership / speaker cards with photo or initials, role and short bio.", "fields": {
        "people": _kR(_kA(_kO({"name": _kR(_kS("name", maxLen=30)), "role": _kS("role", maxLen=45), "text": _kS("bio / responsibility", maxLen=180),
                               "photo": _kS("image URL or file name in the images folder")}, ""), "people", min=1, max=8)),
        "columns": _kN("columns", min=1, max=4)}},
    "raci": {"use": "RACI responsibility matrix: activities × roles with R/A/C/I chips.", "fields": {
        "activities": _kR(_kA(_kS("activity", maxLen=45), "rows", min=1, max=12)), "roles": _kR(_kA(_kS("role", maxLen=18), "columns", min=2, max=8)),
        "matrix": _kR(_kA(_kA(_kS("R, A, C, I, A/R or empty"), ""), "[activity][role]")), "activity_label": _kS("first column header")}},
    "swimlane": {"use": "Cross-functional process across lanes (roles/systems) and phases, with arrows and decision diamonds.", "fields": {
        "lanes": _kR(_kA(_kS("lane", maxLen=22), "lanes top→bottom", min=2, max=6)), "phases": _kA(_kS("phase", maxLen=24), "optional column headers", max=8),
        "steps": _kR(_kA(_kO({"lane": _kR({"type": "any", "desc": "lane index (0-based) or lane name"}), "col": _kR(_kN("column index 0-based")),
                              "title": _kR(_kS("step", maxLen=26)), "text": _kS("detail", maxLen=40), "shape": _kE(["process", "decision"], "default process"),
                              "highlight": _kBo("emphasize")}, ""), "steps", min=2, max=18)),
        "links": _kA(_kA(_K_NUM, ""), "optional [fromStep, toStep] pairs; default links steps in listed order")}},
    "checklist": {"use": "Action plan / next steps / readiness checklist with status icons, owner and due date.", "fields": {
        "items": _kR(_kA(_kO({"text": _kR(_kS("action", maxLen=90)), "status": _kE(["done", "in_progress", "todo", "blocked"], "status (default todo)"),
                              "owner": _kS("owner", maxLen=20), "due": _kS("due date", maxLen=14), "bold": _kBo("bold text")}, ""), "items", min=1, max=12))}},
    "status": {"use": "RAG status report: workstreams with status pill, trend arrow, owner and commentary.", "fields": {
        "items": _kR(_kA(_kO({"name": _kR(_kS("workstream", maxLen=30)), "status": _kR(_kE(["green", "amber", "red", "on_track", "at_risk", "off_track", "done", "not_started"], "status")),
                              "trend": _kE(["up", "flat", "down"], "trend"), "owner": _kS("owner", maxLen=20), "comment": _kS("commentary", maxLen=120),
                              "label": _kS("custom pill text", maxLen=12)}, ""), "rows", min=1, max=10)),
        "name_label": _kS("first column header")}},
    "board": {"use": "Kanban / initiative board: 2-5 columns of cards (backlog → done, now/next/later).", "fields": {
        "columns": _kR(_kA(_kO({"title": _kR(_kS("column", maxLen=20)), "color": _K_COLOR,
                                "cards": _kR(_kA(_kO({"title": _kR(_kS("card", maxLen=40)), "text": _kS("detail", maxLen=70), "tag": _kS("tag", maxLen=20)}, ""), "cards", min=0, max=6))}, ""), "columns", min=2, max=5))}},
    "persona": {"use": "Customer / user persona: profile panel plus goals, pain points, needs, behaviours.", "fields": {
        "name": _kR(_kS("persona name", maxLen=30)), "role": _kS("role", maxLen=40), "quote": _kS("quote", maxLen=140), "photo": _kS("optional image"),
        "attributes": _kA(_kO({"label": _kR(_kS("label", maxLen=14)), "value": _kR(_kS("value", maxLen=30))}, ""), "facts like Age, Location, Tools", max=6),
        "goals": _kA(_kS("goal", maxLen=90), "goals", max=6), "pains": _kA(_kS("pain", maxLen=90), "pain points", max=6),
        "needs": _kA(_kS("need", maxLen=90), "needs", max=6), "behaviours": _kA(_kS("behaviour", maxLen=90), "behaviours", max=6)}},
    "journey": {"use": "Customer / employee journey map: stages × actions/touchpoints/pain points/opportunities with an emotion curve.", "fields": {
        "stages": _kR(_kA(_kO({"title": _kR(_kS("stage", maxLen=20)), "actions": {"type": "any", "desc": "string or list"}, "touchpoints": {"type": "any", "desc": "string or list"},
                               "pains": {"type": "any", "desc": "string or list"}, "opportunities": {"type": "any", "desc": "string or list"},
                               "emotion": _kN("-2 (frustrated) .. 2 (delighted)", min=-2, max=2)}, ""), "stages", min=3, max=7)),
        "emotion_label": _kS("row label for the curve")}},
    "image": {"use": "Photo, screenshot, logo or product visual. src = https URL, file name in the images folder, or data URI.", "fields": {
        "src": _kR(_kS("image source")), "caption": _kS("caption", maxLen=120),
        "fit": _kE(["cover", "contain"], "cover crops to fill; contain shows the whole image (screenshots, logos)"), "alt": _kS("description shown if the image cannot be loaded")}},
}

# Neutral, obviously fictional catalogue examples ("Northwind"): they teach structure only.
_K_EXAMPLES = json.loads(r'''{"kpis":{"type":"kpis","items":[{"value":"61%","label":"Northwind orders shipped same day (fictional)","icon":"truck","trend":[48,52,55,61]},{"value":"€13M","label":"Northwind annual savings target (fictional)","icon":"money"},{"value":"4.2x","label":"Northwind faster claims handling (fictional)","icon":"speed"}]},"timeline":{"type":"timeline","items":[{"date":"Phase A","title":"Northwind pilot","text":"Placeholder milestone text."},{"date":"Phase B","title":"Northwind rollout","text":"Placeholder milestone text."},{"date":"Phase C","title":"Northwind scale-up","text":"Placeholder milestone text.","highlight":true}]},"cards":{"type":"cards","columns":3,"items":[{"icon":"idea","title":"Northwind point one","text":"Placeholder: one or two sentences of researched content."},{"icon":"target","title":"Northwind point two","text":"Placeholder: one or two sentences of researched content."},{"icon":"rocket","title":"Northwind point three","text":"Placeholder: one or two sentences of researched content."}]},"process":{"type":"process","steps":[{"icon":"search","title":"Northwind step 1","text":"Placeholder step text."},{"icon":"settings","title":"Northwind step 2","text":"Placeholder step text."},{"icon":"check","title":"Northwind step 3","text":"Placeholder step text."}]},"chevrons":{"type":"chevrons","stages":[{"title":"Northwind level 1","text":"Placeholder"},{"title":"Northwind level 2","text":"Placeholder"},{"title":"Northwind level 3","text":"Placeholder"}]},"layers":{"type":"layers","tiers":[{"label":"Northwind tier A","sublabel":"Placeholder","items":[{"title":"Component A1","text":"Placeholder detail"},{"title":"Component A2","text":"Placeholder detail"}]},{"label":"Northwind tier B","color":"primary","items":[{"title":"Component B1","bullets":["Placeholder","Placeholder"]},{"title":"Component B2","bullets":["Placeholder"],"highlight":true}],"base":"Northwind foundation placeholder"}],"side":{"title":"Northwind governance","icon":"shield","items":[{"title":"Control 1","text":"Placeholder"}]}},"chart":{"type":"chart","chart_type":"bar","categories":["Northwind depot A","Northwind depot B","Northwind depot C"],"series":[{"name":"On-time rate","values":[71,64,58]}],"unit":"%"},"waterfall":{"type":"waterfall","prefix":"€","unit":"M","steps":[{"label":"Northwind baseline","value":31.7},{"label":"Northwind saving","value":-4.3},{"label":"Northwind new cost","value":1.9},{"label":"Northwind result","total":true,"value":0}]},"table":{"type":"table","columns":["Northwind item","Rating","Note"],"col_widths":[2,1,4],"rows":[["Northwind risk one","High","Placeholder mitigation"],["Northwind risk two","Medium","Placeholder mitigation"]]},"scorecard":{"type":"scorecard","options":["Northwind option X","Northwind option Y"],"show_total":true,"criteria":[{"name":"Criterion placeholder 1","scores":[3,2]},{"name":"Criterion placeholder 2","scores":[1,4]}]},"gantt":{"type":"gantt","periods":{"start":0,"end":12,"step":3,"prefix":"M"},"phases":[{"label":"Northwind phase 1","start":0,"end":4,"text":"Placeholder scope"},{"label":"Northwind phase 2","start":3,"end":12,"text":"Placeholder scope","color":"primary","depends_on":0}],"milestones":[{"at":6,"label":"Northwind gate"}]},"pyramid":{"type":"pyramid","levels":[{"title":"Northwind top tier","text":"Placeholder"},{"title":"Northwind middle tier","text":"Placeholder"},{"title":"Northwind base tier","text":"Placeholder"}]},"funnel":{"type":"funnel","stages":[{"title":"Northwind leads","value":"917"},{"title":"Northwind qualified","value":"233"},{"title":"Northwind won","value":"41"}]},"comparison":{"type":"comparison","columns":[{"title":"Northwind today","bullets":["Placeholder point","Placeholder point"]},{"title":"Northwind target","highlight":true,"bullets":["Placeholder point","Placeholder point"]}]},"matrix":{"type":"matrix","x_axis":{"label":"Effort","low":"Low","high":"High"},"y_axis":{"label":"Value","low":"Low","high":"High"},"highlight":0,"quadrants":[{"title":"Quadrant TL","text":"Placeholder"},{"title":"Quadrant TR","text":"Placeholder"},{"title":"Quadrant BL","text":"Placeholder"},{"title":"Quadrant BR","text":"Placeholder"}],"points":[{"label":"Northwind idea 1","x":0.3,"y":0.8}]},"hub":{"type":"hub","center":{"title":"Northwind core","text":"Placeholder"},"nodes":[{"title":"Northwind node 1","icon":"truck"},{"title":"Northwind node 2","icon":"warehouse"},{"title":"Northwind node 3","icon":"users"}]},"tree":{"type":"tree","root":{"title":"Northwind goal placeholder"},"branches":[{"title":"Northwind driver 1","items":["Placeholder lever","Placeholder lever"]},{"title":"Northwind driver 2","items":["Placeholder lever"]}]},"bullets":{"type":"bullets","items":[{"lead":"Northwind point:","text":"placeholder narrative sentence."},"Northwind second placeholder point."]},"callout":{"type":"callout","kind":"stat","value":"7.7x","text":"Northwind placeholder statistic (fictional)","attribution":"Fictional example"},"agenda":{"type":"agenda","items":[{"title":"Northwind section 1","text":"Placeholder"},{"title":"Northwind section 2","text":"Placeholder"}]},"progress":{"type":"progress","items":[{"label":"Northwind unit 1","value":63,"target":80},{"label":"Northwind unit 2","value":37,"target":70}]},"swot":{"type":"swot","strengths":["Northwind strength placeholder"],"weaknesses":["Northwind weakness placeholder"],"opportunities":["Northwind opportunity placeholder"],"threats":["Northwind threat placeholder"]},"venn":{"type":"venn","sets":[{"title":"Northwind set A"},{"title":"Northwind set B"},{"title":"Northwind set C"}],"overlap":{"title":"Northwind overlap"}},"cycle":{"type":"cycle","center":{"title":"Northwind loop"},"steps":[{"title":"Northwind stage 1","icon":"users"},{"title":"Northwind stage 2","icon":"database"},{"title":"Northwind stage 3","icon":"growth"}]},"value_chain":{"type":"value_chain","support":[{"title":"Northwind support 1","text":"Placeholder"}],"primary":[{"title":"Northwind act 1"},{"title":"Northwind act 2"},{"title":"Northwind act 3"}]},"concentric":{"type":"concentric","rings":[{"title":"Northwind core","text":"Placeholder"},{"title":"Northwind middle","text":"Placeholder"},{"title":"Northwind outer","text":"Placeholder"}]},"staircase":{"type":"staircase","steps":[{"label":"Step 1","title":"Northwind level A"},{"label":"Step 2","title":"Northwind level B"},{"label":"Step 3","title":"Northwind level C"}]},"flow":{"type":"flow","inputs":[{"title":"Northwind input 1","icon":"database"},{"title":"Northwind input 2","icon":"document"}],"core":{"title":"Northwind engine","icon":"settings","bullets":["Placeholder","Placeholder"]},"outputs":[{"title":"Northwind output 1","icon":"check"}]},"heatmap":{"type":"heatmap","rows":["Northwind region 1","Northwind region 2"],"columns":["Dim A","Dim B","Dim C"],"values":[[3,1,2],[2,4,0]],"scale_labels":["None","Low","Medium","High","Leading"]},"maturity":{"type":"maturity","levels":["Level 1","Level 2","Level 3","Level 4","Level 5"],"dimensions":[{"name":"Northwind dimension 1","current":2,"target":4},{"name":"Northwind dimension 2","current":1,"target":3}]},"risk_matrix":{"type":"risk_matrix","risks":[{"label":"Northwind risk placeholder 1","likelihood":4,"impact":3,"owner":"Owner A"},{"label":"Northwind risk placeholder 2","likelihood":2,"impact":5}]},"bubble":{"type":"bubble","x_axis":{"label":"Effort (months)"},"y_axis":{"label":"Value (€M)"},"size_label":"placeholder size measure","bubbles":[{"label":"Northwind item 1","x":3,"y":5.3,"size":470},{"label":"Northwind item 2","x":7,"y":2.9,"size":190}]},"mekko":{"type":"mekko","prefix":"€","unit":"B","part_names":["Northwind","Rival A","Others"],"columns":[{"label":"Region North","width":17,"parts":[{"value":41},{"value":23},{"value":36}]},{"label":"Region South","width":29,"parts":[{"value":27},{"value":33},{"value":40}]}]},"rings":{"type":"rings","items":[{"value":67,"label":"Northwind metric 1"},{"value":39,"label":"Northwind metric 2"}]},"org_chart":{"type":"org_chart","root":{"name":"Northwind Board","role":"Placeholder chair"},"children":[{"name":"Northwind Unit A","role":"Placeholder lead","children":[{"name":"Team A1"}]},{"name":"Northwind Unit B","role":"Placeholder lead"}]},"profiles":{"type":"profiles","people":[{"name":"Alex Example","role":"Placeholder role","text":"Placeholder biography sentence."}]},"raci":{"type":"raci","roles":["Role A","Role B","Role C"],"activities":["Northwind activity 1","Northwind activity 2"],"matrix":[["A","C","R"],["I","A","R"]]},"swimlane":{"type":"swimlane","lanes":["Northwind lane 1","Northwind lane 2","Northwind lane 3"],"steps":[{"lane":0,"col":0,"title":"Northwind start"},{"lane":1,"col":1,"title":"Northwind check"},{"lane":2,"col":2,"title":"Approve?","shape":"decision"},{"lane":1,"col":3,"title":"Northwind finish"}]},"checklist":{"type":"checklist","items":[{"text":"Northwind action placeholder 1","status":"done","owner":"Owner A","due":"Date 1"},{"text":"Northwind action placeholder 2","status":"in_progress","owner":"Owner B","due":"Date 2"}]},"status":{"type":"status","items":[{"name":"Northwind stream 1","status":"amber","trend":"up","owner":"Owner A","comment":"Placeholder commentary"}]},"board":{"type":"board","columns":[{"title":"Now","cards":[{"title":"Northwind card 1","tag":"Tag A"}]},{"title":"Next","cards":[{"title":"Northwind card 2"}]},{"title":"Later","cards":[{"title":"Northwind card 3"}]}]},"persona":{"type":"persona","name":"Northwind persona","role":"Placeholder role","quote":"Placeholder quote from research interviews.","goals":["Placeholder goal"],"pains":["Placeholder pain point"],"needs":["Placeholder need"]},"journey":{"type":"journey","stages":[{"title":"Stage 1","actions":"Placeholder action","pains":"Placeholder pain","emotion":-1},{"title":"Stage 2","actions":"Placeholder action","emotion":0},{"title":"Stage 3","actions":"Placeholder action","opportunities":"Placeholder idea","emotion":1}]},"image":{"type":"image","src":"your-file-name.png","fit":"contain","caption":"Placeholder caption"}}''')
for _k, _ex in _K_EXAMPLES.items():
    if _k in _K_COMPONENTS:
        _K_COMPONENTS[_k]["example"] = _ex

_K_BLOCK_COMMON = {"heading": _kS("optional small heading above the block", maxLen=70), "w": _kN("relative width within the row (default 1)")}

_K_SLIDE_TYPES = {
    "cover": {"title": _kR(_kS("deck title", maxLen=110)), "subtitle": _kS("subtitle", maxLen=160), "kicker": _kS("small line above title", maxLen=80),
              "meta": _kS("presenter · organisation · date", maxLen=100), "blocks": _kA({"type": "object"}, "optional single row of blocks under the title (e.g., kpis)", max=3),
              "notes": _kS("speaker notes")},
    "closing": {"title": _kR(_kS("closing message, e.g. 'Next steps' or 'Thank you'", maxLen=90)), "subtitle": _kS("subtitle", maxLen=160),
                "kicker": _kS("small line above title", maxLen=60), "contact": _kA(_kS("contact line, e.g. 'Jane Doe · jane@firm.com'", maxLen=70), "contact lines", max=4),
                "blocks": _kA({"type": "object"}, "optional single row of blocks (e.g., checklist or cards)", max=3), "notes": _kS("speaker notes")},
    "section": {"title": _kR(_kS("section title", maxLen=80)), "number": _kS("e.g. '02'"), "subtitle": _kS("subtitle", maxLen=160), "notes": _kS("speaker notes")},
    "content": {"title": _kR(_kS("ACTION TITLE: full-sentence takeaway, max ~25 words", maxLen=170)), "kicker": _kS("small section label above the title", maxLen=70),
                "dark": _kBo("dark background (use for exec summary / closing)"),
                "rows": _kR(_kA(_kO({"h": _kN("relative row height (default by component)"), "blocks": _kR(_kA({"type": "object"}, "1-4 blocks side by side", min=1, max=4))}, ""),
                                "rows top to bottom", min=1, max=4)),
                "source": _kS("source line shown at the bottom — REQUIRED whenever the slide has facts or numbers", maxLen=240),
                "notes": _kS("speaker notes: talking points + full source URLs")},
}
_K_DECK = {
    "meta": _kO({"title": _kS("deck title"), "author": _kS("author"), "company": _kS("company"), "footer": _kS("footer text on every slide, e.g. 'Confidential'", maxLen=40)}, "metadata"),
    "theme": {"type": "string|object", "desc": "preset name or {preset, primary:'hex', ...overrides, fonts:{head, body}}"},
    "slides": _kR(_kA({"type": "object"}, "slides", min=1, max=40)),
    "engine": _kS("'kit' to force this engine"), "filename": _kS("optional output file name"),
}


def _k_type(v):
    if isinstance(v, bool):
        return "boolean"
    if isinstance(v, (int, float)):
        return "number"
    if isinstance(v, str):
        return "string"
    if isinstance(v, list):
        return "array"
    if isinstance(v, dict):
        return "object"
    return "null"


def _k_check(spec, v, path, errs, warns):
    if v is None:
        if spec.get("required"):
            errs.append(f"{path}: required {spec['type']}" + (f" ({spec['desc']})" if spec.get("desc") else ""))
        return
    t = _k_type(v)
    st = spec["type"]
    if st == "string":
        if t not in ("string", "number"):
            errs.append(f"{path}: expected string, got {t}")
        elif spec.get("maxLen") and len(_ks(v)) > spec["maxLen"] * 1.25:
            warns.append(f"{path}: {len(_ks(v))} chars, guide is ≤{spec['maxLen']} — shorten to avoid shrunken text")
    elif st == "number":
        if t != "number" or (isinstance(v, float) and v != v):
            errs.append(f"{path}: expected number, got {t} {json.dumps(v)[:30]}")
        else:
            if spec.get("min") is not None and v < spec["min"]:
                errs.append(f"{path}: must be ≥ {spec['min']}")
            if spec.get("max") is not None and v > spec["max"]:
                errs.append(f"{path}: must be ≤ {spec['max']}")
    elif st == "boolean":
        if t != "boolean":
            errs.append(f"{path}: expected true/false")
    elif st == "enum":
        if v not in spec["values"]:
            errs.append(f"{path}: '{v}' invalid; use one of {' | '.join(str(x) for x in spec['values'])}")
    elif st == "array":
        if t != "array":
            errs.append(f"{path}: expected array, got {t}")
            return
        if spec.get("min") is not None and len(v) < spec["min"]:
            errs.append(f"{path}: needs at least {spec['min']} entries (got {len(v)})")
        if spec.get("max") is not None and len(v) > spec["max"]:
            errs.append(f"{path}: at most {spec['max']} entries (got {len(v)}) — split across blocks/slides")
        if spec.get("of"):
            for i, x in enumerate(v):
                _k_check(spec["of"], x, f"{path}[{i}]", errs, warns)
    elif st == "object":
        if t != "object":
            errs.append(f"{path}: expected object, got {t}")
            return
        if spec.get("fields"):
            _k_check_fields(spec["fields"], v, path, errs, warns)
    elif st == "string|object":
        if t == "string":
            return
        if t != "object":
            errs.append(f"{path}: expected string or object")
            return
        if spec.get("fields"):
            _k_check_fields(spec["fields"], v, path, errs, warns)


def _k_check_fields(fields, obj, path, errs, warns, extra=()):
    for k, s in fields.items():
        _k_check(s, obj.get(k), f"{path}.{k}", errs, warns)
    for k in obj:
        if k not in fields and k not in extra and not str(k).startswith("_"):
            warns.append(f"{path}.{k}: unknown field ignored")


def _k_validate_block(b, path, errs, warns):
    if not isinstance(b, dict):
        errs.append(f"{path}: block must be an object")
        return
    comp = _K_COMPONENTS.get(b.get("type"))
    if not comp:
        errs.append(f"{path}.type: '{b.get('type')}' is not a component. Use one of: {', '.join(_K_COMPONENTS)}")
        return
    ty = b["type"]
    _k_check_fields({**comp["fields"], **_K_BLOCK_COMMON}, b, f"{path}({ty})", errs, warns, ("type",))
    L = lambda k: b.get(k) if isinstance(b.get(k), list) else None  # noqa: E731
    if ty == "chart" and L("series") is not None and L("categories") is not None:
        for i, s in enumerate(b["series"]):
            if isinstance(s, dict) and isinstance(s.get("values"), list) and len(s["values"]) != len(b["categories"]):
                errs.append(f"{path}(chart).series[{i}].values: {len(s['values'])} values but {len(b['categories'])} categories")
    if ty == "chart" and b.get("chart_type") == "combo" and L("series") is not None and not any(isinstance(x, dict) and x.get("type") == "line" for x in b["series"]):
        errs.append(f'{path}(chart): combo needs at least one series with "type": "line"')
    if ty == "scorecard" and L("criteria") is not None and L("options") is not None:
        for i, cr in enumerate(b["criteria"]):
            if isinstance(cr, dict) and isinstance(cr.get("scores"), list):
                if len(cr["scores"]) != len(b["options"]):
                    errs.append(f"{path}(scorecard).criteria[{i}].scores: need {len(b['options'])} scores")
                if any(isinstance(s, (int, float)) and (s < 0 or s > 4) for s in cr["scores"]):
                    errs.append(f"{path}(scorecard).criteria[{i}].scores: values must be 0-4")
    if ty == "raci" and L("matrix") is not None and L("activities") is not None:
        if len(b["matrix"]) != len(b["activities"]):
            errs.append(f"{path}(raci).matrix: {len(b['matrix'])} rows but {len(b['activities'])} activities")
        for i, r in enumerate(b["matrix"]):
            if isinstance(r, list) and L("roles") is not None and len(r) != len(b["roles"]):
                errs.append(f"{path}(raci).matrix[{i}]: need {len(b['roles'])} cells")
    if ty == "heatmap" and L("values") is not None and L("rows") is not None:
        if len(b["values"]) != len(b["rows"]):
            errs.append(f"{path}(heatmap).values: {len(b['values'])} rows but {len(b['rows'])} row labels")
        for i, r in enumerate(b["values"]):
            if isinstance(r, list) and L("columns") is not None and len(r) != len(b["columns"]):
                errs.append(f"{path}(heatmap).values[{i}]: need {len(b['columns'])} cells")
    if ty == "maturity" and L("dimensions") is not None and L("levels") is not None:
        for i, d in enumerate(b["dimensions"]):
            for k in ("current", "target"):
                if isinstance(d, dict) and isinstance(d.get(k), (int, float)) and not (1 <= d[k] <= len(b["levels"])):
                    errs.append(f"{path}(maturity).dimensions[{i}].{k}: must be 1..{len(b['levels'])}")
    if ty == "risk_matrix" and L("risks") is not None:
        N = b.get("scale") or 5
        for i, r in enumerate(b["risks"]):
            for k in ("likelihood", "impact"):
                if isinstance(r, dict) and isinstance(r.get(k), (int, float)) and not (1 <= r[k] <= N):
                    errs.append(f"{path}(risk_matrix).risks[{i}].{k}: must be 1..{N}")
    if ty == "swimlane" and L("steps") is not None and L("lanes") is not None:
        for i, st in enumerate(b["steps"]):
            ln = st.get("lane") if isinstance(st, dict) else None
            if (isinstance(ln, (int, float)) and not isinstance(ln, bool) and not (0 <= ln < len(b["lanes"]))) or (isinstance(ln, str) and ln not in b["lanes"]):
                errs.append(f"{path}(swimlane).steps[{i}].lane: '{ln}' is not a lane index or name")
    if ty == "gantt" and isinstance(b.get("periods"), dict) and b["periods"].get("end") is None and not _list(b["periods"].get("labels")):
        errs.append(f"{path}(gantt).periods: give either 'end' (numeric axis) or 'labels'")
    if ty == "table" and L("rows") is not None and L("columns") is not None:
        for i, r in enumerate(b["rows"]):
            if isinstance(r, list) and len(r) != len(b["columns"]):
                errs.append(f"{path}(table).rows[{i}]: {len(r)} cells but {len(b['columns'])} columns")


def _k_validate_deck(deck):
    errs, warns = [], []
    if not isinstance(deck, dict):
        return ["deck must be a JSON object {meta, theme, slides}"], []
    _k_check_fields(_K_DECK, deck, "deck", errs, warns, ("title", "subtitle", "author", "footer", "layout_mode"))
    for i, s in enumerate(_list(deck.get("slides"))):
        p = f"slides[{i}]"
        if not isinstance(s, dict):
            errs.append(f"{p}: must be an object")
            continue
        ty = s.get("type") or "content"
        st = _K_SLIDE_TYPES.get(ty)
        if not st:
            errs.append(f"{p}.type: '{ty}' invalid; use cover | section | content | closing")
            continue
        _k_check_fields(st, s, p, errs, warns, ("type",))
        if ty == "content" and isinstance(s.get("rows"), list):
            nb = 0
            for ri, r in enumerate(s["rows"]):
                for bi, blk in enumerate(_list(r.get("blocks")) if isinstance(r, dict) else []):
                    nb += 1
                    _k_validate_block(blk, f"{p}.rows[{ri}].blocks[{bi}]", errs, warns)
            if nb > 6:
                warns.append(f"{p}: {nb} blocks is very dense — max ~5 recommended")
        if ty in ("cover", "closing") and isinstance(s.get("blocks"), list):
            for bi, blk in enumerate(s["blocks"]):
                _k_validate_block(blk, f"{p}.blocks[{bi}]", errs, warns)
    return errs, warns


def _k_catalogue(names=None) -> dict:
    def fmt(spec):
        st = spec.get("type")
        if st == "object" and spec.get("fields"):
            d = {"_": spec.get("desc", "")} if spec.get("desc") else {}
            d.update({k + ("*" if s.get("required") else ""): fmt(s) for k, s in spec["fields"].items()})
            return d
        if st == "array":
            rng = f"({spec.get('min', 0)}-{spec.get('max', 'n')})"
            of = spec.get("of") or {}
            if of.get("fields"):
                return [fmt(of), f"{spec.get('desc', '')} {rng}".strip()]
            return f"array of {of.get('type', 'any')}: {spec.get('desc', '')} {rng}"
        if st == "string|object":
            d = {"_": spec.get("desc", "")}
            d.update({k: fmt(s) for k, s in (spec.get("fields") or {}).items()})
            return d
        if st == "enum":
            return f"{'|'.join(str(x) for x in spec['values'])} — {spec.get('desc', '')}"
        return f"{st}{(' ≤' + str(spec['maxLen']) + ' chars') if spec.get('maxLen') else ''}: {spec.get('desc', '')}"

    comps = {k: {"use": v["use"], "fields": {f + ("*" if s.get("required") else ""): fmt(s) for f, s in v["fields"].items()}, "example": v.get("example")}
             for k, v in _K_COMPONENTS.items()}
    if names:
        want = [n.strip() for n in names if n.strip()]
        out = {"components": {k: comps[k] for k in want if k in comps}}
        unknown = [n for n in want if n not in comps]
        if unknown:
            out["unknown_components"] = unknown
            out["valid_components"] = list(comps)
        return out
    return {
        "deck_shape": {"meta": {"title": "...", "author": "...", "footer": "Confidential"}, "theme": "midnight_violet",
                       "slides": ["cover slide", "content slides...", "optional section slides", "closing slide"]},
        "slide_types": {k: {f + ("*" if s.get("required") else ""): fmt(s) for f, s in v.items()} for k, v in _K_SLIDE_TYPES.items()},
        "block_common_fields": {"heading": "optional small heading above block", "w": "relative width in its row (default 1)"},
        "components": comps,
        "themes": {k: v["label"] for k, v in _K_PRESETS.items()},
        "icon_aliases": sorted(_KIT_ICON_ALIASES),
        "legend": "* = required. Content slide = title + rows[]; each row = blocks[] side by side; each block = {type:<component>, heading?, w?, ...fields}.",
    }


# ------------------------------------------------------- layout recipes ----
# Proven content-slide skeletons ("templates"). The model copies one, then
# replaces every WRITE marker with researched content.
_K_RECIPES = {
    "executive_summary": {"use": "Exec summary on a dark slide: numbers, path, implications", "slide": {
        "type": "content", "dark": True, "kicker": "Executive summary", "title": "WRITE: the answer in one sentence",
        "rows": [{"h": 0.8, "blocks": [{"type": "kpis", "items": [{"value": "WRITE", "label": "WRITE", "icon": "growth"}, {"value": "WRITE", "label": "WRITE", "icon": "money"}, {"value": "WRITE", "label": "WRITE", "icon": "users"}]}]},
                 {"h": 1.05, "blocks": [{"type": "timeline", "items": [{"date": "WRITE", "title": "WRITE", "text": "WRITE"}, {"date": "WRITE", "title": "WRITE", "text": "WRITE"}, {"date": "WRITE", "title": "WRITE", "text": "WRITE", "highlight": True}]}]},
                 {"h": 0.85, "blocks": [{"type": "cards", "heading": "What it means", "columns": 4, "items": [{"icon": "target", "title": "WRITE", "text": "WRITE"}, {"icon": "shield", "title": "WRITE", "text": "WRITE"}, {"icon": "money", "title": "WRITE", "text": "WRITE"}, {"icon": "rocket", "title": "WRITE", "text": "WRITE"}]}]}]}},
    "architecture": {"use": "Layered architecture / operating model with governance side panel", "slide": {
        "kicker": "Architecture", "title": "WRITE",
        "rows": [{"blocks": [{"type": "layers", "tiers": [{"label": "WRITE", "sublabel": "WRITE", "items": [{"title": "WRITE", "text": "WRITE"}, {"title": "WRITE", "text": "WRITE"}]},
                                                       {"label": "WRITE", "color": "primary", "items": [{"title": "WRITE", "bullets": ["WRITE", "WRITE"]}, {"title": "WRITE", "bullets": ["WRITE"], "highlight": True}], "base": "WRITE"},
                                                       {"label": "WRITE", "color": "neutral", "items": [{"title": "WRITE"}, {"title": "WRITE"}]}],
                              "side": {"title": "Governance", "icon": "shield", "items": [{"title": "WRITE", "text": "WRITE"}, {"title": "WRITE", "text": "WRITE"}]}}]}]}},
    "how_it_works": {"use": "Process + maturity strip + evidence", "slide": {
        "kicker": "How it works", "title": "WRITE",
        "rows": [{"h": 1.15, "blocks": [{"type": "process", "steps": [{"icon": "chat", "title": "WRITE", "text": "WRITE"}, {"icon": "settings", "title": "WRITE", "text": "WRITE"}, {"icon": "check", "title": "WRITE", "text": "WRITE"}]}]},
                 {"h": 0.4, "blocks": [{"type": "chevrons", "stages": [{"title": "WRITE"}, {"title": "WRITE"}, {"title": "WRITE"}]}]},
                 {"h": 1.45, "blocks": [{"type": "chart", "heading": "WRITE", "chart_type": "bar", "categories": ["WRITE", "WRITE"], "series": [{"name": "WRITE", "values": [0, 0]}], "unit": "%"},
                                        {"type": "cards", "heading": "Proof points", "columns": 2, "items": [{"title": "WRITE", "tag": "WRITE", "text": "WRITE"}, {"title": "WRITE", "tag": "WRITE", "text": "WRITE"}]}]}]}},
    "business_case": {"use": "Value bridge + risks/assumptions", "slide": {
        "kicker": "Business case", "title": "WRITE",
        "rows": [{"blocks": [{"type": "waterfall", "w": 1.2, "heading": "WRITE (€M)", "prefix": "€", "unit": "M", "steps": [{"label": "WRITE", "value": 0}, {"label": "WRITE", "value": 0}, {"label": "WRITE", "value": 0}, {"label": "WRITE", "total": True, "value": 0}]},
                             {"type": "table", "heading": "Key risks", "columns": ["Risk", "Exposure", "Mitigation"], "col_widths": [2, 1, 3], "rows": [["WRITE", "High", "WRITE"], ["WRITE", "Medium", "WRITE"]]}]}]}},
    "decision": {"use": "Option scorecard + prioritisation matrix", "slide": {
        "kicker": "Decision", "title": "WRITE",
        "rows": [{"blocks": [{"type": "scorecard", "heading": "Option assessment", "options": ["WRITE", "WRITE"], "show_total": True, "criteria": [{"name": "WRITE", "scores": [3, 2]}, {"name": "WRITE", "scores": [2, 4]}]},
                             {"type": "matrix", "heading": "Prioritisation", "x_axis": {"label": "Effort", "low": "Low", "high": "High"}, "y_axis": {"label": "Value", "low": "Low", "high": "High"}, "highlight": 0,
                              "quadrants": [{"title": "Quick wins"}, {"title": "Strategic bets"}, {"title": "Fill-ins"}, {"title": "Avoid"}], "points": [{"label": "WRITE", "x": 0.3, "y": 0.8}]}]}]}},
    "options_vs": {"use": "As-is vs to-be (or option A vs B) with the so-what callout", "slide": {
        "kicker": "Options", "title": "WRITE",
        "rows": [{"blocks": [{"type": "comparison", "w": 1.5, "vs": True, "columns": [{"title": "WRITE", "bullets": ["WRITE", "WRITE"]}, {"title": "WRITE", "highlight": True, "bullets": ["WRITE", "WRITE"], "footer": "WRITE"}]},
                             {"type": "callout", "kind": "stat", "value": "WRITE", "text": "WRITE"}]}]}},
    "roadmap": {"use": "Gantt with dependencies and milestones + scope cards", "slide": {
        "kicker": "Roadmap", "title": "WRITE",
        "rows": [{"blocks": [{"type": "cards", "w": 0.75, "columns": 1, "items": [{"title": "WRITE", "text": "WRITE"}, {"title": "WRITE", "text": "WRITE"}, {"title": "WRITE", "text": "WRITE"}]},
                             {"type": "gantt", "w": 1.6, "periods": {"labels": ["Q1", "Q2", "Q3", "Q4"]}, "phases": [{"label": "WRITE", "start": 0, "end": 0, "text": "WRITE"}, {"label": "WRITE", "start": 1, "end": 2, "text": "WRITE", "depends_on": 0}, {"label": "WRITE", "start": 2, "end": 3, "text": "WRITE", "color": "primary", "depends_on": 1}],
                              "milestones": [{"at": 1, "label": "WRITE"}]}]}]}},
    "strategy": {"use": "SWOT + overlap of capabilities", "slide": {
        "kicker": "Strategy", "title": "WRITE",
        "rows": [{"blocks": [{"type": "swot", "w": 1.35, "strengths": ["WRITE"], "weaknesses": ["WRITE"], "opportunities": ["WRITE"], "threats": ["WRITE"]},
                             {"type": "venn", "sets": [{"title": "WRITE"}, {"title": "WRITE"}, {"title": "WRITE"}], "overlap": {"title": "WRITE"}}]}]}},
    "assessment": {"use": "Readiness heatmap + current vs target maturity", "slide": {
        "kicker": "Assessment", "title": "WRITE",
        "rows": [{"blocks": [{"type": "heatmap", "heading": "Readiness", "rows": ["WRITE", "WRITE"], "columns": ["Data", "Process", "Skills"], "values": [[3, 2, 1], [2, 3, 2]], "scale_labels": ["None", "Low", "Medium", "High", "Leading"]},
                             {"type": "maturity", "heading": "Maturity", "levels": ["Initial", "Managed", "Defined", "Measured", "Optimized"], "dimensions": [{"name": "WRITE", "current": 2, "target": 4}, {"name": "WRITE", "current": 1, "target": 3}]}]}]}},
    "risks": {"use": "Risk heat grid with numbered legend", "slide": {
        "kicker": "Risks", "title": "WRITE",
        "rows": [{"blocks": [{"type": "risk_matrix", "risks": [{"label": "WRITE", "likelihood": 4, "impact": 4, "owner": "WRITE"}, {"label": "WRITE", "likelihood": 2, "impact": 3}]}]}]}},
    "pmo_status": {"use": "Programme status + now/next/later board", "slide": {
        "kicker": "Programme status", "title": "WRITE",
        "rows": [{"h": 1.45, "blocks": [{"type": "status", "items": [{"name": "WRITE", "status": "green", "trend": "up", "owner": "WRITE", "comment": "WRITE"}, {"name": "WRITE", "status": "amber", "trend": "flat", "owner": "WRITE", "comment": "WRITE"}]}]},
                 {"h": 1.2, "blocks": [{"type": "board", "columns": [{"title": "Now", "cards": [{"title": "WRITE"}]}, {"title": "Next", "cards": [{"title": "WRITE"}]}, {"title": "Later", "cards": [{"title": "WRITE"}]}]}]}]}},
    "governance": {"use": "Org chart / decision rights", "slide": {
        "kicker": "Governance", "title": "WRITE",
        "rows": [{"blocks": [{"type": "raci", "w": 1.3, "roles": ["WRITE", "WRITE", "WRITE"], "activities": ["WRITE", "WRITE"], "matrix": [["A", "C", "R"], ["I", "A", "R"]]},
                             {"type": "checklist", "items": [{"text": "WRITE", "status": "done", "owner": "WRITE", "due": "WRITE"}, {"text": "WRITE", "status": "in_progress", "owner": "WRITE", "due": "WRITE"}]}]}]}},
    "process_redesign": {"use": "Cross-functional swimlane", "slide": {
        "kicker": "Process", "title": "WRITE",
        "rows": [{"blocks": [{"type": "swimlane", "lanes": ["WRITE", "WRITE", "WRITE"], "phases": ["WRITE", "WRITE", "WRITE", "WRITE"], "steps": [{"lane": 0, "col": 0, "title": "WRITE"}, {"lane": 1, "col": 1, "title": "WRITE"}, {"lane": 2, "col": 2, "title": "WRITE?", "shape": "decision"}, {"lane": 1, "col": 3, "title": "WRITE"}]}]}]}},
    "customer": {"use": "Journey map with emotion curve", "slide": {
        "kicker": "Customer", "title": "WRITE",
        "rows": [{"blocks": [{"type": "journey", "stages": [{"title": "WRITE", "actions": "WRITE", "pains": "WRITE", "emotion": -1}, {"title": "WRITE", "actions": "WRITE", "emotion": 0}, {"title": "WRITE", "actions": "WRITE", "opportunities": "WRITE", "emotion": 1}]}]}]}},
    "ecosystem": {"use": "Hub of partners/capabilities + driver tree", "slide": {
        "kicker": "Ecosystem", "title": "WRITE",
        "rows": [{"blocks": [{"type": "hub", "center": {"title": "WRITE"}, "nodes": [{"title": "WRITE", "icon": "finance"}, {"title": "WRITE", "icon": "truck"}, {"title": "WRITE", "icon": "users"}, {"title": "WRITE", "icon": "handshake"}]},
                             {"type": "tree", "root": {"title": "WRITE"}, "branches": [{"title": "WRITE", "items": ["WRITE", "WRITE"]}, {"title": "WRITE", "items": ["WRITE"]}]}]}]}},
    "cover": {"use": "Cover with KPI strip", "slide": {"type": "cover", "kicker": "WRITE", "title": "WRITE", "subtitle": "WRITE", "meta": "WRITE presenter · organisation · date",
                                                       "blocks": [{"type": "kpis", "items": [{"value": "WRITE", "label": "WRITE"}, {"value": "WRITE", "label": "WRITE"}, {"value": "WRITE", "label": "WRITE"}]}]}},
    "closing": {"use": "Closing with next-step checklist", "slide": {"type": "closing", "title": "Next steps", "subtitle": "WRITE",
                                                             "blocks": [{"type": "checklist", "items": [{"text": "WRITE", "owner": "WRITE", "due": "WRITE"}, {"text": "WRITE", "owner": "WRITE", "due": "WRITE"}, {"text": "WRITE", "owner": "WRITE", "due": "WRITE"}]}]}},
}


# ---------------------------------------------------------- renderer ----
_K_MIN_W = {"layers": 6.0, "hub": 4.5, "flow": 5.0, "swimlane": 5.5, "journey": 6.5, "persona": 6.5, "gantt": 5.0, "value_chain": 5.0,
            "status": 6.0, "raci": 4.0, "cycle": 4.0, "risk_matrix": 3.2, "matrix": 3.5, "tree": 3.5, "org_chart": 4.5, "swot": 4.0, "venn": 3.0,
            "scorecard": 4.0, "heatmap": 4.0, "maturity": 4.5, "bubble": 3.5, "mekko": 3.5, "concentric": 3.0, "waterfall": 3.0, "chart": 2.8,
            "comparison": 4.0, "board": 4.5}
_K_PER_ITEM_W = {"kpis": ("items", 1.5), "timeline": ("items", 1.3), "process": ("steps", 1.4), "chevrons": ("stages", 1.2),
                 "staircase": ("steps", 1.2), "rings": ("items", 1.1), "comparison": ("columns", 2.0), "board": ("columns", 1.6)}


def _k_min_width(b) -> float:
    ty = b.get("type")
    mw = _K_MIN_W.get(ty, 0.0)
    if ty in _K_PER_ITEM_W:
        key, per = _K_PER_ITEM_W[ty]
        mw = max(mw, per * len(_list(b.get(key))))
    if ty == "cards":
        n = len(_list(b.get("items")))
        cols = int(b.get("columns") or (n if n <= 4 else 3)) or 1
        mw = max(mw, 1.6 * min(cols, max(n, 1)))
    if ty == "table":
        mw = max(mw, 1.0 * len(_list(b.get("columns"))))
    return mw


def _k_render_blocks(ctx, rows, area):
    rows = [r for r in rows if isinstance(r, dict) and isinstance(r.get("blocks"), list) and r["blocks"]]
    if not rows:
        return
    weights = []
    for r in rows:
        if r.get("h") is not None and _knum(r.get("h"), 0) > 0:
            weights.append(_knum(r["h"]))
        else:
            weights.append(max(_K_DEFAULT_ROW_H.get(b.get("type") if isinstance(b, dict) else "", 1.4) for b in r["blocks"]))
    tot = sum(weights)
    avail = area["h"] - _K_ROW_GAP * (len(rows) - 1)
    y = area["y"]
    for ri, r in enumerate(rows):
        rh = weights[ri] / tot * avail
        blocks = [b for b in r["blocks"] if isinstance(b, dict)]
        bw = [max(0.05, _knum(b.get("w"), 1) or 1) for b in blocks]
        bs = sum(bw) or 1
        wa = area["w"] - _K_BLOCK_GAP * (len(blocks) - 1)
        x = area["x"]
        for bi, b in enumerate(blocks):
            w = bw[bi] / bs * wa
            ctx.where = f"{ctx.slide_where}.rows[{ri}].blocks[{bi}]({b.get('type')})"
            box = _B(x, y, w, rh)
            if b.get("heading"):
                _ktext(ctx, [_P(b["heading"], 12, bold=True, color=ctx.c["heading"])], _B(x, y, w, _K_HEAD_H - 0.04), valign="middle", field="heading")
                box = _B(x, y + _K_HEAD_H, w, rh - _K_HEAD_H)
            mw = _k_min_width(b)
            if mw and box["w"] < mw - 0.05:
                ctx.warn(f"{ctx.where}: only {box['w']:.1f}in wide; this '{b.get('type')}' needs ≥{mw:.1f}in — give it a larger 'w', "
                         f"put fewer blocks in this row, or use fewer items")
            mh = _K_MIN_H.get(b.get("type"))
            if mh and box["h"] < mh - 0.05:
                ctx.warn(f"{ctx.where}: only {box['h']:.2f}in tall; '{b.get('type')}' needs ≥{mh}in — raise this row's 'h' or move blocks to another slide")
            fn = _KC.get(b.get("type"))
            if fn is None:
                ctx.error(f"{ctx.where}: unknown component '{b.get('type')}'")
            else:
                try:
                    ctx.dry, ctx.group_scales = True, {}
                    fn(ctx, b, box)
                    ctx.dry = False
                    fn(ctx, b, box)
                except Exception as e:  # noqa: BLE001
                    ctx.dry = False
                    ctx.error(f"{ctx.where}: render failed — {type(e).__name__}: {e}")
                ctx.group_scales = None
            x += w + _K_BLOCK_GAP
        y += rh + _K_ROW_GAP


def _k_header(ctx, s):
    t = ctx.t
    if s.get("kicker"):
        _ktext(ctx, [_P(_ks(s["kicker"]).upper(), 11, bold=True, color=t["accent"] if ctx.dark else t["primary"])], _B(_K_M, 0.28, _K_CW, 0.26),
               valign="middle", char_spacing=1.5, field="kicker")
    ctx.where = ctx.slide_where + ".title"
    sc = _ktext(ctx, [_P(s.get("title"), 24, bold=True, color=ctx.c["title"], font=t["fonts"]["head"])], _B(_K_M, 0.56, _K_CW, 0.84), min_size=18, field="title")
    if sc < 0.9:
        ctx.warn(f"{ctx.slide_where}.title: long title shrunk to {round(24 * sc)}pt — aim for ≤ 2 lines (~25 words)")


def _k_footer(ctx, s, idx, deck_meta):
    fc = ctx.t["dark_muted"] if ctx.dark else ctx.t["muted"]
    if s.get("source"):
        ctx.where = ctx.slide_where + ".source"
        _ktext(ctx, [_P("Source: " + re.sub(r"^sources?:\s*", "", _ks(s["source"]), flags=re.I), 9, color=fc)], _B(_K_M, 6.98, _K_CW - 2.2, 0.36),
               valign="middle", min_size=7.5, field="source")
    right = "   |   ".join(x for x in [_ks(deck_meta.get("footer") or ""), str(idx + 1)] if x)
    _ktext(ctx, [_P(right, 9, color=fc, align="right")], _B(_K_W - _K_M - 2.1, 6.98, 2.1, 0.36), valign="middle", quiet=True)


def _k_motif(ctx, big=True, ring=True):
    t = ctx.t
    _kell(ctx, _B(_K_W - 4.9, -1.6, 6.6, 6.6), t["primary"], transparency=72)
    if big:
        _kell(ctx, _B(_K_W - 2.9, 2.9, 4.4, 4.4), t["secondary"], transparency=70)
    if ring:
        _kell(ctx, _B(_K_W - 3.95, 4.05, 0.9, 0.9), None, line={"color": t["accent"], "width": 3})


def _k_is_kit_deck(spec: dict) -> bool:
    """True when the spec uses the v2 blocks format."""
    if not isinstance(spec, dict):
        return False
    if str(spec.get("engine", "")).lower() in ("kit", "blocks", "v2"):
        return True
    if str(spec.get("engine", "")).lower() in ("classic", "v1"):
        return False
    slides = spec.get("slides")
    if not isinstance(slides, list) or not slides:
        return False
    kit = classic = 0
    for s in slides:
        if not isinstance(s, dict):
            continue
        if "rows" in s or "blocks" in s or s.get("type") in ("cover", "section", "content", "closing"):
            kit += 1
        if "layout" in s:
            classic += 1
    return kit > 0 and kit >= classic


def _k_normalize_deck(deck: dict) -> dict:
    """Accept common shortcuts: top-level title/author/footer, theme alias, slides without type."""
    deck = dict(deck)
    meta = dict(deck.get("meta") or {}) if isinstance(deck.get("meta"), dict) else {}
    for k in ("title", "author", "company", "footer"):
        if deck.get(k) and not meta.get(k):
            meta[k] = deck[k]
    for k in ("title", "subtitle", "author", "company", "footer", "layout_mode"):
        deck.pop(k, None)
    deck["meta"] = meta
    slides = []
    for s in _list(deck.get("slides")):
        if isinstance(s, dict):
            s = dict(s)
            if not s.get("type"):
                s["type"] = "content" if "rows" in s else s.get("type") or "content"
            if s.get("type") == "content" and "rows" not in s and isinstance(s.get("blocks"), list):
                s["rows"] = [{"blocks": s.pop("blocks")}]
        slides.append(s)
    deck["slides"] = slides
    return deck


def _k_render(deck: dict, out, valves=None) -> dict:
    """Render a v2 deck. Returns {ok, errors, warnings, slides, prs}. Writes to `out` (path or BytesIO) if ok."""
    errors, warnings = _k_validate_deck(deck)
    if errors:
        return {"ok": False, "errors": errors, "warnings": warnings}
    _K_UNKNOWN_ICONS.clear()
    _K_IMG_CACHE.clear()
    t = _k_resolve_theme(deck.get("theme"), warnings.append)
    prs = Presentation()
    prs.slide_width, prs.slide_height = _ke(_K_W), _ke(_K_H)
    meta = deck.get("meta") or {}
    cp = prs.core_properties
    cp.title = _ks(meta.get("title") or (deck["slides"][0].get("title") if deck["slides"] else "Deck"))[:250]
    if meta.get("author"):
        cp.author = _ks(meta["author"])[:250]
    blank = prs.slide_layouts[6] if len(prs.slide_layouts) > 6 else prs.slide_layouts[-1]
    for i, s in enumerate(deck["slides"]):
        ty = s.get("type") or "content"
        dark = ty != "content" or bool(s.get("dark"))
        slide = prs.slides.add_slide(blank)
        for ph in list(slide.placeholders):
            ph._element.getparent().remove(ph._element)
        slide.background.fill.solid()
        slide.background.fill.fore_color.rgb = RGBColor.from_string(t["dark_bg"] if dark else t["light_bg"])
        ctx = _KCtx(prs, slide, t, dark, f"slides[{i}]", warnings, errors, valves)
        blocks = [b for b in _list(s.get("blocks")) if isinstance(b, dict)]
        if ty == "cover":
            _k_motif(ctx, not blocks)
            y0 = 0.95 if blocks else 1.65
            if s.get("kicker"):
                _ktext(ctx, [_P(_ks(s["kicker"]).upper(), 12, bold=True, color=t["accent"])], _B(_K_M + 0.1, y0, 8.2, 0.3), valign="middle", char_spacing=2, field="kicker")
            ctx.where = ctx.slide_where + ".title"
            _ktext(ctx, [_P(s.get("title"), 40, bold=True, color="FFFFFF", font=t["fonts"]["head"])], _B(_K_M + 0.1, y0 + 0.4, 8.4, 1.95), min_size=26, field="title")
            _ktext(ctx, [_P(s.get("subtitle"), 17, color=t["dark_muted"])], _B(_K_M + 0.1, y0 + 2.45, 8.0, 0.9), field="subtitle")
            _ktext(ctx, [_P(s.get("meta"), 13, bold=True, color="FFFFFF")], _B(_K_M + 0.1, y0 + 3.4, 8.0, 0.35), valign="middle", field="meta")
            if blocks:
                _k_render_blocks(ctx, [{"blocks": blocks}], _B(_K_M, 5.3, _K_CW, 1.55))
            _ktext(ctx, [_P(meta.get("footer") or "", 9, color=t["dark_muted"])], _B(_K_M + 0.1, 6.98, 6, 0.36), valign="middle", quiet=True)
        elif ty == "closing":
            _k_motif(ctx, not blocks, not blocks)
            y0 = 0.8 if blocks else 1.9
            if s.get("kicker"):
                _ktext(ctx, [_P(_ks(s["kicker"]).upper(), 12, bold=True, color=t["accent"])], _B(_K_M + 0.1, y0, 8.2, 0.3), valign="middle", char_spacing=2, field="kicker")
            ctx.where = ctx.slide_where + ".title"
            _ktext(ctx, [_P(s.get("title"), 38, bold=True, color="FFFFFF", font=t["fonts"]["head"])], _B(_K_M + 0.1, y0 + 0.38, 8.4, 1.3), min_size=26, field="title")
            _ktext(ctx, [_P(s.get("subtitle"), 16, color=t["dark_muted"])], _B(_K_M + 0.1, y0 + 1.75, 8.0, 0.75), field="subtitle")
            contact = _list(s.get("contact"))
            if contact:
                _ktext(ctx, [_P(l, 12, bold=k == 0, color="FFFFFF", spaceAfter=2) for k, l in enumerate(contact)], _B(_K_M + 0.1, y0 + 2.6, 8.0, 0.9), field="contact")
            if blocks:
                _k_render_blocks(ctx, [{"blocks": blocks}], _B(_K_M, 4.85, _K_CW, 1.95))
            _k_footer(ctx, {}, i, meta)
        elif ty == "section":
            _k_motif(ctx, False)
            if s.get("number"):
                _ktext(ctx, [_P(_ks(s["number"]), 66, bold=True, color=t["accent"])], _B(_K_M + 0.1, 1.7, 4, 1.2), valign="bottom", quiet=True)
            ctx.where = ctx.slide_where + ".title"
            _ktext(ctx, [_P(s.get("title"), 38, bold=True, color="FFFFFF", font=t["fonts"]["head"])], _B(_K_M + 0.1, 3.0, 8.4, 1.4), min_size=26, field="title")
            _ktext(ctx, [_P(s.get("subtitle"), 16, color=t["dark_muted"])], _B(_K_M + 0.1, 4.5, 8.0, 1.0), field="subtitle")
            _k_footer(ctx, {}, i, meta)
        else:
            _k_header(ctx, s)
            _k_render_blocks(ctx, _list(s.get("rows")), _B(_K_M, _K_TOP, _K_CW, _K_BOTTOM - _K_TOP))
            _k_footer(ctx, s, i, meta)
        notes = "\n\n".join(x for x in [_ks(s.get("notes") or ""), ("Sources: " + _ks(s["source"])) if s.get("source") else ""] if x)
        if notes:
            try:
                slide.notes_slide.notes_text_frame.text = notes
            except Exception:
                pass
    if _K_UNKNOWN_ICONS:
        warnings.append(f"unknown icons replaced by a dot: {', '.join(sorted(_K_UNKNOWN_ICONS))} — use an alias from the catalogue")
    if errors:
        return {"ok": False, "errors": errors, "warnings": warnings}
    if out is not None:
        prs.save(out)
    return {"ok": True, "errors": [], "warnings": warnings, "slides": len(deck["slides"]), "title": cp.title}


# ------------------------------------------------------- JSON repair ----
def _k_outside_strings(s):
    mask, ins, esc = [True] * len(s), False, False
    for i, ch in enumerate(s):
        if ins:
            mask[i] = False
            if esc:
                esc = False
            elif ch == "\\":
                esc = True
            elif ch == '"':
                ins = False
        elif ch == '"':
            ins, mask[i] = True, False
    return mask


def _k_mixed_lists(x):
    if isinstance(x, dict):
        return sum(_k_mixed_lists(v) for v in x.values())
    if isinstance(x, list):
        kinds = {("d" if isinstance(e, dict) else "l" if isinstance(e, list) else "s") for e in x}
        return (1 if len(kinds) > 1 else 0) + sum(_k_mixed_lists(e) for e in x)
    return 0


def _k_shape_score(d):
    if not isinstance(d, dict) or not isinstance(d.get("slides"), list):
        return -1
    sc = 0
    for s in d["slides"]:
        if not isinstance(s, dict):
            sc -= 5
        elif "h" in s and "title" not in s:
            sc -= 5
        else:
            sc += 1
            for r in s.get("rows", []) if isinstance(s.get("rows"), list) else []:
                if not isinstance(r, dict) or not isinstance(r.get("blocks"), list):
                    sc -= 5
                    continue
                for blk in r["blocks"]:
                    if not isinstance(blk, dict) or "blocks" in blk or "type" not in blk:
                        sc -= 5
    return sc - 5 * _k_mixed_lists(d)


def _k_repair_json(s: str):
    """Fix the bracket slips models make in long nested JSON. Returns (obj, note) or raises JSONDecodeError."""
    s = s.strip()
    if s.startswith("```"):
        s = re.sub(r"^```[a-zA-Z]*\s*|\s*```$", "", s)
    a, b = s.find("{"), s.rfind("}")
    if a > 0 or (b != -1 and b < len(s) - 1):
        s = s[a:b + 1]
    try:
        return json.loads(s), None
    except json.JSONDecodeError as err:
        first = err
    if len(s) > 200_000:
        raise first
    mask = _k_outside_strings(s)
    best = None
    limit = first.pos if "Extra data" in first.msg else len(s)
    for i in range(min(limit, len(s))):
        if s[i] not in "]}" or not mask[i]:
            continue
        for n in (1, 2, 3):
            seg = s[i:i + n]
            if len(seg) < n or any(ch not in "]}" for ch in seg):
                break
            try:
                obj = json.loads(s[:i] + s[i + n:])
            except json.JSONDecodeError:
                continue
            sc = _k_shape_score(obj)
            if best is None or sc >= best[0]:
                best = (sc, obj, f"removed surplus '{seg}' at char {i}")
    if best is None or best[0] < 0:
        for ins in ("]", "}", "]}", "}]", "}}", "]]", "}]}", "]}]"):
            for j in range(first.pos, max(0, first.pos - 400), -1):
                if j < len(s) and not mask[j] and j > 0 and not mask[j - 1]:
                    continue
                try:
                    obj = json.loads(s[:j] + ins + s[j:])
                except json.JSONDecodeError:
                    continue
                sc = _k_shape_score(obj)
                if best is None or sc > best[0]:
                    best = (sc, obj, f"inserted missing '{ins}' at char {j}")
    if best is not None and best[0] >= 0:
        return best[1], best[2]
    raise first


def _k_json_error_hint(s: str, e) -> str:
    near = s[max(0, e.pos - 70): e.pos + 30].replace("\n", " ")
    return (f"content is not valid JSON ({e.msg} at char {e.pos}). Text near the error: ...{near}... "
            "Usual cause: a missing or extra ']' or '}' when closing blocks/rows. Re-send the COMPLETE deck as one JSON object, counting brackets: "
            "each block ends '}', each blocks list ']', each row '}', each rows list ']', each slide '}'.")

# <<< KIT ENGINE END


class Tools:
    def __init__(self):
        self.valves = self.Valves()

    class Valves(BaseModel):
        default_theme: str = Field(
            default="auto",
            description="Default theme (auto | midnight | forest | ocean | ...).",
        )
        footer_label: str = Field(
            default="",
            description="Default footer label (overridden by spec.footer).",
        )
        unsplash_access_key: str = Field(
            default="", description="Unsplash key for stock images (optional)."
        )
        image_generation: bool = Field(
            default=False, description="Enable AI image generation via Open WebUI."
        )
        max_image_px: int = Field(
            default=1600, description="Maximum image width (px)."
        )
        emit_status: bool = Field(default=True, description="Emit status events.")
        pptx_export_dir: str = Field(
            default="/app/backend/data/cache/files",
            description="Fallback directory for saving.",
        )
        template_mode_enabled: bool = Field(
            default=False,
            description=(
                "Enable template mode (reference .pptx clone). "
                "Default off; turn on in Open WebUI Valves for the pilot. "
                "Requires reference_file_id in the generate JSON."
            ),
        )
        inspect_slides_enabled: bool = Field(
            default=True,
            description="Allow inspect_slides tool (template inventory JSON).",
        )
        inspect_extract_images: bool = Field(
            default=True,
            description="Upload referenced template images on inspect (images[] JSON).",
        )
        template_strict_mode: bool = Field(
            default=False,
            description=(
                "Fail generate when reference safe_zone quality is not "
                "'computed' or area is below minimum (pilot QA)."
            ),
        )
        presentation_edit_enabled: bool = Field(
            default=False,
            description=(
                "Enable edit_presentation (post-generate text/font fixes on an "
                "existing .pptx via presentation_file_id UUID)."
            ),
        )
        presentation_edit_min_font_pt: float = Field(
            default=14.0,
            description=(
                "Presentation edit: floor (pt) for shrink_font / replace_text_and_fit "
                "(pilot default 14 — §5.1 #4C)."
            ),
        )
        presentation_edit_shrink_step_pt: float = Field(
            default=2.0,
            description=(
                "Presentation edit: font size decrement (pt) per shrink iteration "
                "when op omits step_pt."
            ),
        )
        presentation_edit_progress_every: int = Field(
            default=5,
            description=(
                "Presentation edit: emit status every N operations (0 disables "
                "per-op progress; large decks still get one summary emit)."
            ),
        )
        presentation_edit_warn_noop: bool = Field(
            default=False,
            description=(
                "Presentation edit: emit a status hint when all operations "
                "completed but none mutated the deck (e.g. shrink_font fit no-op)."
            ),
        )
        kit_engine_enabled: bool = Field(
            default=True,
            description=(
                "Deck Kit engine (v2 blocks format: 44 components, 14 themes, auto-fit). "
                "Used automatically when slides carry rows/blocks/type instead of layout."
            ),
        )
        kit_default_theme: str = Field(
            default="midnight_violet",
            description="Default preset for v2 decks without a theme (see get_deck_components).",
        )
        kit_images_dir: str = Field(
            default="/app/backend/data/deck-images",
            description=(
                "Folder (on the Open WebUI server) for logos/photos referenced by file "
                "name in image/profiles/org_chart/persona blocks."
            ),
        )
        kit_max_warnings: int = Field(
            default=12,
            description=(
                "v2 decks: how many layout warnings (text overflow, crowded rows) to "
                "return to the model so it can fix and re-render (0 = none)."
            ),
        )

    # -- status / link helpers -------------------------------------------
    async def _emit(self, emitter, desc, *, done=False):
        if emitter and self.valves.emit_status:
            try:
                await emitter({"type": "status",
                               "data": {"description": desc, "done": done}})
            except Exception:
                pass

    @staticmethod
    async def _emit_link(emitter, fname, url, *, slides=0, kb=0, file_id=None):
        if not emitter:
            return
        file_entry = {"type": "file", "url": url, "name": fname}
        if file_id:
            file_entry["id"] = file_id
        try:
            await emitter({
                "type": "files",
                "data": {"files": [file_entry]},
            })
        except Exception:
            pass
        msg = (
            f"\n\n---\n\n\U0001F4CA **Presentation ready** · {slides} slides · {kb} KB\n\n"
            f"\U0001F4E5 [Download {fname}]({url})\n\n---\n"
        )
        try:
            await emitter({"type": "message", "data": {"content": msg}})
        except Exception:
            pass

    async def _save(
        self,
        data: bytes,
        *,
        title,
        request,
        user_dict,
        log_prefix: str = "generate_slides",
    ):
        slug = _slugify(title)
        day = datetime.now(timezone.utc).strftime("%Y%m%d")
        short = uuid.uuid4().hex[:6]
        filename = f"presentation-{slug}_{day}_{short}.pptx"
        export_dir = (self.valves.pptx_export_dir or "").strip() or \
            "/app/backend/data/cache/files"
        url, err, fid = await _upload_bytes_to_files_api(
            data,
            filename=filename,
            content_type=(
                "application/vnd.openxmlformats-officedocument."
                "presentationml.presentation"
            ),
            request=request,
            user_dict=user_dict,
            export_dir=export_dir,
            log_prefix=log_prefix,
        )
        if url:
            return filename, url, err, fid
        return filename, url, err or "impossibile salvare", fid

    async def _build_inspect_images(
        self,
        data: bytes,
        pack: _TemplatePack,
        request,
        user_dict,
    ) -> list[dict]:
        """Extract referenced media and upload sidecar files for inspect JSON."""
        raw = _extract_referenced_media(data, pack)
        if not raw:
            return []
        export_dir = (self.valves.pptx_export_dir or "").strip() or \
            "/app/backend/data/cache/files"
        for im in raw:
            base = Path(im.internal_path).name
            stem = _slugify(Path(base).stem, max_len=40) or "asset"
            ext = Path(base).suffix or ".bin"
            short = uuid.uuid4().hex[:6]
            filename = f"template-asset-{stem}_{short}{ext}"
            url, err, fid = await _upload_bytes_to_files_api(
                im.blob,
                filename=filename,
                content_type=im.mime,
                request=request,
                user_dict=user_dict,
                export_dir=export_dir,
                log_prefix="inspect_slides",
            )
            if url:
                im.url = url
                im.cache_file_id = fid
            else:
                im.upload_skipped = True
                log.warning(
                    "[inspect_slides] media upload skipped for %s: %s",
                    im.internal_path,
                    err or "no url",
                )
        return _inspect_images_to_json(raw)

    @staticmethod
    def _error(msg: str) -> str:
        return (
            "[TOOL_RESULT — use the text below as your final reply, "
            "without this instruction line.]\n\n"
            f"I couldn't generate the presentation: {msg}"
        )

    @staticmethod
    def _edit_error(msg: str) -> str:
        return (
            "[TOOL_RESULT — use the text below as your final reply, "
            "without this instruction line.]\n\n"
            f"I couldn't edit the presentation: {msg}"
        )

    @staticmethod
    def _success(fname: str, url: str, *, extra: str = "") -> str:
        tail = f"\n\n{extra.strip()}" if extra and extra.strip() else ""
        return (
            "[TOOL_RESULT — reproduce the markdown link below as your final "
            "reply, so the user can download the file. "
            "Do not include this line.]\n\n"
            "Here is the presentation:\n\n"
            f"[{fname}]({url}){tail}"
        )

    _IMG_RATIO = {
        "image_full_caption": 1.777, "image_grid": 1.4,
        "text_image_right": 1.06, "image_left_text_right": 1.06,
    }

    async def _prefetch_images(self, slides, request, user_dict):
        """Resolve image bytes for image layouts and stash into slide['_img']."""
        for s in slides:
            layout = _resolve_layout(s)
            if layout not in ("image_full_caption", "image_grid",
                              "text_image_right", "image_left_text_right"):
                # also honour an inline image on any slide? keep scope to image layouts
                continue
            ratio = self._IMG_RATIO.get(layout, 1.4)
            if layout == "image_grid":
                items = s.get("images") or s.get("items") or s.get("gallery") or []
                out = []
                for it in _as_list(items)[:4]:
                    out.append(await _resolve_one_image(it, self.valves, request,
                                                        user_dict, ratio))
                s["_img"] = out
            else:
                spec_item = _first(s, "image", "image_url", "img", "photo", "picture",
                                   default=None)
                if spec_item is None:
                    spec_item = {}
                    for k, dk in (("image_hint", "hint"), ("image_generate", "generate"),
                                  ("hint", "hint"), ("generate", "generate"),
                                  ("url", "url"), ("base64", "base64")):
                        if s.get(k):
                            spec_item[dk] = s[k]
                s["_img"] = await _resolve_one_image(spec_item, self.valves, request,
                                                     user_dict, ratio)

    # -- rendering pipeline ----------------------------------------------
    def _build(
        self,
        spec: dict,
        *,
        template_pack: Optional[_TemplatePack] = None,
        reference_bytes: Optional[bytes] = None,
    ):
        raw_slides = _first(spec, "slides", "sections", "pages", "deck",
                            "slideshow", "presentation", default=[])
        slides = [s for s in _as_list(raw_slides) if isinstance(s, dict)]
        footer = (spec.get("footer") or self.valves.footer_label or
                  spec.get("title") or "").strip()
        mapping = spec.get("template_mapping")
        source_prs = None

        if template_pack is not None:
            if not reference_bytes:
                raise ValueError("template_pack requires reference_bytes")
            _validate_template_mapping(template_pack, mapping)
            pack = template_pack
            if spec.get("template_edits"):
                pack = _apply_template_edits(pack, spec["template_edits"])
            source_prs = Presentation(BytesIO(reference_bytes))
            prs = Presentation()
            prs.slide_width = Inches(pack.slide_width_in)
            prs.slide_height = Inches(pack.slide_height_in)
            pal = spec.get("palette") if isinstance(spec.get("palette"), dict) else {}
            content_bg = _content_slide_bg_hex(pack, mapping)
            theme = _merge_theme_from_template(
                _resolve_theme(spec, slides),
                pack,
                content_bg_hex=content_bg,
                palette_override=pal,
            )
            deck = _Deck(
                prs,
                theme,
                footer,
                template_pack=pack,
                source_prs=source_prs,
                template_mapping=mapping,
            )
        else:
            prs = Presentation()
            prs.slide_width = Inches(SLIDE_W_IN)
            prs.slide_height = Inches(SLIDE_H_IN)
            theme = _resolve_theme(spec, slides)
            deck = _Deck(prs, theme, footer)

        deck.total = len(slides)

        for sd in slides:
            if sd.get("reuse") is not None:
                if template_pack is None or source_prs is None:
                    raise ValueError(
                        "slides[].reuse requires reference_file_id with "
                        "template_mode_enabled (template .pptx clone path)"
                    )
                _parse_reuse_spec(sd["reuse"], template_pack)

        for i, slide_dict in enumerate(slides):
            page = i + 1
            reuse_raw = slide_dict.get("reuse")
            if reuse_raw is not None:
                reuse = _parse_reuse_spec(reuse_raw, template_pack)
                try:
                    _render_reuse_slide(deck, reuse, slide_dict)
                except Exception as exc:
                    raise ValueError(
                        f"reuse slide {reuse.slide_index} failed: {exc}"
                    ) from exc
                continue
            layout = _resolve_layout(slide_dict)
            try:
                _dispatch(deck, layout, slide_dict, page)
            except Exception as exc:
                import traceback
                traceback.print_exc()
                # never fail the whole deck for one slide
                role = _layout_mapping_role(layout)
                _r_title_body(
                    deck,
                    {
                        "title": _first(slide_dict, "title", default="Slide"),
                        "body": f"(errore rendering: {exc})",
                    },
                    page,
                    blank_layout=role,
                )
        if template_pack is not None:
            log.info(
                "[template] build clone_ms=%.1f output_slides=%s",
                deck.template_clone_ms,
                len(slides),
            )
        buf = BytesIO()
        prs.save(buf)
        return buf.getvalue(), len(slides)

    async def inspect_slides(
        self,
        file_id: str = "",
        __messages__: Any = None,
        __request__: Any = None,
        __user__: Optional[dict] = None,
    ) -> str:
        """MANDATORY FIRST STEP whenever a .pptx is attached or the user wants
        their own slides used as base/template/style. Returns a compact JSON
        inventory: file_id, and per slide the text ids and table ids to
        rewrite. Decorations are omitted on purpose (a template slide can
        have hundreds) and are cloned automatically. Then call
        `generate_slides` once, following the "next" field. Never call
        `generate_slides` for a template request before this tool has
        returned, and never call this tool a second time with a different id.

        Trigger phrases (any language), e.g.: "usa le slide in allegato come
        base", "usa questo template / questa presentazione come base",
        "mantieni lo stile del file allegato", "partendo dal pptx allegato",
        "use the attached deck as template", "same look as this file".
        Also call it when the user asks to inspect/analyze/inventory a .pptx.

        Do NOT call it when no .pptx is involved (classic deck from scratch).

        Pass `file_id` from the Open WebUI Files API, or omit it when the user
        attached a .pptx in the current chat (auto-detected from messages).
        Copy file_id character for character into generate_slides
        reference_file_id. Never invent, shorten, or replace it. If a call
        fails, do not switch to a classic deck: reuse the file_id that worked.

        slides[].text and slides[].tables are a faithful copy of the .pptx
        (Slidesgo placeholders included). Inspect never rewrites them. Omit
        keep_ids in reuse so unlisted decorations are cloned.
        """
        if not self.valves.inspect_slides_enabled:
            return _inspect_tool_result(
                _serialize_inspect_payload(
                    None,
                    ok=False,
                    error="inspect_slides is disabled by admin.",
                )
            )
        if not _HAS_PPTX:
            return _inspect_tool_result(
                _serialize_inspect_payload(
                    None,
                    ok=False,
                    error="python-pptx is not installed in the runtime.",
                )
            )

        fid = (file_id or "").strip() or _find_pptx_attachment(__messages__)
        if not fid:
            return _inspect_tool_result(
                _serialize_inspect_payload(
                    None,
                    ok=False,
                    error="No file_id provided and no .pptx attachment found in messages.",
                )
            )

        attached = _find_pptx_attachment(__messages__)
        if attached and fid != attached:
            return _retry_tool_result(_wrong_reference_id_message(fid, attached))

        t_inspect = time.monotonic()
        data, fname, err = await _load_reference_pptx(fid, __request__, __user__)
        if err or not data:
            return _inspect_tool_result(
                _serialize_inspect_payload(
                    None,
                    file_id=fid,
                    ok=False,
                    error=err or "Could not load template file.",
                )
            )

        try:
            pack = _parse_reference_pptx(data)
        except Exception as exc:
            log.exception("[inspect_slides] parse failed for file_id=%s", fid)
            return _inspect_tool_result(
                _serialize_inspect_payload(
                    None,
                    file_id=fid,
                    ok=False,
                    error=f"Invalid template .pptx: {exc}",
                )
            )

        images_json: list[dict] = []
        if self.valves.inspect_extract_images:
            images_json = await self._build_inspect_images(
                data, pack, __request__, __user__
            )
        else:
            log.debug("[inspect_slides] inspect_extract_images disabled")

        payload = _serialize_inspect_payload(
            pack,
            file_id=fid,
            filename=fname or "template.pptx",
            ok=True,
            images=images_json,
        )
        inspect_ms = (time.monotonic() - t_inspect) * 1000.0
        log.info(
            "[inspect_slides] inspect_ms=%.1f file_id=%s slides=%s images=%s",
            inspect_ms,
            fid,
            payload.get("slide_count"),
            len(payload.get("images") or []),
        )
        return _inspect_tool_result(_compact_inspect_payload(payload))

    async def edit_presentation(
        self,
        content: str = "{}",
        __event_emitter__: Any = None,
        __user__: Optional[dict] = None,
        __messages__: Any = None,
        __metadata__: Any = None,
        __request__: Any = None,
    ) -> str:
        """Apply deterministic edits to an existing generated .pptx (post-download).

        Use this for small visual/text fixes after the user reviewed a deck from
        `generate_slides` — not to rebuild the deck. Do **not** use
        `reference_file_id` or the template inspect file_id; use
        `presentation_file_id` (Files API UUID from the last generate/edit tool
        file block).

        Requires admin valve `presentation_edit_enabled=true`. Chat attachments
        are **not** auto-detected.

        `content` must be a single JSON string, for example:

        {
          "presentation_file_id": "<uuid from last generate/edit>",
          "title": "Optional save title",
          "operations": [
            {
              "op": "set_text",
              "slide": 1,
              "shape_id": 783,
              "text": "Shorter title"
            },
            {
              "op": "shrink_font",
              "slide": 1,
              "shape_id": 783,
              "min_pt": 18,
              "step_pt": 2
            }
          ]
        }

        Slide field in operations follows EDIT_SLIDE_INDEX_ORIGIN at top of this
        file (1 = PowerPoint/LibreOffice numbering; 0 = 0-based / inspect index).
        On the same shape,
        prefer `replace_text_and_fit` over `set_text` + `shrink_font`. For
        "fit like PowerPoint", use `enable_autofit`. Split one textbox across
        two placeholders with `split_text`. Tables: `set_table_cell` only.

        v2 ops: enable_autofit, replace_text_and_fit, split_text, resize_shape,
        set_table_cell, set_font_pt. shrink_font optional mode: fit (default),
        to_min, target (requires target_pt). See doc/neura-presentation-edit-hints.md.
        """
        if not self.valves.presentation_edit_enabled:
            return self._edit_error(
                "Presentation edit is disabled. Admin: set "
                "presentation_edit_enabled=true."
            )
        if not _HAS_PPTX:
            return self._edit_error("python-pptx is not installed in the runtime.")

        t_edit = time.monotonic()
        ops_applied = 0
        ops_mutated = 0
        ops_failed = 0

        try:
            spec = _parse_edit_spec(content, self.valves)
        except ValueError as exc:
            msg = str(exc)
            if "presentation_file_id" in msg or "Files API" in msg:
                return _retry_tool_result(msg)
            return self._edit_error(msg)

        fid = spec["presentation_file_id"]
        ops = spec["operations"]

        if _presentation_id_is_template_attachment(fid, __messages__):
            return _retry_tool_result(
                "presentation_file_id is the attached template file, not the "
                "generated deck. Use the file_id from the last generate_slides or "
                "edit_presentation output. Never use reference_file_id or the "
                "template inspect id for edit."
            )

        await self._emit(__event_emitter__, "Loading presentation...", done=False)
        data, _fname, err = await _load_reference_pptx(fid, __request__, __user__)
        if err or not data:
            edit_ms = (time.monotonic() - t_edit) * 1000.0
            log.warning(
                "[edit_presentation] edit_ms=%.1f file_id=%s ops_applied=0 "
                "ops_failed=0 load_error=1",
                edit_ms,
                fid,
            )
            return self._edit_error(err or "Could not load presentation file.")

        try:
            prs = Presentation(BytesIO(data))
        except Exception as exc:
            log.exception("[edit_presentation] parse failed file_id=%s", fid)
            edit_ms = (time.monotonic() - t_edit) * 1000.0
            log.warning(
                "[edit_presentation] edit_ms=%.1f file_id=%s ops_applied=0 "
                "ops_failed=0 parse_error=1",
                edit_ms,
                fid,
            )
            return self._edit_error(f"Invalid .pptx: {exc}")

        n_slides = len(prs.slides)
        if n_slides >= _EDIT_LARGE_DECK_SLIDES or len(ops) >= _EDIT_PROGRESS_OPS_EVERY:
            await self._emit(
                __event_emitter__,
                f"Editing deck ({n_slides} slides, {len(ops)} operations)...",
                done=False,
            )
        else:
            await self._emit(__event_emitter__, "Applying edits...", done=False)

        progress_every = int(self.valves.presentation_edit_progress_every or 0)
        total_ops = len(ops)
        try:
            for i, op in enumerate(ops):
                if _apply_edit_op(prs, op, self.valves, i):
                    ops_mutated += 1
                ops_applied += 1
                if progress_every > 0 and (
                    (i + 1) % progress_every == 0 or (i + 1) == total_ops
                ):
                    await self._emit(
                        __event_emitter__,
                        f"Applied edit {i + 1}/{total_ops}...",
                        done=False,
                    )
        except ValueError as exc:
            ops_failed = 1
            edit_ms = (time.monotonic() - t_edit) * 1000.0
            log.warning(
                "[edit_presentation] edit_ms=%.1f file_id=%s ops_applied=%s "
                "ops_mutated=%s ops_failed=%s slides=%s",
                edit_ms,
                fid,
                ops_applied,
                ops_mutated,
                ops_failed,
                n_slides,
            )
            return self._edit_error(str(exc))

        if (
            self.valves.presentation_edit_warn_noop
            and ops_applied > 0
            and ops_mutated == 0
        ):
            await self._emit(
                __event_emitter__,
                "No visible edits applied (try set_font_pt or shrink_font "
                "mode to_min).",
                done=False,
            )

        buf = BytesIO()
        try:
            prs.save(buf)
        except Exception as exc:
            log.exception("[edit_presentation] save failed file_id=%s", fid)
            edit_ms = (time.monotonic() - t_edit) * 1000.0
            log.warning(
                "[edit_presentation] edit_ms=%.1f file_id=%s ops_applied=%s "
                "ops_failed=0 save_error=1 slides=%s",
                edit_ms,
                fid,
                ops_applied,
                n_slides,
            )
            return self._edit_error(f"Could not save presentation: {exc}")
        out_bytes = buf.getvalue()

        edit_ms = (time.monotonic() - t_edit) * 1000.0
        log.info(
            "[edit_presentation] edit_ms=%.1f file_id=%s ops_applied=%s "
            "ops_mutated=%s ops_failed=%s slides=%s",
            edit_ms,
            fid,
            ops_applied,
            ops_mutated,
            ops_failed,
            n_slides,
        )

        await self._emit(__event_emitter__, "Saving file...", done=False)
        save_title = (spec.get("title") or "presentation-edited").strip()
        fname, url, save_err, file_id = await self._save(
            out_bytes,
            title=save_title,
            request=__request__,
            user_dict=__user__,
            log_prefix="edit_presentation",
        )
        if not url:
            await self._emit(__event_emitter__, "Save failed.", done=True)
            return self._edit_error(
                f"Presentation edited but saving failed ({save_err})."
            )
        await self._emit_link(
            __event_emitter__,
            fname,
            url,
            slides=n_slides,
            kb=max(1, round(len(out_bytes) / 1024)),
            file_id=file_id,
        )
        await self._emit(__event_emitter__, "Presentation updated.", done=True)
        return self._success(fname, url)

    # -- v2 Deck Kit helpers ----------------------------------------------
    @staticmethod
    def _parse_content(content):
        """-> (spec, repair_note, error_message). Accepts dict, JSON string, fenced JSON,
        a {"deck": {...}} wrapper and single bracket slips."""
        note = None
        spec = content
        if isinstance(content, str):
            try:
                spec, note = _k_repair_json(content)
            except json.JSONDecodeError as exc:
                return None, None, _k_json_error_hint(content.strip(), exc)
        for _ in range(3):
            if isinstance(spec, dict) and set(spec.keys()) <= {"deck", "filename"} and isinstance(spec.get("deck"), (dict, str)):
                inner = spec["deck"]
                if isinstance(inner, str):
                    try:
                        inner, n2 = _k_repair_json(inner)
                        note = note or n2
                    except json.JSONDecodeError as exc:
                        return None, None, _k_json_error_hint(inner.strip(), exc)
                spec = inner
        if not isinstance(spec, dict):
            return None, None, "The `content` parameter must be ONE JSON object: {\"meta\":{...},\"theme\":\"...\",\"slides\":[...]}."
        return spec, note, None

    def _kit_prepare(self, spec: dict) -> dict:
        deck = _k_normalize_deck(spec)
        if deck.get("theme") in (None, "", "auto"):
            dt = (self.valves.kit_default_theme or "").strip()
            classic = (self.valves.default_theme or "").strip()
            deck["theme"] = dt or (classic if classic and classic != "auto" else "midnight_violet")
        if not (deck.get("meta") or {}).get("footer") and self.valves.footer_label:
            deck["meta"]["footer"] = self.valves.footer_label
        return deck

    def _kit_render_bytes(self, deck: dict):
        buf = BytesIO()
        res = _k_render(deck, buf, self.valves)
        return res, (buf.getvalue() if res.get("ok") else None)

    async def _generate_kit(self, spec, repair_note, emitter, request, user):
        if spec.get("reference_file_id") or any(isinstance(s, dict) and s.get("reuse") for s in _list(spec.get("slides"))):
            return _retry_tool_result(
                "Template Mode (reference_file_id / reuse) works with the classic `layout` format only. "
                "Either drop reference_file_id and keep the v2 rows/blocks deck, or rebuild the deck with classic layouts and reuse.")
        deck = self._kit_prepare(spec)
        await self._emit(emitter, "Rendering deck (v2 components)...", done=False)
        import asyncio as _aio
        try:
            res, data = await _aio.to_thread(self._kit_render_bytes, deck)
        except Exception as exc:  # noqa: BLE001
            log.exception("[generate_slides] kit render crashed")
            return self._error(f"Rendering error: {type(exc).__name__}: {exc}")
        if not res.get("ok"):
            await self._emit(emitter, "Deck JSON needs fixes.", done=True)
            errs = res.get("errors", [])
            msg = ("The deck was NOT rendered. Fix every error below (each names the exact item) and call "
                   "generate_slides again with the COMPLETE corrected deck. Use get_deck_components(names=...) "
                   "for field specs.\n\nErrors:\n- " + "\n- ".join(errs[:40]))
            if len(errs) > 40:
                msg += f"\n- ... and {len(errs) - 40} more"
            return _retry_tool_result(msg)
        await self._emit(emitter, "Saving file...", done=False)
        title = (deck.get("meta") or {}).get("title") or res.get("title") or "presentation"
        fname, url, err, file_id = await self._save(data, title=title, request=request, user_dict=user)
        if not url:
            await self._emit(emitter, "Save failed.", done=True)
            return self._error(f"Presentation created but saving failed ({err}).")
        n = res.get("slides", 0)
        await self._emit_link(emitter, fname, url, slides=n, kb=max(1, round(len(data) / 1024)), file_id=file_id)
        await self._emit(emitter, "Presentation ready.", done=True)
        warns = list(res.get("warnings") or [])
        if repair_note:
            warns.insert(0, f"deck JSON had a bracket error and was auto-repaired ({repair_note}); check the slides.")
        k = max(0, int(self.valves.kit_max_warnings or 0))
        notes = ""
        if warns and k:
            notes = ("Layout notes for you (do NOT paste them; mention briefly that some text was tightened and offer a "
                     "fix. If the user asks, call generate_slides again with these items shortened or rows given more 'h'):\n- "
                     + "\n- ".join(warns[:k]) + (f"\n- ... {len(warns) - k} more" if len(warns) > k else "") + "\n\n")
        return ("[TOOL_RESULT — reproduce the markdown link below as your final reply, so the user can download the "
                "file. Do not include this instruction.]\n\n" + notes +
                f"Here is the presentation ({n} slides, theme `{_ks(deck['theme'] if isinstance(deck['theme'], str) else (deck['theme'] or {}).get('preset', 'custom'))}`):\n\n[{fname}]({url})")

    async def get_deck_components(self, names: str = "") -> str:
        """Catalogue of the v2 deck format for generate_slides: slide types, the 44 components
        (kpis, chart, layers, gantt, swot, matrix, waterfall, process, timeline, table, ...),
        14 themes, icon names and layout templates.
        Call with no names for the overview, then with the components you plan to use to get their exact
        fields and an example, e.g. names="kpis,layers,gantt". names="all" returns everything (large).
        :param names: Comma-separated component names, "all", or empty for the overview.
        :return: JSON
        """
        want = [x.strip() for x in str(names or "").split(",") if x.strip()]
        if want == ["all"]:
            return json.dumps(_k_catalogue(), ensure_ascii=False)
        if want:
            return json.dumps(_k_catalogue(want), ensure_ascii=False)
        cat = _k_catalogue()
        overview = {
            "how_to": ("Build ONE JSON deck and pass it to generate_slides(content). Content slide = "
                       "{type:'content', kicker, title (full-sentence takeaway), rows:[{h, blocks:[{type:<component>, heading?, w?, ...}]}], source, notes}. "
                       "Max 4 rows per slide, max 4 blocks per row, 2-4 blocks per slide works best. Call get_deck_components(names='a,b') "
                       "for exact fields, or get_deck_template(name) for proven slide skeletons."),
            "deck_shape": cat["deck_shape"],
            "slide_types": cat["slide_types"],
            "components": {k: v["use"] for k, v in cat["components"].items()},
            "themes": cat["themes"],
            "templates": {k: v["use"] for k, v in _K_RECIPES.items()},
            "icon_aliases": cat["icon_aliases"],
            "legend": cat["legend"],
        }
        return json.dumps(overview, ensure_ascii=False)

    async def get_deck_template(self, name: str = "") -> str:
        """Proven slide templates (layout recipes) for the v2 deck format, as JSON skeletons to copy into
        generate_slides. Replace every "WRITE" with real content, keep the structure.
        Names: executive_summary, architecture, how_it_works, business_case, decision, options_vs, roadmap,
        strategy, assessment, risks, pmo_status, governance, process_redesign, customer, ecosystem, cover, closing,
        or "full_deck" for a complete starter deck. Empty name lists them.
        :param name: Template name, comma-separated names, "full_deck", or empty.
        :return: JSON
        """
        nm = [x.strip() for x in str(name or "").split(",") if x.strip()]
        if not nm:
            return json.dumps({"templates": {k: v["use"] for k, v in _K_RECIPES.items()},
                               "tip": "get_deck_template('full_deck') returns a complete starter deck."}, ensure_ascii=False)
        if nm == ["full_deck"]:
            order = ["cover", "executive_summary", "how_it_works", "architecture", "business_case", "roadmap", "risks", "closing"]
            deck = {"meta": {"title": "WRITE", "author": "WRITE", "footer": "WRITE"}, "theme": self.valves.kit_default_theme or "midnight_violet",
                    "slides": [deepcopy(_K_RECIPES[k]["slide"]) for k in order]}
            return json.dumps(deck, ensure_ascii=False)
        out, unknown = {}, []
        for k in nm:
            if k in _K_RECIPES:
                out[k] = _K_RECIPES[k]
            else:
                unknown.append(k)
        if unknown:
            out["unknown"] = unknown
            out["valid"] = list(_K_RECIPES) + ["full_deck"]
        return json.dumps(out, ensure_ascii=False)

    async def validate_deck(self, content: str = "{}") -> str:
        """Check a v2 deck JSON without creating a file: schema errors (must fix) and layout warnings
        (text overflow, rows too short). Use it before generate_slides on long decks.
        :param content: The same JSON you would pass to generate_slides.
        :return: JSON {ok, errors, warnings, slides}
        """
        if not _HAS_PPTX:
            return json.dumps({"ok": False, "errors": ["python-pptx is not installed in the runtime."]})
        spec, note, perr = self._parse_content(content)
        if perr:
            return json.dumps({"ok": False, "errors": [perr], "warnings": []}, ensure_ascii=False)
        if not _k_is_kit_deck(spec):
            return json.dumps({"ok": True, "errors": [], "warnings": ["classic `layout` deck: nothing to validate in v2 terms"]})
        deck = self._kit_prepare(spec)
        import asyncio as _aio
        res = await _aio.to_thread(_k_render, deck, None, self.valves)
        w = list(res.get("warnings") or [])
        if note:
            w.insert(0, f"JSON auto-repaired ({note})")
        return json.dumps({"ok": res.get("ok", False), "errors": res.get("errors", []), "warnings": w, "slides": res.get("slides")}, ensure_ascii=False)

    async def generate_slides(
        self,
        content: str = "{}",
        __event_emitter__: Any = None,
        __user__: Optional[dict] = None,
        __messages__: Any = None,
        __metadata__: Any = None,
        __request__: Any = None,
    ) -> str:
        """Create a high-quality NATIVE PowerPoint (.pptx) and return a download
        link.

        TWO FORMATS. Prefer the v2 BLOCKS format for business / consulting decks:

        v2 BLOCKS FORMAT (Deck Kit engine: 44 components, 14 themes, auto-fit text)
        {"meta": {"title": "...", "author": "...", "footer": "Confidential"},
         "theme": "azure_corporate",   // or {"preset":"mono_slate","primary":"1F4E79","accent":"FFC000"}
         "slides": [
          {"type":"cover","kicker":"...","title":"...","subtitle":"...","meta":"Presenter · Org · Date",
           "blocks":[{"type":"kpis","items":[{"value":"€22.9B","label":"...","icon":"money"}]}]},
          {"type":"section","number":"01","title":"...","subtitle":"..."},
          {"type":"content","kicker":"Value","title":"<one-sentence takeaway>",
           "rows":[{"h":1.1,"blocks":[{"type":"process","steps":[{"icon":"chat","title":"Intent","text":"..."}]}]},
                   {"h":1.4,"blocks":[{"type":"chart","heading":"...","chart_type":"bar","categories":["A","B"],
                                       "series":[{"name":"Impact","values":[90,50]}],"unit":"%"},
                                      {"type":"cards","columns":2,"items":[{"title":"...","text":"..."}]}]}],
           "source":"...", "notes":"speaker notes"},
          {"type":"closing","title":"Next steps","blocks":[{"type":"checklist","items":[{"text":"...","owner":"..."}]}]}]}
        Components: kpis rings progress callout | chart (bar column stacked_bar stacked_column line area pie
        doughnut radar combo) waterfall bubble mekko | process timeline chevrons cycle staircase flow swimlane |
        layers concentric value_chain hub tree pyramid funnel venn | swot matrix scorecard comparison heatmap
        maturity risk_matrix | gantt status board checklist raci | org_chart profiles persona journey |
        cards bullets table agenda image.
        Themes: midnight_violet graphite_gold ocean_teal forest_moss crimson_navy mono_slate azure_corporate
        emerald_finance civic_navy clinic_teal coral_energy royal_plum print_mono high_contrast.
        Before writing, call get_deck_components(names="<the components you plan>") for exact fields, or
        get_deck_template("<recipe>") for proven slide skeletons. Max 4 rows × 4 blocks per slide; 2-4 blocks per
        content slide; every number from a source. If the result lists errors, fix exactly those and call again.

        CLASSIC FORMAT (below): slides with a `layout` field. Required for Template Mode
        (reference_file_id / reuse).

        PRECONDITION — attached .pptx / "usa le slide in allegato come base",
        "usa questo template", "stesso stile del file", "use the attached
        deck": you MUST call `inspect_slides` FIRST and only then call this
        tool with the same Files API id as `reference_file_id` and per-slide
        `reuse` built from inspect slides[].text / slides[].tables (ids only;
        decorations are cloned when keep_ids is omitted). An id that is not
        the attached file is rejected. Do not fall back to a classic deck.

        Do NOT use this tool to analyze or inventory a template .pptx — that
        is `inspect_slides`. If the user asks only for the inspect JSON, do not
        call generate_slides and do not replace the tool JSON with a markdown
        summary.

        Template workflow (when the user wants their .pptx look):
        1) `inspect_slides` → factual inventory JSON (verbatim in the reply).
        2) This tool: same `reference_file_id`, one output slide per template
           slide to mirror, each with `reuse` {slide, keep_ids/drop_ids, text}.

        The `content` parameter MUST be a SINGLE JSON string (no text before or
        after, no markdown fence). Structure:

        {
          "title": "Presentation title",
          "subtitle": "Subtitle (optional)",
          "author": "Author / company (optional)",
          "theme": "auto",              // auto | midnight | forest | ocean |
                                        // coral | terracotta | teal | berry |
                                        // sage | cherry | charcoal | slate
          "accent": "#C99A3B",          // opt: force the accent color
          "footer": "Footer label",      // opt
          "reference_file_id": "<file_id from inspect_slides, copied verbatim>",
          "template_mapping": {          // opt: reference slide index per role
            "default": 0, "cover": 0, "content": 1, "closing": 2
          },
          "template_edits": {            // opt: clone policy (defaults false)
            "defaults": {
              "drop_text": false,
              "drop_placeholders": false,
              "drop_offslide": true
            },
            "slides": [{ "index": 0, "drop_ids": [12, 18] }]
          },
          "slides": [
            {
              "layout": "cover",
              "reuse": {
                "slide": 0,
                "keep_ids": [3, 4, 8, 9],
                "drop_ids": [10],
                "text": { "8": "New title", "9": "Subtitle" }
              }
            },
            {
              "reuse": {
                "slide": 9,
                "text": {
                  "9": "Project tracker",
                  "2": { "headers": ["Project", "Status"],
                         "rows": [["Alpha", "Done"], ["Beta", "WIP"]] }
                }
              }
            },
            { "layout": "title_bullets", "title": "...", "bullets": [] }
          ]
        }

        Per-slide `reuse` (preferred when mirroring a multi-slide template):
        clone template slide `reuse.slide` (0-based, same as inspect
        slides[].index), copy shapes by inspect `shapes[].id` via keep_ids /
        drop_ids, replace text in place via `text` — no classic layout overlay
        on that slide. Omit keep_ids to clone all cloneable shapes minus
        drop_ids. Requires reference_file_id + template_mode_enabled.
        `reference_file_id` is the inspect_slides `file_id` copied verbatim
        (a UUID, 8-4-4-4-12 hex). A number, a filename, or a slide index is
        rejected and you must call this tool again with the real id. Never
        invent one.

        Template tables (inspect `kind: "table"`, with `table.rows/cols/cells`)
        are refilled in place: `text["<table id>"] = {"headers": [...],
        "rows": [[...]]}` keeps the template cell formatting; rows/columns are
        added or removed to fit the data. If the slide also carries the
        classic `headers`/`rows` and the template slide has exactly one table,
        they are routed to that table automatically. A plain string on a table
        id is an error.

        `template_mapping` accepts explicit layout keys (e.g. `"table": 9`,
        `"team": 7`) which win over the coarse cover/section/closing/content
        roles when a slide is not using `reuse`.

        `drop_ids` use shape `id` from inspect JSON, not slide shape order.
        Set `drop_text` / remove text only when the user explicitly wants template
        wording removed (R8); default is to clone template textboxes.

        Requires admin valve `template_mode_enabled` for template rendering
        when `reference_file_id` is set. The valve alone does not clone a chat
        attachment: without `reference_file_id` (plus `template_mapping` and
        `template_edits` when needed) the deck is always the classic theme
        engine, even if the valve is true. If the valve is off, a provided
        reference is ignored and a classic deck is generated (NF8). If the
        valve is on and download, parse, or clone of the reference fails, the
        tool returns an error and does not produce a `.pptx` (F9). Do not
        retry the same request without `reference_file_id` when the user asked
        to keep their template. Template mode does not use chat attachments;
        only an explicit `reference_file_id` in this JSON applies
        the template.

        Each slide has a `layout` and fields consistent with that layout. Common
        fields: `title`, `eyebrow` (kicker, e.g. "PART I"), `subtitle`.

        AVAILABLE LAYOUTS and main fields:
        - "cover":        title, subtitle, author, eyebrow, icon, date, chips[]
        - "section":      number ("01"), eyebrow, title, lead   (chapter divider)
        - "title_bullets":title, eyebrow, bullets[] (or points/items)
        - "title_body":   title, eyebrow, body (paragraphs separated by \\n)
        - "two_column_text"/"comparison_two": left{}, right{} OR columns[];
              each card: {heading, icon, subtitle, description, points[],
              highlight:true, badge:"Most chosen"}
        - "kpi_row":      title, stats[] with {value, label, change}
        - "timeline_horizontal"/"process_flow": steps[] with {when, title, description}
        - "icon_list_vertical": items[] with {icon, title, description}
        - "icon_grid_2x2"/"icon_grid_3"/"pillars": items[] with {icon, title, description}
        - "chart":        chart_type (bar|line|area|pie|doughnut|radar|stacked_bar),
              labels[] and values[]  OR  datasets[]{label,data[]};
              opt. insight[]/insight_title for side text
        - "funnel"/"pyramid"/"cycle"/"quadrant"/"bullseye": nodes[] with
              {label, description}; quadrant accepts x_axis/y_axis and points[]
        - "quote":        quote, author, role
        - "alert":        title, level (info|tip|warning|danger), body or bullets[]
        - "table":        headers[], rows[] (lists or list of dicts)
        - "text_image_right"/"image_left_text_right": title, bullets[]/body,
              image_hint ("stock query") or image_url or base64
        - "image_full_caption": title, subtitle, image_hint/image_url
        - "closing":      title, eyebrow, takeaways[], contact

        DESIGN GUIDELINES (follow them):
        - "Sandwich" structure: start with "cover", use "section" between macro
          topics, end with "closing". Alternate layouts: do NOT repeat the same
          bulleted-list slide back to back.
        - 3-6 short bullets per slide. Prefer data/numbers: use "kpi_row",
          "chart", "funnel"/"pyramid" instead of long lists.
        - Use `eyebrow` to number the parts ("PART I — CONTEXT").
        - Icons: names like target, rocket, lightbulb, shield, lock, users,
          trending-up, bar-chart, cloud, cpu, dollar-sign, check-circle, award,
          layers, scale, workflow, globe, calendar, book-open, settings.
        - Choose `theme:"auto"` if unspecified: it is inferred from the content.
        - 10-18 slides for a full deck; do not generate walls of text.

        Returns a [TOOL_RESULT ...] line with the markdown link to show the user
        so they can download the .pptx.
        """
        if not _HAS_PPTX:
            return self._error("python-pptx is not installed in the runtime.")
        # parse + salvage (fences, prose around the JSON, bracket slips)
        spec, repair_note, perr = self._parse_content(content)
        if perr:
            return _retry_tool_result(perr)

        # ---- v2: Deck Kit blocks engine --------------------------------
        if self.valves.kit_engine_enabled and _k_is_kit_deck(spec):
            return await self._generate_kit(spec, repair_note, __event_emitter__, __request__, __user__)

        if spec.get("theme") in (None, "") and self.valves.default_theme:
            spec["theme"] = self.valves.default_theme

        _pre_slides = [
            s
            for s in _as_list(
                _first(spec, "slides", "sections", "pages", "deck", default=[])
            )
            if isinstance(s, dict)
        ]
        _has_reuse = any(s.get("reuse") is not None for s in _pre_slides)
        ref_id = (spec.get("reference_file_id") or "").strip()
        attached = (
            _find_pptx_attachment(__messages__)
            if self.valves.template_mode_enabled
            else None
        )
        if attached and ref_id != attached:
            # A missing id, a bare number, or a UUID that is not the attachment
            # must not fall through to a classic deck. The model retries.
            return _retry_tool_result(_wrong_reference_id_message(ref_id, attached))
        if _has_reuse and not ref_id:
            return self._error(
                "slides[].reuse requires reference_file_id (same Files API id as "
                "inspect_slides on the template .pptx)."
            )
        if _has_reuse and not self.valves.template_mode_enabled:
            return self._error(
                "slides[].reuse requires admin valve template_mode_enabled=true."
            )

        template_extra = ""
        template_pack = None
        reference_bytes = None
        if ref_id:
            if not self.valves.template_mode_enabled:
                log.warning(
                    "[generate_slides] reference_file_id=%s ignored: "
                    "template_mode_enabled is false",
                    ref_id,
                )
                spec.pop("reference_file_id", None)
                template_extra = (
                    "Template ignored: template_mode_enabled is false."
                )
            else:
                await self._emit(
                    __event_emitter__, "Loading template...", done=False
                )
                reference_bytes, _ref_fname, ref_err = await _load_reference_pptx(
                    ref_id, __request__, __user__
                )
                if ref_err:
                    return _retry_tool_result(
                        f"Could not load reference_file_id {ref_id!r}: {ref_err}. "
                        "Do not generate a classic deck and do not invent another "
                        "file_id. Call generate_slides again with the file_id "
                        "from the successful inspect_slides result, copied exactly, "
                        "and per-slide reuse."
                    )
                await self._emit(
                    __event_emitter__, "Applying template...", done=False
                )
                try:
                    template_pack = _parse_reference_pptx(reference_bytes)
                except Exception as exc:
                    log.exception("[generate_slides] reference parse failed")
                    return self._error(f"Invalid template .pptx: {exc}")
                if self.valves.template_strict_mode:
                    try:
                        _validate_template_strict(
                            template_pack, spec.get("template_mapping")
                        )
                    except ValueError as exc:
                        return self._error(str(exc))

        await self._emit(__event_emitter__, "Generating presentation...", done=False)
        try:
            _pf = [s for s in _as_list(_first(spec, "slides", "sections", "pages",
                                              "deck", default=[])) if isinstance(s, dict)]
            if any(_resolve_layout(s) in ("image_full_caption", "image_grid",
                   "text_image_right", "image_left_text_right") for s in _pf):
                await self._emit(__event_emitter__, "Fetching images...", done=False)
                await self._prefetch_images(_pf, __request__, __user__)
            if template_pack is not None:
                data, n = self._build(
                    spec,
                    template_pack=template_pack,
                    reference_bytes=reference_bytes,
                )
            else:
                data, n = self._build(spec)
        except ValueError as exc:
            return self._error(str(exc))
        except Exception as exc:
            import traceback
            traceback.print_exc()
            detail = str(exc).strip() or type(exc).__name__
            hint = ""
            if template_pack is not None:
                hint = (
                    " Template clone failed. Do not retry generate_slides "
                    "without reference_file_id."
                )
            return self._error(f"Rendering error: {detail}.{hint}")

        await self._emit(__event_emitter__, "Saving file...", done=False)
        fname, url, err, file_id = await self._save(
            data, title=spec.get("title", "presentation"),
            request=__request__, user_dict=__user__,
        )
        if not url:
            await self._emit(__event_emitter__, "Save failed.", done=True)
            return self._error(f"Presentation created but saving failed ({err}).")
        await self._emit_link(__event_emitter__, fname, url, slides=n,
                              kb=max(1, round(len(data) / 1024)),
                              file_id=file_id)
        await self._emit(__event_emitter__, "Presentation ready.", done=True)
        return self._success(fname, url, extra=template_extra)


# ============================================================================
# Layout dispatcher
# ============================================================================

def _dispatch(deck, layout, slide_dict, page):
    if layout == "cover":
        return _r_cover(deck, slide_dict, page)
    if layout == "section":
        return _r_section(deck, slide_dict, page)
    if layout == "title_only":
        return _r_title_only(deck, slide_dict, page)
    if layout in ("title_body", "blank"):
        return _r_title_body(deck, slide_dict, page)
    if layout in ("title_bullets",):
        return _r_title_bullets(deck, slide_dict, page)
    if layout in ("two_column_text", "comparison_two"):
        return _r_comparison(deck, slide_dict, page)
    if layout == "kpi_row":
        return _r_kpi(deck, slide_dict, page)
    if layout == "timeline_horizontal":
        return _r_timeline(deck, slide_dict, page)
    if layout == "process_flow":
        return _r_timeline(deck, slide_dict, page, process=True)
    if layout == "quote":
        return _r_quote(deck, slide_dict, page)
    if layout == "closing":
        return _r_closing(deck, slide_dict, page)
    if layout == "alert":
        return _r_alert(deck, slide_dict, page)
    if layout == "table":
        return _r_table(deck, slide_dict, page)
    if layout == "icon_list_vertical":
        return _r_icon_list(deck, slide_dict, page)
    if layout in ("icon_grid_2x2",):
        return _r_icon_grid(deck, slide_dict, page, cols=2)
    if layout in ("icon_grid_3", "icon_grid"):
        return _r_icon_grid(deck, slide_dict, page, cols=3)
    if layout == "pillars":
        return _r_icon_grid(deck, slide_dict, page, cols=min(3, max(2, len(_icon_items(slide_dict)) or 3)))
    if layout == "chart":
        return _r_chart(deck, slide_dict, page)
    if layout in ("text_image_right", "image_left_text_right", "image_full_caption",
                  "image_grid"):
        return _r_image(deck, slide_dict, page, layout)
    if layout in _DIAGRAM_LAYOUTS or layout == "diagram":
        return _r_diagram(deck, slide_dict, page, layout)
    # content (auto) + unknown → choose bullets vs body
    return _r_content_auto(deck, slide_dict, page)


def _r_content_auto(deck, slide_dict, page):
    bullets = _harvest_bullets(slide_dict)
    if _split_cards_present(slide_dict):
        return _r_comparison(deck, slide_dict, page)
    if bullets:
        return _r_title_bullets(deck, slide_dict, page)
    return _r_title_body(deck, slide_dict, page)


def _split_cards_present(slide_dict):
    if isinstance(slide_dict.get("left"), dict) and isinstance(slide_dict.get("right"), dict):
        return True
    for key in ("columns", "cards", "tiers", "plans", "options"):
        arr = slide_dict.get(key)
        if isinstance(arr, list) and arr and isinstance(arr[0], dict):
            return True
    return False


# ============================================================================
# Native charts
# ============================================================================

_CHART_TYPE_MAP = {
    "bar": "COLUMN_CLUSTERED", "column": "COLUMN_CLUSTERED", "col": "COLUMN_CLUSTERED",
    "bar_horizontal": "BAR_CLUSTERED", "horizontalbar": "BAR_CLUSTERED",
    "hbar": "BAR_CLUSTERED", "barh": "BAR_CLUSTERED",
    "line": "LINE", "area": "AREA", "pie": "PIE", "doughnut": "DOUGHNUT",
    "donut": "DOUGHNUT", "radar": "RADAR",
    "stacked_bar": "COLUMN_STACKED", "stacked": "COLUMN_STACKED",
}


def _num(v):
    if isinstance(v, (int, float)):
        return float(v)
    if isinstance(v, str):
        s = re.sub(r"[^0-9.\-,]", "", v).replace(".", "").replace(",", ".") \
            if v.count(",") == 1 and v.count(".") > 1 else re.sub(r"[^0-9.\-]", "", v.replace(",", "."))
        try:
            return float(s)
        except ValueError:
            return 0.0
    return 0.0


def _normalize_chart(slide: dict):
    """Return (chart_type_str, labels[list], series[list of (name, values, color)])."""
    ch = slide.get("chart") if isinstance(slide.get("chart"), dict) else {}
    raw_type = str(_first(slide, "chart_type", "type", default="") or
                   ch.get("type") or "bar").lower().strip()
    ctype = _CHART_TYPE_MAP.get(raw_type, "COLUMN_CLUSTERED")

    # nested Chart.js style
    data = ch.get("data") if isinstance(ch.get("data"), dict) else None
    labels = (_first(slide, "labels", "categories", "x", default=None)
              or (data or {}).get("labels")
              or (ch or {}).get("labels"))
    datasets = (_first(slide, "datasets", "series", default=None)
                or (data or {}).get("datasets"))

    series = []
    if datasets and isinstance(datasets, list):
        for i, ds in enumerate(datasets):
            if isinstance(ds, dict):
                name = _first(ds, "label", "name", "title", default=f"Serie {i+1}")
                vals = ds.get("data") or ds.get("values") or ds.get("points") or []
                color = ds.get("color") or ds.get("colour") or ds.get("backgroundColor")
                series.append((str(name), [_num(v) for v in _as_list(vals)],
                               _hex(color) if color else None))
            elif isinstance(ds, (list,)):
                series.append((f"Serie {i+1}", [_num(v) for v in ds], None))
    else:
        # single-series shorthands
        vals = _first(slide, "values", "data", "series", default=None)
        if isinstance(vals, list) and vals and isinstance(vals[0], dict):
            # [{label,value}]
            labels = [_first(d, "label", "name", "title", "x", default="") for d in vals]
            series = [(_first(slide, "series_label", "dataset_label", default="Value"),
                       [_num(_first(d, "value", "y", "amount", "count", default=0)) for d in vals],
                       None)]
        elif isinstance(vals, list):
            series = [(_first(slide, "series_label", "dataset_label", default="Value"),
                       [_num(v) for v in vals], None)]

    if not labels:
        # try to derive from content/items array of {label,value}
        for key in ("content", "items", "points", "rows", "entries"):
            arr = slide.get(key)
            if isinstance(arr, list) and arr and isinstance(arr[0], dict):
                labels = [_first(d, "label", "name", "title", "x", default="") for d in arr]
                series = [("Value", [_num(_first(d, "value", "y", "amount", default=0)) for d in arr], None)]
                break

    labels = [str(lb) for lb in _as_list(labels)] if labels else []
    if not series and labels:
        series = [("Value", [0.0] * len(labels), None)]
    return ctype, labels, series


def _series_colors(theme, n):
    base = [theme["accent"], theme["primary"], theme["light_accent"],
            _lighten(theme["accent"], 0.35), _darken(theme["primary"], 0.2),
            _lighten(theme["primary"], 0.4)]
    out = []
    for i in range(max(n, 1)):
        out.append(base[i] if i < len(base) else _lighten(theme["accent"], 0.15 * (i % 5)))
    return out


def _circular_colors(theme, n):
    """Saturated mid/dark slice colours so white % labels stay readable."""
    base = [theme["accent"], theme["primary"], _mix(theme["accent"], theme["primary"], 0.5),
            _darken(theme["accent"], 0.22), _lighten(theme["primary"], 0.28),
            _darken(theme["primary"], 0.18)]
    out = []
    for i in range(max(n, 1)):
        out.append(base[i] if i < len(base) else _mix(theme["accent"], theme["primary"], (i % 5) / 5))
    return out


def _style_chart(chart, theme, ctype, *, single_series, want_labels=True):
    t = theme
    try:
        chart.has_title = False
    except Exception:
        pass
    # legend
    circular = ctype in ("PIE", "DOUGHNUT")
    try:
        if single_series and not circular:
            chart.has_legend = False
        else:
            chart.has_legend = True
            chart.legend.position = XL_LEGEND_POSITION.BOTTOM
            chart.legend.include_in_layout = False
            chart.legend.font.size = Pt(10)
            chart.legend.font.name = t["body_font"]
            chart.legend.font.color.rgb = _rgb(t["muted"])
    except Exception:
        pass
    # series colours + data labels
    colors = _series_colors(t, len(chart.series))
    for i, s in enumerate(chart.series):
        try:
            if circular:
                # colour each point
                cc = _circular_colors(t, len(s.points))
                for j, pt in enumerate(s.points):
                    pt.format.fill.solid()
                    pt.format.fill.fore_color.rgb = _rgb(cc[j])
            else:
                s.format.fill.solid()
                s.format.fill.fore_color.rgb = _rgb(colors[i])
                if ctype in ("LINE",):
                    s.format.line.color.rgb = _rgb(colors[i])
                    s.format.line.width = Pt(2.5)
                    s.smooth = True
        except Exception:
            pass
    # data labels
    try:
        plot = chart.plots[0]
        plot.has_data_labels = want_labels
        if want_labels:
            dl = plot.data_labels
            dl.font.size = Pt(10)
            dl.font.name = t["body_font"]
            dl.font.color.rgb = _rgb(t["ink"] if not circular else "FFFFFF")
            # NOTE: an explicit ``number_format`` MUST be set. Using only
            # ``number_format_is_linked = True`` makes python-pptx emit
            # ``<c:numFmt sourceLinked="1"/>`` WITHOUT the schema-required
            # ``formatCode`` attribute, which corrupts the chart part
            # (PowerPoint reports a content error and drops the chart).
            dl.number_format = "0%" if circular else "General"
            dl.number_format_is_linked = False
            if circular:
                dl.show_percentage = True
                dl.show_value = False
                dl.font.bold = True
                # NOTE: doughnut charts do NOT allow ``<c:dLblPos>``. Emitting it
                # makes PowerPoint report a content error, "repair" the file and
                # drop the chart (the slide then looks empty). Only pie charts
                # accept a data-label position.
                if ctype == "PIE":
                    try:
                        dl.position = XL_LABEL_POSITION.CENTER
                    except Exception:
                        pass
            if ctype == "DOUGHNUT":
                try:
                    chart.plots[0].gap_width = 0
                except Exception:
                    pass
    except Exception:
        pass
    # axes styling (non-circular)
    if not circular:
        try:
            cat = chart.category_axis
            cat.tick_labels.font.size = Pt(10)
            cat.tick_labels.font.name = t["body_font"]
            cat.tick_labels.font.color.rgb = _rgb(t["muted"])
            cat.format.line.color.rgb = _rgb(t["hairline"])
            cat.has_major_gridlines = False
        except Exception:
            pass
        try:
            val = chart.value_axis
            val.tick_labels.font.size = Pt(10)
            val.tick_labels.font.name = t["body_font"]
            val.tick_labels.font.color.rgb = _rgb(t["muted"])
            val.has_major_gridlines = True
            val.major_gridlines.format.line.color.rgb = _rgb(t["hairline"])
            val.major_gridlines.format.line.width = Pt(0.5)
            val.format.line.fill.background()
        except Exception:
            pass


def _r_chart(deck, slide_dict, page):
    t = deck.theme
    # The model sometimes puts a *layout* name in `chart_type` (e.g.
    # "timeline_horizontal") with layout-shaped data (steps/nodes). Honour it
    # by rerouting to the right renderer instead of drawing an empty chart.
    raw_ct = str(_first(slide_dict, "chart_type", "type", default="")).lower().strip()
    if raw_ct and raw_ct not in _CHART_TYPE_MAP:
        rerouted = _LAYOUT_ALIASES.get(raw_ct, raw_ct)
        _RENDERABLE = (
            "timeline_horizontal", "process_flow", "kpi_row", "comparison_two",
            "two_column_text", "title_bullets", "title_body", "icon_list_vertical",
            "icon_grid_2x2", "icon_grid_3", "pillars", "quote", "closing",
            "section", "diagram",
        )
        if rerouted in _RENDERABLE or rerouted in _DIAGRAM_LAYOUTS:
            return _dispatch(deck, rerouted, slide_dict, page)

    ctype_str, labels, series = _normalize_chart(slide_dict)
    if not labels or not series:
        # Not real chart data → render its content as bullets. Return BEFORE
        # creating a slide so we don't emit an empty header-only slide.
        return _r_title_bullets(deck, slide_dict, page)

    slide = deck.blank("content")
    _maybe_set_bg(deck, slide, t["bg_light"])
    top = deck._content_head(slide, slide_dict, page)

    # side insight bullets?
    insight = _first(slide_dict, "insight", "insights", "takeaways", "notes",
                     "bullets", "points", default=None)
    side = _harvest_bullets({"points": insight}) if insight else []
    side_title = _first(slide_dict, "insight_title", "side_title", default="")

    if deck.template_mode:
        fx, fw, y0, avail = _content_layout_band(deck, top)
        y = max(y0 + 0.05, y0)
        h = avail - 0.05
        gap = 0.45
        if side:
            chart_w = fw * 0.55
            chart_x = fx
            text_x = fx + chart_w + gap
            text_w = fw - chart_w - gap
        else:
            chart_x, chart_w = fx, fw
            text_x, text_w = fx, fw
    else:
        y = max(top + 0.2, 1.95)
        h = FOOTER_Y - y - 0.2
        if side:
            cw = 7.1
            chart_x, chart_w = MARGIN, cw
            text_x = MARGIN + cw + 0.5
            text_w = SLIDE_W_IN - MARGIN - text_x
        else:
            chart_x, chart_w = MARGIN, SLIDE_W_IN - 2 * MARGIN
            text_x, text_w = chart_x, chart_w

    cd = CategoryChartData()
    cd.categories = labels
    for name, vals, _c in series:
        # pad/truncate to labels length
        v = (vals + [0.0] * len(labels))[:len(labels)]
        cd.add_series(name, v)

    xl_type = getattr(XL_CHART_TYPE, ctype_str, XL_CHART_TYPE.COLUMN_CLUSTERED)
    try:
        gframe = slide.shapes.add_chart(xl_type, Inches(chart_x), Inches(y),
                                        Inches(chart_w), Inches(h), cd)
        _style_chart(gframe.chart, t, ctype_str,
                     single_series=(len(series) == 1),
                     want_labels=(len(labels) <= 8))
    except Exception:
        import traceback
        traceback.print_exc()

    if side:
        ty = y + 0.1
        if side_title:
            tbt, tft = _textbox(slide, text_x, ty, text_w, 0.6)
            _add_para(tft, _strip_md(str(side_title)), first=True, size=16,
                      color=t["ink"], bold=True, font=t["head_font"], space_after=0)
            ty += 0.7
        _bullet_block(slide, t, side, text_x, ty, text_w, h - (ty - y), size=13)


def _place_image(slide, theme, raw, x, y, w, h, *, caption=""):
    """Place resolved image bytes into a region; graceful gradient placeholder."""
    if raw:
        try:
            slide.shapes.add_picture(BytesIO(raw), Inches(x), Inches(y),
                                     Inches(w), Inches(h))
        except Exception:
            raw = None
    if not raw:
        _card(slide, x, y, w, h, fill=theme["primary"], radius=0.04, shadow=False)
        _draw_icon(slide, theme, "image", x + w / 2 - 0.45, y + h / 2 - 0.45, 0.9,
                   circle_fill=_lighten(theme["primary"], 0.12),
                   glyph_color=theme["light_accent"])
    if caption:
        cy = y + h - 0.6
        _rect(slide, x, cy, w, 0.6, fill="1A1A2E", alpha=55)
        tbc, tfc = _textbox(slide, x + 0.25, cy, w - 0.5, 0.6, anchor="middle")
        _add_para(tfc, _strip_md(str(caption)), first=True, size=11,
                  color="FFFFFF", font=theme["body_font"], space_after=0)


def _slide_images(slide_dict):
    imgs = slide_dict.get("_img")
    if imgs is None:
        return []
    return imgs if isinstance(imgs, list) else [imgs]


def _r_image(deck, slide_dict, page, layout):
    t = deck.theme
    slide = deck.blank("content")
    _maybe_set_bg(deck, slide, t["bg_light"])
    imgs = _slide_images(slide_dict)

    if layout == "image_full_caption":
        raw = imgs[0] if imgs else None
        cap = _first(slide_dict, "caption", "subtitle", "lead", default="")
        if deck.template_mode:
            fx, fw = _frame_x_w(deck)
            fy, fh = deck.frame.y, deck.frame.h
            _place_image(slide, t, raw, fx, fy, fw, fh, caption="")
            if _first(slide_dict, "title", default=""):
                band_h = min(1.5, fh * 0.22)
                _rect(slide, fx, fy + fh - band_h, fw, band_h, fill="0A0F26", alpha=62)
                tb, tf = _textbox(
                    slide, fx, fy + fh - band_h + 0.15, fw, band_h - 0.2, anchor="middle"
                )
                _add_para(
                    tf,
                    _strip_md(str(slide_dict.get("title"))),
                    first=True,
                    size=28,
                    color="FFFFFF",
                    bold=True,
                    font=t["head_font"],
                    space_after=2,
                )
                if cap and cap != slide_dict.get("title"):
                    _add_para(
                        tf,
                        _strip_md(str(cap)),
                        size=13,
                        color="E8EAF0",
                        font=t["body_font"],
                        space_after=0,
                    )
            return
        _place_image(slide, t, raw, 0, 0, SLIDE_W_IN, SLIDE_H_IN, caption="")
        # title band bottom
        if _first(slide_dict, "title", default=""):
            _rect(slide, 0, SLIDE_H_IN - 1.5, SLIDE_W_IN, 1.5, fill="0A0F26", alpha=62)
            tb, tf = _textbox(slide, MARGIN, SLIDE_H_IN - 1.35, SLIDE_W_IN - 2 * MARGIN,
                              1.1, anchor="middle")
            _add_para(tf, _strip_md(str(slide_dict.get("title"))), first=True, size=30,
                      color="FFFFFF", bold=True, font=t["head_font"], space_after=2)
            if cap and cap != slide_dict.get("title"):
                _add_para(tf, _strip_md(str(cap)), size=13, color="E8EAF0",
                          font=t["body_font"], space_after=0)
        return

    if layout == "image_grid":
        top = deck._content_head(slide, slide_dict, page)
        items = slide_dict.get("images") or slide_dict.get("items") or []
        n = max(len(imgs), len(_as_list(items)))
        n = min(max(n, 1), 4)
        cols = 2 if n <= 4 else 3
        rows = (n + cols - 1) // cols
        gap = 0.3
        if deck.template_mode:
            fx, total_w, y0, avail = _content_layout_band(deck, top, y_floor_legacy=2.0)
            ch = (avail - 0.1 - gap * (rows - 1)) / rows
        else:
            fx = MARGIN
            total_w = SLIDE_W_IN - 2 * MARGIN
            y0 = max(top + 0.2, 2.0)
            ch = (FOOTER_Y - y0 - 0.1 - gap * (rows - 1)) / rows
        cw = (total_w - gap * (cols - 1)) / cols
        for i in range(n):
            r, c = divmod(i, cols)
            x = fx + c * (cw + gap)
            yy = y0 + r * (ch + gap)
            raw = imgs[i] if i < len(imgs) else None
            cap = ""
            if i < len(_as_list(items)) and isinstance(_as_list(items)[i], dict):
                cap = _first(_as_list(items)[i], "caption", "title", "label", default="")
            _place_image(slide, t, raw, x, yy, cw, ch, caption=cap)
        return

    # text_image_right / image_left_text_right
    top = deck._content_head(slide, slide_dict, page)
    raw = imgs[0] if imgs else None
    img_left = (layout == "image_left_text_right")
    if deck.template_mode:
        fx, fw, y0, avail = _content_layout_band(deck, top)
        y = max(y0 + 0.05, y0)
        h = avail - 0.05
        iw = fw * 0.42
        gap = 0.45
        if img_left:
            img_x = fx
            txt_x = fx + iw + gap
        else:
            img_x = fx + fw - iw
            txt_x = fx
        txt_w = fw - iw - gap
    else:
        y = max(top + 0.15, 1.95)
        h = FOOTER_Y - y - 0.2
        iw = 5.2
        if img_left:
            img_x = MARGIN
            txt_x = MARGIN + iw + 0.6
        else:
            img_x = SLIDE_W_IN - MARGIN - iw
            txt_x = MARGIN
        txt_w = SLIDE_W_IN - MARGIN - iw - 0.6 - MARGIN
    _place_image(slide, t, raw, img_x, y, iw, h,
                 caption=_first(slide_dict, "caption", default=""))
    bullets = _harvest_bullets(slide_dict)
    body = _first(slide_dict, "body", "content", "text", "description", default="")
    if bullets:
        _bullet_block(slide, t, bullets, txt_x, y + 0.1, txt_w, h, size=14)
    elif body:
        tb, tf = _textbox(slide, txt_x, y + 0.1, txt_w, h)
        for i, para in enumerate([p for p in str(body).split("\n") if p.strip()]):
            _add_para(tf, _md_runs(para), first=(i == 0), size=15, color=t["muted"],
                      font=t["body_font"], line=1.36, space_after=10)


# ============================================================================
# Native diagrams (autoshape compositions)
# ============================================================================

def _diagram_nodes(slide_dict):
    for key in ("nodes", "steps", "stages", "levels", "segments", "layers",
                "items", "phases", "parts", "points", "rings"):
        arr = slide_dict.get(key)
        if isinstance(arr, list) and arr:
            out = []
            for it in arr:
                if isinstance(it, dict):
                    out.append(it)
                else:
                    out.append({"label": str(it)})
            return out
    # bullets fallback
    b = _harvest_bullets(slide_dict)
    return [{"label": _strip_md(x)} for x in b] if b else []


def _node_text(n):
    return _strip_md(str(_first(n, "label", "title", "name", "heading", "text", default="")))


def _node_desc(n):
    return _strip_md(str(_first(n, "description", "detail", "value", "sub", default="")))


def _r_funnel(deck, slide_dict, page, nodes):
    t = deck.theme
    slide = deck.blank("content")
    _maybe_set_bg(deck, slide, t["bg_light"])
    top = deck._content_head(slide, slide_dict, page)
    n = min(len(nodes), 6)
    nodes = nodes[:n]
    dx, y, dw, avail = _diagram_region(deck, top)
    bh = min(0.9, (avail - (n - 1) * 0.12) / n)
    maxw = dw * 0.92 if deck.template_mode else 8.4
    minw = dw * 0.38 if deck.template_mode else 3.6
    cx = dx + dw / 2
    for i, node in enumerate(nodes):
        w = maxw - (maxw - minw) * (i / max(1, n - 1))
        x = cx - w / 2
        yy = y + i * (bh + 0.12)
        color = _mix(t["accent"], t["primary"], i / max(1, n - 1))
        shp = slide.shapes.add_shape(MSO_SHAPE.TRAPEZOID, Inches(x), Inches(yy),
                                     Inches(w), Inches(bh))
        try:
            shp.adjustments[0] = 0.18
        except Exception:
            pass
        shp.rotation = 180
        shp.fill.solid()
        shp.fill.fore_color.rgb = _rgb(color)
        _no_line(shp)
        shp.shadow.inherit = False
        tb, tf = _textbox(slide, x, yy, w, bh, anchor="middle")
        label = _node_text(node)
        val = _node_desc(node)
        _add_para(tf, label + (f"  ·  {val}" if val else ""), first=True, size=14,
                  color=_on(color), bold=True, font=t["body_font"], align="center",
                  space_after=0)


def _r_pyramid(deck, slide_dict, page, nodes):
    t = deck.theme
    slide = deck.blank("content")
    _maybe_set_bg(deck, slide, t["bg_light"])
    top = deck._content_head(slide, slide_dict, page)
    n = min(len(nodes), 5)
    nodes = nodes[:n]
    dx, y, dw, avail = _diagram_region(deck, top)
    bh = min(1.0, (avail - (n - 1) * 0.1) / n)
    maxw = dw * 0.88 if deck.template_mode else 8.0
    minw = dw * 0.28 if deck.template_mode else 2.6
    cx = dx + maxw / 2 + (0.05 if deck.template_mode else 0.2)
    side_x = cx + maxw / 2 + 0.35
    for i, node in enumerate(nodes):
        # top = narrow (i=0), bottom = wide
        w = minw + (maxw - minw) * (i / max(1, n - 1))
        x = cx - w / 2
        yy = y + i * (bh + 0.1)
        color = _mix(t["accent"], t["primary"], 1 - i / max(1, n - 1))
        shp = slide.shapes.add_shape(MSO_SHAPE.TRAPEZOID, Inches(x), Inches(yy),
                                     Inches(w), Inches(bh))
        try:
            shp.adjustments[0] = 0.30 if i == 0 else 0.12
        except Exception:
            pass
        shp.fill.solid()
        shp.fill.fore_color.rgb = _rgb(color)
        _no_line(shp)
        shp.shadow.inherit = False
        tb, tf = _textbox(slide, x, yy, w, bh, anchor="middle")
        _add_para(tf, _node_text(node), first=True, size=13.5, color=_on(color),
                  bold=True, font=t["body_font"], align="center", space_after=0)
        desc = _node_desc(node)
        if desc:
            side_w = (dx + dw - side_x) if deck.template_mode else (SLIDE_W_IN - MARGIN - side_x)
            tbd, tfd = _textbox(slide, side_x, yy, side_w, bh, anchor="middle")
            _add_para(tfd, desc, first=True, size=11, color=t["muted"],
                      font=t["body_font"], line=1.15, space_after=0)


def _r_cycle(deck, slide_dict, page, nodes):
    t = deck.theme
    slide = deck.blank("content")
    _maybe_set_bg(deck, slide, t["bg_light"])
    top = deck._content_head(slide, slide_dict, page)
    import math
    n = min(len(nodes), 6)
    nodes = nodes[:n]
    dx, y, dw, avail = _diagram_region(deck, top)
    if deck.template_mode:
        cx, cy = dx + dw / 2, y + avail / 2
        R = min(dw / 2 - 0.35, avail / 2 - 0.25)
    else:
        cx, cy = SLIDE_W_IN / 2, (top + FOOTER_Y) / 2 + 0.2
        R = min(2.1, (FOOTER_Y - top) / 2 - 0.2)
    nd = 1.5
    for i, node in enumerate(nodes):
        ang = -math.pi / 2 + 2 * math.pi * i / n
        x = cx + R * math.cos(ang) - nd / 2
        y = cy + R * math.sin(ang) - nd / 2
        color = _mix(t["accent"], t["primary"], i / max(1, n - 1))
        _oval(slide, x, y, nd, fill=color)
        tb, tf = _textbox(slide, x, y, nd, nd, anchor="middle")
        _add_para(tf, str(i + 1), first=True, size=18, color=_on(color), bold=True,
                  font=t["head_font"], align="center", space_after=2)
        _add_para(tf, _node_text(node), size=9.5, color=_on(color),
                  font=t["body_font"], align="center", line=1.0, space_after=0)


def _r_quadrant(deck, slide_dict, page, nodes):
    t = deck.theme
    slide = deck.blank("content")
    _maybe_set_bg(deck, slide, t["bg_light"])
    top = deck._content_head(slide, slide_dict, page)
    xa = _first(slide_dict, "x_axis", "xaxis", "x_label", default="")
    ya = _first(slide_dict, "y_axis", "yaxis", "y_label", default="")
    left_gutter = 0.45 if ya else 0.0
    if deck.template_mode:
        fx, fw, y0, avail = _content_layout_band(deck, top, y_floor_legacy=2.0)
        x0 = fx + left_gutter
        y = y0 + 0.1
        bottom = y + avail - (0.35 if xa else 0.12)
        w = fx + fw - x0
        hh = bottom - y
    else:
        x0 = MARGIN + left_gutter
        y = max(top + 0.35, 2.0)
        bottom = FOOTER_Y - (0.4 if xa else 0.15)
        w = SLIDE_W_IN - MARGIN - x0
        hh = bottom - y
    gap = 0.25
    cw = (w - gap) / 2
    ch = (hh - gap) / 2
    quads = (nodes + [{}] * 4)[:4]
    positions = [(x0, y), (x0 + cw + gap, y), (x0, y + ch + gap),
                 (x0 + cw + gap, y + ch + gap)]
    tints = [t["accent_soft"], _lighten(t["primary"], 0.9),
             _lighten(t["accent"], 0.82), t["card_soft"]]
    for i, (qx, qy) in enumerate(positions):
        _card(slide, qx, qy, cw, ch, fill=tints[i], radius=0.05, shadow=False,
              line=t["card_border"], line_w=1.0)
        node = quads[i]
        if not node:
            continue
        pad = 0.32
        tb, tf = _textbox(slide, qx + pad, qy + pad, cw - 2 * pad, ch - 2 * pad,
                          anchor="middle")
        _add_para(tf, _node_text(node), first=True, size=17, color=t["ink"],
                  bold=True, font=t["head_font"], space_after=5)
        items = node.get("points") or node.get("items") or node.get("bullets")
        if items:
            for b in _harvest_bullets({"points": items})[:4]:
                _add_bullet(tf, _clean_bullet(b), size=12, color=t["muted"],
                            accent=t["accent"], font=t["body_font"], space_after=4)
        elif _node_desc(node):
            _add_para(tf, _node_desc(node), size=12.5, color=t["muted"],
                      font=t["body_font"], line=1.28, space_after=0)
    if xa:
        tbx, tfx = _textbox(slide, x0, bottom + 0.05, w, 0.3, anchor="middle")
        _add_para(tfx, _strip_md(str(xa)).upper(), first=True, size=9.5, color=t["faint"],
                  bold=True, font=t["body_font"], align="center", tracking=1.8, space_after=0)
    if ya:
        tby = slide.shapes.add_textbox(Inches(x0 - hh / 2 - 0.15), Inches(y + hh / 2 - 0.15),
                                       Inches(hh), Inches(0.3))
        tfy = tby.text_frame
        tfy.word_wrap = False
        for m in (tfy.margin_left, tfy.margin_right, tfy.margin_top, tfy.margin_bottom):
            pass
        tby.rotation = 270
        _add_para(tfy, _strip_md(str(ya)).upper(), first=True, size=9.5, color=t["faint"],
                  bold=True, font=t["body_font"], align="center", tracking=1.8, space_after=0)


def _r_bullseye(deck, slide_dict, page, nodes):
    t = deck.theme
    slide = deck.blank("content")
    _maybe_set_bg(deck, slide, t["bg_light"])
    top = deck._content_head(slide, slide_dict, page)
    n = min(len(nodes), 4)
    nodes = nodes[:n]
    dx, y, dw, avail = _diagram_region(deck, top)
    if deck.template_mode:
        maxd = min(dw * 0.55, avail - 0.3)
        cx = dx + maxd / 2 + 0.15
        cy = y + avail / 2
        lx = cx + maxd / 2 + 0.35
        leg_w = dx + dw - lx
    else:
        cy = (top + FOOTER_Y) / 2 + 0.1
        cx = MARGIN + 3.0
        maxd = min(4.4, (FOOTER_Y - top) - 0.4)
        lx = cx + maxd / 2 + 0.6
        leg_w = SLIDE_W_IN - MARGIN - lx
    for i in range(n):
        d = maxd * (1 - i / n)
        color = _mix(_lighten(t["accent"], 0.55), t["accent"], i / max(1, n - 1))
        _oval(slide, cx - d / 2, cy - d / 2, d, fill=color)
    if deck.template_mode:
        tb_y, tb_h = y + 0.15, avail - 0.2
    else:
        tb_y, tb_h = top + 0.4, FOOTER_Y - top - 0.4
    tb, tf = _textbox(slide, lx, tb_y, leg_w, tb_h)
    first = True
    for i, node in enumerate(nodes):
        color = _mix(_lighten(t["accent"], 0.55), t["accent"], (n - 1 - i) / max(1, n - 1))
        _add_para(tf, ("\u25CF " if first else "\u25CF ") + _node_text(node),
                  first=first, size=15, color=t["ink"], bold=True,
                  font=t["body_font"], space_after=2)
        if _node_desc(node):
            _add_para(tf, _node_desc(node), size=11.5, color=t["muted"],
                      font=t["body_font"], line=1.2, space_after=10)
        first = False


def _r_diagram(deck, slide_dict, page, layout):
    dtype = str(_first(slide_dict, "diagram_type", "diagram", default="") or layout).lower()
    if layout in _DIAGRAM_LAYOUTS:
        dtype = layout
    nodes = _diagram_nodes(slide_dict)
    if not nodes:
        return _r_title_bullets(deck, slide_dict, page)
    if dtype in ("funnel", "imbuto"):
        return _r_funnel(deck, slide_dict, page, nodes)
    if dtype in ("pyramid", "piramide", "iceberg"):
        return _r_pyramid(deck, slide_dict, page, nodes)
    if dtype in ("cycle", "ciclo", "orbit", "loop"):
        return _r_cycle(deck, slide_dict, page, nodes)
    if dtype in ("quadrant", "matrix", "matrice", "quadrante", "2x2"):
        return _r_quadrant(deck, slide_dict, page, nodes)
    if dtype in ("bullseye", "target", "concentric", "venn"):
        return _r_bullseye(deck, slide_dict, page, nodes)
    if dtype in ("roadmap", "timeline", "process"):
        return _r_timeline(deck, {**slide_dict, "steps": nodes}, page)
    if dtype in ("pillars", "pilastri"):
        return _r_icon_grid(deck, slide_dict, page, cols=min(4, max(2, len(nodes))))
    # graceful degradation → icon list / bullets
    return _r_icon_list(deck, {**slide_dict, "items": nodes}, page)

