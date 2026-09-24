"""Fase 5: multi-layout deck in template mode (mixed safe zone).

    pip install -r examples/requirements-dev.txt
    python examples/template_mode_mixed_deck.py
"""
from __future__ import annotations

import importlib.util
import sys
from io import BytesIO
from pathlib import Path

BASE = Path(__file__).resolve().parent.parent
FIXTURES = BASE / "examples" / "fixtures"
OUT_DIR = BASE / "examples" / "output"
TOL_IN = 0.1


def _load_mod():
    spec = importlib.util.spec_from_file_location(
        "generate_slides", BASE / "generate_slides.py"
    )
    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)
    return mod


def _bbox_in(shape, mod) -> mod._BBox:
    return mod._BBox(
        x=mod._emu_in(shape.left),
        y=mod._emu_in(shape.top),
        w=mod._emu_in(shape.width),
        h=mod._emu_in(shape.height),
    )


def _contains(outer: mod._BBox, inner: mod._BBox, tol: float) -> bool:
    return (
        inner.x >= outer.x - tol
        and inner.y >= outer.y - tol
        and inner.x + inner.w <= outer.x + outer.w + tol
        and inner.y + inner.h <= outer.y + outer.h + tol
    )


def _intersects(a: mod._BBox, b: mod._BBox) -> bool:
    return not (
        a.x + a.w <= b.x
        or b.x + b.w <= a.x
        or a.y + a.h <= b.y
        or b.y + b.h <= a.y
    )


def main() -> None:
    mod = _load_mod()
    if not mod._HAS_PPTX:
        print("python-pptx not available", file=sys.stderr)
        sys.exit(2)

    from pptx import Presentation
    from pptx.enum.shapes import MSO_SHAPE_TYPE

    corners = FIXTURES / "template_corners_one_slide.pptx"
    if not corners.is_file():
        print(f"Missing {corners}", file=sys.stderr)
        sys.exit(1)

    source_bytes = corners.read_bytes()
    pack = mod._apply_template_edits(mod._parse_reference_pptx(source_bytes), None)
    sz = pack.slides[0].safe_zone
    assert sz is not None

    spec = {
        "title": "Mixed template deck",
        "template_mapping": {
            "default": 0,
            "cover": 0,
            "content": 0,
            "section": 0,
            "closing": 0,
        },
        "slides": [
            {
                "layout": "cover",
                "title": "Quarterly review",
                "subtitle": "Template mode mixed layouts",
            },
            {
                "layout": "title_bullets",
                "title": "Highlights",
                "bullets": ["Alpha", "Beta", "Gamma"],
            },
            {
                "layout": "kpi_row",
                "title": "KPIs",
                "stats": [
                    {"value": "42%", "label": "Growth"},
                    {"value": "1.2M", "label": "Users"},
                    {"value": "98", "label": "NPS"},
                ],
            },
            {
                "layout": "chart",
                "title": "Trend",
                "chart_type": "bar",
                "labels": ["Q1", "Q2", "Q3"],
                "datasets": [{"name": "Revenue", "values": [10, 14, 18]}],
            },
            {
                "layout": "comparison_two",
                "title": "Options",
                "left": {"heading": "Plan A", "points": ["Fast", "Simple"]},
                "right": {"heading": "Plan B", "points": ["Rich", "Flexible"]},
            },
            {
                "layout": "closing",
                "title": "Thank you",
                "takeaways": ["Questions welcome"],
            },
        ],
    }

    data, n = mod.Tools()._build(
        spec, template_pack=pack, reference_bytes=source_bytes
    )
    assert n == len(spec["slides"])

    OUT_DIR.mkdir(parents=True, exist_ok=True)
    out_path = OUT_DIR / "template_mode_mixed.pptx"
    out_path.write_bytes(data)

    prs = Presentation(BytesIO(data))
    assert len(prs.slides) == n
    for slide in prs.slides:
        pics = [
            sh
            for sh in slide.shapes
            if mod._shape_kind(sh) == "picture"
        ]
        assert len(pics) >= 4, "each slide should clone corner pictures"

    kpi_slide = prs.slides[2]
    kpi_text = [
        sh
        for sh in kpi_slide.shapes
        if getattr(sh, "has_text_frame", False) and "42%" in (sh.text_frame.text or "")
    ]
    assert kpi_text, "KPI value text not found"

    chart_slide = prs.slides[3]
    charts = [
        sh for sh in chart_slide.shapes if sh.shape_type == MSO_SHAPE_TYPE.CHART
    ]
    assert charts, "chart shape missing on chart slide"
    cb = _bbox_in(charts[0], mod)
    assert _contains(sz, cb, TOL_IN), "chart outside safe zone"

    print(f"OK: wrote {out_path} ({n} slides, KPI+chart in safe zone)")
    sys.exit(0)


if __name__ == "__main__":
    main()
