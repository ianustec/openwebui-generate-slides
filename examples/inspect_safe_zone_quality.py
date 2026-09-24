"""Fase 3.3: safe_zone.quality + NF1 determinism (offline).

    pip install -r examples/requirements-dev.txt
    python examples/inspect_safe_zone_quality.py
"""
from __future__ import annotations

import importlib.util
import sys
from pathlib import Path

BASE = Path(__file__).resolve().parent.parent
FIXTURES = BASE / "examples" / "fixtures"


def _load_mod():
    spec = importlib.util.spec_from_file_location(
        "generate_slides", BASE / "generate_slides.py"
    )
    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)
    return mod


def _sz_tuple(sz) -> tuple:
    if sz is None:
        return ()
    return (
        round(sz.x, 4),
        round(sz.y, 4),
        round(sz.w, 4),
        round(sz.h, 4),
        sz.method,
        sz.quality,
    )


def main() -> None:
    mod = _load_mod()
    if not mod._HAS_PPTX:
        print("python-pptx not available", file=sys.stderr)
        sys.exit(2)

    failed = False
    for name in ("template_logo_confidential.pptx", "template_corners_one_slide.pptx"):
        path = FIXTURES / name
        if not path.is_file():
            print(f"SKIP missing {name}")
            continue
        data = path.read_bytes()
        a = mod._parse_reference_pptx(data)
        b = mod._parse_reference_pptx(data)
        for st_a, st_b in zip(a.slides, b.slides):
            if _sz_tuple(st_a.safe_zone) != _sz_tuple(st_b.safe_zone):
                print(f"FAIL: NF1 safe_zone mismatch on {name} slide {st_a.index}", file=sys.stderr)
                failed = True
            if st_a.safe_zone and st_a.safe_zone.quality != mod.SAFE_ZONE_QUALITY_COMPUTED:
                print(
                    f"FAIL: {name} slide {st_a.index} expected computed, "
                    f"got {st_a.safe_zone.quality}",
                    file=sys.stderr,
                )
                failed = True

    # Synthetic: many tiny decorations + one large blocker; edges simplify → computed grid.
    slide_w, slide_h = 13.333, 7.5
    tiny = [
        mod._Decoration(
            shape_id=i,
            kind="autoshape",
            name=f"t{i}",
            bbox=mod._BBox(x=0.2 + (i % 20) * 0.05, y=0.2, w=0.04, h=0.04),
            z_order=i,
        )
        for i in range(80)
    ]
    large = mod._Decoration(
        shape_id=999,
        kind="autoshape",
        name="block",
        bbox=mod._BBox(x=0.0, y=0.0, w=3.0, h=7.5),
        z_order=100,
    )
    sz = mod._compute_safe_zone(slide_w, slide_h, tiny + [large])
    if sz.quality != mod.SAFE_ZONE_QUALITY_COMPUTED:
        print(
            f"FAIL: synthetic dense slide expected computed, got {sz.quality}",
            file=sys.stderr,
        )
        failed = True
    elif sz.w <= 0 or sz.h <= 0:
        print("FAIL: synthetic safe_zone has non-positive size", file=sys.stderr)
        failed = True
    else:
        print("OK synthetic dense slide: quality=computed")

    if failed:
        sys.exit(1)
    print("OK: inspect_safe_zone_quality")
    sys.exit(0)


if __name__ == "__main__":
    main()
