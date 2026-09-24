"""Fase 3.1: extract referenced media from template fixtures (offline, no upload).

    pip install -r examples/requirements-dev.txt
    python examples/inspect_extract_media_fixtures.py

Writes a summary JSON to examples/output/inspect_media_summary.json
"""
from __future__ import annotations

import importlib.util
import json
import sys
from pathlib import Path

BASE = Path(__file__).resolve().parent.parent
FIXTURES = BASE / "examples" / "fixtures"
OUT = BASE / "examples" / "output"


def _load_mod():
    spec = importlib.util.spec_from_file_location(
        "generate_slides", BASE / "generate_slides.py"
    )
    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)
    return mod


def main() -> None:
    mod = _load_mod()
    if not mod._HAS_PPTX:
        print("python-pptx not available", file=sys.stderr)
        sys.exit(2)

    OUT.mkdir(parents=True, exist_ok=True)
    summary: list[dict] = []
    failed = False

    for path in sorted(FIXTURES.glob("*.pptx")):
        data = path.read_bytes()
        pack = mod._parse_reference_pptx(data)
        images = mod._extract_referenced_media(data, pack)
        row = {
            "file": path.name,
            "unique_parts": len(images),
            "usage_count": sum(len(im.usages) for im in images),
            "bytes_total": sum(len(im.blob) for im in images),
            "paths": [im.internal_path for im in images],
        }
        summary.append(row)
        print(
            f"OK {path.name}: parts={row['unique_parts']} "
            f"usages={row['usage_count']} bytes={row['bytes_total']}"
        )

    corners = FIXTURES / "template_corners_one_slide.pptx"
    corners_row = next(r for r in summary if r["file"] == corners.name)
    if corners_row["unique_parts"] < 1:
        print("FAIL: expected media in template_corners_one_slide.pptx", file=sys.stderr)
        failed = True

    out_path = OUT / "inspect_media_summary.json"
    out_path.write_text(json.dumps(summary, indent=2) + "\n", encoding="utf-8")
    print(f"\nWrote {out_path.relative_to(BASE)}")
    sys.exit(1 if failed else 0)


if __name__ == "__main__":
    main()
