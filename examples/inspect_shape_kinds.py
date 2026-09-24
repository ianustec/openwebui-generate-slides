"""Fase 3.2: shape kind + cloneable smoke tests (offline).

    pip install -r examples/requirements-dev.txt
    python examples/inspect_shape_kinds.py
"""
from __future__ import annotations

import importlib.util
import json
import sys
from pathlib import Path

BASE = Path(__file__).resolve().parent.parent
FIXTURES = BASE / "examples" / "fixtures"
ONLINE = BASE / "doc" / "online_templates"


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

    failed = False

    corners = FIXTURES / "template_corners_one_slide.pptx"
    data = corners.read_bytes()
    pack = mod._parse_reference_pptx(data)
    pics = [s for s in pack.slides[0].shapes if s.kind == "picture"]
    if len(pics) < 4:
        print(f"FAIL: corners expected >=4 picture kinds, got {len(pics)}", file=sys.stderr)
        failed = True
    for s in pack.slides[0].shapes:
        if s.cloneable != mod._shape_cloneable(s.kind):
            print(f"FAIL: cloneable mismatch id={s.id}", file=sys.stderr)
            failed = True

    payload = mod._serialize_inspect_payload(
        pack, file_id="fixture:corners", filename=corners.name, ok=True, images=[]
    )
    if "uncloneable" not in payload.get("hints", {}):
        print("FAIL: hints.uncloneable missing", file=sys.stderr)
        failed = True
    for sh in payload["slides"][0]["shapes"]:
        if "cloneable" not in sh:
            print("FAIL: shapes[].cloneable missing", file=sys.stderr)
            failed = True
            break

    progetto = ONLINE / "Progetto senza titolo.pptx"
    if progetto.is_file():
        pdata = progetto.read_bytes()
        ppack = mod._parse_reference_pptx(pdata)
        s0 = ppack.slides[0].shapes
        by_id = {s.id: s for s in s0}
        pic = by_id.get(2)
        if not pic or pic.kind != "picture":
            print(
                f"FAIL: Progetto slide0 id=2 expected kind=picture, "
                f"got {pic.kind if pic else None}",
                file=sys.stderr,
            )
            failed = True
        elif not pic.cloneable:
            print("FAIL: Progetto full-bleed picture should be cloneable", file=sys.stderr)
            failed = True
        else:
            print("OK Progetto senza titolo.pptx: id=2 picture cloneable")
    else:
        print("SKIP Progetto senza titolo.pptx (not on disk)")

    pack_a = mod._parse_reference_pptx(data)
    pack_b = mod._parse_reference_pptx(data)
    sig_a = [(s.id, s.kind, s.cloneable) for st in pack_a.slides for s in st.shapes]
    sig_b = [(s.id, s.kind, s.cloneable) for st in pack_b.slides for s in st.shapes]
    if sig_a != sig_b:
        print("FAIL: NF1 kind/cloneable not stable between parses", file=sys.stderr)
        failed = True

    print(f"OK corners: {len(pics)} picture shapes, hints={json.dumps(payload['hints']['uncloneable'])}")
    sys.exit(1 if failed else 0)


if __name__ == "__main__":
    main()
