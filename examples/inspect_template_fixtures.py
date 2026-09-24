"""E2E dev check: parse fixtures → inspect JSON (Fase 2, offline).

    pip install -r examples/requirements-dev.txt
    python examples/inspect_template_fixtures.py

Full inspect_slides with Files API download requires Open WebUI runtime.
"""
from __future__ import annotations

import importlib.util
import json
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


def _shape_ids(pack) -> list[tuple[int, int]]:
    out = []
    for st in pack.slides:
        for sh in st.shapes:
            out.append((st.index, sh.id))
    return out


def _safe_zone_fingerprint(pack) -> list[tuple]:
    out = []
    for st in pack.slides:
        sz = st.safe_zone
        if sz is None:
            out.append((st.index, None))
            continue
        out.append(
            (
                st.index,
                round(sz.x, 4),
                round(sz.y, 4),
                round(sz.w, 4),
                round(sz.h, 4),
                sz.quality,
            )
        )
    return out


_VALID_QUALITIES = frozenset({"computed", "grid_capped", "admissible_fallback"})


def main() -> None:
    mod = _load_mod()
    if not mod._HAS_PPTX:
        print("python-pptx not available", file=sys.stderr)
        sys.exit(2)

    pptx_files = sorted(FIXTURES.glob("*.pptx"))
    if not pptx_files:
        print(f"No fixtures in {FIXTURES}", file=sys.stderr)
        sys.exit(1)

    for path in pptx_files:
        data = path.read_bytes()
        pack_a = mod._parse_reference_pptx(data)
        pack_b = mod._parse_reference_pptx(data)
        ids_a = _shape_ids(pack_a)
        ids_b = _shape_ids(pack_b)
        if ids_a != ids_b:
            print(f"FAIL: shape ids differ between runs for {path.name}", file=sys.stderr)
            sys.exit(1)
        sz_a = _safe_zone_fingerprint(pack_a)
        sz_b = _safe_zone_fingerprint(pack_b)
        if sz_a != sz_b:
            print(
                f"FAIL: safe_zone fingerprint differs between runs for {path.name}",
                file=sys.stderr,
            )
            sys.exit(1)

        media = mod._extract_referenced_media(data, pack_a)
        images_json = mod._inspect_images_to_json(media)
        payload = mod._serialize_inspect_payload(
            pack_a,
            file_id=f"fixture:{path.name}",
            filename=path.name,
            ok=True,
            images=images_json,
        )
        raw = json.dumps(payload)
        parsed = json.loads(raw)
        for key in ("ok", "slides", "theme", "decorations_source", "hints", "images"):
            if key not in parsed:
                print(f"FAIL: missing key {key!r} in {path.name}", file=sys.stderr)
                sys.exit(1)
        assert parsed["ok"] is True
        assert isinstance(parsed["images"], list)
        assert "uncloneable" in parsed["hints"]
        for slide in parsed["slides"]:
            if slide.get("safe_zone"):
                assert slide["safe_zone"].get("method"), "safe_zone.method required"
                q = slide["safe_zone"].get("quality")
                if q not in _VALID_QUALITIES:
                    print(
                        f"FAIL: invalid safe_zone.quality {q!r} in {path.name}",
                        file=sys.stderr,
                    )
                    sys.exit(1)
                if q != "computed":
                    print(
                        f"FAIL: fixture {path.name} slide {slide['index']} "
                        f"expected quality=computed, got {q!r}",
                        file=sys.stderr,
                    )
                    sys.exit(1)
            for sh in slide["shapes"]:
                assert "cloneable" in sh, "shapes[].cloneable required"
                assert sh["cloneable"] == mod._shape_cloneable(sh["kind"])

        img_n = len(parsed["images"])
        print(
            f"OK {path.name}: slides={parsed['slide_count']} "
            f"shapes={len(ids_a)} images={img_n}"
        )

    logo = FIXTURES / "template_logo_confidential.pptx"
    pack = mod._parse_reference_pptx(logo.read_bytes())
    payload = mod._serialize_inspect_payload(
        pack, file_id="fixture:logo", filename=logo.name, ok=True
    )
    texts = payload["slides"][0]["text_verbatim"]
    if not any("Confidential" in t for t in texts):
        print(f"FAIL: Confidential not in text_verbatim: {texts}", file=sys.stderr)
        sys.exit(1)
    conf_shapes = [
        sh
        for sh in payload["slides"][0]["shapes"]
        if sh.get("text") and "Confidential" in sh["text"]
    ]
    if not conf_shapes:
        print("FAIL: Confidential missing from shapes[].text (R8)", file=sys.stderr)
        sys.exit(1)
    pics = [s for s in payload["slides"][0]["shapes"] if s["kind"] == "picture"]
    if not pics:
        print("FAIL: expected picture shape in logo fixture", file=sys.stderr)
        sys.exit(1)

    print("\nOK: inspect JSON serialization and stable shape ids")
    sys.exit(0)


if __name__ == "__main__":
    main()
