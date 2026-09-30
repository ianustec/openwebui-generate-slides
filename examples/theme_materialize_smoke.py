"""Smoke test: reuse slides materialize a:schemeClr to a:srgbClr from template theme.

    pip install -r examples/requirements-dev.txt
    python examples/theme_materialize_smoke.py
"""
from __future__ import annotations

import importlib.util
import re
import sys
import zipfile
from io import BytesIO
from pathlib import Path

BASE = Path(__file__).resolve().parent.parent
INDIAN = BASE / "doc" / "online_templates" / "Indian Doctors Day _ by Slidesgo.pptx"
MARKETING = BASE / "doc" / "online_templates" / "Marketing.pptx"


def _load_mod():
    spec = importlib.util.spec_from_file_location(
        "generate_slides", BASE / "generate_slides.py"
    )
    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)
    return mod


def _slide_xml(pptx_bytes: bytes, slide_index: int) -> str:
    z = zipfile.ZipFile(BytesIO(pptx_bytes))
    return z.read(f"ppt/slides/slide{slide_index + 1}.xml").decode("utf-8", "replace")


def _assert_bg_materialized(xml: str, expected_hex: str) -> None:
    bg = re.search(r"<p:bg>.*?</p:bg>", xml, re.DOTALL)
    if bg is None:
        raise AssertionError("missing p:bg on slide")
    block = bg.group(0)
    hx = expected_hex.upper()
    if f'val="{hx}"' not in block and f'val="{hx.lower()}"' not in block:
        raise AssertionError(f"p:bg missing srgb {hx}: {block[:300]!r}")
    if "schemeClr" in block:
        raise AssertionError("p:bg still contains schemeClr after materialize")


def _assert_no_scheme_on_slide(xml: str) -> None:
    if re.search(r"<a:schemeClr\b", xml):
        raise AssertionError("slide XML still contains a:schemeClr")


def _indian_smoke(mod) -> None:
    if not INDIAN.is_file():
        print(f"SKIP Indian (missing {INDIAN})")
        return
    ref = INDIAN.read_bytes()
    pack = mod._parse_reference_pptx(ref)
    assert pack.theme.get("lt1") == "E2EEEA", pack.theme.get("lt1")
    assert pack.theme.get("dk1") == "0B352D", pack.theme.get("dk1")
    spec = {
        "title": "theme-smoke",
        "slides": [
            {
                "reuse": {
                    "slide": 0,
                    "text": {"783": "TITLE TEST", "784": "SUB TEST"},
                }
            }
        ],
    }
    data, n = mod.Tools()._build(
        spec, template_pack=pack, reference_bytes=ref
    )
    assert n == 1
    xml = _slide_xml(data, 0)
    _assert_bg_materialized(xml, "E2EEEA")
    _assert_no_scheme_on_slide(xml)
    assert "0B352D" in xml.upper()
    print("OK: Indian Doctors slide 0 theme materialized (bg lt1, dk1/accent srgb)")

    # Slide index 1: no p:bg on slide in template; mint comes from layout (flatten).
    spec_layout = {
        "title": "theme-smoke-layout-bg",
        "slides": [{"reuse": {"slide": 1, "text": {}}}],
    }
    data_layout, n_layout = mod.Tools()._build(
        spec_layout, template_pack=pack, reference_bytes=ref
    )
    assert n_layout == 1
    xml_layout = _slide_xml(data_layout, 0)
    _assert_bg_materialized(xml_layout, "E2EEEA")
    print("OK: Indian Doctors slide 1 layout background flattened to slide p:bg")


def _marketing_smoke(mod) -> None:
    if not MARKETING.is_file():
        print(f"SKIP Marketing (missing {MARKETING})")
        return
    ref = MARKETING.read_bytes()
    pack = mod._parse_reference_pptx(ref)
    lt1 = pack.theme.get("lt1")
    assert lt1, "Marketing theme lt1 missing"
    spec = {
        "title": "theme-smoke-mkt",
        "slides": [
            {
                "reuse": {
                    "slide": 0,
                    "text": {"8": "IANUSTEC", "9": "THEME SMOKE"},
                }
            }
        ],
    }
    data, _n = mod.Tools()._build(
        spec, template_pack=pack, reference_bytes=ref
    )
    xml = _slide_xml(data, 0)
    bg = re.search(r"<p:bg>.*?</p:bg>", xml, re.DOTALL)
    if bg is not None and "schemeClr" in bg.group(0):
        _assert_bg_materialized(xml, lt1)
    # Cover uses scheme text colors heavily; after materialize none should remain.
    scheme_count = len(re.findall(r"<a:schemeClr\b", xml))
    if scheme_count:
        raise AssertionError(f"Marketing slide 0 still has {scheme_count} schemeClr")
    print(f"OK: Marketing slide 0 materialized (lt1={lt1}, schemeClr=0)")


def main() -> None:
    mod = _load_mod()
    _indian_smoke(mod)
    _marketing_smoke(mod)
    print("PASS: theme_materialize_smoke")
    sys.exit(0)


if __name__ == "__main__":
    main()
