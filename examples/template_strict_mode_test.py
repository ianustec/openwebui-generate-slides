"""Fase 8: template_strict_mode fail-closed and pass on computed safe zone.

    pip install -r examples/requirements-dev.txt
    python examples/template_strict_mode_test.py
"""
from __future__ import annotations

import asyncio
import importlib.util
import json
import sys
import tempfile
from dataclasses import replace
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


async def _test_generate_strict_fail(mod) -> None:
    fixture_bytes = FIXTURES / "template_corners_one_slide.pptx"
    if not fixture_bytes.is_file():
        print("SKIP strict generate (missing fixture)")
        return

    async def fake_load(_fid, _req, _user):
        return fixture_bytes.read_bytes(), "template.pptx", None

    original = mod._load_reference_pptx
    mod._load_reference_pptx = fake_load
    try:
        pack = mod._parse_reference_pptx(fixture_bytes.read_bytes())
        issues = mod._template_strict_issues(pack, {"default": 0})
        assert not issues, issues

        bad = replace(
            pack.slides[0],
            safe_zone=mod._BBox(x=1, y=1, w=0.2, h=0.2, quality="computed"),
        )
        bad_pack = replace(pack, slides=[bad])
        try:
            mod._validate_template_strict(bad_pack, {"default": 0})
        except ValueError:
            pass
        else:
            raise AssertionError("expected strict failure on tiny safe zone")

        spec = {
            "title": "Strict fail",
            "reference_file_id": "mock",
            "template_mapping": {"default": 0},
            "slides": [{"layout": "title_bullets", "title": "T", "bullets": ["a"]}],
        }
        with tempfile.TemporaryDirectory() as tmp:
            tools = mod.Tools()
            tools.valves.pptx_export_dir = tmp
            tools.valves.template_mode_enabled = True
            tools.valves.template_strict_mode = True
            reply = await tools.generate_slides(
                json.dumps(spec),
                __user__={"id": "strict-user"},
            )
            assert "couldn't generate" not in reply.lower(), reply

            bad_pack2 = mod._parse_reference_pptx(fixture_bytes.read_bytes())
            st = bad_pack2.slides[0]
            st = replace(
                st,
                safe_zone=mod._BBox(
                    x=st.safe_zone.x,
                    y=st.safe_zone.y,
                    w=st.safe_zone.w,
                    h=st.safe_zone.h,
                    quality=mod.SAFE_ZONE_QUALITY_ADMISSIBLE_FALLBACK,
                ),
            )
            bad_pack2 = replace(bad_pack2, slides=[st])

            async def fake_load_bad(_fid, _req, _user):
                return fixture_bytes.read_bytes(), "x.pptx", None

            mod._load_reference_pptx = fake_load_bad
            orig_parse = mod._parse_reference_pptx

            def parse_return_bad(_data):
                return bad_pack2

            mod._parse_reference_pptx = parse_return_bad
            reply2 = await tools.generate_slides(
                json.dumps(spec),
                __user__={"id": "strict-user"},
            )
            mod._parse_reference_pptx = orig_parse
            assert "couldn't generate" in reply2.lower()
            assert "template_strict_mode" in reply2
    finally:
        mod._load_reference_pptx = original


def _test_inspect_hints(mod) -> None:
    data = (FIXTURES / "template_corners_one_slide.pptx").read_bytes()
    pack = mod._parse_reference_pptx(data)
    payload = mod._serialize_inspect_payload(
        pack, file_id="fixture", filename="corners.pptx", ok=True, images=[]
    )
    hints = payload.get("hints") or {}
    assert "strict_would_fail" in hints
    assert hints["strict_would_fail"] == []


async def main_async() -> None:
    mod = _load_mod()
    _test_inspect_hints(mod)
    await _test_generate_strict_fail(mod)


def main() -> None:
    asyncio.run(main_async())
    print("PASS: template_strict_mode tests")


if __name__ == "__main__":
    main()
