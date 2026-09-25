"""NF9 / NF5 / F10 / F9 / template E2E offline for async generate_slides.

    pip install -r examples/requirements-dev.txt
    python examples/generate_slides_async_smoke.py
"""
from __future__ import annotations

import asyncio
import hashlib
import importlib.util
import json
import sys
import tempfile
from pathlib import Path

BASE = Path(__file__).resolve().parent.parent
DECK_JSON = BASE / "examples" / "deck.json"
FIXTURE = BASE / "examples" / "fixtures" / "template_corners_one_slide.pptx"


def _load_mod():
    spec = importlib.util.spec_from_file_location(
        "generate_slides", BASE / "generate_slides.py"
    )
    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)
    return mod


def _latest_pptx(export_dir: Path) -> Path | None:
    files = sorted(export_dir.glob("*.pptx"), key=lambda p: p.stat().st_mtime)
    return files[-1] if files else None


async def _run_classic_generate(mod, export_dir: str) -> bytes:
    tools = mod.Tools()
    tools.valves.pptx_export_dir = export_dir
    tools.valves.template_mode_enabled = False
    content = DECK_JSON.read_text()
    reply = await tools.generate_slides(
        content,
        __user__={"id": "smoke-user"},
    )
    if "couldn't generate" in reply.lower():
        raise AssertionError(f"NF9 failed: {reply[:400]}")
    if "Here is the presentation" not in reply:
        raise AssertionError(f"Unexpected reply: {reply[:400]}")
    path = _latest_pptx(Path(export_dir))
    assert path is not None, "no pptx written to cache"
    return path.read_bytes()


async def test_nf9_and_nf5(mod) -> None:
    with tempfile.TemporaryDirectory() as tmp:
        gen_bytes = await _run_classic_generate(mod, tmp)
        spec = json.loads(DECK_JSON.read_text())
        sync_bytes, n = mod.Tools()._build(spec)
        assert n == 13, n
        if hashlib.sha256(gen_bytes).hexdigest() != hashlib.sha256(sync_bytes).hexdigest():
            raise AssertionError("NF5: generate_slides bytes != _build(spec)")
    print("PASS: NF9 async generate + NF5 hash match _build")


async def test_f10(mod) -> None:
    messages = [
        {
            "role": "user",
            "files": [
                {
                    "id": "attached-template-id",
                    "name": "user_template.pptx",
                    "content_type": "application/vnd.openxmlformats-officedocument.presentationml.presentation",
                }
            ],
        }
    ]
    spec = json.loads(DECK_JSON.read_text())
    assert "reference_file_id" not in spec or not spec.get("reference_file_id")
    sync_bytes, _ = mod.Tools()._build(spec)
    with tempfile.TemporaryDirectory() as tmp:
        tools = mod.Tools()
        tools.valves.pptx_export_dir = tmp
        reply = await tools.generate_slides(
            json.dumps(spec),
            __messages__=messages,
            __user__={"id": "smoke-user"},
        )
        assert "couldn't generate" not in reply.lower()
        path = _latest_pptx(Path(tmp))
        assert path is not None
        if hashlib.sha256(path.read_bytes()).hexdigest() != hashlib.sha256(sync_bytes).hexdigest():
            raise AssertionError("F10: attachment must not change classic output")
    print("PASS: F10 attachment ignored without reference_file_id")


async def test_f9_fail_closed(mod) -> None:
    async def fail_load(_fid, _req, _user):
        return None, None, "File not found or not accessible."

    original = mod._load_reference_pptx
    mod._load_reference_pptx = fail_load
    try:
        spec = {
            "title": "Should fail",
            "reference_file_id": "missing-id",
            "template_mapping": {"default": 0},
            "slides": [{"layout": "title_bullets", "title": "T", "bullets": ["a"]}],
        }
        with tempfile.TemporaryDirectory() as tmp:
            tools = mod.Tools()
            tools.valves.pptx_export_dir = tmp
            tools.valves.template_mode_enabled = True
            reply = await tools.generate_slides(
                json.dumps(spec),
                __user__={"id": "smoke-user"},
            )
            assert "couldn't generate" in reply.lower(), reply
            assert "File not found" in reply
            assert _latest_pptx(Path(tmp)) is None, "F9 must not write pptx"
    finally:
        mod._load_reference_pptx = original
    print("PASS: F9 fail closed on reference load error")


async def test_template_e2e_offline(mod) -> None:
    if not FIXTURE.is_file():
        print(f"SKIP template E2E: missing {FIXTURE}", file=sys.stderr)
        return
    fixture_bytes = FIXTURE.read_bytes()

    async def fake_load(_fid, _req, _user):
        return fixture_bytes, "template_corners_one_slide.pptx", None

    original = mod._load_reference_pptx
    mod._load_reference_pptx = fake_load
    try:
        spec = {
            "title": "Template smoke",
            "reference_file_id": "fixture-mock-id",
            "template_mapping": {"default": 0, "content": 0},
            "slides": [
                {
                    "layout": "title_bullets",
                    "title": "Points",
                    "bullets": ["One", "Two"],
                }
            ],
        }
        with tempfile.TemporaryDirectory() as tmp:
            tools = mod.Tools()
            tools.valves.pptx_export_dir = tmp
            tools.valves.template_mode_enabled = True
            reply = await tools.generate_slides(
                json.dumps(spec),
                __user__={"id": "smoke-user"},
            )
            assert "couldn't generate" not in reply.lower(), reply
            path = _latest_pptx(Path(tmp))
            assert path is not None
            from pptx import Presentation

            prs = Presentation(str(path))
            assert len(prs.slides) == 1
    finally:
        mod._load_reference_pptx = original
    print("PASS: template E2E offline (mock load + valve ON)")


async def main_async() -> None:
    mod = _load_mod()
    if not mod._HAS_PPTX:
        print("python-pptx not available", file=sys.stderr)
        sys.exit(2)
    await test_nf9_and_nf5(mod)
    await test_f10(mod)
    await test_f9_fail_closed(mod)
    await test_template_e2e_offline(mod)


def main() -> None:
    asyncio.run(main_async())
    print("\n=== generate_slides async smoke OK ===")


if __name__ == "__main__":
    main()
