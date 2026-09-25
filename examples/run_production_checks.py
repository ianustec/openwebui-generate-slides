"""Pre-production test battery (local / CI).

    pip install -r examples/requirements-dev.txt httpx
    python examples/run_production_checks.py

Runs NF5 baseline, template-mode smoke scripts, syntax check, and a few
offline unit checks for the Files API helpers and cache fallback save path.
Exit code 0 only if every step passes.
"""
from __future__ import annotations

import asyncio
import importlib.util
import subprocess
import sys
import tempfile
from pathlib import Path

BASE = Path(__file__).resolve().parent.parent
GENERATE = BASE / "generate_slides.py"


def _load_module():
    spec = importlib.util.spec_from_file_location("generate_slides", GENERATE)
    mod = importlib.util.module_from_spec(spec)
    assert spec.loader
    spec.loader.exec_module(mod)
    return mod


def _run(label: str, argv: list[str]) -> None:
    print(f"\n--- {label} ---")
    proc = subprocess.run(
        [sys.executable, *argv],
        cwd=BASE,
        capture_output=False,
    )
    if proc.returncode != 0:
        print(f"FAIL: {label} (exit {proc.returncode})", file=sys.stderr)
        sys.exit(proc.returncode)
    print(f"PASS: {label}")


async def _test_maybe_await(mod) -> None:
    async def coro():
        return 42

    assert await mod._maybe_await(coro()) == 42
    assert await mod._maybe_await(7) == 7


def _test_extract_file_id(mod) -> None:
    class M:
        id = "abc-123"

    assert mod._extract_file_id(M()) == "abc-123"
    assert mod._extract_file_id({"id": "x"}) == "x"
    assert mod._extract_file_id({}) is None
    assert mod._extract_file_id(None) is None


async def _test_save_cache_fallback(mod) -> None:
    tools = mod.Tools()
    with tempfile.TemporaryDirectory() as tmp:
        tools.valves.pptx_export_dir = tmp
        data = b"PK\x03\x04fake-pptx-bytes"
        fname, url, err, fid = await tools._save(
            data,
            title="Prod Check",
            request=None,
            user_dict={"id": "user-1"},
        )
        assert err is None, err
        assert fid is None
        assert url and url.startswith("/cache/files/")
        assert fname.endswith(".pptx")
        path = Path(tmp) / fname
        assert path.is_file() and path.read_bytes() == data


def main() -> None:
    print("Production checks — repo:", BASE)

    _run("py_compile generate_slides.py", ["-m", "py_compile", str(GENERATE)])

    mod = _load_module()
    asyncio.run(_test_maybe_await(mod))
    _test_extract_file_id(mod)
    asyncio.run(_test_save_cache_fallback(mod))
    print("PASS: offline helpers + _save cache fallback")

    # check_baseline already smoke-tests _build(deck.json); skip build.py here
    # so the committed demo_tech_deck.pptx is not overwritten before the hash check.
    scripts = [
        ("NF5 check_baseline", ["examples/check_baseline.py"]),
        ("fixtures/build_fixtures", ["examples/fixtures/build_fixtures.py"]),
        ("parse_template_fixtures", ["examples/parse_template_fixtures.py"]),
        ("inspect_template_fixtures", ["examples/inspect_template_fixtures.py"]),
        ("clone_template_fixture", ["examples/clone_template_fixture.py"]),
        (
            "template_mode_title_bullets",
            ["examples/template_mode_title_bullets.py"],
        ),
        ("template_mode_mixed_deck", ["examples/template_mode_mixed_deck.py"]),
        (
            "generate_slides_async_smoke",
            ["examples/generate_slides_async_smoke.py"],
        ),
        ("template_mode_nf6_15_slides", ["examples/template_mode_nf6_15_slides.py"]),
        ("clone_nested_group", ["examples/clone_nested_group.py"]),
        ("clone_blip_fill", ["examples/clone_blip_fill.py"]),
        ("reuse_slide_text", ["examples/reuse_slide_text.py"]),
        ("ianustec_neura_reuse_local", ["examples/ianustec_neura_reuse_local.py"]),
        ("template_strict_mode_test", ["examples/template_strict_mode_test.py"]),
        ("inspect_master_p1_smoke", ["examples/inspect_master_p1_smoke.py"]),
    ]
    for label, argv in scripts:
        _run(label, argv)

    print("\n=== ALL PRODUCTION CHECKS PASSED ===")


if __name__ == "__main__":
    main()
