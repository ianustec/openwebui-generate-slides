"""ACL / access checks for edit_presentation file load (offline).

    pip install -r examples/requirements-dev.txt httpx
    python examples/edit_presentation_acl_smoke.py
"""
from __future__ import annotations

import asyncio
import importlib.util
import sys
from pathlib import Path
from unittest.mock import AsyncMock, MagicMock, patch

BASE = Path(__file__).resolve().parent.parent
VALID_UUID = "8f3c2a1b-1234-5678-9abc-def012345678"


def _load_mod():
    spec = importlib.util.spec_from_file_location(
        "generate_slides", BASE / "generate_slides.py"
    )
    mod = importlib.util.module_from_spec(spec)
    assert spec.loader
    spec.loader.exec_module(mod)
    return mod


async def _test_missing_user(mod) -> None:
    data, _fname, err = await mod._load_reference_pptx(
        VALID_UUID, MagicMock(), None
    )
    assert data is None
    assert err and "user" in err.lower()
    print("OK: load rejected without user context")


async def _test_http_403(mod) -> None:
    if not mod._HAS_HTTPX:
        print("SKIP HTTP ACL (_HAS_HTTPX false)")
        return
    req = MagicMock()
    req.base_url = "http://localhost:8080"
    req.headers = MagicMock()
    req.headers.get = MagicMock(return_value=None)

    mock_resp = MagicMock()
    mock_resp.status_code = 403
    mock_resp.content = b""
    mock_resp.headers = {}

    mock_client = AsyncMock()
    mock_client.get = AsyncMock(return_value=mock_resp)
    mock_client.__aenter__ = AsyncMock(return_value=mock_client)
    mock_client.__aexit__ = AsyncMock(return_value=None)

    with patch.object(mod.httpx, "AsyncClient", return_value=mock_client):
        data, _fname, err = await mod._load_reference_pptx(
            VALID_UUID, req, {"id": "user-a"}
        )
    assert data is None
    assert err and "not accessible" in err.lower()
    print("OK: HTTP 403 → not accessible (same path as edit download)")


async def _test_owui_file_not_for_user(mod) -> None:
    if not mod._HAS_OWUI_FILE_DOWNLOAD:
        print("SKIP in-process OWUI ACL (Open WebUI models not importable)")
        return

    async def _no_file(_fid, _uid):
        return None

    fake_files = MagicMock()
    fake_files.get_file_by_id_and_user_id = _no_file
    fake_files.get_file_by_id = AsyncMock(return_value=None)

    with patch.object(mod, "_OwuiFiles", fake_files), patch.object(
        mod, "_HAS_OWUI_FILE_DOWNLOAD", True
    ):
        data, _fname, err = await mod._load_reference_pptx(
            VALID_UUID, MagicMock(), {"id": "user-a"}
        )
    assert data is None
    assert err and "not accessible" in err.lower()
    print("OK: OWUI get_file_by_id_and_user_id miss → not accessible")


def _test_template_attachment_guard(mod) -> None:
    msgs = [
        {
            "files": [
                {
                    "id": "template-uuid-1111-2222-3333-444455556666",
                    "name": "Marketing.pptx",
                }
            ]
        }
    ]
    fid = "template-uuid-1111-2222-3333-444455556666"
    assert mod._presentation_id_is_template_attachment(fid, msgs)
    assert not mod._presentation_id_is_template_attachment(VALID_UUID, msgs)
    print("OK: #7B template attachment id detection helper")


async def main_async() -> None:
    mod = _load_mod()
    await _test_missing_user(mod)
    await _test_http_403(mod)
    await _test_owui_file_not_for_user(mod)
    _test_template_attachment_guard(mod)
    print("edit_presentation_acl_smoke: all checks passed")


def main() -> None:
    asyncio.run(main_async())


if __name__ == "__main__":
    try:
        main()
    except Exception as exc:
        print(f"FAIL: {exc}", file=sys.stderr)
        raise SystemExit(1) from exc
