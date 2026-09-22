"""Compare baseline deck output to the committed golden (NF5).

    pip install -r examples/requirements-dev.txt
    python examples/check_baseline.py

The golden SHA256 locks ``examples/demo_tech_deck.pptx`` in git. A fresh
``_build(deck.json)`` may differ slightly in bytes (python-pptx chart packaging)
while still being valid; this script always checks the committed artifact, then
smoke-checks an in-memory build (slide count + size band).

    python examples/check_baseline.py --strict
        Also require in-memory SHA256 == golden (may fail until chart output is stable).

    python examples/check_baseline.py --write
        Regenerates examples/demo_tech_deck.pptx from _build (does not update golden).
"""
from __future__ import annotations

import argparse
import hashlib
import importlib.util
import json
import sys
from pathlib import Path

BASE = Path(__file__).resolve().parent.parent
GOLDEN_FILE = BASE / "examples" / "golden" / "demo_tech_deck.sha256"
DECK_JSON = BASE / "examples" / "deck.json"
OUT_PPTX = BASE / "examples" / "demo_tech_deck.pptx"
EXPECTED_SLIDES = 13
# Allow small byte drift between python-pptx runs (charts); golden file is canonical.
SIZE_TOLERANCE_BYTES = 512


def _load_expected_hash() -> str:
    if not GOLDEN_FILE.is_file():
        print(f"Missing golden file: {GOLDEN_FILE}", file=sys.stderr)
        sys.exit(2)
    for line in GOLDEN_FILE.read_text().splitlines():
        line = line.strip()
        if not line or line.startswith("#"):
            continue
        return line.lower()
    print(f"No hash in {GOLDEN_FILE}", file=sys.stderr)
    sys.exit(2)


def _build_bytes() -> tuple[bytes, int]:
    spec = json.loads(DECK_JSON.read_text())
    mod_spec = importlib.util.spec_from_file_location(
        "generate_slides", BASE / "generate_slides.py"
    )
    mod = importlib.util.module_from_spec(mod_spec)
    mod_spec.loader.exec_module(mod)
    data, n = mod.Tools()._build(spec)
    print(f"Built {n} slides, {len(data)} bytes (in memory)")
    return data, n


def main() -> None:
    parser = argparse.ArgumentParser(description="NF5 baseline hash check")
    parser.add_argument(
        "--write",
        action="store_true",
        help="Write demo_tech_deck.pptx from build (does not update golden file)",
    )
    parser.add_argument(
        "--strict",
        action="store_true",
        help="Require in-memory _build SHA256 to match golden exactly",
    )
    args = parser.parse_args()

    expected = _load_expected_hash()
    if not OUT_PPTX.is_file():
        print(f"Missing {OUT_PPTX}; run python examples/build.py first.", file=sys.stderr)
        sys.exit(2)

    on_disk = hashlib.sha256(OUT_PPTX.read_bytes()).hexdigest()
    if on_disk != expected:
        print("FAIL: committed demo_tech_deck.pptx does not match golden", file=sys.stderr)
        print(f"  golden:   {expected}", file=sys.stderr)
        print(f"  on-disk:  {on_disk}", file=sys.stderr)
        sys.exit(1)

    data, n = _build_bytes()
    if args.write:
        OUT_PPTX.write_bytes(data)
        print(f"Wrote {OUT_PPTX}")

    digest = hashlib.sha256(data).hexdigest()
    if n != EXPECTED_SLIDES:
        print(f"FAIL: expected {EXPECTED_SLIDES} slides, got {n}", file=sys.stderr)
        sys.exit(1)

    ref_size = OUT_PPTX.stat().st_size
    if abs(len(data) - ref_size) > SIZE_TOLERANCE_BYTES:
        print(
            f"FAIL: _build size {len(data)} differs from artifact {ref_size} "
            f"by more than {SIZE_TOLERANCE_BYTES} bytes",
            file=sys.stderr,
        )
        sys.exit(1)

    if digest == expected:
        print(f"OK: SHA256 matches golden ({digest[:16]}…)")
        sys.exit(0)

    if args.strict:
        print("FAIL: strict mode — in-memory SHA256 mismatch", file=sys.stderr)
        print(f"  golden:  {expected}", file=sys.stderr)
        print(f"  _build:  {digest}", file=sys.stderr)
        sys.exit(1)

    print(
        "OK: golden artifact verified; _build smoke passed "
        f"({n} slides, size within {SIZE_TOLERANCE_BYTES} B of committed file)"
    )
    print(
        f"  note: in-memory SHA256 differs ({digest[:16]}…) — "
        "expected with python-pptx charts; use --strict when byte-stable"
    )
    sys.exit(0)


if __name__ == "__main__":
    main()
