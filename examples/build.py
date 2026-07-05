"""Render examples/deck.json into a .pptx locally (no Open WebUI required).

    pip install python-pptx pillow
    python examples/build.py
"""
import importlib.util
import json
from pathlib import Path

BASE = Path(__file__).resolve().parent.parent

spec = json.loads((BASE / "examples" / "deck.json").read_text())

_spec = importlib.util.spec_from_file_location(
    "generate_slides", BASE / "generate_slides.py"
)
mod = importlib.util.module_from_spec(_spec)
_spec.loader.exec_module(mod)

data, n = mod.Tools()._build(spec)
out = BASE / "examples" / "demo_tech_deck.pptx"
out.write_bytes(data)
print(f"OK: {n} slides, {len(data)} bytes -> {out}")
