"""v2 Deck Kit engine smoke test (offline).

    python examples/kit_engine_smoke.py

Renders the kit example decks through generate_slides (Tools API), checks the
catalogue / templates / validate_deck tools, JSON bracket repair, schema error
reporting and every component in every theme (light + dark).
"""
from __future__ import annotations

import asyncio
import importlib.util
import json
import sys
import tempfile
from pathlib import Path

BASE = Path(__file__).resolve().parent.parent
KIT = BASE / "examples" / "kit"


def _load():
    spec = importlib.util.spec_from_file_location("generate_slides", BASE / "generate_slides.py")
    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)
    return mod


async def main() -> int:
    mod = _load()
    t = mod.Tools()
    tmp = tempfile.mkdtemp()
    t.valves.pptx_export_dir = tmp
    t.valves.kit_images_dir = str(KIT / "images")
    fails = []

    def check(cond, label):
        print(("PASS " if cond else "FAIL ") + label)
        if not cond:
            fails.append(label)

    ov = json.loads(await t.get_deck_components())
    check(len(ov["components"]) == 44 and len(ov["themes"]) == 14, "catalogue: 44 components, 14 themes")
    sub = json.loads(await t.get_deck_components("gantt,swot"))
    check(set(sub["components"]) == {"gantt", "swot"}, "catalogue subset")
    full = json.loads(await t.get_deck_template("full_deck"))
    check(len(full["slides"]) == 8, "full_deck template")

    for name, n in (("showcase", 35), ("sap-business-ai", 4)):
        res = await t.generate_slides((KIT / f"{name}.json").read_text())
        check(".pptx](" in res and f"({n} slides" in res, f"render {name}")
        v = json.loads(await t.validate_deck((KIT / f"{name}.json").read_text()))
        check(v["ok"] and not v["warnings"], f"validate {name} clean")

    s = json.dumps(json.loads((KIT / "sap-business-ai.json").read_text()), separators=(",", ":"))
    spec, note, err = t._parse_content("```json\n" + s.replace("}]}", "}]", 1) + "\n```")
    check(spec is not None and note, "bracket repair")
    bad = {"slides": [{"type": "content", "title": "x", "rows": [{"blocks": [{"type": "chart", "chart_type": "bar",
           "categories": ["a", "b"], "series": [{"name": "s", "values": [1]}]}]}]}]}
    res = await t.generate_slides(json.dumps(bad))
    check("1 values but 2 categories" in res and "NOT rendered" in res, "schema error reported for retry")

    ex = {k: v["example"] for k, v in mod._K_COMPONENTS.items()}
    for theme in mod._K_PRESETS:
        slides = [{"type": "content", "dark": d, "title": k, "rows": [{"blocks": [ex[k]]}]} for k in ex for d in (False, True)]
        mod._K_DECK["slides"]["max"] = 200
        r = mod._k_render({"theme": theme, "slides": slides}, None)
        check(r["ok"], f"all components, theme {theme}")
    return 1 if fails else 0


if __name__ == "__main__":
    sys.exit(asyncio.run(main()))
