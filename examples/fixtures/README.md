# Template mode fixtures (Fase 0)

Small `.pptx` files for future template-engine tests. Regenerate with:

```bash
pip install -r examples/requirements-dev.txt
python examples/fixtures/build_fixtures.py
```

| File | Type | Purpose |
|------|------|---------|
| `template_corners_one_slide.pptx` | (a) | Picture shapes at corners + white central area |
| `template_bg_fill_one_slide.pptx` | (b) | Slide background fill + corner picture |
| `template_two_slides.pptx` | — | Two slides: cover-style + content-style (layout mapping tests) |
| `template_logo_confidential.pptx` | (c) | Logo picture + “Confidential” textbox + divider |
| `template_master_heavy.pptx` | (d) | Placeholder note: slide-master cloning is not in template v1 |

Slide size matches the engine: **13.333 × 7.5 in** (16:9).
