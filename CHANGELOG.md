# Changelog

All notable changes to the Generate Slides Open WebUI tool.

## [2.0.0] — Deck Kit engine (v2 blocks format)

Classic `layout` decks, Template Mode and `edit_presentation` behave exactly as in 1.0.4. The new engine is used only when a deck uses the blocks format.

### New engine
- The Deck Kit renderer, ported from pptxgenjs to python-pptx inside the single file: **44 components**, **14 theme presets** (incl. industry, print and high-contrast), cover/section/content/closing slide types, rows × blocks grid.
- Measured auto-fit: text is sized from Calibri/Carlito metrics, and sibling items share one font size.
- Native, editable charts including **combo** (columns + line on a secondary axis, still editable in PowerPoint's data sheet), radar, stacked and doughnut. Waterfall, bubble, mekko, Harvey balls, Gantt dependency arrows and journey emotion curves are drawn as native shapes.
- Images and avatars from URLs, data URIs or an admin images folder (`kit_images_dir`).
- 100 Font Awesome Free icons embedded (CC BY 4.0), with ~150 friendly aliases; the classic Lucide set is used as a fallback.

### New tools
- `get_deck_components(names)` — LLM-facing catalogue (overview, or fields + example per component).
- `get_deck_template(name)` — 17 layout recipes plus a `full_deck` starter.
- `validate_deck(content)` — schema + layout dry run without creating a file.

### Model guardrails
- JSON repair for one missing or extra bracket, code fences, surrounding prose and `{"deck": …}` wrappers (applies to classic decks too).
- Schema validation with exact error paths, returned as a "call again" tool result.
- Layout notes (overflow, rows too short, **new: blocks too narrow for their component**) returned with the link.

### Docs and tests
- `prompts/deck-builder-system-prompt.md` — system prompt for a "Deck Builder" model.
- `examples/kit/` — the 35-slide component showcase and an SAP Business AI brief.
- `examples/kit_engine_smoke.py` — added to `run_production_checks.py`.

## [1.0.4] — Templates and presentation edit

With the new valves left at their defaults, `generate_slides` still builds a deck from scratch, as in v1.0.3.

### Templates

Attach a `.pptx` and produce a new deck in that file's look. Backgrounds, logos and layout come from the template. The new copy replaces the text and the tables.

- **`inspect_slides`** reads the uploaded file and returns a short inventory: the file id, and on each slide the text boxes and tables that can be rewritten. Decorative shapes are left out of that list and cloned unchanged. On by default (`inspect_slides_enabled`).
- **`generate_slides`** writes the deck. The id from inspect is passed as `reference_file_id`. Each output slide can reuse one template slide and set the new text and table content. Cloning stays off until an admin sets `template_mode_enabled`.

### Edits on a deck already generated

**`edit_presentation`** changes a file the tool has already saved. It does not call `generate_slides` again, so a correction does not discard the template clone or the deck the user just opened.

It takes the generated file (`presentation_file_id`) and a batch of operations: replace text, set or shrink the font size, fit text to the box, PowerPoint autofit, split one text box into a title and a body, resize a shape, edit a table cell. The tool saves a new `.pptx`. The template id is not a valid target. If one operation is invalid, the batch is rejected and nothing is saved.

Off by default (`presentation_edit_enabled`), independent of template mode.

### Docs and tests

- [Template Mode hints](doc/neura-template-mode-model-hints.md)
- [Presentation edit hints](doc/neura-presentation-edit-hints.md)
- `examples/edit_presentation_smoke.py`, `examples/edit_presentation_acl_smoke.py`

## [1.0.3] — Files API baseline

- Stable Files API save path and classic layout engine (NF5 golden baseline in `examples/golden/`).
