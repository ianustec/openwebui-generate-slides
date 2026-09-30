# Changelog

All notable changes to the Generate Slides Open WebUI tool.

## [1.2.1] — Presentation Edit (font fixes)

**Compatibility:** Default `shrink_font` behavior unchanged (`mode=fit`). No change to generate/template paths.

### Added

- **`set_font_pt`** — set absolute `font_pt` or relative `font_pt_delta` on text shapes.
- **`shrink_font` / `fit_text` modes:** `fit` (default), `to_min`, `target` (+ `target_pt`).
- **`_edit_shape_font_pt_effective`** — reads run `sz` plus paragraph `endParaRPr` / `defRPr` fallbacks.
- Log field **`ops_mutated`** (distinct from `ops_applied`).
- Valve **`presentation_edit_warn_noop`** (default `false`) — optional status when no op mutated the deck.

### Tests

- Smoke: endParaRPr effective font, `set_font_pt`, PE1b `to_min`.
- `examples/montessori_edit_font_regression.py` (skip-safe if fixture deck missing targets).

## [1.2.0] — Presentation Edit

**Compatibility:** No change to `generate_slides` / Template Mode when `presentation_edit_enabled=false` (default) and when `edit_presentation` is not called.

### Added

- **`edit_presentation`** — post-generate edits on an existing `.pptx` via `presentation_file_id` (Files API UUID) and `operations[]`.
- Operations v1: `set_text`, `shrink_font`, `fit_text`.
- Operations v2: `enable_autofit`, `replace_text_and_fit`, `split_text`, `resize_shape`, `set_table_cell`.
- Admin valves: `presentation_edit_enabled` (default `false`), `presentation_edit_min_font_pt`, `presentation_edit_shrink_step_pt`, `presentation_edit_progress_every`.

### Behavior

- Edit requires explicit UUID (`presentation_file_id`); cache-only URLs are rejected (#1A).
- Batch operations: **fail total** on first invalid op (#8A); partial save (#8B) not implemented.
- Rejects `presentation_file_id` equal to the chat-attached template file (#7B).
- File load uses the same ACL path as template reference download.

### Hardening

- Status emits for large decks / long operation batches.
- Logs: `edit_ms`, `ops_applied`, `ops_failed`.
- Upload/save log prefix `edit_presentation`.

### Docs and tests

- [Presentation edit plan](doc/presentation-edit-implementation-plan.md)
- [Model hints (Neura)](doc/neura-presentation-edit-hints.md)
- `examples/edit_presentation_smoke.py`, `examples/edit_presentation_acl_smoke.py`

## [1.1.0] — Template Mode

**Compatibility:** No behavior change when `template_mode_enabled=false` (default) and when the JSON spec has no `reference_file_id`. Classic deck generation matches the v1.0.3 path (NF5 / NF8).

### Added

- **`inspect_slides`** — factual inventory JSON for an uploaded template `.pptx` (shape ids, bbox, `safe_zone`, `text_verbatim`, optional `images[]`).
- **Template Mode in `generate_slides`** — clone decorative shapes from a reference file and overlay semantic content in the reference safe zone.
- Spec fields: `reference_file_id`, `template_mapping`, `template_edits` (including `drop_ids` from inspect).
- Admin valves: `template_mode_enabled` (default `false`), `inspect_slides_enabled`, `inspect_extract_images`.

### Behavior

- Valve **ON** + invalid or inaccessible reference → clear error, **no** output `.pptx` (F9).
- Valve **OFF** + `reference_file_id` in JSON → classic deck; reference ignored with a visible note (NF8).
- **`generate_slides` does not** auto-detect chat attachments for templates; only an explicit `reference_file_id` applies Template Mode.

### Hardening (1.1.x)

- Inspect: `clone_reason` on non-cloneable shapes; `hints.master_decorations`, `hints.strict_would_fail`.
- Template: flattened leaf decorations, recursive shape lookup, layout/master decoration merge (P1).
- Valve `template_strict_mode` for pilot QA on `safe_zone` quality.
- Timing logs: `parse_ms`, `inspect_ms`, template `build clone_ms`.

### Docs and examples

- [Template Mode implementation plan](doc/template-mode-implementation-plan.md)
- [Example spec with reference](examples/template-deck-with-reference.json)
- [Model hints (Neura)](doc/neura-template-mode-model-hints.md)

## [1.0.3] — Files API baseline

- Stable Files API save path and classic layout engine (NF5 golden baseline in `examples/golden/`).
