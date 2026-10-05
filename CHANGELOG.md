# Changelog

All notable changes to the Generate Slides Open WebUI tool.

## [1.0.4] — Inspect slides

**Compatibility:** No behavior change when `template_mode_enabled=false` and `presentation_edit_enabled=false` (both default) and when the JSON spec has no `reference_file_id`. Classic deck generation matches the v1.0.3 path.

### Added

- **`inspect_slides`** — factual inventory JSON for an uploaded template `.pptx` (shape ids, bbox, `safe_zone`, `text_verbatim`, optional `images[]`). Call it before `generate_slides` when a `.pptx` is the base.
- **Template Mode in `generate_slides`** — clone decorative shapes from a reference file and overlay semantic content in the reference safe zone.
- Spec fields: `reference_file_id`, `template_mapping`, `template_edits` (including `drop_ids` from inspect).
- Admin valves: `template_mode_enabled` (default `false`), `inspect_slides_enabled`, `inspect_extract_images`.
- **`edit_presentation`** — post-generate edits on an existing `.pptx` via `presentation_file_id` (Files API UUID) and `operations[]`.
- Operations: `set_text`, `shrink_font`, `fit_text`, `enable_autofit`, `replace_text_and_fit`, `split_text`, `resize_shape`, `set_table_cell`, `set_font_pt`.
- **`shrink_font` / `fit_text` modes:** `fit` (default), `to_min`, `target` (+ `target_pt`).
- Admin valves: `presentation_edit_enabled` (default `false`), `presentation_edit_min_font_pt`, `presentation_edit_shrink_step_pt`, `presentation_edit_progress_every`, `presentation_edit_warn_noop` (default `false`).

### Behavior

- Valve **ON** + invalid or inaccessible reference → clear error, **no** output `.pptx`.
- Valve **OFF** + `reference_file_id` in JSON → classic deck; reference ignored with a visible note.
- **`generate_slides` does not** auto-detect chat attachments for templates; only an explicit `reference_file_id` applies Template Mode.
- Edit requires an explicit UUID (`presentation_file_id`); cache-only URLs are rejected. The whole batch fails on the first invalid operation.
- Font size reads the run `sz` plus paragraph `endParaRPr` / `defRPr` fallbacks.

### Docs and tests

- [Template Mode implementation plan](doc/template-mode-implementation-plan.md)
- [Model hints (Neura)](doc/neura-template-mode-model-hints.md)
- [Presentation edit plan](doc/presentation-edit-implementation-plan.md)
- [Presentation edit hints](doc/neura-presentation-edit-hints.md)
- `examples/edit_presentation_smoke.py`, `examples/edit_presentation_acl_smoke.py`

## [1.0.3] — Files API baseline

- Stable Files API save path and classic layout engine (NF5 golden baseline in `examples/golden/`).
