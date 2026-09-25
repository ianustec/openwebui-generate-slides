# Changelog

All notable changes to the Generate Slides Open WebUI tool.

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
- [Model workflow hints (Neura)](doc/neura-template-mode-model-hints.md)

## [1.0.3] — Files API baseline

- Stable Files API save path and classic layout engine (NF5 golden baseline in `examples/golden/`).
