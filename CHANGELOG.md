# Changelog

All notable changes to the Generate Slides Open WebUI tool.

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
