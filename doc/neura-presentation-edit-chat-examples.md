# Neura — chat examples for Presentation Edit

Worked examples for the model after the user **opened or downloaded** a deck from `generate_slides`. Use **`edit_presentation`**, not a full **`generate_slides`** rerun.

Model hints: [neura-presentation-edit-hints.md](./neura-presentation-edit-hints.md) · Technical plan §12: [presentation-edit-implementation-plan.md](./presentation-edit-implementation-plan.md#12-appendice-esempio-end-to-end).

**Requires:** admin valve `presentation_edit_enabled=true` on the pilot workspace and a Files API **UUID** from the last generate/edit (`presentation_file_id`).

---

## Scenario A — Slide 1 title too large

### Premessa

- Chat already produced a deck; tool message includes `file_id` (UUID), e.g. `e4f5a6b7-c8d9-4012-e345-678901234567`.
- User opened the `.pptx` and sees overlap on **slide 1** (cover).

### User (chat)

> Sulla slide 1 il titolo è troppo grande e copre il sottotitolo. Riduci il titolo e lascia il sottotitolo com’è.

### Model (reasoning, no tool yet)

1. User slide **1** → JSON `"slide": 1` (PowerPoint numbering; `EDIT_SLIDE_INDEX_ORIGIN=1` in `generate_slides.py`).
2. `presentation_file_id` = UUID from the **last** `generate_slides` or `edit_presentation` response — copy verbatim.
3. **Do not** call `generate_slides` or use `reference_file_id` from template inspect.
4. Optional: `inspect_slides` with **`presentation_file_id`** (output deck) to find title vs subtitle `shape_id` values. Do not paste full inspect JSON to the user.

### Model → `edit_presentation`

Example with **`replace_text_and_fit`** (preferred when only the title size must change):

```json
{
  "presentation_file_id": "e4f5a6b7-c8d9-4012-e345-678901234567",
  "title": "My deck (title fix)",
  "operations": [
    {
      "op": "replace_text_and_fit",
      "slide": 1,
      "shape_id": 783,
      "text": "Shorter main title"
    }
  ]
}
```

Alternative (explicit shrink after setting text):

```json
{
  "presentation_file_id": "e4f5a6b7-c8d9-4012-e345-678901234567",
  "operations": [
    {
      "op": "set_text",
      "slide": 1,
      "shape_id": 783,
      "text": "Shorter main title"
    },
    {
      "op": "shrink_font",
      "slide": 1,
      "shape_id": 783,
      "min_pt": 18,
      "step_pt": 2
    }
  ]
}
```

Shape id **783** is **illustrative** — take ids from inspect on the **generated** file.

### After the tool

- User downloads the edited file.
- Further tweaks in the same chat → call `edit_presentation` again with the **new** `file_id` from the edit response (PE7 chain).

### Anti-patterns

- Regenerating the whole deck to shrink one title.
- Using template inspect `file_id` as `presentation_file_id`.

---

## Scenario B — Split title and body on slide 4

### Premessa

- Same UUID discipline as scenario A.
- User refers to **slide 4** → JSON `"slide": 3`.

### User (chat)

> Sulla slide 4 separa titolo e paragrafo: titolo in alto, testo descrittivo nel riquadro sotto.

### Model (reasoning)

1. Inspect the **output** deck if unsure: two text placeholders vs one combined text box.
2. **Two placeholders** (two `shape_id` from inspect): two `set_text` operations.
3. **One shape** with title+body: `split_text` into a second placeholder, then optional `resize_shape` for small overlap.

### Model → `edit_presentation` (two placeholders)

```json
{
  "presentation_file_id": "e4f5a6b7-c8d9-4012-e345-678901234567",
  "title": "My deck (slide 4 split)",
  "operations": [
    {
      "op": "set_text",
      "slide": 4,
      "shape_id": 801,
      "text": "SECTION TITLE"
    },
    {
      "op": "set_text",
      "slide": 4,
      "shape_id": 802,
      "text": "Body paragraph stays in the lower box. Edit only wording if the user asked for text changes."
    }
  ]
}
```

### Model → `edit_presentation` (single combined shape)

```json
{
  "presentation_file_id": "e4f5a6b7-c8d9-4012-e345-678901234567",
  "operations": [
    {
      "op": "split_text",
      "slide": 4,
      "from_shape_id": 801,
      "to_shape_id": 802,
      "mode": "first_paragraph"
    },
    {
      "op": "resize_shape",
      "slide": 4,
      "shape_id": 801,
      "delta_height_in": -0.15
    }
  ]
}
```

Ids **801/802** are **illustrative**. `mode` may be `first_line` if the title is a single line without a paragraph break.

### PE7 — second edit on new file id

After scenario B succeeds, if the user asks for a small wording tweak on the same slide:

1. Use `presentation_file_id` from the **edit** tool message (not the original generate id).
2. One more `edit_presentation` with e.g. a single `set_text` on `shape_id` 802, `"slide": 3`.

This validates the generate → edit → edit chain in production (manual smoke in rollout runbook).

---

## Scenario C — Many titles smaller (Montessori-style)

### User (chat)

> Sulle slide 2, 3, 5, 11, 14 e 19 riduci i titoli. Sulle slide 8, 16, … riduci i sottotitoli verde acqua.

### Model

- Use **`set_font_pt`** (or `shrink_font` with `"mode": "to_min"`) — **not** default `shrink_font` alone (no-op when text already fits).
- One op per `(slide, shape_id)` from inspect on the **output** UUID; verify each id exists on that slide.

```json
{
  "presentation_file_id": "<uuid>",
  "operations": [
    { "op": "set_font_pt", "slide": 2, "shape_id": 1100, "font_pt": 28 },
    { "op": "set_font_pt", "slide": 3, "shape_id": 1105, "font_pt": 28 },
    { "op": "set_font_pt", "slide": 5, "shape_id": 1462, "font_pt": 28 }
  ]
}
```

Alternative for the same shapes:

```json
{
  "op": "shrink_font",
  "slide": 2,
  "shape_id": 1100,
  "mode": "to_min",
  "min_pt": 22,
  "step_pt": 2
}
```

---

## Quick mapping (user language → JSON)

| User says (PowerPoint) | JSON `slide` (origin=1) |
|------------------------|-------------------------|
| slide 1 | 1 |
| slide 4 | 4 |
| slide 12 | 12 |

From inspect: JSON `slide` = `slides[].index + 1` when origin is 1.
