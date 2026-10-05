# Neura — model hints for Presentation Edit (post-generate)

Copy the block below into the Open WebUI **system prompt** or tool instructions for chats where the user **already received** a deck from `generate_slides` and asks for **small visual/text fixes** (title too large, overlapping text, split title vs body).

This workflow is **separate** from [Template Mode](./template-mode-implementation-plan.md) (inspect → generate). It is **not** step 4 of that flow. For building decks from a template, use [neura-template-mode-model-hints.md](./neura-template-mode-model-hints.md).

Technical plan: [presentation-edit-implementation-plan.md](./presentation-edit-implementation-plan.md).

**Tool name (Open WebUI):** `edit_presentation` only. Do **not** use `edit_slides` (deprecated alias).

**Admin:** `presentation_edit_enabled=true` is required on pilot workspaces once the tool is deployed (v1.0.4). Independent of `template_mode_enabled`.

**Chat examples (training):** [neura-presentation-edit-chat-examples.md](./neura-presentation-edit-chat-examples.md) — slide 1 title too large; split slide 4.

---

## Neura rollout (2026) — mandatory one-liner

Paste this into the workspace **system prompt** (alongside the block below) when Presentation Edit is enabled on the pilot:

> After the user has downloaded or opened a deck you already generated, use **only** `edit_presentation` for visual or text fixes on specific slides. **Do not** call `generate_slides` again unless they explicitly ask for a **new** deck.

---

## Copy-paste (system prompt)

When the user has **downloaded or opened** a presentation you already generated and asks to **fix layout or wording on specific slides** (e.g. “slide 1 title too big”, “split title and description on slide 4”):

1. **Do not** call `generate_slides` again for these fixes. Regenerating risks losing the cloned look and wastes work. Use **`edit_presentation`** when the tool is available and the valve is on.

2. **`presentation_file_id` (required)** — Use the Files API **UUID** of the **generated output** deck (from the last successful `generate_slides` or `edit_presentation` tool message: `file_id` on the file/link). Copy it **character for character**.
   - **Never** use `reference_file_id` or the template’s inspect `file_id` for edit.
   - **Never** invent, shorten, or substitute a filename or slide index for the id.
   - If the user only has a `/cache/files/...` link **without** a UUID, **do not** call `edit_presentation`. Explain that edit needs Files API upload (production) or a regenerated deck that returns a `file_id`.

3. **Optional inspect on the output deck** — If you need `shape_id` or current text on the **generated** file, call `inspect_slides` with **`presentation_file_id`** (the output UUID), **not** the template id. Do not paste the full inspect JSON to the user. Reuse the same inspect tool as template mode; the file being inspected is different.

4. **Slide numbers** — In `edit_presentation` JSON, use the **same number as PowerPoint / LibreOffice** (slide 1 → `"slide": 1`, slide 4 → `"slide": 4`). The tool maps this via `EDIT_SLIDE_INDEX_ORIGIN` in `generate_slides.py` (currently **1**). **Do not subtract 1** when the user quotes UI slide numbers. `inspect_slides` still reports `slides[].index` **0-based**; for edit JSON use **`index + 1`** when origin is 1.

5. **Build edit JSON** — Single JSON string in `content` for `edit_presentation`:

```json
{
  "presentation_file_id": "<uuid from last generate/edit>",
  "title": "Optional save title",
  "operations": [
    {
      "op": "set_text",
      "slide": 1,
      "shape_id": 783,
      "text": "Shorter title"
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

6. **Font size intent** — Do **not** use default **`shrink_font`** when the user only wants text **smaller** but it already fits in the box (typical after template reuse):

   | User intent | Operation |
   |-------------|-----------|
   | Overlap / text does not fit | `shrink_font` (default `mode`: **fit**) or `replace_text_and_fit` |
   | “Titolo/sottotitolo più piccolo” (aesthetic) | **`set_font_pt`** with `font_pt` (e.g. 28 / 14), or **`shrink_font`** with `"mode": "to_min"` |
   | Shrink to a specific size | `shrink_font` with `"mode": "target"` and `target_pt` |

   **`shape_id` must be on the same PowerPoint slide number** as in JSON (`inspect.slides[].index + 1` when `EDIT_SLIDE_INDEX_ORIGIN=1`).

7. **Same shape: text + fit** — Prefer **`replace_text_and_fit`** (one op) instead of `set_text` then `shrink_font`. For “adatta al riquadro / come PowerPoint”, use **`enable_autofit`** on that `shape_id` (PowerPoint adjusts on open).

8. **Split title / body** — If the template has **two** placeholders, use two `set_text` ops with two ids from inspect on the **output** file. If **one** shape holds title+body, use **`split_text`** with `from_shape_id`, `to_shape_id`, and `mode`: `first_line` or `first_paragraph`. Optional **`resize_shape`** with `delta_height_in` for small vertical overlap (conservative cap). Tables: **`set_table_cell`** with `row`, `col`, `text` (0-based).

9. **After edit** — Use the **new** `file_id` from the edit tool response for any further edits in the same chat.

10. **When edit is unavailable** — If the tool is missing from the workspace, say edit is not installed yet. If the tool exists but **`presentation_edit_enabled` is false**, say an admin must enable the valve on this workspace; do **not** silently rerun `generate_slides` for a layout tweak. Only use `generate_slides` when the user clearly wants a **new** deck.

---

## Tool choice (quick reference)

| User intent | Tool |
|-------------|------|
| Build or rebuild deck (with or without template) | `generate_slides` (+ template workflow as in template hints) |
| Inventory template before first generate | `inspect_slides` on **template** `file_id` |
| Fix title size / text / overlap **after** user saw the output | `edit_presentation` with **`presentation_file_id`** |
| Need shape ids on the **generated** deck | `inspect_slides` on **`presentation_file_id`** |

---

## Anti-patterns

- Calling `generate_slides` to shrink a title or move one paragraph.
- Using `reference_file_id` or template inspect id as `presentation_file_id`.
- Calling `inspect_slides` on the template when the user is editing the **output** deck.
- Using `set_text` / `shrink_font` on table shapes — use `set_table_cell` instead.
- Using only `shrink_font` (fit) for “make titles smaller” on many slides — use `set_font_pt` or `mode: to_min`.
- Trying to split across one shape without `split_text` when two placeholders do not exist.

---

## Rollout pilot Neura (admin + ops)

Use this checklist when enabling Presentation Edit on an **internal pilot** workspace. Keep `presentation_edit_enabled=false` on customer instances until product sign-off.

### Pre-deploy (repo / CI)

1. Plugin version **≥ 1.0.4** ([`generate_slides.py`](../generate_slides.py)).
2. `examples/edit_presentation_smoke.py` PASS (Docker):  
   `docker run --rm -v "$PWD:/w" -w /w python:3.12-slim bash -lc 'pip install -q -r examples/requirements-dev.txt && python examples/edit_presentation_smoke.py'`
3. Review [§13 checklist pre-merge](./presentation-edit-implementation-plan.md#13-checklist-pre-merge-produzione-neura) in the implementation plan.

### Deploy on Open WebUI

1. **Workspace → Tools** → paste updated `generate_slides.py` → **Save** (same plugin as generate/inspect).
2. Set valve **`presentation_edit_enabled=true`** on the pilot workspace only.
3. Paste [Neura rollout one-liner](#neura-rollout-2026--mandatory-one-liner) + [Copy-paste (system prompt)](#copy-paste-system-prompt) into model instructions for that workspace.

### Manual smoke (PE7)

1. `generate_slides` → confirm response includes Files API **`file_id`** (UUID).
2. `edit_presentation` with one op (e.g. `shrink_font` or `set_text`) using that UUID as `presentation_file_id`.
3. Second `edit_presentation` using the **new** `file_id` from step 2 → download both outputs.

See [chat examples — PE7](./neura-presentation-edit-chat-examples.md#pe7--second-edit-on-new-file-id).

### Rollback

- Set **`presentation_edit_enabled=false`**, or pin the previous tool file version (same as Template Mode rollback in README).

### Log monitoring (weeks 1–2)

Search Open WebUI / tool container logs for `[edit_presentation]`.

| Signal | What to look for |
|--------|------------------|
| Healthy calls | `log.info` lines with `ops_failed=0` and `ops_applied` > 0 |
| Failed op | `log.warning` with `ops_failed=1`; user/tool message often contains `operations[N]` |
| Bad slide index | `slide index … out of range` (remind model: user slide N → JSON N−1) |
| Bad shape id | messages mentioning `shape` / id not found — rerun `inspect_slides` on **output** UUID |
| Template id misuse | retry text about template / `reference_file_id` (#7B) — reinforce [chat examples](./neura-presentation-edit-chat-examples.md) |

**Weekly log review template** (copy into your ops notes):

| Week ending | Edit calls (approx.) | Failures | Top error messages | Action |
|-------------|----------------------|----------|--------------------|--------|
| YYYY-MM-DD | | | | hint / training / bug |
