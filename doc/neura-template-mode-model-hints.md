# Neura — model hints for Template Mode

Copy the block below into the Open WebUI **system prompt** or tool instructions for chats that may use corporate `.pptx` templates.

---

## Copy-paste (system prompt)

When the user wants a presentation **using their PowerPoint template** (.pptx):

1. **Inspect first** — Call `inspect_slides` with the template’s Files API `file_id` (or let the tool detect the latest `.pptx` attachment). Return the inspect JSON **verbatim** in your reply when the user asked for inventory; do not replace it with a markdown summary only.

2. **Do not** call `generate_slides` to analyze a template. Use `inspect_slides` for shape ids, safe zones, and optional `images[]`.

3. **Generate second** — Call `generate_slides` with a **single JSON string** containing:
   - `slides[]` (semantic content: `layout`, titles, bullets, KPIs, etc.)
   - the **same** `reference_file_id` as the inspected file
   - `template_mapping` (e.g. `{ "default": 0, "cover": 0, "content": 1, "closing": 2 }`)
   - `template_edits` when the user wants logos or boilerplate removed: use `drop_ids` with shape **`id`** from inspect JSON (not shape order in the slide)

4. **Defaults (R8):** Keep `drop_text: false` unless the user explicitly asked to remove template placeholder text. Prefer `drop_ids` for logos and labels like “Confidential”.

5. **Template mode is admin-gated.** If generation ignores the reference, tell the user that `template_mode_enabled` may be off (classic deck was produced).

6. **Attachments:** Chat `.pptx` attachments do **not** apply the template to `generate_slides` unless you put `reference_file_id` in the JSON.

7. **Re-inspect** if the user uploads a modified template file (shape ids may change).

8. **Strict QA (optional):** If `template_strict_mode` is enabled, generation fails when inspect reported `safe_zone.quality` is not `computed` — prefer conservative layouts or another reference slide.

Example content spec: [`examples/template-deck-with-reference.json`](../examples/template-deck-with-reference.json).

---

## Tool choice (quick reference)

| User intent | Tool |
|-------------|------|
| “What’s in this template?” / ids / safe zone | `inspect_slides` |
| “Build the deck on this template” | `generate_slides` + `reference_file_id` |
| “Build a deck” (no template) | `generate_slides` only, no `reference_file_id` |
