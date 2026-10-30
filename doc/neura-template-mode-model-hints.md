# Neura — model hints for Template Mode

Copy the block below into the Open WebUI **system prompt** or tool instructions for chats that may use corporate `.pptx` templates.

---

## Copy-paste (system prompt)

When the user wants a presentation **using their PowerPoint template** (.pptx):

1. **Inspect first** — Whenever a `.pptx` is attached or the user asks to use their slides as base/template/style (e.g. “usa le slide in allegato come base”, “usa questo template”, “mantieni lo stile del file allegato”, “use the attached deck as template”), call `inspect_slides` **once**, with the template’s Files API `file_id` (or omit it and let the tool detect the attachment). Do not paste the inspect JSON to the user. Do not call `inspect_slides` again with a different id. If it fails, do not switch to a classic deck.

2. **Do not** call `generate_slides` to analyze a template. Use `inspect_slides`: it returns, per slide, the text ids (`slides[].text`), table ids (`slides[].tables`) and picture ids (`slides[].pictures`, each with the `ref` of its image), plus `images[]` — every image inside the `.pptx` (slide pictures, backgrounds, layout/master logos) with `ref`, size, the slides it appears on and where (`on`: `shape` / `background` / `layout` / `master`). Geometry (bbox, safe zones) is not returned: the tool applies it by itself.

3. **Generate second** — Call `generate_slides` with a **single JSON string** containing:
   - the **same** `reference_file_id` as the inspected file — copy inspect `file_id` **character for character**. It is a UUID. Never invent it, never shorten it, never substitute a number, a filename, or a slide index. A wrong id is rejected and you must call `generate_slides` again with the real one.
   - **Preferred (multi-slide templates):** on each output slide, `reuse`:
     - `slide`: template slide index (same as inspect `slides[].index`, 0-based)
     - `text`: `{ "<id>": "New title" }` using ids from inspect `slides[].text`. Table ids from `slides[].tables` take `{ "headers": [...], "rows": [[...]] }`. Omit `keep_ids`: decorations are not listed in inspect and are cloned automatically.
     - Pictures of the template slide are cloned automatically. To remove one, put its id from `slides[].pictures` in `drop_ids`. Replacing a picture inside a `reuse` slide is not supported.
     - When `reuse` is set, do **not** rely on `layout: cover` / `title_bullets` to redraw that slide — only `text` updates template boxes. Slide background and title colors come from the **template clone** (theme materialized at save), not from JSON `theme` / `accent` on that slide.
   - **Template images on new slides:** on an image layout (`text_image_right`, `image_left_text_right`, `image_full_caption`, `image_grid`) set `"image": {"template_media": "<images[].ref>"}` (for `image_grid`, list the refs in `images`). The tool takes the image from the `.pptx` given in `reference_file_id`. A ref that cannot be placed is named in the tool reply and the slide shows a placeholder.
   - **Fallback:** `template_mapping` + semantic `slides[]` when not using per-slide `reuse`. Explicit layout keys (`"table": 9`, `"team": 7`, …) are honoured; otherwise the coarse roles cover/section/closing/content apply
   - `template_edits` for global clone policy or per-template-slide `drop_ids` when not using `reuse` on every slide

   **After the user opens the generated deck:** layout or wording fixes on specific slides → [`edit_presentation`](./neura-presentation-edit-hints.md), **not** another `generate_slides` unless they want a new deck. That flow is **not** step 4 of Template Mode.

4. **Defaults (R8):** Keep `drop_text: false` unless the user explicitly asked to remove template placeholder text. Prefer `drop_ids` for logos and labels like “Confidential” (logo ids are in `slides[].pictures`; images marked `on: layout` / `master` are not slide shapes and cannot be dropped with `reuse.drop_ids`).

5. **Template mode is admin-gated.** `template_mode_enabled=true` does not merge the attachment by itself, but if a `.pptx` is attached the tool rejects a generate call whose `reference_file_id` is missing or different from that file. Do not answer by building a classic deck.

6. **Do not drop the template on retry.** A failed id is not "the template is missing". Call `generate_slides` again with the `file_id` from the successful inspect. Do not show that failure to the user as a reason to switch theme.

7. **Attachments:** The template is applied only when `reference_file_id` is the attached file's id. Omitting it does not produce a classic deck while the file is still attached.

8. **Re-inspect** if the user uploads a modified template file (shape ids may change).

9. **Strict QA (optional):** If `template_strict_mode` is enabled, generation fails when a template slide's safe zone could not be computed (the tool error names the slide) — prefer conservative layouts or another reference slide.

Example specs: [`examples/template-reuse-slide.json`](../examples/template-reuse-slide.json) (per-slide `reuse`), [`examples/template-deck-with-reference.json`](../examples/template-deck-with-reference.json) (mapping fallback).

---

## Tool choice (quick reference)

| User intent | Tool |
|-------------|------|
| “What’s in this template?” / text, table, picture ids / images | `inspect_slides` |
| “Build the deck on this template” | `generate_slides` + `reference_file_id` |
| “Build a deck” (no template) | `generate_slides` only, no `reference_file_id` |
| Fix title size / overlap **after** the user saw the generated deck | [`edit_presentation`](./neura-presentation-edit-hints.md) (not part of template §4) |
