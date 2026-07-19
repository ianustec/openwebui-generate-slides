# Generate Slides — Native PPTX engine for Open WebUI

> Part of **[NEURA Office](https://github.com/ianustec/neura-office)**: native Word, PowerPoint and Excel for Open WebUI.

A [Open WebUI](https://github.com/open-webui/open-webui) **Tool** that generates
**native PowerPoint (.pptx)** presentations from a JSON spec produced by the model.
It doesn't export HTML or images: it builds the slides directly with `python-pptx`,
with a coherent visual system, layered decorative shapes, **native charts**, icons
inside circles and rich layouts.

The resulting file is saved through Open WebUI's **Files API** (with a `/cache/files`
fallback) and a clickable **download link** appears in the chat.

> License: MIT · Author: [IANUSTEC](https://ianustec.com)

![Preview](assets/hero.png)

*Slides generated from the example [`examples/deck.json`](examples/deck.json) → [`examples/demo_tech_deck.pptx`](examples/demo_tech_deck.pptx).*

## Features

- **Native, editable .pptx** (text, charts and shapes are editable in PowerPoint/Keynote/LibreOffice).
- **Native Office charts**: `bar`, `line`, `area`, `pie`, `doughnut`, `radar`, `stacked_bar`.
- **~25 ready-made layouts**: cover, sections, bullets, columns/comparison, KPIs, timeline,
  process flow, icon lists, icon grids, pillars, quote, alert, tables,
  diagrams (funnel, pyramid, cycle, quadrant, bullseye) and image layouts.
- **Curated themes** + custom accent (`theme:"auto"` infers the theme from the content).
- **Lucide-style icons** bundled in the file (no network dependency for icons).
- **Optional images** from Unsplash (with a key) or generated via Open WebUI.
- **Single-file**: one self-contained `.py`, ready to paste into the Tools registry.

## Requirements

- Open WebUI `>= 0.4.0`
- Python: `python-pptx`, `pillow` (declared in the frontmatter → Open WebUI installs them automatically)
- Optional: `httpx` (fetch images from URL/Unsplash)

## Installation

### Option A — from the Open WebUI community
1. Open the tool page on the Open WebUI community site.
2. Click **Get** / **Import** to your instance.

### Option B — manual
1. In your Open WebUI instance go to **Workspace → Tools → +**.
2. Paste the contents of [`generate_slides.py`](generate_slides.py).
3. Save. The declared dependencies are installed on first use.
4. Enable the tool for the model (or chat) that should use it.

## Usage

The model calls `generate_slides(content)`, where `content` is a **single JSON string**.
Minimal structure:

```json
{
  "title": "Presentation title",
  "subtitle": "Optional subtitle",
  "author": "Author / company",
  "theme": "auto",
  "footer": "Footer label",
  "slides": [
    { "layout": "cover", "title": "...", "subtitle": "...", "icon": "cpu" },
    { "layout": "kpi_row", "title": "...", "stats": [ { "value": "-40%", "label": "..." } ] }
  ]
}
```

See the full example in [`examples/deck.json`](examples/deck.json).

### Themes
`auto` (default, inferred from content) · `midnight` · `forest` · `ocean` · `coral`
· `terracotta` · `teal` · `berry` · `sage` · `cherry` · `charcoal` · `slate`.
You can force the accent with `"accent": "#C99A3B"`.

### Available layouts (main fields)

| Layout | Main fields |
|---|---|
| `cover` | `title`, `subtitle`, `author`, `eyebrow`, `icon`, `date`, `chips[]` |
| `section` | `number` (`"01"`), `eyebrow`, `title`, `lead` |
| `title_bullets` | `title`, `eyebrow`, `bullets[]` |
| `title_body` | `title`, `eyebrow`, `body` (paragraphs separated by `\n`) |
| `two_column_text` / `comparison_two` | `left{}`, `right{}` or `columns[]` (`heading`, `icon`, `points[]`, `highlight`, `badge`) |
| `kpi_row` | `stats[]` with `{value, label, change}` |
| `timeline_horizontal` / `process_flow` | `steps[]` with `{when, title, description}` |
| `icon_list_vertical` | `items[]` with `{icon, title, description}` |
| `icon_grid_2x2` / `icon_grid_3` / `pillars` | `items[]` with `{icon, title, description}` |
| `chart` | `chart_type`, `labels[]`, `values[]` or `datasets[]{label,data[]}`, `insight[]` |
| `funnel` / `pyramid` / `cycle` / `quadrant` / `bullseye` | `nodes[]` with `{label, description}` |
| `quote` | `quote`, `author`, `role` |
| `alert` | `title`, `level` (`info`\|`tip`\|`warning`\|`danger`), `body` or `bullets[]` |
| `table` | `headers[]`, `rows[]` |
| `text_image_right` / `image_left_text_right` | `title`, `bullets[]`/`body`, `image_hint` or `image_url` or `base64` |
| `image_full_caption` | `title`, `subtitle`, `image_hint`/`image_url` |
| `closing` | `title`, `eyebrow`, `takeaways[]`, `contact` |

## Screenshots

| Cover | KPI row |
|---|---|
| ![Cover](assets/cover.png) | ![KPI](assets/kpi.png) |
| **Native chart** | **Funnel** |
| ![Chart](assets/chart.png) | ![Funnel](assets/funnel.png) |

## Valves (configuration)

| Valve | Default | Description |
|---|---|---|
| `default_theme` | `auto` | Default theme when the spec doesn't set one |
| `footer_label` | `""` | Default footer (overridden by `spec.footer`) |
| `unsplash_access_key` | `""` | Unsplash key for stock images (optional) |
| `image_generation` | `false` | Enable AI image generation via Open WebUI |
| `max_image_px` | `1600` | Maximum image width |
| `emit_status` | `true` | Emit status events in chat |
| `pptx_export_dir` | `/app/backend/data/cache/files` | Fallback directory for saving |

## How it works

1. The model produces the JSON spec and calls `generate_slides`.
2. The engine resolves theme/accent, prefetches images (if any image layouts are present),
   and for each slide invokes the corresponding layout renderer.
3. Shapes, text and charts are written as native OOXML objects.
4. The `.pptx` is saved via the Files API (fallback `/cache/files`) and the link is returned in chat.

## Local development / testing

Requires `python-pptx` and (optionally) `pillow`/`httpx`:

```bash
pip install python-pptx pillow httpx
python examples/build.py   # → examples/demo_tech_deck.pptx
```

The file is designed to run inside Open WebUI: the `open_webui.*` imports are optional
and the tool degrades gracefully when they're missing (handy for isolated render tests).

## Contributing

Issues and PRs welcome. Please keep the file **single-file** and free of mandatory
network dependencies for the core features.

## License

[MIT](LICENSE) © IANUSTEC
