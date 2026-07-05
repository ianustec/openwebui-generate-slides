# Assets

Screenshots used by the README and the Reddit post.

- `cover.png` — cover slide
- `kpi.png` — `kpi_row` layout
- `chart.png` — `chart` layout
- `funnel.png` — `funnel` diagram
- `hero.png` — 2x2 preview grid (used as the README banner and Reddit image)

Regenerate them from the example deck:

```bash
python examples/build.py                                   # -> examples/demo_tech_deck.pptx
soffice --headless --convert-to pdf examples/demo_tech_deck.pptx
pdftoppm -png -r 150 demo_tech_deck.pdf slide              # one PNG per slide
```
