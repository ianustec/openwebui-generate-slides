# Template Mode — Piano di implementazione

> **Issue:** [#2 Support custom themes](https://github.com/ianustec/openwebui-generate-slides/issues/2) (estesa)  
> **Versione base:** `generate_slides.py` v1.0.3  
> **Stato:** Progettazione approvata — implementazione da avviare  
> **Ultimo aggiornamento:** 2026-09-22  
> **Contesto deploy:** Neura è **già in produzione** con clienti attivi — Template Mode deve essere **opt-in esplicito** e **inattivo di default** fino ad abilitazione admin.

---

## Indice

1. [Executive summary](#1-executive-summary)
2. [Contesto e requisiti](#2-contesto-e-requisiti)
3. [Stato attuale vs obiettivo](#3-stato-attuale-vs-obiettivo)
4. [Architettura: flusso a 3 passi](#4-architettura-flusso-a-3-passi)
5. [Principi di design](#5-principi-di-design) (incl. [§5.1 Strategia refactor](#51-strategia-refactor-geometria-anti-regression))
6. [Contratti dati](#6-contratti-dati)
7. [Componenti software](#7-componenti-software)
8. [Cosa NON implementare (v1)](#8-cosa-non-implementare-v1)
9. [Rischi e mitigazioni](#9-rischi-e-mitigazioni)
10. [Criteri di accettazione](#10-criteri-di-accettazione)
11. [Piano di implementazione a fasi](#11-piano-di-implementazione-a-fasi)
12. [Appendice: esempio end-to-end](#12-appendice-esempio-end-to-end)
13. [Checklist pre-merge (produzione Neura)](#13-checklist-pre-merge-produzione-neura)

---

## 1. Executive summary

### Obiettivo

Permettere all’utente di allegare un file **`.pptx` di esempio** (slide vuota a livello testuale, solo elementi grafici) e chiedere a Neura di generare una **nuova presentazione** sull’argomento richiesto, **mantenendo lo stesso look** (decorazioni, logo, sfondo, posizioni) e **popolando** l’area centrale con contenuto generato dall’AI.

### Soluzione adottata

Un **flusso a 3 passi** sulla stessa classe `Tools`, senza tool OWUI separati e senza cache in memoria:

| Passo | Attore | Output |
|-------|--------|--------|
| **1** | `inspect_slides` | Inventario JSON **fattuale** del file (id shape, bbox, testo, safe zone) |
| **2** | AI / utente | JSON di contenuto + `reference_file_id` + `template_mapping` + `template_edits` |
| **3** | `generate_slides` | Clone selettivo delle shape dal file originale + overlay contenuto in safe zone |

### Insight centrale

> Il `.pptx` allegato **è la cache**. L’AI passa solo **testo e decisioni** (`file_id`, `drop_ids`). Loghi, foglie e immagini embedded **non transitano nel prompt** — vengono ricopiati dal binario in `generate_slides`.

### Regola produzione (Neura / clienti esistenti)

| Regola | Motivo |
|--------|--------|
| Template Mode **solo** se `template_mode_enabled` (valve) **e** `reference_file_id` esplicito nel JSON | Evita attivazione accidentale |
| **Mai** auto-detect `.pptx` in `__messages__` dentro `generate_slides` | Un allegato “contesto” non deve cambiare il render classico |
| Auto-detect attachment **solo** in `inspect_slides` (opzionale) | Isola il rischio al passo 1 |
| Reference **invalido** con valve **ON** → **errore** (F9), nessun deck | Fail closed: l’utente deve sapere che il template non è stato applicato |
| `reference_file_id` presente ma valve **OFF** → path **classico** + log (NF8); opzionale messaggio visibile in risposta tool | Non confondere con F9: non è “reference invalido”, è feature disabilitata |
| Senza `reference_file_id` (o valve off) → output **byte-identico** (o equivalente) a v1.0.3 | R10 / NF5 |

---

## 2. Contesto e requisiti

### Origine

- **GitHub Issue #2:** “Support custom themes” — uso di temi esistenti o file `.potx`/`.pptx`.
- **Use case NEURA:** l’utente allega una slide template (es. cornice botanica/acquerello con area centrale bianca) e chiede un deck su argomento X con lo stesso template.

### Requisiti funzionali (decisioni prese)

| # | Requisito | Decisione |
|---|-----------|-----------|
| R1 | Formato input template | Solo **`.pptx`** (v1); no PNG/JPG come template |
| R2 | Slide di riferimento | **Una o più** slide nel file allegato |
| R3 | Mapping slide reference → ruolo | **Richiesta utente** (cover/content/closing); AI traduce in `template_mapping` |
| R4 | Layout ricchi (chart, KPI, funnel) | **Opzione B:** esistono ancora, **ridimensionati** dentro la safe zone; decorazioni angolari intatte |
| R5 | Footer e numeri pagina | **Mantenuti** (con adattamento posizione se overlap decorazioni) |
| R6 | Colori testo | **Automatici** per leggibilità sullo sfondo della safe zone |
| R7 | Safe zone | **Calcolo automatico** (max rettangolo vuoto centrato) |
| R8 | Testo placeholder nel template | **Ignorato e sovrascritto** (`drop_text` default) |
| R9 | Elementi da rimuovere (logo, watermark) | AI specifica `drop_ids` dopo `inspect`; tool applica meccanicamente |
| R10 | Modalità senza template | Comportamento **identico a v1.0.3** (nessuna breaking change) |
| R11 | Abilitazione feature | Valve admin `template_mode_enabled` — default **`false`** fino a rollout pilot |
| R12 | Attivazione template | **`reference_file_id` obbligatorio** nel JSON per `generate_slides`; niente inferenza da attachment in generate |

**Nota R1 — `.potx`:** in v1 il contratto documenta **`.pptx`**. I file `.potx` sono lo stesso formato OOXML e python-pptx li apre come presentazione; possono essere trattati come input equivalente se caricati su Files API (estensione diversa, stesso parser).

### Requisiti non funzionali

- File **single-file** (`generate_slides.py`) — convenzione repo.
- Degradazione graceful se Files API non disponibile (come v1.0.3) — **solo per `_save` output**, non per download reference (reference fail = errore).
- Compatibilità OWUI v0.6.x e v0.11.x (`_maybe_await` già presente).
- Log strutturato su errori parse/clone/fallback.
- Download reference: **solo file accessibili all’utente** corrente (ACL multi-tenant Neura).
- `_save`, `_prefetch_images`, `_ai_image` (issue #4): **non modificare** il comportamento esistente in v1.1.0 template work.

---

## 3. Stato attuale vs obiettivo

### Comportamento attuale (`generate_slides.py` v1.0.3)

```
JSON spec → _resolve_theme() → Presentation() vuota
           → slide_layouts[6] blank per ogni slide
           → _set_bg() + shape programmatiche (ovali, card, icone)
           → _save() via Files API
```

**Limiti rilevanti:**

- Nessun input `reference_file_id` o attachment `.pptx`.
- `__messages__` accettato ma **non usato**.
- Geometria fissa: `MARGIN`, `CONTENT_TOP`, `FOOTER_Y`, `SLIDE_W_IN`/`SLIDE_H_IN`.
- Theme solo da `_PALETTES` + override JSON (`palette{}`, `accent` — poco documentati).
- `_set_bg()` **sovrascrive** sempre lo sfondo slide.

### Comportamento target (template mode)

Attivo **solo se** (stessa condizione di §6.7):

`template_mode_enabled == true` **AND** `reference_file_id` non vuoto **AND** download + parse OK.

```
JSON spec + reference_file_id + template_edits  (+ valve ON)
  → download .pptx da Files API
  → parse shape + safe zone + theme XML
  → Presentation() nuova (dimensioni dal template)
  → per ogni slide output: clone background + decorazioni (filtrate da edits)
  → overlay contenuto in deck.frame (safe zone della slide reference scelta)
  → _save()
```

Se valve **OFF** o assenza di `reference_file_id`: pipeline §3 “Comportamento attuale” (v1.0.3).

---

## 4. Architettura: flusso a 3 passi

### Diagramma sequenza

```
Utente: allega template.pptx
        "PP su argomento X, usa questo template, togli logo e Confidential"

┌─────────────────────────────────────────────────────────────────┐
│  PASSO 1 — inspect_slides(file_id | attachment)                 │
│  → GET /api/v1/files/{id}/content                               │
│  → inventario JSON: shape id, kind, bbox, text_verbatim,        │
│    safe_zone, theme colors, slide_count                         │
└────────────────────────────┬────────────────────────────────────┘
                             │
                             ▼
┌─────────────────────────────────────────────────────────────────┐
│  PASSO 2 — AI                                                   │
│  → legge inventario + richiesta utente                          │
│  → produce: content JSON (slides[])                             │
│            + reference_file_id                                  │
│            + template_mapping                                   │
│            + template_edits.drop_ids                            │
│  → NON passa immagini/loghi (solo testo e id)                   │
└────────────────────────────┬────────────────────────────────────┘
                             │
                             ▼
┌─────────────────────────────────────────────────────────────────┐
│  PASSO 3 — generate_slides(content_json, reference_file_id)   │
│  → riapre STESSO file (source of truth binaria)                 │
│  → clona shape ammesse (es. 4,5,6 foglie; esclude 12,18)       │
│  → overlay titolo/KPI/chart in safe_zone                        │
│  → salva .pptx output via Files API                             │
└─────────────────────────────────────────────────────────────────┘
```

### Separazione responsabilità

| Attore | Responsabilità | Non fa |
|--------|----------------|--------|
| **`inspect_slides`** | Verità del file (OOXML → JSON fatti) | Clone, generazione, classificazione layout canonica |
| **AI** | Contenuto semantico, `drop_ids`, `template_mapping` | Trasporto pixel, ridisegno decorazioni |
| **`generate_slides`** | Download file, clone shape, overlay, save | Interpretazione “è un logo?” senza `drop_ids` |

### Perché NON un tool inverso `pptx → JSON generate_slides`

Il JSON di `generate_slides` descrive **contenuto semantico** (`layout: kpi_row`, `stats[]`), non la scenografia (posizioni foglie, PNG embedded). Invertire non è fedele al 100% e **perde la grafica**. `inspect` produce un **verbale**, non una spec di rendering.

### Perché NON cache/variabile temporanea

OWUI istanzia `Tools` per ogni request — `self._pack` muore a fine chiamata. Il file su **Files API** è l’unica cache persistente necessaria; `generate_slides` riapre `reference_file_id`.

---

## 5. Principi di design

1. **File = source of truth grafica.** Media in `ppt/media/`; clone via OOXML + relationship copy.
2. **AI = source of truth semantica.** Testi, argomento, quali id droppare.
3. **Inspect = solo fatti.** Niente `layout: kpi_row` nel body (solo in `hints` opzionali non vincolanti).
4. **Edit esplicito.** `drop_ids` / `keep_ids` — niente “capisce da solo” cosa è logo.
5. **Regression zero.** Senza `reference_file_id` (o valve off) → path attuale invariato.
6. **Safe zone geometrica.** Algoritmo deterministico; non vision AI in v1.
7. **Stesso repo, stessa classe.** `inspect_slides` + `generate_slides` con parser condiviso.
8. **Opt-in esplicito.** Template Mode = valve ON + `reference_file_id`; mai side-effect da attachment in generate.
9. **Fail closed su reference.** Download/parse reference fallito → `_error()`, non deck “classico” mascherato.
10. **Geometria classica preservata.** `_default_frame()` deve replicare esattamente `MARGIN` / `CONTENT_TOP` / `FOOTER_Y`; refactor anche `_eyebrow`, `_title`, `_footer` (non solo `_r_*`).
11. **Layout blank classico invariato.** `slide_layouts[6]` **solo** quando `template_mode=False`; con template usare `_find_blank_layout(prs)`.
12. **Due path geometria (produzione).** Path classico: codice legacy con `MARGIN` / `FOOTER_Y` / `SLIDE_W_IN` **invariato** finché NF5 è verde. Path template: `deck.frame` + helper dedicati. Evitare refactor “unificato” senza golden hash a ogni commit.

### 5.1 Strategia refactor geometria (anti-regression)

| Approccio | Descrizione | Raccomandazione |
|-----------|-------------|-----------------|
| **A — Dual path** | `if deck.template_mode:` usa `frame`; `else:` lascia costanti esistenti | **Preferito** per Neura in produzione |
| **B — Path unico** | Tutti i renderer usano sempre `deck.frame` via `_default_frame()` | Solo se NF5 (hash) gira **dopo ogni** modifica geometria |

**Nota:** molti renderer calcolano altezza con `top = _content_head(...)` e `h = FOOTER_Y - y - 0.35`, non solo `MARGIN`. Sostituire solo `MARGIN → frame.x` **non** garantisce equivalenza byte-a-byte.

Helper consigliato:

```python
def _content_bounds(deck, top_y: float) -> tuple[float, float, float, float]:
    """x, y, w, h per corpo contenuto sotto il titolo."""
    if deck.template_mode:
        return (deck.frame.x, top_y, deck.frame.w, deck.frame.y + deck.frame.h - top_y)
    # legacy — identico a v1.0.3
    return (MARGIN, top_y, SLIDE_W_IN - 2 * MARGIN, FOOTER_Y - top_y - 0.35)
```

---

## 6. Contratti dati

### 6.1 Output `inspect_slides`

```json
{
  "ok": true,
  "file_id": "abc123",
  "filename": "template.pptx",
  "slide_width_in": 13.333,
  "slide_height_in": 7.5,
  "slide_count": 2,
  "theme": {
    "dk1": "1A1A1A",
    "lt1": "FFFFFF",
    "accent1": "7BAE9E",
    "major_font": "Calibri",
    "minor_font": "Calibri"
  },
  "slides": [
    {
      "index": 0,
      "shape_count": 7,
      "background": { "type": "solid", "color": "FFFFFF" },
      "shapes": [
        {
          "id": 4,
          "kind": "picture",
          "name": "Picture 2",
          "bbox": { "x": 0.0, "y": 0.0, "w": 4.2, "h": 3.1 },
          "has_text": false,
          "text": null
        },
        {
          "id": 18,
          "kind": "textbox",
          "name": "TextBox 3",
          "bbox": { "x": 10.0, "y": 6.8, "w": 2.5, "h": 0.4 },
          "has_text": true,
          "text": "Confidential"
        }
      ],
      "text_verbatim": ["Click to add title", "Confidential"],
      "picture_count": 4,
      "safe_zone": {
        "x": 1.35,
        "y": 1.1,
        "w": 10.6,
        "h": 5.2,
        "method": "max_empty_rect"
      }
    }
  ],
  "hints": {
    "likely_roles": [{ "index": 0, "role": "content", "confidence": "low" }]
  }
}
```

**Regole:**

- Campi in root/`slides[]` = fatti misurabili.
- `hints` = interpretazione non vincolante per l’AI.
- `safe_zone.method` sempre presente se calcolata.

### 6.2 Input `generate_slides` (estensione spec JSON)

Campi **nuovi** (opzionali):

| Campo | Tipo | Descrizione |
|-------|------|-------------|
| `reference_file_id` | string | ID Files API del `.pptx` template |
| `template_mapping` | object | Mappa ruolo → indice slide reference (`cover`, `content`, `default`, …) |
| `template_edits` | object | Policy clone + `drop_ids` / `keep_ids` per slide |

Campi **esistenti** invariati: `title`, `slides[]`, `theme`, `footer`, layout fields, ecc.

### 6.3 Schema `template_edits`

```json
{
  "template_edits": {
    "defaults": {
      "drop_text": true,
      "drop_placeholders": true,
      "drop_offslide": true
    },
    "slides": [
      {
        "index": 0,
        "drop_ids": [12, 18],
        "keep_ids": null
      }
    ]
  }
}
```

**Ordine applicazione:**

1. Applica `defaults` (es. escludi tutte le shape con testo).
2. Se `keep_ids` definito → clona **solo** quella lista (meno eventuali drop).
3. Altrimenti → clona tutte le shape grafiche ammesse, meno `drop_ids`.
4. Escludi sempre shape fuori slide se `drop_offslide: true`.

### 6.4 Schema `template_mapping`

```json
{
  "template_mapping": {
    "default": 0,
    "cover": 0,
    "section": 0,
    "content": 1,
    "closing": 2
  }
}
```

Mapping layout interno → ruolo:

| Layout `generate_slides` | Ruolo mapping |
|--------------------------|---------------|
| `cover` | `cover` |
| `section` | `section` |
| `closing` | `closing` |
| Tutti gli altri | `content` (fallback: `default`) |

### 6.5 Identificatori shape (`drop_ids`)

- In inspect e in `template_edits`, gli id sono **`shape.shape_id`** di python-pptx (campo JSON `"id"`), **non** l’indice nell’iterazione `for shape in slide.shapes`.
- Stesso file, stessa sessione: id **stabili** tra `inspect_slides` e `generate_slides`.
- Se l’utente modifica e ri-carica il template, gli id possono cambiare → **richiedere nuovo inspect**.

### 6.6 `template_edits.slides[].index`

- `index` è l’indice della slide **nel file reference** (0-based), **non** l’indice della slide nel deck output.
- Esempio: 1 slide reference, 12 slide output → tutte le clonazioni usano regole con `"index": 0`.
- Le regole per `index: 0` si applicano **ogni volta** che `_pick_template_slide` sceglie quella slide reference.

### 6.7 Attivazione Template Mode (condizione AND)

Template Mode è attivo **solo se**:

```
self.valves.template_mode_enabled == True
AND spec.get("reference_file_id") non vuoto
AND download + parse reference OK
```

Altrimenti: pipeline classica v1.0.3.

**Valve OFF + `reference_file_id` nel JSON:** ignorare template; generare deck classico; **log warning** + (consigliato) breve nota nella risposta tool all’AI/utente: *“Template ignored: template_mode_enabled is false.”* (NF8 UX).

**Valve ON + download/parse fallito:** `_error()` — **nessun** `.pptx` (F9). Non fare fallback silenzioso a modalità classica.

### 6.8 Precedenza colori / theme

| Situazione | Precedenza |
|------------|------------|
| Nessun `reference_file_id` | Comportamento attuale: `_PALETTES` + `palette{}` / `accent` / `primary` |
| Template Mode attivo | Colori/font da theme XML reference → `_derive_readable_colors(safe_zone_bg)` → merge; `theme: "midnight"` nel JSON **non** sovrascrive sfondo/decorazioni clonate |
| Template Mode + `palette{}` esplicito | Opzionale v1: `palette{}` può affinare **solo** ink testo, non le shape clonate |

### 6.9 Safe zone per slide reference (multi-mapping)

Ogni slide **output** usa la safe zone della slide **reference** scelta da `_pick_template_slide(pack, layout, mapping)` — non una safe zone globale unica.

Esempio: `cover → 0`, `content → 1` → cover e KPI usano safe zone diverse se le due slide reference differiscono.

### 6.10 Edge case `template_edits` / mapping

| Caso | Comportamento v1 |
|------|------------------|
| `drop_ids` contiene id inesistente | Ignorare id + log debug (no crash) |
| `template_mapping` indice ≥ `slide_count` | Fail closed: `_error("Invalid template_mapping index")` |
| `keep_ids: []` per errore AI | Log warning; inspect può suggerire di non usare keep vuoto |

---

## 7. Componenti software

### 7.1 Nuova sezione: Template engine

Posizione: `generate_slides.py`, prima di `# Deck builder + renderers`.

| Simbolo | Tipo | Ruolo |
|---------|------|-------|
| `_BBox` | dataclass | Rettangolo in pollici (x, y, w, h) |
| `_Decoration` | dataclass | Shape da clonare (element XML, bbox, z_order) |
| `_SlideTemplate` | dataclass | Una slide reference parsata |
| `_TemplatePack` | dataclass | Intero file reference + mapping metadata |

| Funzione | Ruolo |
|----------|-------|
| `_load_reference_pptx(file_id, request, user)` | Download bytes da Files API con credenziali request; verifica ACL utente |
| `_find_pptx_attachment(messages)` | Scan attachment `.pptx` — **solo** per `inspect_slides`, **non** per `generate_slides` |
| `_inspect_master_decorations(prs, slide_index)` | (P1) Sfondo/logo su slide master se slide reference “vuota” |
| `_parse_reference_pptx(data: bytes)` | → `_TemplatePack` |
| `_extract_decorations(slide)` | Shape non-testo + bbox |
| `_is_text_shape(shape)` | Euristica textbox/placeholder |
| `_compute_safe_zone(slide, decorations)` | Max empty rect centrato |
| `_apply_template_edits(pack, edits)` | Filtra decorations per slide |
| `_clone_background(target, source)` | Copia fill/immagine sfondo |
| `_clone_decorations(target, decorations)` | Deep-copy XML + media parts |
| `_pick_template_slide(pack, layout, mapping)` | Sceglie slide reference |
| `_find_blank_layout(prs)` | Per nome ("Blank"), non indice 6 |
| `_derive_readable_colors(bg_hex)` | ink/muted per leggibilità |
| `_merge_theme_from_template(theme, pack)` | Colori/font da theme XML |

### 7.2 Nuovo metodo: `inspect_slides`

```python
async def inspect_slides(
    self,
    file_id: str = "",
    __messages__: Any = None,
    __request__: Any = None,
    __user__: Optional[dict] = None,
) -> str:
    """Restituisce inventario JSON fattuale del .pptx."""
```

### 7.3 Modifiche `_Deck`

```python
class _Deck:
    def __init__(self, prs, theme, footer_label, *,
                 template_pack: _TemplatePack | None = None):
        self.template_pack = template_pack
        self.template_mode = template_pack is not None
        self.frame = _default_frame()

    def blank(self, layout: str = "default"):
        # Se template_mode: clone bg + decorations, set self.frame
        # Altrimenti: comportamento attuale
```

Helper geometria:

```python
def _default_frame() -> _BBox:
    return _BBox(x=MARGIN, y=CONTENT_TOP,
                 w=SLIDE_W_IN - 2 * MARGIN,
                 h=FOOTER_Y - CONTENT_TOP - 0.35)
```

### 7.4 Modifiche renderer (~20 funzioni)

Pattern meccanico:

| Prima | Dopo (template mode) |
|-------|----------------------|
| `_set_bg(slide, t["bg_light"])` | Skip se `deck.template_mode` |
| `MARGIN` | `deck.frame.x` |
| `SLIDE_W_IN - 2 * MARGIN` | `deck.frame.w` |
| `FOOTER_Y - y - 0.35` | `deck.frame.y + deck.frame.h - y` |
| Decorazioni programmatiche (ovali cover) | Skip o ridotte se `template_mode` |

Renderer prioritari per refactor: `_r_cover`, `_r_title_bullets`, `_r_title_body`, `_r_kpi`, `_r_chart`, `_r_section`, `_r_closing`, `_dispatch` path diagrammi.

### 7.5 Modifiche `_build()` e `generate_slides()`

```python
def _build(self, spec: dict, *, template_pack=None):
    if template_pack:
        prs = Presentation()
        prs.slide_width = template_pack.slide_width
        prs.slide_height = template_pack.slide_height
    else:
        prs = Presentation()
        prs.slide_width = Inches(SLIDE_W_IN)
        prs.slide_height = Inches(SLIDE_H_IN)
    # ...
```

Entry point (**produzione-safe**):

```python
template_pack = None
ref_id = (spec.get("reference_file_id") or "").strip()
if ref_id:
    if not self.valves.template_mode_enabled:
        log.warning("[generate_slides] reference_file_id ignored (template_mode_enabled=false)")
    else:
        ref_bytes = await _load_reference_pptx(ref_id, __request__, __user__)
        if not ref_bytes:
            return self._error("Template file not found or not accessible.")
        try:
            template_pack = _parse_reference_pptx(ref_bytes)
        except Exception as exc:
            return self._error(f"Invalid template .pptx: {exc}")
        if spec.get("template_edits"):
            _apply_template_edits(template_pack, spec["template_edits"])
data, n = self._build(spec, template_pack=template_pack)
```

**Non usare** `_find_pptx_attachment(__messages__)` in `generate_slides`.

### 7.6 Valves (nuove)

| Valve | Default | Descrizione |
|-------|---------|-------------|
| `template_mode_enabled` | **`false`** | Abilita download reference + clone; se false, ignora `reference_file_id` |

| `inspect_slides_enabled` | **`true`** (opzionale v1) | Se `false`, nasconde/disabilita il secondo tool OWUI (solo `generate_slides`) |

Valve opzionale post-v1: `template_strict_mode` (fail se safe zone troppo piccola).

### 7.6.1 Wiring `deck.blank(layout)` (~21 call site)

Oggi ogni `_r_*` chiama `deck.blank()` **senza** argomento. In template mode ogni renderer deve passare il **ruolo** per `template_mapping`:

| Renderer | Chiamata |
|----------|----------|
| `_r_cover` | `deck.blank("cover")` |
| `_r_section` | `deck.blank("section")` |
| `_r_closing` | `deck.blank("closing")` |
| Altri layout contenuto | `deck.blank("content")` o default `"default"` |

**Fallback errore in `_build`:** se `_dispatch` fallisce e si usa `_r_title_body` di recupero, chiamare comunque `blank(...)` con ruolo coerente al layout richiesto, altrimenti slide di errore **senza** decorazioni template.

### 7.7 Download Files API (gap attuale)

Il repo implementa **upload** (`upload_file_handler` in `_save`). Il download reference **non esiste ancora**.

Task implementativi:

- Verificare endpoint OWUI v0.6 / v0.11 (es. `GET /api/v1/files/{id}/content`) con cookie/session da `__request__`
- Rifiutare file non `.pptx`/`.potx` o non presentation MIME
- **403/404** se file di altro utente (Neura multi-tenant)
- Timeout e limite dimensione file (es. log warning > 25 MB)

### 7.8 Decorazioni: slide vs master

| Sorgente | Use case | v1 |
|----------|----------|-----|
| Shape sulla **slide** reference | Template botanico, slide “vuota” con immagini | Clone da `_extract_decorations(slide)` |
| **Slide master** / layout master | Corporate `.potx`, logo nel master | P1: `_inspect_master_decorations` + `decorations_source` in inspect JSON |

Inspect (estensione schema):

```json
"decorations_source": "slide"
```

Valori possibili: `"slide"`, `"master"`, `"both"` (quando implementato).

### 7.9 Clone: z-order e media duplicati

- Inserire shape clonate **prima** del contenuto programmatico (decorazioni sotto testo/chart).
- Ogni slide output che riusa le stesse picture **duplica** le part in `ppt/media/` → deck lunghi più pesanti; accettabile v1, monitorare dimensione.
- Gestire relationship id (`rId`) senza collisioni al `save()`.
- Test: deck 15 slide + apertura in PowerPoint/LibreOffice senza “repair file”.

### 7.10 Layout immagine contenuto in Template Mode

Layout `_r_image`, `image_full_caption`, ecc. oggi usano **full-bleed** (`0, 0, SLIDE_W_IN, SLIDE_H_IN`).

**Policy v1:** con `template_mode=True`, immagini **contenuto** (URL/hint/base64 nel JSON) rispettano `deck.frame` — **no** full-bleed che copre decorazioni angolari.

### 7.11 Cover / section / closing con template chiaro

Renderer cover/section/closing usano `dark=True` e colori `on_dark` (testo bianco). Su template chiaro:

- Skip `_set_bg` e ovali programmatici ✅
- Forzare **`dark=False`** (o colori da `_derive_readable_colors`) su `_title`, `_content_head`, `_footer` in template mode

### 7.12 Chrome condiviso da refactorare

Oltre ai `_r_*`, aggiornare per `deck.frame` / template mode:

- `_eyebrow`, `_title`, `_footer` (oggi hardcoded `MARGIN`, `FOOTER_Y`, `SLIDE_W_IN`)
- `_Deck._content_head()`

Senza questo, anche **senza template** un refactor parziale può introdurre micro-shift (violazione NF5). Preferire **dual path** (§5.1) sui helper condivisi.

### 7.13 Timeout e feedback utente

Clone + deck lunghi (15+ slide) possono avvicinarsi al timeout HTTP del tool OWUI.

- Emettere `_emit(..., "Applying template…", done=False)` prima di parse/clone pesante.
- Docstring: consigliare 10–18 slide; warning in log se clone > N slide con stesse media duplicate.

### 7.14 Runtime Python

Verificare in container Neura/OWUI: `dataclass`, sintassi `X | None` (Python **3.10+** consigliato). Fase 0: registrare versione Python dell’immagine deploy.

### 7.15 Struttura file (target)

```
generate_slides.py
├── [esistente] Colour, palettes, layout aliases, shape helpers
├── [NUOVO]     Template engine (~400 righe)
├── [NUOVO]     inspect_slides (~150 righe)
├── [MOD]       _Deck, _build, generate_slides
├── [MOD]       Renderer geometry (~30 punti)
└── [MOD]       Docstring tool + version bump

docs/
└── template-mode-implementation-plan.md   ← questo documento

examples/
├── deck.json                              (invariato)
└── [NUOVO] template-deck-with-reference.json
└── [NUOVO] fixtures/                      (.pptx test, fase QA)
```

---

## 8. Cosa NON implementare (v1)

- [ ] Template da **PNG/JPG** (solo background ripetuto).
- [ ] Tool OWUI **separato** `analyze_slides` / `pptx_to_json`.
- [ ] **Cache sidecar** con PNG estratte (ridondante).
- [ ] **Variabile in-memory** tra inspect e generate.
- [ ] **Full template mode** (placeholder master nativi al posto dei renderer).
- [ ] Edit geometrici via AI (`resize`, `move`, `rotate` shape per id).
- [ ] Classificazione **layout canonico** in inspect come fatto vincolante.
- [ ] Supporto **`.thmx`** standalone.
- [ ] **Inject automatico** inventario all’upload (enhancement Neura/OWUI — fase 2 ops).
- [ ] Auto-detect `.pptx` in `generate_slides` via `__messages__` (**vietato** — rischio regression clienti).

---

## 9. Rischi e mitigazioni

| Rischio | Prob. | Impatto | Mitigazione |
|---------|-------|---------|-------------|
| **Attachment `.pptx` in chat attiva template per errore** | **Alta** | **Critico** | **No** auto-detect in `generate_slides`; solo `reference_file_id` esplicito |
| LLM salta `inspect_slides` | Media | Alto | Docstring; system prompt Neura; inject inventario (fase 2) |
| Nuovo tool `inspect_slides` confonde il modello | Media | Medio | Docstring “solo se utente chiede template”; valve off di default |
| `slide_layouts[6]` cambiato in path classico | Media | Alto | `[6]` **solo** se `template_mode=False` |
| Refactor geometria parziale | Media | Alto | NF5 hash baseline; refactor chrome + `_r_*`; `_default_frame()` calibrato |
| Download API / ACL non verificati | Media | Alto | Test su OWUI v0.11.3 NEURA; 403 cross-user |
| Reference fail → fallback classico | Media | Alto | **Fail closed** con `_error()` |
| `slide_layouts[6]` ≠ blank su template custom | Alta | Crash | `_find_blank_layout()` **solo** in template mode |
| Font corporate assente in container | Media | Medio | Fallback Calibri + log warning |
| Safe zone non ottimale vs occhio umano | Media | Basso | Margine 0.15"; documentare limite |
| Footer overlap decorazioni basso-dx | Media | Basso | Footer dentro safe zone |
| Cover con `dark=True` su template chiaro | Alta | Medio | `dark=False` + readable colors in template mode |
| Gruppi / SmartArt complessi | Media | Medio | Clone gruppo XML; `unsupported` in inspect |
| Decorazioni solo su master | Media | Alto | P1 master inspection; documentare limite v1 |
| File reference eliminato prima di generate | Bassa | Alto | Errore esplicito |
| Template 4:3 vs engine 16:9 | Media | Deformazione | Dimensioni da reference |
| z-order / rId / file corrupt | Media | Medio | Test deck lungo; NF6 |
| Regression modalità classica | Bassa | **Critico** | NF5 + F10 + valve default false |
| Tool non aggiornato in UI OWUI | Media | Alto | Checklist deploy (lezione v1.0.3) |
| Refactor geometria unificato senza NF5 | Media | **Critico** | Dual path §5.1; golden hash ogni PR |
| `blank()` senza ruolo layout / errore `_dispatch` | Media | Medio | §7.6.1; clone mancante su slide fallback |
| Timeout tool su clone deck lungo | Bassa | Medio | `_emit` progress; §7.13 |
| AI passa `reference_file_id` con valve off | Media | Medio | NF8 messaggio visibile; admin abilita valve |

---

## 10. Criteri di accettazione

### Funzionali

- [ ] **F1** — Slide template solo grafica → output con stesse decorazioni e posizioni.
- [ ] **F2** — `drop_ids: [12]` → shape id=12 assente nell’output.
- [ ] **F3** — Testo placeholder reference assente nell’output.
- [ ] **F4** — Contenuto generato solo in safe zone, non sopra decorazioni.
- [ ] **F5** — `kpi_row`, `chart`, `funnel` presenti e ridimensionati in safe zone.
- [ ] **F6** — Footer e numero pagina visibili.
- [ ] **F7** — Multi-slide reference + `template_mapping` rispettato.
- [ ] **F8** — Senza `reference_file_id` → output identico a v1.0.3.
- [ ] **F9** — `reference_file_id` presente ma download/parse fallisce → errore chiaro, **nessun** `.pptx` generato.
- [ ] **F10** — Utente allega `.pptx` in chat ma JSON **senza** `reference_file_id` → stesso output di v1.0.3 (F8/F10).

### Non funzionali

- [ ] **NF1** — Due inspect sullo stesso file → stessi id/bbox (deterministico).
- [ ] **NF2** — Nessuno stato condiviso tra due tool call.
- [ ] **NF3** — Log su errori parse, clone fallito, Files API miss.
- [ ] **NF4** — `examples/build.py` continua a funzionare senza modifiche.
- [ ] **NF5** — Stesso `examples/deck.json`, nessun reference, valve off → hash SHA256 (o byte-identico) uguale a baseline **v1.0.3**.
- [ ] **NF6** — Deck 15 slide in template mode → file apre in PowerPoint/LibreOffice senza repair.
- [ ] **NF7** — `reference_file_id` di altro utente → 403, nessun leak contenuto.
- [ ] **NF8** — `template_mode_enabled=false` + JSON con `reference_file_id` → ignore reference, output = classico (con log).
- [ ] **NF9** — Smoke test percorso **async** `generate_slides()` (valve off, stesso `deck.json`) oltre a `examples/build.py` / `_build` sync.
- [x] **NF10** — Docstring `inspect_slides`: invocare **solo** se l’utente chiede esplicitamente template da `.pptx`.

### Funzionali (template)

- [ ] **F11** — `template_mapping` con slide reference diverse → safe zone **per ruolo** (cover ≠ content se ref diverse).

---

## 11. Piano di implementazione a fasi

> Usare le checkbox `[ ]` per tracciare avanzamento.  
> **Stima complessiva:** 4–5 giorni dev + QA (include safety gate e NF5).

**Ordine consigliato:** Fase 0 → **0.5** → 1 → 2 → 3 → 4 → **6 (MVP 1 layout)** → 5 (resto renderer) → 7 → 8.

---

### Fase 0 — Preparazione e baseline

**Obiettivo:** ambiente test e baseline regression.

- [x] Creare directory `examples/fixtures/` per `.pptx` di test
- [x] Aggiungere fixture minima: template 1 slide (decorazioni angoli + area bianca)
- [x] Aggiungere fixture: template 2 slide (cover + content)
- [x] Aggiungere fixture: template con logo + textbox "Confidential"
- [x] Verificare `examples/build.py` su baseline v1.0.3 (regression reference)
- [x] Salvare hash SHA256 di `examples/demo_tech_deck.pptx` come **golden file** (NF5)
- [x] Documentare come eseguire test locali (`pip install python-pptx pillow`)
- [x] Registrare versione **Python** container OWUI Neura (≥3.10 per typing `|`) — dev locale: **3.11** (Docker `python:3.11-slim`); container Neura: _da registrare su deploy_
- [x] Fixture tipi: (a) picture angoli + bg bianco, (b) background fill slide + picture, (c) logo+Confidential textbox, (d) opzionale master-heavy (limite v1)

**Deliverable:** fixture pronte, baseline verde, golden hash registrato.

---

### Fase 0.5 — Safety gate (prima di esporre template in produzione)

**Obiettivo:** codice template presente ma **inattivo** per clienti esistenti.

- [x] Aggiungere valve `template_mode_enabled` (default **`false`**) in `Tools.Valves`
- [x] Condizione AND per template pipeline (sezione 6.7)
- [x] **Non** implementare `_find_pptx_attachment` in `generate_slides`
- [x] Se `reference_file_id` con valve off → log warning + path classico (NF8); messaggio visibile in risposta tool (§6.7)
- [x] Documentare strategia **dual path** geometria (§5.1) nel codice / commento blocco template
- [x] Verificare che nessuna modifica a `_Deck.blank()` alteri path classico (`slide_layouts[6]`)
- [x] CI/manuale: rigenerare `demo_tech_deck.pptx` → hash = golden (NF5)

**Deliverable:** merge sicuro su Neura con feature **spenta**; zero impatto clienti.

---

### Fase 1 — Template engine core (parse)

**Obiettivo:** leggere `.pptx` e produrre `_TemplatePack` in memoria.

- [x] Definire dataclass `_BBox`, `_Decoration`, `_SlideTemplate`, `_TemplatePack`
- [x] Implementare `_parse_reference_pptx(data: bytes) -> _TemplatePack`
- [x] Implementare `_extract_decorations(slide) -> list[_Decoration]`
- [x] Implementare `_is_text_shape(shape) -> bool`
- [x] Estrarre bbox shape in pollici (EMU → inches)
- [x] Estrarre testo verbatim per inspect (`text_verbatim`)
- [x] Estrarre colori theme XML (`dk1`, `lt1`, `accent1`, font major/minor)
- [x] Leggere `slide_width` / `slide_height` dal file (non forzare 16:9)
- [x] Implementare `_compute_safe_zone(slide, decorations) -> _BBox`
- [x] Algoritmo: max empty rect centrato, margine 0.15", riserva footer 0.35"
- [x] Gestire slide senza decorazioni (safe zone ≈ full slide minus margin)
- [x] Log diagnostico: shape count, safe zone, errori parse
- [x] Campo inspect `decorations_source: "slide"` (default v1)
- [x] (P1) Bozza `_inspect_master_decorations` per template corporate — opzionale post-MVP
- [x] Test unitario/manuale: parse fixture 1 slide, stampa inventario
- [x] Test `_is_text_shape` su gruppi e placeholder "Click to add title"

**Deliverable:** parse affidabile, no scrittura PPTX ancora.

---

### Fase 2 — `inspect_slides`

**Obiettivo:** esporre inventario JSON all’AI.

- [x] Implementare `_load_reference_pptx(file_id, request, user) -> bytes`
- [x] Integrare download Files API — verificare v0.6 / **v0.11.3 NEURA** (sezione 7.7) — _download HTTP + Storage in-process; verifica manuale Neura v0.11.3 da fare su deploy_
- [x] ACL: rifiutare file non appartenenti a `__user__`
- [x] Implementare `_find_pptx_attachment(__messages__) -> Optional[str]` **solo per inspect**
- [x] Implementare `async def inspect_slides(...)` su classe `Tools`
- [x] Serializzare `_TemplatePack` → JSON schema (sezione 6.1)
- [x] Campo `ok: false` + messaggio errore su parse/download fallito
- [x] Blocco `hints.likely_roles` (non vincolante, confidence low)
- [x] Aggiornare docstring `inspect_slides` con istruzioni per l’AI
- [x] Test E2E: inspect su fixture → JSON valido, id stabili tra run

**Deliverable:** tool inspect invocabile in OWUI, output JSON documentato.

---

### Fase 3 — Clone shape e background

**Obiettivo:** copiare decorazioni da reference a slide nuova.

- [x] Implementare `_find_blank_layout(prs)` (nome "Blank" / "Vuota" / fallback ultimo)
- [x] Implementare `_clone_background(target_slide, source_slide)`
- [x] Implementare `_clone_decorations(target_slide, decorations: list[_Decoration])`
- [x] Copia relationship media parts (`ppt/media/imageN.png`) nel package dest
- [x] Deep-copy elementi `spTree` (picture, shape, group)
- [x] Preservare posizione/dimensione (bbox) identica
- [x] Implementare `_apply_template_edits(pack, edits)` (filtra per drop_ids/keep_ids/defaults)
- [x] Ordine inserimento: decorazioni **sotto** contenuto (z-order)
- [x] Test visivo: 1 slide output = stessa grafica della fixture (senza contenuto)
- [x] Test: `keep_ids` vuoto per errore → warning in log / inspect hint

**Deliverable:** clone visivo corretto su singola slide.

---

### Fase 4 — Integrazione `_Deck` e geometria

**Obiettivo:** template mode attiva `deck.frame` e disabilita override sfondo.

- [ ] Aggiungere `template_pack` e `template_mode` a `_Deck.__init__`
- [ ] Refactor `blank(layout="default")` → chiama clone se template_mode; legacy `blank()` = stesso comportamento di oggi se `layout` default e non template
- [ ] Implementare `_pick_template_slide(pack, layout, mapping)` + validazione indice (§6.10)
- [ ] Aggiornare **tutti** i `_r_*` (~21) per passare ruolo a `blank()` (§7.6.1)
- [ ] `_build` loop: fallback `_r_title_body` su exception usa `blank` con ruolo layout corretto
- [ ] Helper `_content_bounds(deck, top_y)` dual path (§5.1)
- [ ] Implementare `_default_frame() -> _BBox`
- [ ] Property `deck.frame` usata dai renderer
- [ ] Skip `_set_bg()` quando `deck.template_mode`
- [ ] Implementare `_derive_readable_colors(bg_hex)` e merge theme (sezione 6.8)
- [ ] `_find_blank_layout(prs)` **solo** se `template_mode=True`; altrimenti `slide_layouts[6]`
- [ ] Test: `title_bullets` con template → testo dentro safe zone
- [ ] Dopo ogni modifica: NF5 hash senza reference

**Deliverable:** una slide contenuto renderizzata correttamente in template mode.

---

### Fase 5 — Refactor renderer (Opzione B)

**Obiettivo:** layout ricchi ridimensionati in safe zone.

**Priorità alta (5 renderer + chrome):**

- [ ] `_r_title_bullets` → usa `deck.frame`
- [ ] `_r_title_body` → usa `deck.frame`
- [ ] `_r_cover` → skip ovali programmatici se template; contenuto in frame; **`dark=False`** + readable colors
- [ ] `_r_kpi` → ridimensiona card in `deck.frame.w`
- [ ] `_r_chart` → chart width/height proporzionali a frame
- [ ] `_Deck._content_head()` / `_eyebrow` / `_title` / `_footer` → coordinate relative a `deck.frame`
- [ ] Footer: fallback dentro safe zone se overlap decorazione basso

**Priorità media (resto layout):**

- [ ] `_r_section` / `_r_closing` → stessa policy `dark=False` in template mode
- [ ] `_r_comparison`
- [ ] `_r_timeline` / `_r_process`
- [ ] `_r_quote`
- [ ] `_r_alert`
- [ ] `_r_table`
- [ ] `_r_icon_list` / `_r_icon_grid`
- [ ] `_r_image` → **no full-bleed** in template mode; rispetta `deck.frame` (sezione 7.10)
- [ ] `_r_funnel` / `_r_diagram` / diagrammi

**Deliverable:** deck multi-slide con mix layout in template mode.

---

### Fase 6 — Integrazione `generate_slides` end-to-end

**Obiettivo:** flusso 3 passi completo in produzione.

**MVP (può precedere completamento Fase 5):** valve ON + reference + solo `title_bullets` / `title_body` in template mode.

- [ ] Estendere `_build(spec, template_pack=None)`
- [ ] Parse **solo** `reference_file_id` da spec JSON (entry point sezione 7.5)
- [ ] **Vietato:** fallback attachment in `generate_slides`
- [ ] Parse `template_mapping` e `template_edits` da spec
- [ ] Pipeline: load → parse → edits → build → save; fail closed su errori reference (F9)
- [ ] `_emit` stati: “Loading template…”, “Applying template…” durante download/clone
- [ ] Aggiornare docstring `generate_slides` (campi nuovi + workflow inspect + valve)
- [ ] Test NF9: smoke `generate_slides` async (mock/minimo OWUI) con valve off
- [ ] Creare `examples/template-deck-with-reference.json` di esempio
- [ ] Test E2E manuale OWUI v0.11.3: inspect → generate con drop_ids
- [ ] Test E2E: generate senza reference → NF5
- [ ] Test F10: attachment in chat senza `reference_file_id` → NF5
- [ ] Test NF6: deck 15 slide template
- [ ] Checklist deploy: aggiornare tool in UI OWUI (lezione release v1.0.3)

**Deliverable:** feature invocabile su Neura pilot con `template_mode_enabled=true`.

---

### Fase 7 — Documentazione e release

**Obiettivo:** docs utente/admin e bump versione.

- [ ] Aggiornare `README.md` — sezione Template Mode + workflow 3 passi
- [ ] Documentare `palette{}`, `heading_font`, `body_font` (quick win anche senza template)
- [ ] Aggiungere tabella campi JSON nuovi (`reference_file_id`, `template_edits`, …)
- [ ] Bump version frontmatter `1.0.3` → `1.1.0` (minor feature)
- [ ] Changelog: **“Nessun cambiamento comportamento se `template_mode_enabled=false` (default) e senza `reference_file_id`”**
- [ ] Nota admin Neura: abilitare valve solo per pilot; aggiornare tool in UI
- [ ] System prompt / istruzioni modello per workflow inspect → generate (opzionale Neura)
- [ ] Commento su GitHub Issue #2 con link doc e istruzioni
- [ ] Screenshot before/after (template vs output) in `assets/` (opzionale)

**Deliverable:** release documentata.

---

### Fase 8 — Hardening e enhancement (post-v1)

**Obiettivo:** robustezza produzione; non bloccante per MVP.

- [ ] Gestione esplicita SmartArt / OLE → `unsupported` in inspect
- [ ] Gestione shape group annidate (clone ricorsivo)
- [ ] Valve `template_strict_mode` (fail se safe zone troppo piccola)
- [ ] Master slide decorations (completamento P1)
- [ ] Inject automatico inventario all’upload `.pptx` (integrazione Neura/OWUI)
- [ ] Preview thumbnail URL in output inspect (senza base64)
- [ ] Supporto PNG/JPG come template (background mode — fase separata)
- [ ] Metriche/log: tempo parse, tempo clone, slide count

**Deliverable:** production-hardening iterativo.

---

## 12. Appendice: esempio end-to-end

### Input utente

```
Allegato: botanical-template.pptx (1 slide, solo grafica angoli)
Prompt: "Generami una presentazione sull'AI in medicina usando questo template.
         Togli il logo in alto a destra e la scritta Confidential."
```

**Prerequisito deploy (Passo 3):** valve admin `template_mode_enabled=true` sul workspace pilot Neura. Senza valve ON, `reference_file_id` viene ignorato (NF8).

### Passo 1 — `inspect_slides("abc123")`

```json
{
  "ok": true,
  "file_id": "abc123",
  "slide_count": 1,
  "slides": [{
    "index": 0,
    "shapes": [
      { "id": 4, "kind": "picture", "bbox": { "x": 0, "y": 0, "w": 4.2, "h": 3.1 } },
      { "id": 5, "kind": "picture", "bbox": { "x": 9, "y": 5, "w": 4, "h": 2.5 } },
      { "id": 12, "kind": "picture", "bbox": { "x": 11.2, "y": 0.2, "w": 1.4, "h": 0.8 } },
      { "id": 18, "kind": "textbox", "text": "Confidential" }
    ],
    "safe_zone": { "x": 1.35, "y": 1.1, "w": 10.6, "h": 5.2, "method": "max_empty_rect" }
  }]
}
```

### Passo 2 — AI → `generate_slides`

```json
{
  "title": "AI in Medicina",
  "reference_file_id": "abc123",
  "template_mapping": { "default": 0 },
  "template_edits": {
    "defaults": { "drop_text": true },
    "slides": [{ "index": 0, "drop_ids": [12, 18] }]
  },
  "slides": [
    {
      "layout": "cover",
      "title": "AI in Medicina",
      "subtitle": "Opportunità e sfide cliniche",
      "author": "NEURA"
    },
    {
      "layout": "kpi_row",
      "title": "Impatto numerico",
      "stats": [
        { "value": "30%", "label": "Riduzione tempi diagnosi", "change": "studi 2025" },
        { "value": "24/7", "label": "Supporto decisionale", "change": "always-on" }
      ]
    },
    {
      "layout": "closing",
      "title": "Grazie",
      "takeaways": ["L'AI augmenta il clinico", "Serve governance dei dati"]
    }
  ]
}
```

### Passo 3 — Risultato

- Shape **4, 5** (foglie) clonate su **ogni** slide output.
- Shape **12** (logo) e **18** (Confidential) **non** clonate.
- Titoli, KPI, testi in **safe zone** (1.35, 1.1, 10.6×5.2).
- Footer e numeri pagina presenti.
- File salvato via Files API; link in chat.

---

## 13. Checklist pre-merge (produzione Neura)

Prima di abilitare Template Mode per clienti (oltre al merge codice con valve **off**):

- [ ] `template_mode_enabled=false` su **tutti** gli ambienti clienti fino a decisione esplicita
- [ ] NF5 verde: `examples/build.py` → hash = golden v1.0.3
- [ ] F10 verde: simulazione attachment `.pptx` + JSON senza reference
- [ ] Nessuna modifica comportamentale a `_save`, `_emit_link`, Files API upload
- [ ] `_prefetch_images` / immagini contenuto invariati per path classico
- [ ] Test download reference + ACL su istanza NEURA (v0.11.3)
- [ ] Tool `generate_slides.py` **ricaricato** in Open WebUI admin (post-deploy)
- [ ] Pilot interno: valve ON su workspace test → flusso inspect + generate + drop_ids
- [ ] Rollback plan: disabilitare valve (instant) o pin versione tool v1.0.3
- [ ] NF9 eseguito (non solo `build.py` sync)
- [ ] Strategia dual path geometria rispettata; NF5 verde post-merge

---

## Riferimenti

| Risorsa | Path / URL |
|---------|------------|
| Codice attuale | `generate_slides.py` |
| Esempio JSON contenuto | `examples/deck.json` |
| Test locale render | `examples/build.py` |
| Issue GitHub | https://github.com/ianustec/openwebui-generate-slides/issues/2 |
| Release Files API fix | v1.0.3 |
| Chrome hardcoded (refactor) | `_eyebrow`, `_title`, `_footer` ~L703–737 |

---

*Documento di implementazione Template Mode — IANUSTEC / NEURA. Revisione sicurezza produzione: 2026-09-22 (audit integrato: dual path, blank(layout), NF9/F11, edge case).*
