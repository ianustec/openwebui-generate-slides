# Presentation Edit — Piano di implementazione

> **Relazione:** complementare a [template-mode-implementation-plan.md](./template-mode-implementation-plan.md) — **non** fa parte del flusso inspect → generate (§4 di quel documento).  
> **Versione base:** `generate_slides.py` v1.1.x (post template mode + reuse + materialize theme + layout bg flatten)  
> **Stato:** **non implementato** — piano di lavoro  
> **Ultimo aggiornamento:** 2026-09-30 (decisioni gap §5.1)  
> **Contesto deploy:** Neura — tool **on-demand** invocato solo quando l’utente chiede correzioni visive/testuali su un deck **già generato**.

---

## Indice

1. [Executive summary](#1-executive-summary)
2. [Contesto e requisiti](#2-contesto-e-requisiti)
3. [Stato attuale vs obiettivo](#3-stato-attuale-vs-obiettivo)
4. [Architettura: flusso edit (standalone)](#4-architettura-flusso-edit-standalone)
5. [Principi di design](#5-principi-di-design) (incl. [§5.1 Decisioni gap](#51-decisioni-gap-e-trade-off))
6. [Contratti dati](#6-contratti-dati)
7. [Componenti software](#7-componenti-software)
8. [Cosa NON implementare (v1 edit)](#8-cosa-non-implementare-v1-edit)
9. [Rischi e mitigazioni](#9-rischi-e-mitigazioni)
10. [Criteri di accettazione](#10-criteri-di-accettazione)
11. [Piano di implementazione a fasi](#11-piano-di-implementazione-a-fasi)
12. [Appendice: esempio end-to-end](#12-appendice-esempio-end-to-end)
13. [Checklist pre-merge (produzione Neura)](#13-checklist-pre-merge-produzione-neura)

---

## 1. Executive summary

### Obiettivo

Dopo che `generate_slides` ha prodotto un `.pptx`, l’utente lo apre, nota problemi di **presentabilità** (titolo troppo grande, testo che esce dal box, titolo e descrizione sovrapposti) e chiede in chat correzioni mirate. Neura analizza la richiesta e invoca un tool dedicato che **modifica in place** il file già salvato, senza rigenerare il deck da template.

### Problema che risolve

In modalità **reuse**, il testo viene sostituito con `_replace_txBody_text`, che **conserva** `a:rPr` (incluso `sz` fisso). Testo più lungo del placeholder template → overflow, a capo brutti, overlap con altri shape. Non è un bug del clone: è assenza di **post-processing layout testo** dopo la generazione.

### Soluzione adottata

| Passo | Attore | Output |
|-------|--------|--------|
| **A** | `generate_slides` (già esistente) | `.pptx` + `file_id` / link download |
| **B** | Utente | Feedback qualitativo (“slide 1 titolo troppo grande”, …) |
| **C** | Neura | Interpretazione → JSON operazioni strutturate |
| **D** | **`edit_presentation`** (nuovo) | Download deck generato → applica `operations[]` → `_save` → nuovo link |

**Non** è un quarto passo del template mode: nessun `reference_file_id`, nessun clone, nessun overlay safe zone.

### Insight centrale

> **`presentation_file_id`** = Files API id del **output** di generate (o di un edit precedente), non del template di inspect.
>
> Neura può chiamare **`inspect_slides(presentation_file_id)`** in parallelo al bisogno (shape id, bbox, testo attuale), ma **edit** non dipende obbligatoriamente da inspect se id e operazioni sono già noti dalla conversazione.

### Regola produzione

| Regola | Motivo |
|--------|--------|
| Edit **solo** se valve `presentation_edit_enabled` (default **`false`**) | Opt-in pilot |
| Input obbligatorio: `presentation_file_id` (UUID Files API) — **non** accettare solo URL cache (§5.1 #1) | Stesso pattern ACL; edit richiede id restituito da upload Files API |
| **Mai** NLP / parsing linguaggio naturale **dentro** il tool Python | Neura traduce utente → JSON; edit resta deterministico |
| **Mai** riaprire template o rifare `_build` / reuse | Evita regressioni tema/sfondo/clone |
| Fallimento download / file non `.pptx` → `_error()`, nessun file parziale | Fail closed |
| Path `generate` / template mode **invariato** se edit disabilitato o non chiamato | Regression zero |

---

## 2. Contesto e requisiti

### Origine

- Feedback produzione: deck reuse fedeli al template ma **non sempre presentabili** senza micro-aggiustamenti font/testo.
- Workflow desiderato: generate → review umana → richiesta modifica → edit → download aggiornato.

### Requisiti funzionali

| # | Requisito | Decisione |
|---|-----------|-----------|
| E1 | Input | `presentation_file_id` del deck da modificare |
| E2 | Trigger | **Solo** su richiesta esplicita utente post-generate; non automatico |
| E3 | Interpretazione intent | **Neura** (chat); tool esegue solo JSON |
| E4 | Indice slide | JSON **`slide` 0-based**; Neura mappa “slide 1” utente → `0` |
| E5 | Target shape | `shape_id` come in inspect (`shapes[].id` sul file **generato**) |
| E6 | Operazioni v1 | `set_text`, `shrink_font` / `fit_text` (euristica bbox + valve `min_pt` / `step_pt` — §5.1 #4) |
| E7 | Operazioni v2 | **`enable_autofit`** (priorità), `split_text`, `replace_text_and_fit`, `resize_shape` (opzionale) |
| E8 | Formattazione | `set_text` riusa `_set_shape_text_preserve_font` **senza modificarne il comportamento** (§5.1 #9) |
| E9 | Output | Nuovo upload Files API via `_save`; catena edit richiede **nuovo UUID** (§5.1 #1) |
| E10 | Template mode | Edit funziona su **qualsiasi** `.pptx` generato (reuse, mapping classico, deck NF5) |
| E11 | Tabelle | **v1:** solo shape con text frame — **no** table id (§5.1 #6); v2: `set_table_cell` o reuse `_fill_table_preserve_format` |

### Requisiti non funzionali

- Stesso file `generate_slides.py` (classe `Tools`), nuovo metodo tool OWUI.
- Riutilizzo `_load_reference_pptx` (o alias `_load_presentation_bytes`) + `_save` + ACL utente.
- Log: `presentation_file_id`, numero operazioni, slide/shape toccati, errori per-op (fail parziale vs fail totale — v1: **fail totale** su prima op invalida).
- Test offline: script `examples/` con bytes locali + mock senza OWUI.

---

## 3. Stato attuale vs obiettivo

### Comportamento attuale

```
generate_slides → _build → reuse: _replace_txBody_text (sz invariato)
                → prs.save → _save → file_id / cache URL
```

Nessun tool per modificare un `.pptx` esistente. L’unica alternative oggi è **rigenerare** con JSON diverso (costoso, rischio drift grafico).

### Comportamento target

```
Utente → Neura → edit_presentation(JSON)
  → GET presentation_file_id
  → Presentation(BytesIO(data))
  → for op in operations: _apply_edit_op(prs, op)
  → prs.save → _save → link + file_id
```

---

## 4. Architettura: flusso edit (standalone)

### Diagramma sequenza

```
Utente (dopo aver scaricato/aperto il deck):
        "Nella slide 1 il titolo è troppo grande;
         nella slide 4 titolo e descrizione vanno separati"

┌─────────────────────────────────────────────────────────────────┐
│  (opzionale) inspect_slides(presentation_file_id)               │
│  → inventario slide/shape/bbox/testo sul FILE GENERATO          │
│  → Neura usa id per costruire operations                        │
└────────────────────────────┬────────────────────────────────────┘
                             │
                             ▼
┌─────────────────────────────────────────────────────────────────┐
│  Neura — analisi intent                                         │
│  → NON chiama generate_slides                                   │
│  → produce JSON edit: presentation_file_id + operations[]       │
└────────────────────────────┬────────────────────────────────────┘
                             │
                             ▼
┌─────────────────────────────────────────────────────────────────┐
│  edit_presentation(content_json)                                │
│  → download presentation_file_id (ACL utente)                   │
│  → applica operazioni deterministiche su OOXML                  │
│  → salva .pptx aggiornato via Files API                         │
└─────────────────────────────────────────────────────────────────┘
```

### Separazione responsabilità

| Attore | Responsabilità | Non fa |
|--------|----------------|--------|
| **Utente** | Giudizio visivo, richiesta in linguaggio naturale | Passare shape id (salvo power user) |
| **Neura** | Tradurre richiesta → `operations[]`; slide 1-based → 0-based; inspect output se serve | Modificare OOXML |
| **`inspect_slides`** (opz.) | Verità file **generato** | Edit / generate |
| **`generate_slides`** | Creazione deck | Fix post-hoc titolo/font |
| **`edit_presentation`** | Download, mutate, save | Clone template, NLP, rigenerare contenuto |

### Relazione con Template Mode (§4 altro doc)

| Domanda | Risposta |
|---------|----------|
| Edit entra nel diagramma inspect → generate? | **No** |
| Serve `reference_file_id`? | **No** |
| Serve `template_mode_enabled`? | **No** (edit ha valve propria) |
| Stesso `file_id` del template? | **No** — solo output generate (o catena edit) |

---

## 5. Principi di design

1. **Deck generato = source of truth** per edit; template reference irrilevante.
2. **AI = intent; tool = esecuzione** — stesso pattern di `reuse.text` + shape id.
3. **Micro-edit only** — font, testo, split tra placeholder esistenti; niente redesign slide.
4. **Riuso helper esistenti** — `_shape_by_id`, `_set_shape_text_preserve_font`, `_replace_txBody_text`, parse `_parse_reference_pptx` solo se serve metriche bbox (opz.).
5. **Regression zero** — codice generate/inspect/template clone non cambia comportamento se edit non invocato.
6. **Opt-in valve** — `presentation_edit_enabled` default false.
7. **Fail closed** — file illeggibile, id shape mancante, slide out of range → errore chiaro, nessun save silenzioso incompleto (v1).
8. **Idempotenza parziale** — rieseguire `shrink_font` su testo già ridotto può fermarsi a `min_pt` (documentato).
9. **No attachment auto-detect in edit** — solo `presentation_file_id` esplicito (coerenza NF8 template).
10. **Documentazione Neura separata** — `doc/neura-presentation-edit-hints.md` (o sezione dedicata), **non** fusione nel flusso template §4.
11. **Isolamento codice reuse** — nessuna modifica a `_build`, clone, `_replace_txBody_text` per edit; helper font solo additivi (§5.1 #9).

### 5.1 Decisioni gap e trade-off

Decisioni approvate (2026-09-30) per bug/gap noti in review del piano. Implementazione e hint Neura devono allinearsi a questa tabella.

| # | Gap | Scelta | Implicazione |
|---|-----|--------|--------------|
| **1** | `file_id` assente dopo generate (cache fallback) | **A** | `edit_presentation` accetta **solo** UUID Files API. Se generate restituisce solo `/cache/files/...` senza id: hints dicono di non chiamare edit finché Files API non restituisce UUID (fix deploy / rigenerare in prod). PE7 valido **solo** con Files API. |
| **2** | Ordine `set_text` vs `shrink_font` | **A** (+ **B** v2) | **Hints:** sullo stesso shape → prima `set_text`, poi `shrink_font`. **v2:** op `replace_text_and_fit` (testo + shrink atomico) per meno errori modello. |
| **3** | Shape id dentro gruppi | **A** | Risoluzione shape in edit **sempre** con `_shape_by_id_recursive`. |
| **4** | Euristica shrink ≠ PowerPoint | **A + C** MVP; **B** v2 | MVP: shrink euristico + doc “verificare in PowerPoint”; valve **`presentation_edit_min_font_pt`** / `step_pt` conservativi. **v2 priorità:** `enable_autofit` (`normAutofit`). |
| **5** | Split titolo / descrizione | **B** (+ **C** opz.) | **v2 obbligatorio in piano:** `split_text`. **`resize_shape`** opzionale v2 per overlap verticale. v1: due `set_text` su **due id** da inspect se esistono. |
| **6** | Tabelle | **A** MVP | v1 **non** target table shape. v2: **6B** `set_table_cell` o **6C** refill parziale via `_fill_table_preserve_format`. |
| **7** | Confusione template id vs deck output | **A** (+ **7B** se prod) | Hints: `presentation_file_id` = ultimo output generate/edit; **mai** `reference_file_id`. **7B (opz.):** rifiuto esplicito se id = template allegato in chat — solo se errori in produzione. |
| **8** | Batch `operations[]` con un errore | **A** MVP; **B** v2 | v1: **fail totale**, nessun save, messaggio con indice op. v2 (power user): save parziale + report ops ok/fail. |
| **9** | Regression reuse/generate | **A** (+ **9B**) | **Vietato** cambiare semantica helper reuse. Nuova logica font in helper **solo edit** (es. `_shrink_txBody_rPr`) se necessario. |
| **10** | Inspect sul file generato | **A** (+ **B** fallback) | Default: stesso `inspect_slides` + `_compact_inspect_payload` (id, text, bbox). Se in test mancano campi per edit → **10B** parametro/detail inspect dedicato (solo post-MVP). |

---

## 6. Contratti dati

### 6.1 Payload root `edit_presentation`

Parametro tool: **`content`** — stringa JSON unica (stesso stile `generate_slides`).

```json
{
  "presentation_file_id": "8f3c2a1b-....",
  "title": "Deck Rita (edited)",
  "operations": []
}
```

| Campo | Tipo | Obbligatorio | Note |
|-------|------|--------------|------|
| `presentation_file_id` | string (UUID) | sì | Id Files API da generate/edit (**non** URL cache — §5.1 #1) |
| `title` | string | no | Nome file in `_save` |
| `operations` | array | sì | Almeno un’operazione in v1 |

### 6.2 Operazione comune

Ogni elemento di `operations`:

| Campo | Tipo | Note |
|-------|------|------|
| `op` | string | Nome operazione (vedi §6.3) |
| `slide` | int | Indice slide **0-based** |

Altri campi dipendono da `op`.

### 6.3 Operazioni (v1 / v2)

#### `set_text` (v1)

```json
{
  "op": "set_text",
  "slide": 0,
  "shape_id": 783,
  "text": "Titolo più corto\nSeconda riga"
}
```

- `\n` → paragrafi (come reuse).
- Shape deve avere **text frame** (v1: **no** table — §5.1 #6).

**Ordine consigliato (Neura hints §5.1 #2A):** sullo stesso `shape_id`, eseguire **`set_text` prima di `shrink_font`**.

#### `shrink_font` (v1)

```json
{
  "op": "shrink_font",
  "slide": 0,
  "shape_id": 783,
  "min_pt": 18,
  "step_pt": 2,
  "max_iterations": 20
}
```

- Legge `a:rPr/@sz` (centesimi di punto); riduce finché euristica “testo entra in bbox” o `min_pt`.
- Euristica v1: rapporto lunghezza testo vs area shape (EMU) + numero righe `\n`; **non** layout engine PowerPoint (§5.1 #4A).
- Rispettare `min_pt` da op o valve `presentation_edit_min_font_pt` (§5.1 #4C).

#### `fit_text` (v1 alias)

Stesso comportamento di `shrink_font` con parametri default da valve (`min_pt`, `step_pt`).

#### `replace_text_and_fit` (v2 — §5.1 #2B)

```json
{
  "op": "replace_text_and_fit",
  "slide": 0,
  "shape_id": 783,
  "text": "Titolo più corto",
  "min_pt": 18,
  "step_pt": 2
}
```

Sostituisce testo e applica shrink **in un solo passaggio** (ordine interno fisso). Preferire questa op quando il modello deve evitare array `set_text` + `shrink_font` sullo stesso id.

#### `enable_autofit` (v2 — priorità §5.1 #4B)

Imposta `a:bodyPr` → `normAutofit` / rimuove `noAutofit` dove presente; PowerPoint adatta all’apertura. Preferire quando l’utente chiede “adatta al riquadro” / “come in PowerPoint”.

#### `split_text` (v2 — §5.1 #5B)

```json
{
  "op": "split_text",
  "slide": 3,
  "from_shape_id": 10,
  "to_shape_id": 11,
  "mode": "first_paragraph"
}
```

- Sposta porzione testo da shape A a B; entrambi devono esistere sul file generato.

#### `resize_shape` (v2 opzionale — §5.1 #5C)

Micro-aggiustamento `a:xfrm` (altezza/larghezza textbox) entro limiti conservativi documentati. Utile per overlap verticale titolo/body; **non** sostituisce `split_text` se il testo è nello stesso shape.

### 6.4 Convenzione slide “utente vs JSON”

| Utente dice | JSON `slide` |
|-------------|----------------|
| slide 1 | `0` |
| slide 4 | `3` |

Documentare in model hints Neura.

### 6.5 Risposta tool

Allineata a generate: messaggio markdown + link download + `file_id` se Files API restituisce id (come `_save`).

---

## 7. Componenti software

Posizione: [`generate_slides.py`](../generate_slides.py), classe `Tools`.

### 7.1 Nuovo tool

| Simbolo | Tipo | Ruolo |
|---------|------|--------|
| `edit_presentation` | async method | Entry OWUI; parse JSON, valve, download, apply, save |
| `_parse_edit_spec(raw) -> dict` | function | Validazione schema minimo |
| `_apply_edit_operations(prs, operations) -> None` | function | Loop operazioni |
| `_apply_edit_op(prs, op) -> None` | function | Dispatch per `op` |
| `_shrink_shape_font(shape, ...) -> bool` | function | Mutazione XML `a:rPr/@sz` |
| `_text_fits_shape_heuristic(shape) -> bool` | function | Euristica fit v1 |
| `_shape_bbox_emu(shape) -> tuple` | function | Da inspect logic / EMU esistente |

### 7.2 Riutilizzo

| Esistente | Uso in edit |
|-----------|-------------|
| `_load_reference_pptx` | Download bytes (rinominare mentalmente “presentation load”) |
| `_save` | Upload output |
| `_shape_by_id_recursive` | Risoluzione shape (**obbligatorio** in edit — §5.1 #3A) |
| `_set_shape_text_preserve_font` | `set_text` (chiamata read-only; **non** modificare per edit — §5.1 #9A) |
| `_shrink_txBody_rPr` (nuovo, opz.) | Logica font solo edit se serve isolamento (§5.1 #9B) |
| `_is_files_api_id` | Validazione id |
| `_wrong_reference_id_message` | Pattern messaggio errore (adapt per presentation_file_id) |

### 7.3 Valve

```python
presentation_edit_enabled: bool = Field(default=False, description="...")
presentation_edit_min_font_pt: float = Field(default=14.0, ...)  # conservativo §5.1 #4C
presentation_edit_shrink_step_pt: float = Field(default=2.0, ...)
```

### 7.4 Struttura file (target)

```
generate_slides.py
├── [esistente] Template engine, generate, inspect
├── [NUOVO]     Presentation edit engine (~200–350 righe)
└── [MOD]       Valves, version bump, description header

doc/
├── template-mode-implementation-plan.md   (invariato §4)
├── presentation-edit-implementation-plan.md   ← questo documento
└── neura-presentation-edit-hints.md       (nuovo, Fase 5)

examples/
├── edit_presentation_smoke.py             (nuovo)
└── run_production_checks.py               (+ voce edit smoke)
```

---

## 8. Cosa NON implementare (v1 edit)

- [ ] Parsing linguaggio naturale dentro `edit_presentation`.
- [ ] Rigenerazione deck / chiamata interna a `_build` o reuse.
- [ ] Modifica template `reference_file_id` o clone decorazioni.
- [ ] Spostamento libero shape (`move`, `rotate`, z-order) via AI.
- [ ] Edit immagini (sostituzione blip, crop).
- [ ] Edit master/layout/theme del package.
- [ ] Batch “fix automatico tutte le slide” senza lista operazioni esplicita (v3 opzionale).
- [ ] Tool OWUI separato file — restare single-file repo.
- [ ] Integrazione obbligatoria in workflow template §4.
- [ ] `set_text` / shrink su **table** shape (v1 — §5.1 #6A).
- [ ] Modificare `_replace_txBody_text` o clone/reuse per “far comodo” a edit (§5.1 #9A).

---

## 9. Rischi e mitigazioni

| Rischio | Prob. | Impatto | Mitigazione |
|---------|-------|---------|-------------|
| Euristica fit ≠ PowerPoint | Alta | Medio | §5.1 #4A doc; v2 priorità `enable_autofit` (#4B) |
| Neura passa shape id sbagliato | Media | Medio | `inspect_slides` su file **generato** (#10A); `_shape_by_id_recursive` (#3A) |
| Confusione template file_id vs output file_id | Media | Alto | Hints #7A; opz. rifiuto id template #7B |
| Solo URL cache, no UUID | Media | Alto | §5.1 #1A: edit non invocabile; messaggio chiaro in hints |
| Edit su file altrui | Bassa | Alto | Stessa ACL di `_load_reference_pptx` |
| Ordine op sbagliato (shrink prima di set_text) | Media | Medio | Hints #2A; v2 `replace_text_and_fit` #2B |
| Font size sotto leggibilità | Media | Medio | Valve `min_pt` / `step_pt` #4C |
| Batch ops: una op invalida | Bassa | Medio | v1 fail totale #8A; v2 parziale #8B |
| Rigenerare invece di edit (Neura) | Media | Medio | System prompt: fix puntuali → edit |
| Regression generate / reuse | Bassa | Critico | §5.1 #9A: codice additivo; NF5 invariato |
| Inspect compatto insufficiente | Bassa | Medio | Verifica #10A; fallback param inspect #10B |

---

## 10. Criteri di accettazione

### Funzionali

- [ ] **PE1** — Dato un pptx reuse con titolo lungo, `shrink_font` riduce `sz` e il testo non supera grossolanamente il bbox (test fixture).
- [ ] **PE2** — `set_text` su shape id valido cambia solo testo, preserva colore/grassetto/`spc` del primo run template.
- [ ] **PE3** — `presentation_file_id` invalido → errore, nessun file.
- [ ] **PE4** — Valve off → tool rifiuta con messaggio admin.
- [ ] **PE5** — Slide index out of range → errore chiaro.
- [ ] **PE6** — Shape id assente → errore chiaro (op index citata).
- [ ] **PE7** — Catena: generate (Files API **con UUID**) → edit → secondo edit sul **nuovo** `file_id` funziona (§5.1 #1A; non richiesto con solo URL cache).
- [ ] **PE8** — `presentation_file_id` non-UUID o assente → errore esplicito, nessun download (§5.1 #1A).

### Non funzionali

- [ ] **PEN1** — `generate_slides` / template mode tests esistenti verdi senza chiamare edit.
- [ ] **PEN2** — Script smoke edit in Docker con python-pptx.
- [ ] **PEN3** — Log operazioni applicate (slide, shape_id, op).
- [ ] **PEN4** — Doc Neura edit separata da template hints.

---

## 11. Piano di implementazione a fasi

### Fase 0 — Documentazione e contratto (pre-code)

**Obiettivo:** allineare Neura e repo prima del codice.

- [ ] Approvare questo documento in review interna.
- [ ] Creare `doc/neura-presentation-edit-hints.md` (workflow utente → Neura → edit; **no** inserimento in template §4; includere §5.1 #1A, #2A, #7A).
- [ ] Aggiungere cross-link da README (voce “Post-generate edit”) — opzionale breve.
- [ ] Definire nome tool OWUI definitivo: `edit_presentation` (alias doc `edit_slides` deprecato).

**Deliverable:** hint copiabili in system prompt Neura.

---

### Fase 1 — MVP engine + tool shell

**Obiettivo:** download, `set_text`, `shrink_font`, save.

- [ ] Valve `presentation_edit_enabled` (default false).
- [ ] `_parse_edit_spec` + validazione UUID `presentation_file_id` (`_is_files_api_id`; rifiuto se manca — §5.1 #1A).
- [ ] `async def edit_presentation(self, content: str, ...)` — specchio pattern `generate_slides` (emitter, user, request).
- [ ] Riutilizzo download: `_load_reference_pptx(presentation_file_id, ...)`.
- [ ] `_apply_edit_operations` + dispatch `set_text` | `shrink_font`; fail totale prima op invalida (§5.1 #8A).
- [ ] Lookup shape via **`_shape_by_id_recursive`** (§5.1 #3A).
- [ ] `_shrink_shape_font` (helper **solo edit**, §5.1 #9B) mutando `a:rPr/@sz`; rispettare valve `min_pt` / `step_pt`.
- [ ] `_text_fits_shape_heuristic` v1 (caratteri × stima vs `shape.width/height`).
- [ ] `_save` output + `_emit_link`.
- [ ] Docstring tool con schema JSON esempio.
- [ ] Messaggio errore se valve disabilitata.

**Deliverable:** edit manuale via JSON funzionante **con Files API** (UUID obbligatorio — §5.1 #1A).

---

### Fase 2 — Test e production checks

**Obiettivo:** regressione automatica.

- [ ] `examples/edit_presentation_smoke.py`: costruisce o usa pptx fixture (textbox piccolo + testo lungo), applica shrink, assert `sz` diminuito.
- [ ] Estendere fixture reuse local / Marketing: opzionale assert post-edit su una slide.
- [ ] Registrare smoke in `examples/run_production_checks.py`.
- [ ] Eseguire battery template esistente (nessuna regressione).

**Deliverable:** CI locale verde con edit smoke.

---

### Fase 3 — Operazioni v2 (opzionale post-MVP)

**Obiettivo:** meno errori modello + split overlap + autofit PowerPoint.

- [ ] **`enable_autofit`** — priorità §5.1 #4B; XML `a:bodyPr` / `normAutofit`.
- [ ] **`replace_text_and_fit`** — §5.1 #2B (testo + shrink atomico).
- [ ] **`split_text`** — obbligatorio in v2 §5.1 #5B; `first_paragraph` | `first_line` tra due `shape_id`.
- [ ] **`resize_shape`** — opzionale §5.1 #5C; delta altezza textbox conservativo.
- [ ] Tabelle: **`set_table_cell`** (6B) o refill parziale (6C) — una scelta in implementazione.
- [ ] Test dedicati split + autofit (Indian/Marketing se disponibili).
- [ ] (Opz.) Validazione inspect compatto su output reuse — se insufficiente, **10B** param inspect.

**Deliverable:** casi overlap titolo/body e “adatta al riquadro” senza rigenerare.

---

### Fase 4 — Hardening

**Obiettivo:** produzione Neura.

- [ ] `_emit` progress su deck molti slide / molte operazioni.
- [ ] Valve `presentation_edit_min_font_pt` / default `step_pt`.
- [ ] Metriche log: `edit_ms`, `ops_applied`, `ops_failed`.
- [ ] (v2 opz.) Save parziale + report se batch ops (#8B).
- [ ] (Opz. prod) Rifiuto `presentation_file_id` = template allegato (#7B).
- [ ] Test ACL cross-user (stesso pattern reference).
- [ ] Version bump `generate_slides.py` header + CHANGELOG voce edit.

**Deliverable:** pilot abilitabile con valve ON.

---

### Fase 5 — Rollout Neura

**Obiettivo:** adozione senza confusione template mode.

- [ ] Aggiornare system prompt: “fix visivo post-download → `edit_presentation`, non rigenerare”.
- [ ] Esempi chat: titolo troppo grande; split slide 4.
- [ ] Abilitare valve su workspace pilot.
- [ ] Monitor log errori shape id / slide index (prime settimane).

**Deliverable:** flusso utente descritto in §12 operativo in produzione.

---

## 12. Appendice: esempio end-to-end

### Premessa

L’utente ha già ricevuto un deck da `generate_slides` (template Indian, 35 slide reuse). Ha aperto il file e vede overlap titolo/body sulla slide “Formazione educatori”.

### Passo 1 — Utente (chat)

> “Nella slide 12 il titolo FORMAZIONE EDUCATORI è troppo grande e copre il paragrafo sotto. Riduci il titolo e lascia il testo descrittivo nel riquadro sotto.”

### Passo 2 — Neura (ragionamento, non tool)

- Slide utente **12** → JSON `"slide": 11`.
- Verificare che l’ultimo generate abbia restituito un **UUID** (`presentation_file_id` — §5.1 #1A); non usare `reference_file_id` del template (#7A).
- Opzionale: `inspect_slides(presentation_file_id)` sul file **generato** (#10A) per `shape_id` titolo vs body.
- Costruisce JSON (esempio v1 — ordine **set_text poi shrink** sul titolo, §5.1 #2A):

```json
{
  "presentation_file_id": "a1b2c3d4-e5f6-7890-abcd-ef1234567890",
  "title": "Rita Levi-Montalcini (edited)",
  "operations": [
    {
      "op": "set_text",
      "slide": 11,
      "shape_id": 805,
      "text": "FORMAZIONE EDUCATORI"
    },
    {
      "op": "shrink_font",
      "slide": 11,
      "shape_id": 805,
      "min_pt": 20,
      "step_pt": 2
    },
    {
      "op": "set_text",
      "slide": 11,
      "shape_id": 806,
      "text": "I corsi di formazione AMI e ONM preparano educatori qualificati per ogni fascia d'età, dal nido alla scuola superiore."
    }
  ]
}
```

(Gli id 805/806 sono **esemplificativi** — da inspect sul deck generato. Se titolo e body sono **uno stesso shape**, v2 **`split_text`** §5.1 #5B o **`resize_shape`** #5C.)

**Nota v2:** stesso caso con una sola op: `replace_text_and_fit` su shape 805 (#2B).

### Passo 3 — `edit_presentation`

- Scarica il pptx, applica operazioni in ordine, salva.
- Restituisce nuovo link download.

### Passo 4 — Utente

Scarica la versione editata; eventuali ulteriori tweak → nuova chiamata edit sul **nuovo** `file_id`.

---

## 13. Checklist pre-merge (produzione Neura)

- [ ] Valve `presentation_edit_enabled` default **false** in repo.
- [ ] Nessuna modifica al diagramma §4 di `template-mode-implementation-plan.md` (solo cross-link opzionale in README o in §1 di questo doc).
- [ ] `examples/edit_presentation_smoke.py` PASS in Docker.
- [ ] Battery `run_production_checks` (escluso NF5 se golden noto) PASS.
- [ ] Docstring `edit_presentation` esplicita: **non** sostituisce generate; **non** usa `reference_file_id`.
- [ ] `doc/neura-presentation-edit-hints.md` pubblicato per team modello.
- [ ] Tool registrato in UI Open WebUI (stesso plugin `generate_slides.py`).
- [ ] CHANGELOG entry “Presentation Edit (opt-in)”.

---

*Documento di implementazione Presentation Edit — IANUSTEC / NEURA. Indipendente dal Template Mode §4; rev. 2026-09-30 (§5.1 decisioni gap).*
