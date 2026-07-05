# Guida al rilascio (checklist riusabile)

Processo per pubblicare questo tool (e i prossimi) come open-source. La **fonte di
verità resta il monorepo privato**; questa cartella è l'area di *packaging* da cui
si esporta la versione pubblica.

## 0. Sanificazione (già fatta per questo tool)
- [x] Frontmatter `author` / `author_url` / `license` / `version` corretti
- [x] Nessun segreto, token, URL interno o path privato
- [x] Nessun import verso altri tool non pubblicati
- [x] Nessun riferimento in commenti a tool interni
- [x] `python3 -m py_compile` ok

## 1. Repo GitHub dedicato
1. Crea un repo pubblico, es. `openwebui-generate-slides` sotto l'org **ianustec**.
2. Copia il contenuto di questa cartella nella root del repo.
3. `git init && git add . && git commit -m "feat: native PPTX slides engine for Open WebUI"`
4. `git branch -M main && git remote add origin git@github.com:ianustec/openwebui-generate-slides.git && git push -u origin main`
5. Aggiungi topic al repo: `open-webui`, `openwebui`, `pptx`, `python-pptx`, `powerpoint`, `ai-tools`.
6. Carica gli screenshot in `assets/` e scommenta la riga `![Anteprima]` nel README.

## 2. Community Open WebUI
1. Vai su https://openwebui.com/ → accedi → **Tools → Create**.
2. Incolla l'intero contenuto di `generate_slides.py` (il frontmatter popola titolo/descrizione/versione).
3. Aggiungi lo stesso set di screenshot e, nella descrizione, il link al repo GitHub.
4. Pubblica. La scheda diventa importabile con **Get** dalle istanze Open WebUI.

> Nota: sulla community conta il **singolo file .py**. Il repo GitHub è per issue,
> stelle, changelog e link dal post Reddit.

## 3. Reddit
- Subreddit primario: **r/OpenWebUI**. Secondari: **r/LocalLLaMA**, **r/selfhosted**.
- Titolo esempio: *"I built a native PowerPoint (.pptx) generator tool for Open WebUI — layouts, native charts, themes [MIT]"*.
- Corpo: 1 frase sul problema, 3-4 bullet di feature, 2-3 screenshot, link GitHub + link community.
- Rispondi ai commenti nelle prime ore (aiuta il ranking).

## 4. Manutenzione
- Bump `version` nel frontmatter ad ogni release e aggiorna `CHANGELOG` nel repo.
- Le fix nel monorepo privato vanno **riportate qui** e ri-esportate (la copia non è un symlink).

---

## Prossimi tool da rilasciare (uno per volta, repo separati)
Candidati nel monorepo, in ordine di "pronti al pubblico":
- `generate_word.py` → documenti DOCX nativi
- `openwebui_dashboards.py` → dashboard BI HTML server-side
- (altri) `openwebui_todos.py`, `openwebui_subagents.py`, `neura_knowledge_*`

Per ognuno: duplica questa cartella come template, sanifica, adatta README/esempi.
