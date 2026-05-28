# AI Debate Pipeline — Dokumentáció

> Multi-model Expert Council vita rendszer: több AI modell strukturált vitája dokumentumokról,  
> végső üzleti terv, megvalósítási terv és meeting report automatikus generálásával.

---

## Tartalomjegyzék

1. [Áttekintés](#1-áttekintés)
2. [Telepítés](#2-telepítés)
3. [Gyors start](#3-gyors-start)
4. [Hogyan működik — a folyamat](#4-hogyan-működik--a-folyamat)
5. [Scenariók](#5-scenariók)
6. [Modellek és szerepek](#6-modellek-és-szerepek)
7. [Bemeneti forrástípusok](#7-bemeneti-forrástípusok)
8. [Kimeneti fájlok](#9-kimeneti-fájlok)
9. [Kapcsolók referencia](#9-kapcsolók-referencia)
10. [Konfiguráció — .env](#10-konfiguráció--env)
11. [Tipikus use case-ek](#11-tipikus-use-case-ek)
12. [Hibaelhárítás](#12-hibaelhárítás)
13. [Architektúra](#13-architektúra)

---

## 1. Áttekintés

Az `ai_debate.py` egy Python CLI eszköz, amely több AI modellt (Claude, GPT, Gemini, DeepSeek, Grok) szervezett, többfázisú vitába rendez egy vagy több forrásanyag elemzéséhez.

A vita nem egyszerű "mindenki reagál mindenkire" lánc — hanem strukturált **Expert Council** folyamat:

- **Független vélemények** — a modellek NEM látják egymás válaszát az első körben (anchoring elleni védelem)
- **Moderátor issue matrix** — nem összefoglaló, hanem vita-térkép: konszenzusok, konfliktusok, top kérdések
- **Célzott rebuttal** — csak a vitatott pontokra reagálnak, nem egymás teljes szövegére
- **Álláspont-frissítés** — strukturált JSON: ki változtatott véleményt, mennyire biztos, mi maradt vita
- **Külön Final Judge** — nem ugyanaz a modell szintetizál, aki moderált (elfogultság ellen)

### Miért jobb ez, mint egy sima AI-kérdés?

| Módszer | Probléma |
|---------|----------|
| Egyetlen AI-tól kérdezni | Egy perspektíva, nincs ütköztetés |
| Sorban kérdezni több AI-t | Anchoring: mindenki az előző után alakítja véleményét |
| Egyszerű "vitáztasd" prompt | Az utolsó modell narratív előnyt kap, nincsenek strukturált döntések |
| **AI Debate Pipeline** | Független vélemények → ütköztetés → konszenzus → külön bíró |

---

## 2. Telepítés

### Python csomagok

```powershell
pip install anthropic openai google-genai python-docx openpyxl python-dotenv
```

### Node.js (Word fájl generáláshoz)

```powershell
npm install -g docx
# vagy projekt-szinten:
cd D:\Munka\AI2AI
npm init -y
npm install docx
```

### API kulcsok — `.env` fájl

Másold a `.env.example`-t és töltsd ki:

```powershell
copy .env.example .env
```

`.env` tartalma:

```env
ANTHROPIC_API_KEY=sk-ant-...
OPENAI_API_KEY=sk-...
GOOGLE_API_KEY=...
DEEPSEEK_API_KEY=sk-...
XAI_API_KEY=xai-...

DEFAULT_SCENARIO=quick
DEFAULT_QUALITY=balanced
```

> A script gracefulan kihagyja azt a modellt amelynek nincs API kulcsa — a többi fut tovább.

---

## 3. Gyors start

```powershell
# Legegyszerűbb futtatás (quick scenario, 3 modell)
python ai_debate.py dokumentum.docx --prompt "Értékeld ezt üzleti szempontból"

# Mappa összes fájlja
python ai_debate.py --folder ./docs --prompt "..."

# ZIP projekt + tervdokumentumok
python ai_debate.py projekt.zip terv.docx --prompt-file feladat.txt

# Előbb becsüld meg a költséget
python ai_debate.py ... --estimate

# Teljes expert council
python ai_debate.py ... --scenario expert-council --quality best --parallel

# Megszakadt futás folytatása
python ai_debate.py ... --resume eredmenyek\debate_log.json
```

---

## 4. Hogyan működik — a folyamat

### `quick` scenario (default, 3 modell)

```
Forrásanyagok
     │
     ▼
[PHASE 0] Evidence Pack
     Claude Sonnet tömöríti a forrásokat strukturált csomaggá
     → Ezt kapják a vitázók, nem a nyers fájlokat (token és anchoring védelem)
     │
     ▼
[PHASE 1] Független vélemények (párhuzamosan futhat)
     GPT-4o      → saját álláspont  ┐
     Gemini      → saját álláspont  ┘ NEM látják egymást
     │
     ▼
[PHASE 5] Final Judge
     GPT szintetizál → üzleti terv + megvalósítási terv + AI kontextus
     │
     ▼
Kimenetek: .docx + szintézis.md + meeting_report.md
```

### `expert-council` scenario (5 fázis, 6 modell)

```
[PHASE 0] Evidence Pack         ← Claude Sonnet moderátor
[PHASE 1] Független vélemények  ← GPT + Claude + Gemini + DeepSeek + Grok párhuzamosan
[PHASE 2] Issue Matrix          ← Claude Sonnet azonosítja: konszenzus / vita / top 3 kérdés
[PHASE 3] Célzott Rebuttal      ← Mindenki CSAK a top 3 vitakérdésre reagál
[PHASE 4] Álláspont-frissítés   ← JSON: changed_mind, confidence, final_recommendation
[PHASE 5] Final Judge           ← GPT-best (nem a moderátor!) hozza a végső döntést
```

### Miért nem látják egymást Phase 1-ben?

Ha a modellek látják az előző válaszokat, **anchoring** lép fel: a 3. modell már a 2. keretezésében gondolkodik. A független vélemények garantálják, hogy minden perspektíva valóban önálló.

### Miért külön a Final Judge?

Ha Claude moderál és Claude szintetizál is, az eredmény Anthropic-elfogult lehet. A Final Judge mindig **más modell**, mint a moderátor.

### Az Evidence Pack szerepe

Ahelyett hogy minden modell megkapná az összes nyers fájlt (drága és zajos), a moderátor először elkészít egy strukturált összefoglalót:

```
## USER_GOAL        ← mit akar a felhasználó
## KEY_FACTS        ← 10-15 legfontosabb tény a forrásokból
## CONSTRAINTS      ← korlátok, feltételek
## OPEN_QUESTIONS   ← megválaszolatlan kérdések
## EVIDENCE_SNIPPETS← kulcsrészletek kontextussal
## TASK_FOR_MODELS  ← mit várunk a vitázóktól
```

Ez ~70-80%-kal csökkenti a token-használatot és javítja a fókuszt.

---

## 5. Scenariók

### `quick` — Default, gyors értékelés

```powershell
python ai_debate.py dok.docx --prompt "..." --scenario quick
```

- **3 modell:** Moderátor + 2 debater + Judge
- **Fázisok:** Evidence → Független → Judge
- **Mikor használd:** Gyors dokumentum review, email draft értékelés, rövid döntés
- **Kb. költség:** $0.05–0.15

### `expert-council` — Alapértelmezett komoly elemzéshez

```powershell
python ai_debate.py projekt.zip terv.docx --prompt "..." --scenario expert-council
```

- **6 modell:** Moderátor + Stratégista + Mérnök + Piac-elemző + Szkeptikus + Döntéshozó + Judge
- **Fázisok:** Mind az 5 (evidence → independent → matrix → rebuttal → revision → judge)
- **Mikor használd:** Üzleti terv, architektúra döntés, termékötlet, fejlesztési terv
- **Kb. költség:** $0.50–2.00

### `red-team` — Kritikus döntések, magas tétű helyzetek

```powershell
python ai_debate.py ... --scenario red-team --quality best
```

- **6 modell:** 2 szkeptikus / red-team szerepben
- **Mikor használd:** Befektetői anyag, biztonsági kérdés, jogi kockázat, enterprise termék
- **Különlegesség:** Két független szkeptikus — az egyik üzleti, a másik technikai kritika

### `build-plan` — Implementációs roadmap

```powershell
python ai_debate.py ... --scenario build-plan
```

- **Szerepek:** Product Manager + Senior Architect + Security + QA + PM Judge
- **Mikor használd:** Sprint terv, developer promptok, Claude Code / Codex workflow
- **Kimenet:** Sprint bontás, definition of done, technikai kockázatok

---

## 6. Modellek és szerepek

### Modell katalógus

| Kulcs | Modell (balanced) | Szerep | Erősség |
|-------|------------------|--------|---------|
| `claude-sonnet` | claude-sonnet-4-6 | Moderátor, Engineer | Analitikus, architektúra |
| `claude-opus` | claude-opus-4-7 | Best engineer/strategist | Komplex reasoning |
| `gpt` | gpt-5.4 | Stratégista, Judge | GTM, üzleti logika |
| `gpt-best` | gpt-5.5 | Best Judge | Komplex döntés |
| `gemini` | gemini-3.1-pro-preview | Piac-elemző | Nagy kontextus, trend |
| `deepseek` | deepseek-v4-flash | Cost reasoner | Olcsó, erős reasoning |
| `deepseek-pro` | deepseek-v4-pro | Best reasoner | Komplex elemzés |
| `grok` | grok-4.3 | Szkeptikus | Kritikai gondolkodás |
| `grok-best` | grok-4.20 | Red team | Agresszív kritika |

### Quality presets

| Quality | Claude | GPT | Gemini | DeepSeek | Grok | Kb. ár |
|---------|--------|-----|--------|----------|------|--------|
| `fast` | Haiku | gpt-5.4-mini | gemini-2.5-flash | v4-flash | grok-4.3 | ~$0.02–0.10 |
| `balanced` | Sonnet | gpt-5.4 | gemini-3.1-pro | v4-flash | grok-4.3 | ~$0.20–1.00 |
| `best` | Opus | gpt-5.5 | gemini-3.1-pro | v4-pro | grok-4.20 | ~$1.00–5.00 |

### Szerepek felülírása

```powershell
# Grok legyen a judge, DeepSeek a szkeptikus
python ai_debate.py ... --roles judge=grok,skeptic=deepseek

# Claude Opus legyen az engineer
python ai_debate.py ... --roles engineer=claude-opus
```

---

## 7. Bemeneti forrástípusok

A script automatikusan felismeri és feldolgozza:

| Típus | Kiterjesztések | Feldolgozás |
|-------|---------------|-------------|
| **ZIP / TAR** | .zip .tar .gz .tgz | Kicsomagol, rekurzívan feldolgoz mindent |
| **Word** | .docx .doc | Szöveg + táblázatok kinyerése |
| **PDF** | .pdf | Szöveg kinyerés (PyMuPDF vagy pypdf) |
| **Excel / CSV** | .xlsx .xls .csv .tsv | Táblázat → szöveges összefoglaló |
| **Kód** | .ts .tsx .py .js .go stb. | Közvetlen szöveg |
| **Markdown** | .md .txt .rst | Közvetlen szöveg |
| **JSON / YAML** | .json .yaml .toml | Közvetlen szöveg |
| **Képek** | .png .jpg .webp .gif | Base64 → Vision API (Claude + GPT látja) |
| **Mappa** | — | Összes fájl rekurzívan |

### Automatikusan kizárva (boilerplate)

- `node_modules/`, `.git/`, `dist/`, `build/`
- `components/ui/` (shadcn/ui boilerplate)
- `package-lock.json`, `yarn.lock`, `bun.lockb`
- `.min.js`, `.min.css` (minified fájlok)
- Bináris fájlok

### Több forrás egyszerre

```powershell
# ZIP + dokumentumok + kép vegyesen
python ai_debate.py projekt.zip terv1.docx terv2.pdf diagram.png `
    --prompt "Ezek együtt alkotják a projektet..."
```

---

## 8. Kimeneti fájlok

Minden futás az `--output-dir` mappába ír:

| Fájl | Tartalom |
|------|----------|
| `*.docx` | Formázott Word dokumentum (szintézis + döntések) |
| `*_szintezis.md` | **Főbb kimenet:** Üzleti terv + Megvalósítási terv + AI kontextus blokk |
| `*_meeting_report.md` | **Ki mit mondott:** Fázisok átirata, döntési nyomvonal, álláspont-frissítések |
| `debate_log.json` | Teljes vita gépi formátumban (resume-hoz és feldolgozáshoz) |
| `debate_transcript.txt` | Teljes vita emberi olvasásra |

### Szintézis MD szerkezete

```markdown
# Projekt — Végső Dokumentáció

## ⚖️ Végső ítélet

# VÉGSŐ ÜZLETI TERV
## 1. Piaci helyzet és lehetőség
## 2. Termék és értékajánlat
## 3. Jelenlegi állapot és korlátok
## 4. Bevételi modell (Freemium/Pro/Team/Enterprise)
## 5. Go-to-Market stratégia (30/60/90 nap)
## 6. Versenyelőnyök
## 7. Versenytársak
## 8. Kockázatok
## 9. Sikerkritériumok

# VÉGSŐ MEGVALÓSÍTÁSI TERV
## 1. Jelenlegi architektúra
## 2. Célarchitektúra
## 3. Azonnali teendők (0-2 hét)
## 4. Sprint 1 — Backend alapok
## 5. Sprint 2 — AI integráció
## 6. Sprint 3 — GitHub App és CLI
## 7. Sprint 4+ — Enterprise
## 8. Technikai kockázatok

# AI KONTEXTUS BLOKK
> Claude/GPT/Gemini system promptba csatolható összefoglaló

## ✅ Konszenzus-pontok
## ❓ Nyitott kérdések
## ⚠️ Kockázati napló
## 📋 Döntési napló
## 📌 Ajánlások
## 🗺️ Implementációs terv
```

### Meeting Report szerkezete

```markdown
# Expert Council — Meeting Report

## 👥 Résztvevők és szerepek

## ⚡ PHASE 0 — Evidence Pack
   Claude Sonnet [moderator]: [mit talált a forrásokban]

## 🔍 PHASE 1 — Független álláspontok
   GPT-4o [strategist]: [üzleti elemzés — első 800 szó]
   Claude Sonnet [engineer]: [technikai elemzés]
   Gemini [market]: [piaci elemzés]
   ...

## 📊 PHASE 2 — Issue Matrix
   [Moderátor vita-térképe: konszenzusok, konfliktusok, top 3 kérdés]

## ⚔️ PHASE 3 — Rebuttal
   [Ki hogyan reagált a vitapontokra]

## 🔄 PHASE 4 — Álláspont-frissítések
   | Modell | Változtatott? | Bizalom | Mit változtatott |

## ⚖️ PHASE 5 — Final Judge döntése
   [Verdict + konszenzusok + döntési napló + ajánlások]

## 📌 Összesített konklúzió
```

---

## 9. Kapcsolók referencia

```
python ai_debate.py [FORRÁS...] [KAPCSOLÓK]

FORRÁS:
  Fájlok, mappák, ZIP-ek vegyesen megadhatók

FORRÁS KAPCSOLÓK:
  --folder MAPPA        Forrásmappa (a FORRÁS listához adódik)
  --prompt SZÖVEG       Feladat / kontextus szöveg
  --prompt-file FÁJL    Prompt betöltése .txt fájlból (hosszabb feladathoz)

SCENARIO:
  --scenario SCENARIO   quick | expert-council | red-team | build-plan
                        Default: quick (vagy DEFAULT_SCENARIO env változó)
  --quality QUALITY     fast | balanced | best
                        Default: balanced (vagy DEFAULT_QUALITY env változó)
  --roles ROLE=MODEL    Szerepek felülírása: --roles judge=gpt-best,skeptic=grok

FUTTATÁS:
  --parallel            Phase 1 párhuzamos API hívások (gyorsabb, de streaming nem látható)
  --rounds N|auto       Vita körök száma (auto: max 8, konszenzusnál leáll)
  --estimate            Csak token/költség becslés, futtatás nélkül
  --resume LOG.json     Megszakadt futás folytatása a mentett logból

KIMENET:
  --output FÁJL.docx    Kimeneti Word fájl neve (default: synthesis_output.docx)
  --output-dir MAPPA    Kimeneti könyvtár (default: jelenlegi mappa)
  --lang hu|en          Válasz nyelve (default: hu)
  --no-docx             Word generálás kihagyása (csak MD és log)

FORRÁS KEZELÉS:
  --max-chars N         Max karakter/forrás (default: 10000)
```

---

## 10. Konfiguráció — .env

A script soha nem tartalmaz hardkódolt modell neveket — mindent az `.env`-ből tölt be:

```env
# API kulcsok
ANTHROPIC_API_KEY=
OPENAI_API_KEY=
GOOGLE_API_KEY=
DEEPSEEK_API_KEY=
XAI_API_KEY=

# Alapértelmezések (felülírhatók CLI kapcsolókkal)
DEFAULT_SCENARIO=quick
DEFAULT_QUALITY=balanced
DEFAULT_PARALLEL=false
MAX_SOURCE_CHARS=10000
MAX_OUTPUT_TOKENS=2000

# OpenAI modellek (Responses API, reasoning_effort támogatással)
OPENAI_FAST_MODEL=gpt-5.4-mini
OPENAI_BALANCED_MODEL=gpt-5.4
OPENAI_BEST_MODEL=gpt-5.5
OPENAI_REASONING_EFFORT_FAST=low
OPENAI_REASONING_EFFORT_BALANCED=medium
OPENAI_REASONING_EFFORT_BEST=high

# Claude modellek
CLAUDE_FAST_MODEL=claude-sonnet-4-6
CLAUDE_BALANCED_MODEL=claude-sonnet-4-6
CLAUDE_BEST_MODEL=claude-opus-4-7

# Gemini modellek
GEMINI_FAST_MODEL=gemini-2.5-flash
GEMINI_BALANCED_MODEL=gemini-3.1-pro-preview
GEMINI_BEST_MODEL=gemini-3.1-pro-preview

# DeepSeek modellek (deepseek-chat/reasoner retire: 2026-07-24)
DEEPSEEK_FAST_MODEL=deepseek-v4-flash
DEEPSEEK_BALANCED_MODEL=deepseek-v4-flash
DEEPSEEK_BEST_MODEL=deepseek-v4-pro

# Grok modellek (grok-4/grok-4-fast retired: 2026-05-15)
GROK_FAST_MODEL=grok-4.3
GROK_BALANCED_MODEL=grok-4.3
GROK_BEST_MODEL=grok-4.20
```

> **Fontos:** A `.env` soha ne kerüljön git-be. Csak `.env.example` commitolható.

---

## 11. Tipikus use case-ek

### Projekt audit és üzleti terv

```powershell
python ai_debate.py `
    projekt.zip uzleti_terv.docx megvalositasi_terv.docx `
    --prompt-file feladat.txt `
    --scenario expert-council `
    --quality balanced `
    --parallel `
    --output-dir eredmenyek
```

**feladat.txt tartalma:**
```
Ez a mappa egy SaaS termék terveit tartalmazza.
Vizsgáljátok meg a forráskódot és mindkét tervdokumentumot.
Készítsetek:
1. Végső üzleti terv — piac, GTM, bevételi modell
2. Végső megvalósítási terv — sprint bontás, architektúra
3. AI kontextus markdown — system promptba csatolható
```

### Architektúra döntés

```powershell
python ai_debate.py `
    architektura_A.md architektura_B.md `
    --prompt "Melyik architektúra alkalmasabb 100k+ felhasználóra? Indokoljátok." `
    --scenario expert-council `
    --roles engineer=claude-opus,judge=gpt-best
```

### Gyors dokumentum review

```powershell
python ai_debate.py ajanlat.docx `
    --prompt "Értékeld ezt az ajánlatot ügyfél szemszögéből." `
    --scenario quick `
    --quality fast
```

### Befektetői anyag red-team

```powershell
python ai_debate.py pitch_deck.pdf financials.xlsx `
    --prompt "Legyetek befektetők. Mit nem hisztek el? Mi a gyenge pont?" `
    --scenario red-team `
    --quality best
```

### Sprint terv generálás

```powershell
python ai_debate.py projekt.zip prd.md `
    --prompt "Készítsetek 4 sprintre bontott fejlesztési tervet." `
    --scenario build-plan `
    --quality balanced
```

---

## 12. Hibaelhárítás

### `Cannot find module 'docx'` (Node.js hiba)

```powershell
# Globális telepítés
npm install -g docx

# Ha nem találja, projekt-szintű telepítés
cd D:\Munka\AI2AI
npm init -y
npm install docx
```

### `credit balance too low` (Anthropic/OpenAI hiba)

A vita részlegesen lefutott. A log mentve van — `--resume`-mal folytatható:

```powershell
python ai_debate.py ... --resume eredmenyek\debate_log.json
```

### `[Model] hiba: Error code: 403`

Az adott modell API kulcsa érvénytelen vagy nincs kredit. A script kihagyja és fut tovább a többi modellel. Ellenőrizd az adott platform console-ját.

### Gemini nem válaszol

```powershell
pip install google-genai
```

### PDF nem olvasható

```powershell
pip install pymupdf
```

### JSON parse hiba a szintézisnél

A Final Judge néha nem ad vissza tiszta JSON-t. A script ilyenkor fallback struktúrát használ. `--quality best` és `--no-docx` kombinációval próbáld újra — a markdown kimenet robusztusabb.

---

## 13. Architektúra

```
ai_debate.py
│
├── Konfiguráció
│   ├── CATALOG          ← modell definíciók (.env-ből töltve)
│   └── SCENARIOS        ← scenario + quality_map definíciók
│
├── Fájlkezelés
│   ├── load_sources()   ← ZIP kicsomagolás, rekurzív betöltés
│   ├── _proc_file()     ← típusonkénti feldolgozás
│   └── DocItem          ← egységes forráselem (text vagy image)
│
├── AI Engine
│   └── AIEngine         ← egységes interfész minden modellhez
│       ├── anthropic    ← streaming messages.stream()
│       ├── openai_responses ← GPT-5.x Responses API + reasoning_effort
│       ├── openai_compat ← DeepSeek, Grok (OpenAI-kompatibilis)
│       └── gemini       ← google-genai SDK
│
├── Fázisok
│   ├── phase_evidence()     ← Evidence Pack generálás
│   ├── phase_independent()  ← Párhuzamos/szekvenciális vélemények
│   ├── phase_issue_matrix() ← Moderátor vita-térkép
│   ├── phase_rebuttal()     ← Célzott rebuttal
│   ├── phase_revision()     ← JSON álláspont-frissítés
│   └── phase_judge()        ← Final Judge szintézis
│
├── Scenariók
│   ├── run_quick()          ← 3 fázis
│   └── run_expert_council() ← 6 fázis (expert-council + red-team + build-plan)
│
├── Kimenetek
│   ├── save_markdown()      ← Szintézis + üzleti terv + megvalósítási terv
│   ├── save_meeting_report() ← Ki mit mondott, döntési nyomvonal
│   ├── create_docx()        ← Word fájl (Node.js docx csomag)
│   └── save_logs()          ← JSON log + szöveges átirat
│
└── Segédeszközök
    ├── estimate_cost()      ← Token/költség becslő
    ├── build_roles()        ← Szerepek + model override kezelése
    └── _autosave()          ← Kör utáni automatikus mentés (crash protection)
```

### Adatfolyam

```
Forrásanyagok (ZIP/docx/pdf/kód/kép)
        │
        ▼ load_sources()
    [DocItem lista]
        │
        ▼ phase_evidence()
    [Evidence Package — moderátor által tömörítve]
        │
        ├──► phase_independent() → [vélemények dict]
        │         │
        │         ▼ phase_issue_matrix()
        │    [issue matrix szöveg]
        │         │
        │         ▼ phase_rebuttal()
        │    [rebuttals dict]
        │         │
        │         ▼ phase_revision()
        │    [revisions dict — JSON struktúra]
        │         │
        └──────── ▼ phase_judge()
              [synthesis dict — üzleti terv + megvalósítási terv + AI ctx]
                   │
        ┌──────────┼──────────┐
        ▼          ▼          ▼
   save_markdown  create_docx  save_meeting_report
   save_logs
```

---

*Verzió: v5 | 2026-05-27*  
*Modellek: Claude Opus 4.7, GPT-5.5, Gemini 3.1 Pro, DeepSeek V4, Grok 4.3/4.20*
