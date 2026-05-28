# 🤖 AI Debate Pipeline — Claude vs GPT-4o  (v2)

Claude és GPT-4o vitáznak egy **mappa** tartalmáról (vagy két fájlról), a te promptod
által meghatározott cél mentén — majd közösen szintetizálnak egy Word dokumentumot.

## Telepítés

```bash
pip install anthropic openai python-docx
npm install -g docx
```

## API kulcsok

```powershell
# Windows PowerShell
$env:ANTHROPIC_API_KEY = "sk-ant-..."
$env:OPENAI_API_KEY    = "sk-..."
```

---

## Használat

### 1. MAPPA MÓD — ajánlott ✅

```powershell
python ai_debate.py --folder D:\Munka\projektek\docs `
    --prompt "Ez a mappa egy SaaS termék terveit tartalmazza. Célom: értékeljétek az architektúrát és a go-to-market stratégiát, hozzátok fel a gyengeségeket, de legyetek kompromisszumkészek a technikai megvalósítás kérdéseiben."
```

Támogatott fájltípusok a mappában: `.docx`, `.txt`, `.md`, `.pdf`

### 2. PROMPT FÁJLBÓL — hosszabb feladatleíráshoz

```powershell
# feladat.txt tartalma: részletes feladatleírás
python ai_debate.py --folder ./sources --prompt-file feladat.txt --rounds 4
```

### 3. KÉT EXPLICIT FÁJL

```powershell
python ai_debate.py tervezet_A.docx tervezet_B.docx `
    --prompt "Melyik üzleti terv életképesebb magyar KKV piacon?"
```

### 4. ANGOL KIMENET

```powershell
python ai_debate.py --folder ./docs --prompt "Evaluate these for enterprise readiness." --lang en
```

---

## Kapcsolók

| Kapcsoló | Leírás | Alapértelmezett |
|----------|--------|-----------------|
| `--folder MAPPA` | Forrásmappa összes dokumentuma | — |
| `--prompt "..."` | Feladat / kontextus szöveg | (üres) |
| `--prompt-file f.txt` | Prompt betöltése fájlból | — |
| `--rounds N` | Vita körök száma | 3 |
| `--output fajl.docx` | Kimeneti fájl neve | synthesis_output.docx |
| `--output-dir ./dir` | Kimeneti könyvtár | . (jelenlegi) |
| `--lang hu\|en` | Válasz nyelve | hu |
| `--max-chars N` | Max karakter/forrás (token védelem) | 8000 |

---

## A folyamat

```
[--folder ./docs]  +  [--prompt "célom: ..."]
        │
        ▼
  Összes .docx/.txt/.md/.pdf beolvasása
        │
        ▼
  ┌─────────────────────────────────────────┐
  │  KÖR 1..N                               │
  │  Claude elemez & érvel                  │
  │      ↕  (láncolt üzenet história)       │
  │  GPT-4o reagál & ellenérvel             │
  └─────────────────────────────────────────┘
        │
        ▼
  Claude szintetizálja a vita eredményét
        │
        ▼
  synthesis_output.docx  ✅
  debate_log.json        ✅
  debate_transcript.txt  ✅
```

---

## Kimenet

| Fájl | Tartalom |
|------|----------|
| `synthesis_output.docx` | Cél + Forrásanyagok + Fejezetek + Ajánlások + Összefoglalás + Vita kivonat |
| `debate_transcript.txt` | Claude↔GPT teljes párbeszéd olvashatóan |
| `debate_log.json` | Gépi formátum (prompt + vita) |

---

## Prompt írási tippek

A `--prompt` a vita irányát és mélységét szabja meg. Érdemes beleírni:

- **Mit tartalmaz a mappa**: "Ez a mappa a DocuAgent termék 3 tervezési dokumentumát tartalmazza..."
- **Mi a cél**: "Szeretném, ha értékelnétek az architektúrát KKV szempontból..."
- **Mire fókuszáljanak**: "Különös figyelmet fordítsatok a GDPR megfelelőségre és az árazási modellre."
- **Milyen kompromisszumot várunk**: "A technikai megoldásokban legyetek rugalmasak, az üzleti modellben szigorúak."

### Példa prompt (feladat.txt):

```
Ez a mappa az Agentify DocuAgent termékének tervdokumait tartalmazza:
- egy technikai architektúra leírást
- egy go-to-market stratégiai dokumentumot
- egy konkurenciaelemzést

CÉLOM:
Értékeljétek mindhárom dokumentumot együtt. Vitassátok meg:
1. Mennyire koherens a technikai és az üzleti stratégia?
2. Hol vannak hiányosságok vagy ellentmondások?
3. Mi az, amiben mindketten egyetértetek mint erősség?

Magyar KKV piac (főleg könyvelők, ügyvédek) a célcsoport. Legyetek
kritikusak, de kompromisszumkészek. A végső dokumentum legyen actionable.
```
