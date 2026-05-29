# AI Debate Pipeline

AI Debate Pipeline is a Python CLI tool that lets multiple AI providers debate a task, compare perspectives, and synthesize final artifacts.

The current pipeline supports:
- provider/model health checks and fallback,
- failed-response exclusion,
- TaskProfile-based task understanding,
- OutputContract-based final document structure,
- multi-step synthesis with validation and repair,
- markdown synthesis, meeting report, transcript, `debate_log.json`, and optional docx output.

## CLI Usage

Basic quick run:

```powershell
python -B ai_debate.py --folder .\source --prompt-file .\feladat.txt --scenario quick --quality fast --no-docx --output-dir .\eredmenyek
```

Task profile only:

```powershell
python -B ai_debate.py --folder .\source --prompt-file .\feladat.txt --scenario quick --quality fast --task-profile-only --output-dir .\eredmenyek_profile
```

Run with an OutputContract:

```powershell
python -B ai_debate.py --folder .\source --prompt-file .\feladat.txt --scenario quick --quality fast --contract-file .\contracts\business_master_plan.json --no-docx --output-dir .\eredmenyek_business
```

Smoke test:

```powershell
python -B ai_debate.py --smoke-test
```

## Common CLI Options

| Option | Purpose |
| --- | --- |
| `--folder PATH` | Load source material from a folder. |
| `--prompt TEXT` | Provide the task directly. |
| `--prompt-file PATH` | Load the task prompt from a file. |
| `--scenario quick` | Select the debate scenario. |
| `--quality fast` | Select model quality tier. |
| `--roles role=model` | Override role model mapping. |
| `--output-dir PATH` | Write generated artifacts to a directory. |
| `--no-docx` | Skip optional docx generation. |
| `--task-profile-only` | Generate only TaskProfile artifacts. |
| `--health-check-only` | Run provider/model health checks only. |
| `--contract-file PATH` | Load a JSON OutputContract. |
| `--synthesis-max-output-tokens N` | Set synthesis output token budget. |
| `--smoke-test` | Run local smoke tests without provider calls. |

## Output Files

Typical full runs write:

| File | Purpose |
| --- | --- |
| `synthesis_output.md` | Final human-readable synthesis. |
| `synthesis_output_meeting_report.md` | Meeting-style process and decision report. |
| `debate_log.json` | Machine-readable run log and metadata. |
| `debate_transcript.txt` | Readable debate transcript. |
| `synthesis_output.docx` | Optional docx output when enabled. |

`--task-profile-only` also writes `task_profile.json`.

## Current Architecture

`ai_debate.py` is intentionally thin. It is the public CLI entry point and delegates to `ai2ai.cli.main()`.

Implementation lives under `ai2ai/`:

```text
ai2ai/
  cli.py
  config.py
  model_catalog.py
  core/
  providers/
  ingestion/
  profiling/
  contracts/
  debate/
  synthesis/
  renderers/
  utils/
```

Development should add new logic to the relevant package area:
- scenario and role orchestration: `ai2ai/core/`
- provider adapters: `ai2ai/providers/`
- source loading and summaries: `ai2ai/ingestion/`
- TaskProfile: `ai2ai/profiling/`
- OutputContract: `ai2ai/contracts/`
- debate phases: `ai2ai/debate/`
- final synthesis and validation: `ai2ai/synthesis/`
- artifact writing: `ai2ai/renderers/`
- shared helpers: `ai2ai/utils/`

Do not put new large logic into `ai_debate.py`.

## Completed Sprint State

Verified through Sprint 3.5:
- Sprint 0: provider health checks, fallback, failed-response exclusion, smoke path.
- Sprint 1: multi-step synthesis, validation, repair, structured metadata handling.
- Sprint 2: TaskProfile.
- Sprint 3: OutputContract and contract-aware synthesis.
- Sprint 3.5: safe modularization into the `ai2ai/` package.
- Sprint 3.6: documentation sync after modularization.

## Regression Commands

Run after behavior changes:

```powershell
python -B ai_debate.py --smoke-test
```

```powershell
python -B ai_debate.py --folder .\source --prompt-file .\feladat.txt --scenario quick --quality fast --no-docx --output-dir .\eredmenyek_regression_default --synthesis-max-output-tokens 8000
```

```powershell
python -B ai_debate.py --folder .\source --prompt-file .\feladat.txt --scenario quick --quality fast --contract-file .\contracts\business_master_plan.json --no-docx --output-dir .\eredmenyek_regression_business --synthesis-max-output-tokens 8000
```

```powershell
python -B ai_debate.py --folder .\source --prompt-file .\feladat.txt --scenario quick --quality fast --contract-file .\contracts\technical_audit.json --no-docx --output-dir .\eredmenyek_regression_technical --synthesis-max-output-tokens 8000
```

For documentation-only changes, the smoke test is usually sufficient.

## Future Sprint Prompts

Future sprint prompts live in:

```text
docs/SPRINT_PROMPTS.md
```

They currently cover:
- Sprint 4: config-based scenarios and roles,
- Sprint 5: source ingestion, repo map, secret scrubber, nested ZIP,
- Sprint 6: optional backend API preparation.
