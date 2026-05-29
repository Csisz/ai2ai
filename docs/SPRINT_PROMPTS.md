# Sprint Prompts

Use these prompts as starting points for future Codex tasks. They assume Sprint 3.5 modularization is complete and `ai_debate.py` remains a thin CLI wrapper.

## Sprint 4 - Config-Based Scenarios And Roles

Read `AGENTS.md` first. Start Sprint 4 only.

Goal:
Move scenario and role definitions toward configuration files without changing CLI behavior.

Important:
- Do not put scenario logic back into `ai_debate.py`.
- `ai_debate.py` must remain a thin wrapper around `ai2ai.cli.main()`.
- Keep the existing CLI options compatible.
- Preserve current default, business contract, and technical contract behavior.
- Do not add FastAPI or frontend.
- Do not change synthesis behavior unless needed for scenario integration.

Tasks:
1. Add scenario config loading under `ai2ai/core/scenarios.py` or `ai2ai/config/scenario_loader.py`.
2. Add config files under:
   - `config/scenarios/`
   - `config/roles/`
3. Keep built-in scenario defaults as a deterministic fallback.
4. Support at least the existing scenarios:
   - `quick`
   - `expert-council`
   - `red-team`
5. Preserve the existing role mappings for `fast`, `balanced`, and `best`.
6. Validate loaded scenario config with clear errors:
   - missing role
   - unknown model key
   - missing judge
   - insufficient debate participants
7. Keep provider fallback and health checks working through the existing modular role/health path.
8. Store the resolved scenario config summary in `debate_log.json`.
9. Include a short scenario config summary in the meeting report.
10. Add/update smoke tests for:
    - built-in fallback config
    - valid scenario config load
    - invalid scenario config error
    - role override compatibility
    - quick/default regression
    - business and technical contract regressions

Regression commands:

```powershell
python -B ai_debate.py --smoke-test

python -B ai_debate.py --health-check-only --scenario quick --quality fast

python -B ai_debate.py --folder .\source --prompt-file .\feladat.txt --scenario quick --quality fast --no-docx --output-dir .\eredmenyek_sprint4_default_test --synthesis-max-output-tokens 8000

python -B ai_debate.py --folder .\source --prompt-file .\feladat.txt --scenario quick --quality fast --contract-file .\contracts\business_master_plan.json --no-docx --output-dir .\eredmenyek_sprint4_business_test --synthesis-max-output-tokens 8000

python -B ai_debate.py --folder .\source --prompt-file .\feladat.txt --scenario quick --quality fast --contract-file .\contracts\technical_audit.json --no-docx --output-dir .\eredmenyek_sprint4_technical_test --synthesis-max-output-tokens 8000
```

Expected result:
Scenario and role behavior can be configured without editing Python logic, while the old CLI commands still work.

After changes, summarize:
- modified files,
- config file format,
- fallback behavior,
- exact test results.

## Sprint 5 - Source Ingestion, Repo Map, Secret Scrubber, Nested ZIP

Read `AGENTS.md` first. Start Sprint 5 only.

Goal:
Improve source ingestion quality and source hygiene while preserving the existing orchestration and CLI workflow.

Important:
- Target `ai2ai/ingestion/` and `ai2ai/utils/`.
- Integrate through the existing Evidence Pack/orchestration path.
- Do not add frontend or API changes.
- Do not change final synthesis behavior unless needed to pass richer source summaries.
- Do not put ingestion logic into `ai_debate.py`.
- Keep `ai_debate.py` thin.

Tasks:
1. Add or improve repo map generation in `ai2ai/ingestion/source_summary.py`.
2. Add nested ZIP handling in `ai2ai/ingestion/loader.py`.
3. Add secret scrubbing utilities in `ai2ai/utils/`.
4. Ensure detected secrets are masked before:
   - Evidence Pack prompts
   - debate logs
   - transcript
   - synthesis input
5. Add source statistics to `debate_log.json`:
   - file counts by type
   - skipped files
   - nested archives detected
   - scrubbed secret counts by category
   - repo map summary
6. Include a short source ingestion summary in the meeting report.
7. Keep source material private; do not add local source artifacts to docs or tests.
8. Add/update smoke tests for:
   - nested ZIP discovery
   - secret masking
   - binary skip behavior
   - repo map summary
   - Evidence Pack receives scrubbed content
9. Preserve current default/business/technical full-run behavior.

Regression commands:

```powershell
python -B ai_debate.py --smoke-test

python -B ai_debate.py --folder .\source --prompt-file .\feladat.txt --scenario quick --quality fast --no-docx --output-dir .\eredmenyek_sprint5_default_test --synthesis-max-output-tokens 8000

python -B ai_debate.py --folder .\source --prompt-file .\feladat.txt --scenario quick --quality fast --contract-file .\contracts\business_master_plan.json --no-docx --output-dir .\eredmenyek_sprint5_business_test --synthesis-max-output-tokens 8000

python -B ai_debate.py --folder .\source --prompt-file .\feladat.txt --scenario quick --quality fast --contract-file .\contracts\technical_audit.json --no-docx --output-dir .\eredmenyek_sprint5_technical_test --synthesis-max-output-tokens 8000
```

Expected result:
The pipeline builds a safer and more useful source context without changing the user-facing CLI workflow.

After changes, summarize:
- modified files,
- secret categories scrubbed,
- repo map structure,
- exact test results.

## Sprint 6 - Optional Backend API Preparation

Read `AGENTS.md` first. Start Sprint 6 only.

Goal:
Prepare an optional backend/API layer around the existing orchestrator without replacing the CLI.

Important:
- CLI remains primary and must keep working.
- Do not add a frontend.
- Add API/server code only if the sprint explicitly asks for it.
- Keep orchestration reusable through `ai2ai/core/orchestrator.py`.
- Keep artifact writing in `ai2ai/renderers/`.
- Do not move business logic into API route handlers.
- Do not break OutputContract, TaskProfile, provider fallback, or synthesis validation.

Tasks:
1. Identify the smallest orchestration boundary needed for API use in `ai2ai/core/orchestrator.py`.
2. Ensure CLI can call the same orchestration boundary as before.
3. If adding an optional API layer, place it under `api/` or `server/`.
4. Keep API request/response models thin and separate from core domain logic.
5. Ensure renderers still own output writing:
   - markdown
   - meeting report
   - transcript
   - docx
6. Add a no-network local smoke path for the orchestration boundary.
7. Add tests for:
   - CLI still works
   - orchestrator can be imported without provider calls
   - renderers can be called from orchestrator output
   - API layer, if present, does not own synthesis logic
8. Preserve current regression behavior for:
   - default run
   - business contract run
   - technical contract run

Regression commands:

```powershell
python -B -c "import ai_debate; print('ai_debate import ok')"
python -B -c "import ai2ai.cli; print('ai2ai.cli import ok')"
python -B -c "import ai2ai.core.orchestrator; print('orchestrator import ok')"

python -B ai_debate.py --smoke-test

python -B ai_debate.py --folder .\source --prompt-file .\feladat.txt --scenario quick --quality fast --no-docx --output-dir .\eredmenyek_sprint6_default_test --synthesis-max-output-tokens 8000

python -B ai_debate.py --folder .\source --prompt-file .\feladat.txt --scenario quick --quality fast --contract-file .\contracts\business_master_plan.json --no-docx --output-dir .\eredmenyek_sprint6_business_test --synthesis-max-output-tokens 8000

python -B ai_debate.py --folder .\source --prompt-file .\feladat.txt --scenario quick --quality fast --contract-file .\contracts\technical_audit.json --no-docx --output-dir .\eredmenyek_sprint6_technical_test --synthesis-max-output-tokens 8000
```

Expected result:
The project has a clean API-ready orchestration boundary, while the CLI remains the verified primary interface.

After changes, summarize:
- modified files,
- API/orchestrator boundary,
- what remains CLI-owned,
- exact test results.
