# AI Debate Pipeline - Codex Master Context

## Project Goal

This project is an AI Debate Pipeline: a Python CLI tool that lets multiple AI models discuss, challenge, and refine a user-provided task, then generate final artifacts such as markdown synthesis, meeting report, `debate_log.json`, transcript, and optional docx output.

The long-term goal is to turn this into a general-purpose AI Debate Orchestrator, not only a fixed business-plan generator.

## Current Architecture

`ai_debate.py` is now only the public CLI wrapper / entry point. It should stay thin and import `ai2ai.cli.main()`.

Most implementation lives under the `ai2ai/` package. Future work should extend the package modules, not move large logic back into `ai_debate.py`.

Main phases:
1. Load source files: zip, folder, docx, pdf, markdown, code, json, yaml, images.
2. Build a TaskProfile.
3. Resolve OutputContract.
4. Generate an Evidence Pack.
5. Ask models for independent opinions.
6. Create an Issue Matrix when the scenario requires it.
7. Run targeted rebuttals and structured revisions when the scenario requires them.
8. Use the SynthesisEngine to produce contract-aware final artifacts.
9. Save artifacts: markdown synthesis, meeting report, `debate_log.json`, transcript, and optionally docx.

## Module Structure

```text
ai_debate.py              # thin CLI wrapper only

ai2ai/
  __init__.py
  cli.py                  # CLI parsing and top-level command flow
  config.py               # shared config/defaults
  model_catalog.py        # model catalog and model metadata

  core/
    orchestrator.py       # orchestration-facing helpers
    roles.py              # role mapping/fallback concepts
    health.py             # provider/model health checks
    phases.py             # phase-level coordination helpers

  providers/
    base.py
    openai_responses.py
    anthropic.py
    gemini.py
    openai_compatible.py

  ingestion/
    loader.py             # source loading
    source_summary.py     # source summaries/repo summaries

  profiling/
    task_profile.py       # TaskProfile structure and profiling

  contracts/
    output_contract.py    # OutputContract structure
    contract_loader.py    # contract file loading
    contract_validator.py # contract validation/normalization

  debate/
    evidence.py
    independent.py
    rebuttal.py
    revision.py

  synthesis/
    engine.py             # SynthesisEngine
    prompts.py
    assembly.py
    validation.py
    metadata.py
    markdown_utils.py

  renderers/
    markdown_renderer.py
    meeting_report.py
    transcript.py
    docx_renderer.py

  utils/
    json_utils.py
    text_utils.py
    logging_utils.py
    file_utils.py
```

The modular package exists to keep future changes localized. Some modules may currently re-export or host migrated logic while Sprint 3.5 settles; new work should still target the module area named above.

## Ownership Guide

- CLI compatibility and command parsing: `ai2ai/cli.py`.
- Scenario, phase, role, and orchestration work: `ai2ai/core/` and config-related modules.
- Provider calls and provider-specific adapters: `ai2ai/providers/`.
- Provider/model health checks and fallback behavior: `ai2ai/core/health.py`, `ai2ai/core/roles.py`, and provider modules.
- Source loading and source summaries: `ai2ai/ingestion/`.
- TaskProfile logic: `ai2ai/profiling/`.
- OutputContract loading, validation, and normalization: `ai2ai/contracts/`.
- Debate phase logic: `ai2ai/debate/`.
- Final synthesis, markdown assembly, metadata extraction, and validation: `ai2ai/synthesis/`.
- Artifact writing, meeting reports, transcripts, and docx output: `ai2ai/renderers/`.
- Shared parsing/text/log/file helpers: `ai2ai/utils/`.

## Current Completed State

The verified pipeline includes:
- provider health checks before debate execution,
- provider/model fallback,
- exclusion of failed model responses,
- configurable synthesis token budget,
- multi-step SynthesisEngine,
- truncation/section validation and repair,
- structured metadata extraction with safe fallback,
- TaskProfile,
- OutputContract,
- contract-aware synthesis,
- default, `business_master_plan`, and `technical_audit` runs passing with `final_validation_status=ok` and `human_artifact_status=ok`.

## Development Strategy

Do not rewrite the whole project at once.

Work in safe, small sprints.

Priority order:
1. Stabilize the current implementation.
2. Improve final synthesis quality.
3. Generalize scenarios and outputs.
4. Improve source ingestion.
5. Prepare an optional backend API structure while keeping CLI primary.

## Engineering Rules

- Keep the existing CLI compatible unless explicitly asked otherwise.
- Keep `ai_debate.py` as a thin public wrapper. Do not put new large logic there.
- Do not add a frontend unless the current sprint explicitly asks for it.
- Do not add FastAPI/backend behavior unless the current sprint explicitly asks for it.
- Avoid large rewrites.
- Prefer small, testable changes.
- Preserve existing output files:
  - `synthesis_output.md`
  - `synthesis_output_meeting_report.md`
  - `debate_log.json`
  - `debate_transcript.txt`
  - optional docx output
- Preserve `debate_log.json` keys unless the sprint explicitly changes the schema.
- New scenario/role loading should go into `ai2ai/core/` or config-related modules.
- New contract changes should go into `ai2ai/contracts/` and `ai2ai/synthesis/`.
- New source ingestion changes should go into `ai2ai/ingestion/`.
- New output writing should go into `ai2ai/renderers/`.
- Do not commit secrets or `.env`.
- If API keys, tokens, or credentials are found, mask them.
- Always explain modified files and provide test commands.

## Regression Discipline

After behavior changes, run the relevant regression set:

```powershell
python -B ai_debate.py --smoke-test

python -B ai_debate.py --regression-test
```

The regression gate runs the smoke test, task-profile-only mode, default full run, business contract run, technical contract run, artifact checks, validation-status checks, ingestion metadata checks, meeting-report source summary checks, and the technical `## Audit verdict` heading-count check.

For documentation-only changes, `python -B ai_debate.py --smoke-test` is usually enough unless the docs change command examples in a way that needs validation.

## Current Sprint Discipline

Only implement the sprint requested by the user.

If the user asks for Sprint 3.6, update documentation only. Do not start Sprint 4.

For future work:
- Sprint 4 should introduce config-based scenarios and roles in the modular package.
- Sprint 5 should improve ingestion and source hygiene in the modular package.
- Sprint 6 should prepare an optional backend/API layer without making it primary.

Do not jump ahead to scenario YAML, API/backend, or frontend unless explicitly requested.

## Definition of Done

A task is done only if:
- the code still runs from CLI,
- existing behavior is preserved where possible,
- failed model calls are handled safely,
- output degradation is explicit, not silent,
- documentation reflects the current modular architecture,
- the user receives clear test commands.
