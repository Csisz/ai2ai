# AI Debate Pipeline — Codex Master Context

## Project goal

This project is an AI Debate Pipeline: a Python CLI tool that lets multiple AI models discuss, challenge and refine a user-provided task, then generate final artifacts such as a business plan, implementation plan, AI context block, meeting report and logs.

The long-term goal is to turn this into a general-purpose AI Debate Orchestrator, not only a fixed business-plan generator.

## Current architecture

The current implementation is mainly in `ai_debate.py`.

Main phases:
1. Load source files: zip, folder, docx, pdf, markdown, code, json, yaml, images.
2. Generate an Evidence Pack.
3. Ask multiple models for independent opinions.
4. Create an Issue Matrix.
5. Run targeted rebuttals.
6. Run structured revisions.
7. Use a Final Judge to synthesize the final result.
8. Save artifacts: markdown synthesis, meeting report, debate_log.json, transcript and optionally docx.

## Current problem

The final answer generation is unreliable.

Known issues:
- If the judge model is unavailable, the system may fall back to weak output.
- Failed model responses may be included in the debate as if they were valid opinions.
- Final synthesis currently depends too much on one large judge output.
- The system is too hardcoded toward business plan / implementation plan outputs.
- The project should become more general-purpose through TaskProfile, OutputContract and configurable scenarios.

## Development strategy

Do not rewrite the whole project at once.

Work in safe, small sprints.

Priority order:
1. Stabilize the current implementation.
2. Improve final synthesis quality.
3. Generalize scenarios and outputs.
4. Improve source ingestion.
5. Prepare backend API structure for a future frontend.

## Engineering rules

- Keep the existing CLI compatible unless explicitly asked otherwise.
- Do not introduce a frontend yet.
- Do not migrate to FastAPI yet unless the current sprint asks for it.
- Avoid large rewrites.
- Prefer small, testable changes.
- Preserve existing output files:
  - synthesis markdown
  - meeting report
  - debate_log.json
  - transcript
  - optional docx
- Do not commit secrets or `.env`.
- If API keys, tokens or credentials are found, mask them.
- Always explain modified files and provide test commands.

## Current sprint discipline

Only implement the sprint requested by the user.

If the user asks for Sprint 0, only do stabilization.
If the user asks for Sprint 1, only improve final synthesis.
Do not jump ahead to TaskProfile, OutputContract or FastAPI unless explicitly requested.

## Definition of done

A task is done only if:
- the code still runs from CLI,
- existing behavior is preserved where possible,
- failed model calls are handled safely,
- output degradation is explicit, not silent,
- the user receives clear test commands.