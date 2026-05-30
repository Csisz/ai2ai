# Optional API

Sprint 6 adds a minimal optional API layer. The command-line interface remains the primary supported interface.

## Dependencies

The API requires optional dependencies:

```powershell
pip install -r requirements-api.txt
```

If they are missing, CLI commands still work. Starting API mode prints:

```text
FastAPI is not installed. Install API dependencies first.
```

## Start Server

```powershell
python -B ai_debate.py --api --host 127.0.0.1 --port 8000
```

Do not run this for normal CLI use. It is only for local API integration.

## Endpoints

`GET /health`

Returns API status, package version, and `cli_available: true`.

`POST /sessions`

Runs a debate synchronously through the existing CLI path and records local session metadata.

Example body:

```json
{
  "folder": ".\\source",
  "prompt_file": ".\\feladat.txt",
  "scenario": "quick",
  "quality": "fast",
  "contract_file": ".\\contracts\\technical_audit.json",
  "no_docx": true,
  "synthesis_max_output_tokens": 8000
}
```

`GET /sessions/{session_id}`

Returns stored session metadata.

`GET /sessions/{session_id}/artifacts`

Returns available artifact paths such as `synthesis_output.md`, `synthesis_output_meeting_report.md`, `debate_log.json`, `debate_transcript.txt`, and `task_profile.json` when present.

`GET /sessions/{session_id}/log`

Returns the parsed validation summary from `debate_log.json` when available.

## Session Storage

Session metadata and API-owned outputs are stored under:

```text
.ai2ai_sessions/
```

This directory is ignored by git.

## Scope

The API is intentionally small and synchronous. Sprint 6 does not add authentication, queues, workers, streaming, websockets, database persistence, production deployment, or frontend UI.
