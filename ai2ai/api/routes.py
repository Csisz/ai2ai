"""Route registration for the optional FastAPI app."""

from __future__ import annotations

from pathlib import Path

from ai2ai import __version__
from ai2ai.api.models import DebateSessionRequest, request_to_dict
from ai2ai.api.session_store import SessionStore
from ai2ai.core.orchestrator import (
    DebateRunRequest,
    discover_artifacts,
    read_validation_summary,
    run_debate_via_cli,
)


def register_routes(app, store: SessionStore) -> None:
    from fastapi import HTTPException

    @app.get("/health")
    def health():
        return {
            "status": "ok",
            "version": __version__,
            "cli_available": True,
        }

    @app.post("/sessions")
    def create_session(request: DebateSessionRequest):
        # TODO: Replace synchronous execution with a background job runner in a later sprint.
        payload = request_to_dict(request)
        _validate_request(payload)
        session = store.create(payload, payload.get("output_dir"))
        store.update(session["session_id"], status="running")
        run_request = DebateRunRequest(**payload)
        run_request.output_dir = session["output_dir"]
        result = run_debate_via_cli(run_request)
        error = result.stderr_tail or result.stdout_tail or "CLI run failed" if result.returncode != 0 else None
        updated = store.update(
            session["session_id"],
            status=result.status,
            output_dir=result.output_dir,
            artifacts=result.artifacts,
            validation=result.validation,
            final_validation_status=result.validation.get("final_validation_status"),
            human_artifact_status=result.validation.get("human_artifact_status"),
            returncode=result.returncode,
            stdout_tail=result.stdout_tail,
            stderr_tail=result.stderr_tail,
            command=result.command,
            error=error,
        )
        response = updated or session
        return response

    @app.get("/sessions/{session_id}")
    def get_session(session_id: str):
        session = store.get(session_id)
        if not session:
            raise HTTPException(status_code=404, detail="session not found")
        return session

    @app.get("/sessions/{session_id}/artifacts")
    def get_artifacts(session_id: str):
        session = store.get(session_id)
        if not session:
            raise HTTPException(status_code=404, detail="session not found")
        output_dir = session.get("output_dir")
        artifacts = discover_artifacts(output_dir) if output_dir else {}
        if artifacts != session.get("artifacts"):
            store.update(session_id, artifacts=artifacts)
        return {
            "session_id": session_id,
            "status": session.get("status"),
            "output_dir": output_dir,
            "artifacts": artifacts,
        }

    @app.get("/sessions/{session_id}/log")
    def get_log(session_id: str):
        session = store.get(session_id)
        if not session:
            raise HTTPException(status_code=404, detail="session not found")
        output_dir = session.get("output_dir")
        log_path = Path(output_dir or "") / "debate_log.json"
        if not log_path.exists():
            return {
                "session_id": session_id,
                "status": session.get("status"),
                "log_available": False,
                "validation": session.get("validation", {}),
            }
        validation = read_validation_summary(output_dir)
        return {
            "session_id": session_id,
            "status": session.get("status"),
            "log_available": True,
            "log_path": str(log_path.resolve()),
            "validation": validation,
            "final_validation_status": validation.get("final_validation_status"),
            "human_artifact_status": validation.get("human_artifact_status"),
        }


def _validate_request(payload: dict) -> None:
    from fastapi import HTTPException

    if not payload.get("folder") and not payload.get("sources") and not payload.get("task_profile_only"):
        raise HTTPException(status_code=400, detail="folder or sources are required unless task_profile_only=true")
    if not payload.get("prompt_text") and not payload.get("prompt_file"):
        raise HTTPException(status_code=400, detail="prompt_text or prompt_file is required")
