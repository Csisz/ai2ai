"""Simple file-based session metadata store for the optional API."""

from __future__ import annotations

import json
import uuid
from datetime import datetime, timezone
from pathlib import Path
from typing import Any


class SessionStore:
    def __init__(self, root: str | Path | None = None):
        self.root = Path(root or ".ai2ai_sessions").resolve()
        self.sessions_dir = self.root / "sessions"
        self.outputs_dir = self.root / "outputs"
        self.sessions_dir.mkdir(parents=True, exist_ok=True)
        self.outputs_dir.mkdir(parents=True, exist_ok=True)

    def create(self, request: dict[str, Any], output_dir: str | None = None) -> dict[str, Any]:
        session_id = uuid.uuid4().hex
        now = _now()
        resolved_output = str(Path(output_dir).resolve()) if output_dir else str((self.outputs_dir / session_id).resolve())
        data = {
            "session_id": session_id,
            "request": request,
            "status": "pending",
            "created_at": now,
            "updated_at": now,
            "scenario": request.get("scenario", "quick"),
            "quality": request.get("quality", "fast"),
            "contract": request.get("contract_file"),
            "output_dir": resolved_output,
            "artifacts": {},
            "validation": {},
            "final_validation_status": None,
            "human_artifact_status": None,
            "error": None,
        }
        self.save(data)
        return data

    def save(self, session: dict[str, Any]) -> dict[str, Any]:
        session["updated_at"] = _now()
        path = self._path(session["session_id"])
        path.write_text(json.dumps(session, ensure_ascii=False, indent=2), encoding="utf-8")
        return session

    def get(self, session_id: str) -> dict[str, Any] | None:
        path = self._path(session_id)
        if not path.exists():
            return None
        try:
            return json.loads(path.read_text(encoding="utf-8"))
        except Exception:
            return None

    def update(self, session_id: str, **updates: Any) -> dict[str, Any] | None:
        session = self.get(session_id)
        if not session:
            return None
        session.update(updates)
        return self.save(session)

    def _path(self, session_id: str) -> Path:
        clean = "".join(ch for ch in session_id if ch.isalnum() or ch in ("-", "_"))
        return self.sessions_dir / f"{clean}.json"


def _now() -> str:
    return datetime.now(timezone.utc).isoformat()
