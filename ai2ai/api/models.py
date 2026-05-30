"""Request and response schemas for the optional API."""

from __future__ import annotations

from typing import Any

try:
    from pydantic import BaseModel, Field
except Exception:  # pragma: no cover - exercised in FastAPI-free envs
    BaseModel = object  # type: ignore[assignment]

    def Field(default=None, **kwargs):  # type: ignore[no-redef]
        if "default_factory" in kwargs:
            return kwargs["default_factory"]()
        return default


class DebateSessionRequest(BaseModel):
    prompt_text: str | None = None
    prompt_file: str | None = None
    folder: str | None = None
    sources: list[str] = Field(default_factory=list)
    scenario: str = "quick"
    quality: str = "fast"
    contract_file: str | None = None
    roles: str | dict[str, str] | None = None
    no_docx: bool = True
    output_dir: str | None = None
    task_profile_only: bool = False
    synthesis_max_output_tokens: int | None = None
    max_source_files: int | None = None
    max_source_chars: int | None = None
    max_nested_archive_depth: int | None = None


class SessionResponse(BaseModel):
    session_id: str
    status: str
    output_dir: str | None = None
    artifacts: dict[str, str] = Field(default_factory=dict)
    validation: dict[str, Any] = Field(default_factory=dict)


def request_to_dict(request: DebateSessionRequest) -> dict[str, Any]:
    if hasattr(request, "model_dump"):
        return request.model_dump()
    if hasattr(request, "dict"):
        return request.dict()
    return dict(getattr(request, "__dict__", {}))
