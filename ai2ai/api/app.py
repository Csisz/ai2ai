"""Optional FastAPI application factory."""

from __future__ import annotations

from ai2ai.api import FASTAPI_MISSING_MESSAGE
from ai2ai.api.session_store import SessionStore

try:
    from fastapi import FastAPI
except Exception:  # pragma: no cover - depends on local optional deps
    FastAPI = None  # type: ignore[assignment]


def create_app(session_root: str | None = None):
    if FastAPI is None:
        raise RuntimeError(FASTAPI_MISSING_MESSAGE)
    from ai2ai.api.routes import register_routes

    app = FastAPI(
        title="AI2AI Debate API",
        version=_version(),
        description="Optional synchronous API wrapper around the AI Debate CLI.",
    )
    register_routes(app, SessionStore(session_root))
    return app


def run_api_server(host: str = "127.0.0.1", port: int = 8000) -> None:
    if FastAPI is None:
        raise RuntimeError(FASTAPI_MISSING_MESSAGE)
    try:
        import uvicorn
    except Exception as exc:
        raise RuntimeError("Uvicorn is not installed. Install API dependencies first.") from exc
    uvicorn.run(create_app(), host=host, port=port)


def _version() -> str:
    try:
        from ai2ai import __version__
        return __version__
    except Exception:
        return "unknown"
