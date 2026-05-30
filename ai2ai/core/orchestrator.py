"""Reusable orchestration wrappers for CLI and optional API callers."""

from __future__ import annotations

import json
import os
import subprocess
import sys
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any


ARTIFACT_NAMES = [
    "synthesis_output.md",
    "synthesis_output_meeting_report.md",
    "debate_log.json",
    "debate_transcript.txt",
    "task_profile.json",
]


@dataclass
class DebateRunRequest:
    prompt_text: str | None = None
    prompt_file: str | None = None
    folder: str | None = None
    sources: list[str] = field(default_factory=list)
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


@dataclass
class DebateRunResult:
    status: str
    returncode: int
    output_dir: str
    artifacts: dict[str, str]
    validation: dict[str, Any]
    stdout_tail: str = ""
    stderr_tail: str = ""
    command: list[str] = field(default_factory=list)


def run_debate_via_cli(
    request: DebateRunRequest,
    repo_root: Path | None = None,
) -> DebateRunResult:
    """Run the public CLI synchronously and summarize generated artifacts.

    The optional API intentionally uses this wrapper instead of duplicating
    debate logic. That keeps the CLI path as the primary execution surface.
    """

    root = (repo_root or Path(__file__).resolve().parents[2]).resolve()
    output_dir = Path(request.output_dir or root / ".ai2ai_sessions" / "outputs" / "manual").resolve()
    command = _build_cli_command(request, output_dir)
    env = os.environ.copy()
    env.setdefault("PYTHONIOENCODING", "utf-8")
    proc = subprocess.run(
        command,
        cwd=root,
        text=True,
        encoding="utf-8",
        errors="replace",
        capture_output=True,
        env=env,
    )
    artifacts = discover_artifacts(output_dir)
    validation = read_validation_summary(output_dir)
    return DebateRunResult(
        status="completed" if proc.returncode == 0 else "failed",
        returncode=proc.returncode,
        output_dir=str(output_dir),
        artifacts=artifacts,
        validation=validation,
        stdout_tail=_tail(proc.stdout),
        stderr_tail=_tail(proc.stderr),
        command=command,
    )


def discover_artifacts(output_dir: str | Path) -> dict[str, str]:
    base = Path(output_dir)
    found = {}
    for name in ARTIFACT_NAMES:
        path = base / name
        if path.exists():
            found[name] = str(path.resolve())
    return found


def read_validation_summary(output_dir: str | Path) -> dict[str, Any]:
    log_path = Path(output_dir) / "debate_log.json"
    if not log_path.exists():
        return {}
    try:
        data = json.loads(log_path.read_text(encoding="utf-8"))
    except Exception as exc:
        return {"log_error": str(exc)}
    synthesis = data.get("metadata", {}).get("synthesis", {})
    output_contract = data.get("metadata", {}).get("output_contract", {})
    return {
        "final_validation_status": synthesis.get("final_validation_status"),
        "human_artifact_status": synthesis.get("human_artifact_status"),
        "structured_metadata_status": synthesis.get("structured_metadata_status"),
        "failed_steps": synthesis.get("failed_steps", []),
        "truncation_warnings": synthesis.get("truncation_warnings", []),
        "contract_id": output_contract.get("contract_id"),
    }


def _build_cli_command(request: DebateRunRequest, output_dir: Path) -> list[str]:
    command = [sys.executable, "-B", "ai_debate.py"]
    for source in request.sources or []:
        if source:
            command.append(str(source))
    if request.folder:
        command += ["--folder", str(request.folder)]
    if request.prompt_file:
        command += ["--prompt-file", str(request.prompt_file)]
    elif request.prompt_text:
        command += ["--prompt", request.prompt_text]
    command += ["--scenario", request.scenario or "quick"]
    command += ["--quality", request.quality or "fast"]
    command += ["--output-dir", str(output_dir)]
    if request.contract_file:
        command += ["--contract-file", str(request.contract_file)]
    roles = _format_roles(request.roles)
    if roles:
        command += ["--roles", roles]
    if request.no_docx:
        command.append("--no-docx")
    if request.task_profile_only:
        command.append("--task-profile-only")
    if request.synthesis_max_output_tokens:
        command += ["--synthesis-max-output-tokens", str(request.synthesis_max_output_tokens)]
    if request.max_source_files:
        command += ["--max-source-files", str(request.max_source_files)]
    if request.max_source_chars:
        command += ["--max-source-chars", str(request.max_source_chars)]
    if request.max_nested_archive_depth is not None:
        command += ["--max-nested-archive-depth", str(request.max_nested_archive_depth)]
    return command


def _format_roles(roles: str | dict[str, str] | None) -> str:
    if not roles:
        return ""
    if isinstance(roles, str):
        return roles
    return ",".join(f"{key}={value}" for key, value in roles.items())


def _tail(text: str, max_lines: int = 80) -> str:
    return "\n".join((text or "").splitlines()[-max_lines:])


# Compatibility exports for older imports.
from ai2ai.cli import main, run_expert_council, run_quick  # noqa: E402,F401
