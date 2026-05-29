"""Lightweight regression runner for the public CLI.

This module intentionally exercises ``ai_debate.py`` as a subprocess so the
quality gate covers the same entry point users run manually.
"""

from __future__ import annotations

import json
import os
import shutil
import subprocess
import sys
from dataclasses import dataclass
from pathlib import Path


PROFILE_DIR = "eredmenyek_regression_profile"
DEFAULT_DIR = "eredmenyek_regression_default"
BUSINESS_DIR = "eredmenyek_regression_business"
TECHNICAL_DIR = "eredmenyek_regression_technical"

BAD_SYNTHESIS_PATTERNS = [
    "Min\u0151s\u00e9gi megjegyz\u00e9s",
    "Failed steps",
    "heuristic_fallback",
    "Ez a szakasz nem k\u00e9sz\u00fclt el teljesen",
]

FAILED_STATUS_PATTERNS = [
    '"final_validation_status": "failed"',
    '"human_artifact_status": "failed"',
]


@dataclass
class CheckResult:
    name: str
    passed: bool
    reason: str = ""
    path: str = ""
    pattern: str = ""


class RegressionRunner:
    def __init__(self, repo_root: Path | None = None):
        self.repo_root = (repo_root or Path(__file__).resolve().parents[2]).resolve()
        self.results: list[CheckResult] = []

    def run(self) -> int:
        print("AI2AI regression test")
        print(f"Repo root: {self.repo_root}")
        print("")

        self._check_preconditions()
        if self._has_failures():
            return self._finish()

        self._clean_output_dirs()
        self._run_cli_step(
            "smoke test",
            ["--smoke-test"],
        )
        self._run_cli_step(
            "task-profile-only run",
            [
                "--folder", ".\\source",
                "--prompt-file", ".\\feladat.txt",
                "--scenario", "quick",
                "--quality", "fast",
                "--task-profile-only",
                "--output-dir", f".\\{PROFILE_DIR}",
            ],
        )
        self._run_cli_step(
            "default full quick run",
            self._base_full_args(DEFAULT_DIR),
        )
        self._run_cli_step(
            "business_master_plan contract run",
            self._base_full_args(BUSINESS_DIR)
            + ["--contract-file", ".\\contracts\\business_master_plan.json"],
        )
        self._run_cli_step(
            "technical_audit contract run",
            self._base_full_args(TECHNICAL_DIR)
            + ["--contract-file", ".\\contracts\\technical_audit.json"],
        )

        self._validate_profile_outputs()
        self._validate_full_outputs(DEFAULT_DIR, "default")
        self._validate_full_outputs(BUSINESS_DIR, "business")
        self._validate_full_outputs(TECHNICAL_DIR, "technical")
        self._validate_ingestion_metadata(DEFAULT_DIR)
        self._validate_meeting_report(DEFAULT_DIR)
        self._validate_audit_verdict_count(TECHNICAL_DIR)
        return self._finish()

    def _base_full_args(self, output_dir: str) -> list[str]:
        return [
            "--folder", ".\\source",
            "--prompt-file", ".\\feladat.txt",
            "--scenario", "quick",
            "--quality", "fast",
            "--no-docx",
            "--output-dir", f".\\{output_dir}",
            "--synthesis-max-output-tokens", "8000",
        ]

    def _check_preconditions(self) -> None:
        required = [
            "ai_debate.py",
            "source",
            "feladat.txt",
            "contracts/business_master_plan.json",
            "contracts/technical_audit.json",
        ]
        for rel in required:
            self._record(
                f"precondition: {rel}",
                (self.repo_root / rel).exists(),
                path=rel,
                reason="required regression input is missing",
            )

    def _clean_output_dirs(self) -> None:
        for name in (PROFILE_DIR, DEFAULT_DIR, BUSINESS_DIR, TECHNICAL_DIR):
            target = (self.repo_root / name).resolve()
            if not self._is_under_repo(target):
                self._record(
                    f"clean output folder: {name}",
                    False,
                    path=str(target),
                    reason="refusing to remove a path outside the repository",
                )
                continue
            if target.exists():
                shutil.rmtree(target)
            self._record(f"clean output folder: {name}", True, path=str(target))

    def _run_cli_step(self, name: str, args: list[str]) -> None:
        command = [sys.executable, "-B", "ai_debate.py", *args]
        print(f"[RUN] {name}")
        env = os.environ.copy()
        env.setdefault("PYTHONIOENCODING", "utf-8")
        proc = subprocess.run(
            command,
            cwd=self.repo_root,
            text=True,
            encoding="utf-8",
            errors="replace",
            capture_output=True,
            env=env,
        )
        if proc.returncode == 0:
            self._record(name, True)
            return
        reason = f"exit code {proc.returncode}"
        detail = self._tail(proc.stdout, proc.stderr)
        if detail:
            reason = f"{reason}\n{detail}"
        self._record(name, False, reason=reason)

    def _validate_profile_outputs(self) -> None:
        base = self.repo_root / PROFILE_DIR
        self._require_file(base / "task_profile.json", "profile task_profile.json exists")
        log_path = base / "debate_log.json"
        self._require_file(log_path, "profile debate_log.json exists")
        data = self._load_json(log_path, "profile debate_log.json parses")
        if data is None:
            return
        metadata = data.get("metadata", {})
        self._record(
            "profile metadata contains task_profile",
            isinstance(metadata.get("task_profile"), dict),
            path=str(log_path),
            reason="metadata.task_profile missing",
        )

    def _validate_full_outputs(self, output_dir: str, label: str) -> None:
        base = self.repo_root / output_dir
        synthesis_path = base / "synthesis_output.md"
        report_path = base / "synthesis_output_meeting_report.md"
        log_path = base / "debate_log.json"
        transcript_path = base / "debate_transcript.txt"

        self._require_file(synthesis_path, f"{label} synthesis_output.md exists")
        self._require_file(report_path, f"{label} meeting report exists")
        self._require_file(log_path, f"{label} debate_log.json exists")
        self._require_file(transcript_path, f"{label} transcript exists")
        self._validate_bad_patterns(synthesis_path, f"{label} synthesis")
        self._validate_log_status(log_path, label)

    def _validate_bad_patterns(self, path: Path, label: str) -> None:
        text = self._read_text(path, f"{label} synthesis reads")
        if text is None:
            return
        for pattern in BAD_SYNTHESIS_PATTERNS:
            self._record(
                f"{label} has no bad pattern: {pattern}",
                pattern not in text,
                path=str(path),
                pattern=pattern,
                reason="bad output pattern found",
            )

    def _validate_log_status(self, path: Path, label: str) -> None:
        text = self._read_text(path, f"{label} debate_log.json reads")
        if text is None:
            return
        for pattern in FAILED_STATUS_PATTERNS:
            self._record(
                f"{label} debate_log has no failed status: {pattern}",
                pattern not in text,
                path=str(path),
                pattern=pattern,
                reason="failed validation status found",
            )
        data = self._load_json(path, f"{label} debate_log.json parses")
        if data is None:
            return
        synthesis = data.get("metadata", {}).get("synthesis", {})
        self._record(
            f"{label} final_validation_status is ok",
            synthesis.get("final_validation_status") == "ok",
            path=str(path),
            reason=f"actual={synthesis.get('final_validation_status')!r}",
        )
        self._record(
            f"{label} human_artifact_status is ok",
            synthesis.get("human_artifact_status") == "ok",
            path=str(path),
            reason=f"actual={synthesis.get('human_artifact_status')!r}",
        )

    def _validate_ingestion_metadata(self, output_dir: str) -> None:
        log_path = self.repo_root / output_dir / "debate_log.json"
        data = self._load_json(log_path, "default ingestion debate_log.json parses")
        if data is None:
            return
        metadata = data.get("metadata", {})
        ingestion = metadata.get("ingestion")
        self._record(
            "default debate_log contains metadata.ingestion",
            isinstance(ingestion, dict),
            path=str(log_path),
            reason="metadata.ingestion missing",
        )
        if not isinstance(ingestion, dict):
            return
        for key in ("repo_map", "redaction_summary", "important_files"):
            self._record(
                f"default ingestion contains {key}",
                key in ingestion,
                path=str(log_path),
                pattern=key,
                reason=f"metadata.ingestion.{key} missing",
            )

    def _validate_meeting_report(self, output_dir: str) -> None:
        report_path = self.repo_root / output_dir / "synthesis_output_meeting_report.md"
        text = self._read_text(report_path, "default meeting report reads")
        if text is None:
            return
        required = ["Source / Ingestion Summary", "Repo Map", "redaction_count"]
        for pattern in required:
            self._record(
                f"default meeting report contains {pattern}",
                pattern in text,
                path=str(report_path),
                pattern=pattern,
                reason="meeting report ingestion summary missing expected text",
            )

    def _validate_audit_verdict_count(self, output_dir: str) -> None:
        path = self.repo_root / output_dir / "synthesis_output.md"
        text = self._read_text(path, "technical synthesis reads for Audit verdict")
        if text is None:
            return
        count = sum(1 for line in text.splitlines() if line == "## Audit verdict")
        self._record(
            "technical output contains exactly one Audit verdict heading",
            count == 1,
            path=str(path),
            pattern="^## Audit verdict$",
            reason=f"actual_count={count}",
        )

    def _require_file(self, path: Path, name: str) -> None:
        self._record(name, path.is_file(), path=str(path), reason="file missing")

    def _read_text(self, path: Path, name: str) -> str | None:
        if not path.exists():
            self._record(name, False, path=str(path), reason="file missing")
            return None
        try:
            text = path.read_text(encoding="utf-8")
        except Exception as exc:
            self._record(name, False, path=str(path), reason=str(exc))
            return None
        self._record(name, True, path=str(path))
        return text

    def _load_json(self, path: Path, name: str) -> dict | None:
        text = self._read_text(path, name)
        if text is None:
            return None
        try:
            data = json.loads(text)
        except Exception as exc:
            self._record(name, False, path=str(path), reason=f"invalid JSON: {exc}")
            return None
        if isinstance(data, dict):
            return data
        self._record(name, False, path=str(path), reason="JSON root is not an object")
        return None

    def _record(
        self,
        name: str,
        passed: bool,
        reason: str = "",
        path: str = "",
        pattern: str = "",
    ) -> None:
        self.results.append(CheckResult(name, passed, reason, path, pattern))
        status = "PASS" if passed else "FAIL"
        print(f"[{status}] {name}")
        if not passed:
            if path:
                print(f"       file: {path}")
            if pattern:
                print(f"       pattern: {pattern}")
            if reason:
                print(f"       reason: {reason}")

    def _finish(self) -> int:
        total = len(self.results)
        failed = [r for r in self.results if not r.passed]
        print("")
        print(f"Summary: {total - len(failed)}/{total} checks passed.")
        if failed:
            print("REGRESSION TEST FAILED")
            for item in failed:
                detail = item.reason or "no detail"
                suffix = f" ({item.path})" if item.path else ""
                print(f"- {item.name}{suffix}: {detail}")
            return 1
        print("REGRESSION TEST PASSED")
        return 0

    def _has_failures(self) -> bool:
        return any(not r.passed for r in self.results)

    def _is_under_repo(self, path: Path) -> bool:
        try:
            path.relative_to(self.repo_root)
            return True
        except ValueError:
            return False

    @staticmethod
    def _tail(stdout: str, stderr: str, max_lines: int = 80) -> str:
        lines = []
        if stdout.strip():
            lines += ["--- stdout tail ---", *stdout.splitlines()[-max_lines:]]
        if stderr.strip():
            lines += ["--- stderr tail ---", *stderr.splitlines()[-max_lines:]]
        return "\n".join(lines)


def main() -> int:
    return RegressionRunner().run()


if __name__ == "__main__":
    raise SystemExit(main())
