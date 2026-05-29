"""Repository map generation for folders and archives."""

from __future__ import annotations

import json
import re
from collections import Counter
from pathlib import Path

from ai2ai.utils.file_utils import language_for_path, normalized_rel

PACKAGE_FILES = {
    "package.json", "pyproject.toml", "requirements.txt", "requirements-dev.txt",
    "poetry.lock", "Pipfile", "Dockerfile", "docker-compose.yml", "docker-compose.yaml",
    "Cargo.toml", "go.mod", "pom.xml", "build.gradle", "Makefile", "tsconfig.json",
    "vite.config.ts", "vite.config.js", "next.config.js", "next.config.mjs",
}
DOC_NAMES = {"readme.md", "readme", "agents.md", "contributing.md", "changelog.md", "license"}
ENTRY_NAMES = {
    "main.py", "app.py", "manage.py", "server.py", "index.js", "index.ts",
    "main.ts", "main.tsx", "app.tsx", "cli.py", "ai_debate.py",
}
IMPORTANT_HINTS = (
    "src/", "app/", "pages/", "routes/", "api/", "config/", "configs/",
    "core/", "orchestrator", "scoring", "pipeline", ".github/workflows/",
)


def importance_score(rel_path: str) -> tuple[int, str]:
    rel = normalized_rel(rel_path)
    lower = rel.lower()
    name = Path(lower).name
    score = 10
    reasons: list[str] = []
    if name in DOC_NAMES:
        score += 90
        reasons.append("documentation")
    if lower.startswith("docs/"):
        score += 70
        reasons.append("docs")
    if name in PACKAGE_FILES:
        score += 85
        reasons.append("package_or_config")
    if lower.startswith(".github/workflows/"):
        score += 75
        reasons.append("ci_workflow")
    if name in ENTRY_NAMES or re.search(r"(^|/)(main|app|server|cli)\.(py|js|ts|tsx)$", lower):
        score += 80
        reasons.append("entrypoint")
    if lower.startswith(("src/", "app/")):
        score += 45
        reasons.append("source")
    if any(hint in lower for hint in ("route", "api", "controller", "endpoint")):
        score += 45
        reasons.append("api_or_route")
    if any(hint in lower for hint in ("orchestrator", "scoring", "pipeline", "core")):
        score += 45
        reasons.append("core_logic")
    if lower.startswith(("test/", "tests/")) or name.startswith("test_") or name.endswith(".test.ts"):
        score += 40
        reasons.append("tests")
    if "example" in lower or lower.endswith(".env.example"):
        score += 35
        reasons.append("config_example")
    return score, ", ".join(reasons) if reasons else "ordinary"


def _detect_frameworks_from_text(rel_path: str, text: str) -> set[str]:
    frameworks: set[str] = set()
    lower_path = rel_path.lower()
    lower_text = (text or "").lower()
    if lower_path.endswith("package.json"):
        try:
            data = json.loads(text)
            deps = {}
            for key in ("dependencies", "devDependencies", "peerDependencies"):
                deps.update(data.get(key) or {})
            names = set(deps)
            mapping = {
                "react": "React", "next": "Next.js", "vite": "Vite",
                "express": "Express", "nestjs": "NestJS", "vue": "Vue",
                "svelte": "Svelte", "tailwindcss": "Tailwind CSS",
                "typescript": "TypeScript", "playwright": "Playwright",
                "vitest": "Vitest", "jest": "Jest",
            }
            for dep, label in mapping.items():
                if dep in names:
                    frameworks.add(label)
        except Exception:
            pass
    text_markers = {
        "fastapi": "FastAPI", "django": "Django", "flask": "Flask",
        "pytest": "pytest", "sqlalchemy": "SQLAlchemy", "pydantic": "Pydantic",
        "openai": "OpenAI SDK", "anthropic": "Anthropic SDK",
        "google-genai": "Google GenAI SDK", "docker": "Docker",
    }
    if lower_path.endswith(("pyproject.toml", "requirements.txt", "requirements-dev.txt")):
        for marker, label in text_markers.items():
            if marker in lower_text:
                frameworks.add(label)
    if "dockerfile" in lower_path or lower_path.endswith(("docker-compose.yml", "docker-compose.yaml")):
        frameworks.add("Docker")
    if lower_path.startswith(".github/workflows/"):
        frameworks.add("GitHub Actions")
    return frameworks


def build_repo_map(records: list[dict], skipped_files: list[dict] | None = None,
                   truncated_files: list[dict] | None = None,
                   limits: dict | None = None) -> dict:
    skipped_files = skipped_files or []
    truncated_files = truncated_files or []
    limits = limits or {}
    language_counts: Counter[str] = Counter()
    frameworks: set[str] = set()
    package_files: list[str] = []
    docs: list[str] = []
    tests: list[str] = []
    entrypoints: list[str] = []
    important: list[dict] = []
    tree_entries: list[str] = []

    for rec in records:
        rel = normalized_rel(rec.get("display_path") or rec.get("path") or "")
        if not rel:
            continue
        lang = rec.get("language") or language_for_path(rel)
        if rec.get("kind") == "text":
            language_counts[lang] += 1
        name = Path(rel.lower()).name
        if name in PACKAGE_FILES:
            package_files.append(rel)
        if name in DOC_NAMES or rel.lower().startswith("docs/"):
            docs.append(rel)
        if rel.lower().startswith(("test/", "tests/")) or name.startswith("test_") or ".test." in name:
            tests.append(rel)
        if name in ENTRY_NAMES:
            entrypoints.append(rel)
        frameworks.update(_detect_frameworks_from_text(rel, rec.get("text_for_detection", "")))
        score, reason = importance_score(rel)
        if score >= 50:
            important.append({"path": rel, "score": score, "reason": reason, "language": lang})
        tree_entries.append(rel)

    important.sort(key=lambda x: (-x["score"], x["path"]))
    tree_entries = sorted(tree_entries)[:200]
    file_count = len(records) + len(skipped_files)
    root_summary = (
        f"{file_count} files scanned; {len(records)} readable source/image files; "
        f"{len(skipped_files)} skipped; {len(truncated_files)} truncated."
    )
    if package_files:
        root_summary += f" Package/config files detected: {', '.join(package_files[:5])}."

    return {
        "root_summary": root_summary,
        "file_tree_summary": tree_entries,
        "detected_languages": [k for k, _ in language_counts.most_common()],
        "detected_language_counts": dict(language_counts.most_common()),
        "detected_frameworks": sorted(frameworks),
        "package_config_files": sorted(package_files),
        "likely_entrypoints": sorted(entrypoints),
        "docs_readme_files": sorted(docs),
        "test_files": sorted(tests)[:60],
        "important_source_files": important[:80],
        "ignored_skipped_files_summary": {
            "count": len(skipped_files),
            "examples": skipped_files[:30],
        },
        "size_limits_and_truncation": {
            **limits,
            "truncated_file_count": len(truncated_files),
            "truncated_files": truncated_files[:50],
        },
    }
