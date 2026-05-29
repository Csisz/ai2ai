"""Shared file classification helpers for source ingestion."""

from __future__ import annotations

from pathlib import Path

TEXT_EXT = {
    ".txt", ".md", ".markdown", ".rst", ".json", ".yaml", ".yml", ".xml",
    ".html", ".htm", ".css", ".js", ".ts", ".tsx", ".jsx", ".py", ".java",
    ".cs", ".cpp", ".c", ".h", ".go", ".rs", ".php", ".rb", ".sh", ".bat",
    ".ps1", ".sql", ".toml", ".ini", ".cfg", ".env", ".gitignore", ".kt",
    ".swift", ".mjs", ".cjs", ".vue", ".svelte", ".gradle", ".properties",
    ".dockerfile", ".tf", ".tfvars", ".graphql", ".proto", ".lock",
}
WORD_EXT = {".docx", ".doc"}
EXCEL_EXT = {".xlsx", ".xls", ".xlsm", ".csv", ".tsv"}
PDF_EXT = {".pdf"}
IMG_EXT = {".png", ".jpg", ".jpeg", ".gif", ".webp", ".bmp", ".tiff"}
ZIP_EXT = {".zip", ".tar", ".gz", ".tgz"}

EXCL_DIRS = {
    "node_modules", ".git", "dist", "build", ".next", "__pycache__", "vendor",
    "venv", ".venv", ".pytest_cache", ".mypy_cache", ".ruff_cache", ".tox",
    ".idea", ".vscode", "coverage", ".coverage", ".turbo",
}
EXCL_FILES = {"package-lock.json", "yarn.lock", "bun.lockb", "pnpm-lock.yaml"}
EXCL_PATS = ["components/ui/", "components\\ui\\", ".min.js", ".min.css"]

LANG_BY_EXT = {
    ".py": "Python", ".js": "JavaScript", ".jsx": "JavaScript/React",
    ".ts": "TypeScript", ".tsx": "TypeScript/React", ".json": "JSON",
    ".md": "Markdown", ".markdown": "Markdown", ".yml": "YAML", ".yaml": "YAML",
    ".toml": "TOML", ".html": "HTML", ".css": "CSS", ".scss": "SCSS",
    ".go": "Go", ".rs": "Rust", ".java": "Java", ".cs": "C#",
    ".php": "PHP", ".rb": "Ruby", ".sh": "Shell", ".ps1": "PowerShell",
    ".sql": "SQL", ".xml": "XML", ".dockerfile": "Dockerfile",
}


def normalized_rel(path: Path | str) -> str:
    return str(path).replace("\\", "/")


def is_excluded_rel(rel_path: str | Path) -> bool:
    rel_s = normalized_rel(rel_path)
    p = Path(rel_s)
    for part in p.parts:
        if part in EXCL_DIRS or part.startswith("~"):
            return True
    if p.name in EXCL_FILES:
        return True
    return any(pattern.replace("\\", "/") in rel_s for pattern in EXCL_PATS)


def _excluded(path: Path, base: Path) -> bool:
    try:
        rel = path.relative_to(base)
    except ValueError:
        rel = path
    return is_excluded_rel(rel)


def language_for_path(path: str | Path) -> str:
    p = Path(str(path))
    name = p.name.lower()
    if name == "dockerfile" or name.startswith("dockerfile."):
        return "Dockerfile"
    return LANG_BY_EXT.get(p.suffix.lower(), p.suffix.lower().lstrip(".").upper() or "text")


def is_probably_binary(data: bytes) -> bool:
    if not data:
        return False
    sample = data[:4096]
    if b"\x00" in sample:
        return True
    text_chars = bytearray({7, 8, 9, 10, 12, 13, 27} | set(range(0x20, 0x100)))
    non_text = sample.translate(None, text_chars)
    return len(non_text) / max(1, len(sample)) > 0.30


def decode_text(data: bytes) -> str:
    for enc in ("utf-8", "utf-8-sig", "cp1250", "latin-1"):
        try:
            return data.decode(enc)
        except UnicodeDecodeError:
            continue
    return data.decode("utf-8", errors="replace")
