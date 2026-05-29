"""Secret masking for source text before it reaches model prompts."""

from __future__ import annotations

import re
from collections import Counter


class RedactionStats:
    def __init__(self):
        self.counts: Counter[str] = Counter()

    def add(self, kind: str, count: int):
        if count:
            self.counts[kind] += count

    def to_dict(self) -> dict:
        total = sum(self.counts.values())
        return {"total_redactions": total, "by_type": dict(sorted(self.counts.items()))}


PRIVATE_KEY_RE = re.compile(
    r"-----BEGIN [A-Z0-9 ]*PRIVATE KEY-----.*?-----END [A-Z0-9 ]*PRIVATE KEY-----",
    re.DOTALL,
)
BEARER_RE = re.compile(r"(?i)\bBearer\s+[A-Za-z0-9._~+/=-]{16,}")
JWT_RE = re.compile(r"\beyJ[A-Za-z0-9_-]{8,}\.[A-Za-z0-9_-]{8,}\.[A-Za-z0-9_-]{8,}\b")
TOKEN_PATTERNS = [
    ("openai_api_key", re.compile(r"\bsk-(?:proj-)?[A-Za-z0-9_-]{20,}\b")),
    ("anthropic_api_key", re.compile(r"\bsk-ant-[A-Za-z0-9_-]{20,}\b")),
    ("github_token", re.compile(r"\b(?:ghp|gho|ghu|ghs|ghr)_[A-Za-z0-9_]{20,}\b")),
    ("github_token", re.compile(r"\bgithub_pat_[A-Za-z0-9_]{20,}\b")),
    ("gemini_api_key", re.compile(r"\bAIza[0-9A-Za-z_-]{20,}\b")),
    ("xai_api_key", re.compile(r"\bxai-[A-Za-z0-9_-]{20,}\b")),
    ("generic_api_token", re.compile(r"\b(?:api[_-]?key|token|secret)_[A-Za-z0-9_-]{24,}\b", re.I)),
]
DB_URL_RE = re.compile(
    r"\b(?:postgres(?:ql)?|mysql|mongodb(?:\+srv)?|redis|mssql|sqlserver)://[^\s'\"<>]+",
    re.I,
)
ASSIGNMENT_RE = re.compile(
    r"(?im)^(\s*[A-Z0-9_.-]*(?:API[_-]?KEY|TOKEN|SECRET|PASSWORD|PASS|PWD|DATABASE_URL|DB_URL|CONNECTION_STRING|WEBHOOK)[A-Z0-9_.-]*\s*[:=]\s*)([\"']?)([^\n#]+?)(\2)(\s*(?:#.*)?)$"
)
JSON_KEY_RE = re.compile(
    r"(?i)([\"']?(?:api[_-]?key|token|secret|password|database_url|db_url|connection_string|webhook_secret|authorization)[\"']?\s*:\s*)([\"'])(.*?)(\2)"
)


def _sub_count(pattern: re.Pattern, repl, text: str) -> tuple[str, int]:
    return pattern.subn(repl, text)


def scrub_text(text: str, stats: RedactionStats | None = None) -> tuple[str, dict]:
    """Return text with likely secret values replaced by ``[REDACTED]``."""
    stats = stats or RedactionStats()
    scrubbed = text or ""

    scrubbed, count = _sub_count(PRIVATE_KEY_RE, "[REDACTED_PRIVATE_KEY]", scrubbed)
    stats.add("private_key", count)

    def repl_assignment(match: re.Match) -> str:
        return f"{match.group(1)}[REDACTED]{match.group(5)}"

    scrubbed, count = _sub_count(ASSIGNMENT_RE, repl_assignment, scrubbed)
    stats.add("env_or_config_assignment", count)

    def repl_json(match: re.Match) -> str:
        return f"{match.group(1)}{match.group(2)}[REDACTED]{match.group(2)}"

    scrubbed, count = _sub_count(JSON_KEY_RE, repl_json, scrubbed)
    stats.add("json_or_yaml_secret", count)

    scrubbed, count = _sub_count(BEARER_RE, "Bearer [REDACTED]", scrubbed)
    stats.add("bearer_token", count)

    scrubbed, count = _sub_count(JWT_RE, "[REDACTED_JWT]", scrubbed)
    stats.add("jwt", count)

    scrubbed, count = _sub_count(DB_URL_RE, lambda m: m.group(0).split("://", 1)[0] + "://[REDACTED]", scrubbed)
    stats.add("database_url", count)

    for kind, pattern in TOKEN_PATTERNS:
        scrubbed, count = _sub_count(pattern, "[REDACTED]", scrubbed)
        stats.add(kind, count)

    return scrubbed, stats.to_dict()
