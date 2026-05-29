"""Shared text helpers used by ingestion and compatibility imports."""

from __future__ import annotations


def trunc_middle(text: str, max_chars: int, marker: str | None = None) -> str:
    """Truncate text from the middle while preserving both ends."""
    text = text or ""
    if max_chars <= 0 or len(text) <= max_chars:
        return text
    marker = marker or f"\n[...{len(text) - max_chars} omitted...]\n"
    if len(marker) >= max_chars:
        return text[:max_chars]
    head = (max_chars - len(marker)) // 2
    tail = max_chars - len(marker) - head
    return text[:head] + marker + text[-tail:]


def safe_excerpt(text: str, max_chars: int = 240) -> str:
    text = " ".join((text or "").split())
    if len(text) <= max_chars:
        return text
    return text[: max_chars - 3].rstrip() + "..."


def _trunc(text: str, n: int) -> str:
    """Backward-compatible Hungarian truncation helper."""
    return trunc_middle(text, n, marker=f"\n[...{max(0, len(text or '') - n)} kihagyva...]\n")


def _is_failed_response(text: str) -> bool:
    if not text or not str(text).strip():
        return True
    t = str(text).strip().lower()
    if t.startswith("[error"):
        return True
    failure_bits = (
        " hiba:", "hiba:", "api kulcs", "missing api key",
        "not available", "unavailable", "nem el", "nem el",
        "invalid api key", "authentication", "permission denied",
        "rate limit", "quota", "timeout", "timed out",
    )
    return t.startswith("[") and any(bit in t for bit in failure_bits)


def _valid_response(text: str) -> bool:
    return not _is_failed_response(text)
