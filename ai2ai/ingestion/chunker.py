"""Line-aware text chunking for source ingestion."""

from __future__ import annotations

from ai2ai.utils.file_utils import language_for_path


def _line_span(lines: list[str], max_chars: int, from_end: bool = False) -> tuple[list[str], int, int]:
    selected: list[str] = []
    total = 0
    iterable = range(len(lines) - 1, -1, -1) if from_end else range(len(lines))
    for idx in iterable:
        line = lines[idx]
        if total + len(line) + 1 > max_chars and selected:
            break
        selected.append(line)
        total += len(line) + 1
    if from_end:
        selected.reverse()
        start = max(1, len(lines) - len(selected) + 1)
        end = len(lines)
    else:
        start = 1
        end = len(selected)
    return selected, start, end


def chunk_text_for_ingestion(path: str, text: str, max_chars: int) -> tuple[str, dict]:
    """Return text that preserves structure and metadata under a per-file limit."""
    text = text or ""
    language = language_for_path(path)
    lines = text.splitlines()
    meta = {
        "path": path,
        "language": language,
        "line_count": len(lines),
        "start_line": 1 if lines else 0,
        "end_line": len(lines),
        "truncated": False,
        "chars_before": len(text),
        "chars_after": len(text),
        "truncation_reason": "",
    }
    if max_chars <= 0 or len(text) <= max_chars:
        return text, meta

    header_budget = min(max_chars // 3, 1800)
    tail_budget = max(800, max_chars - header_budget - 360)
    first, start_a, end_a = _line_span(lines, header_budget, from_end=False)
    last, start_b, end_b = _line_span(lines, tail_budget, from_end=True)
    omitted = max(0, start_b - end_a - 1)
    summary = (
        f"[File: {path} | language: {language} | lines: 1-{len(lines)} | truncated]\n"
        f"[Summary: large file; kept header lines {start_a}-{end_a} and tail lines "
        f"{start_b}-{end_b}; omitted about {omitted} middle lines.]\n\n"
    )
    chunked = (
        summary
        + "\n".join(first)
        + f"\n\n[... omitted {omitted} middle lines ...]\n\n"
        + "\n".join(last)
    )
    meta.update(
        {
            "end_line": end_b,
            "truncated": True,
            "chars_after": len(chunked),
            "truncation_reason": "per_file_char_limit",
            "kept_ranges": [[start_a, end_a], [start_b, end_b]],
        }
    )
    return chunked, meta
