"""Source and ingestion summaries for Evidence Pack and reports."""

from __future__ import annotations

from ai2ai.utils.text_utils import safe_excerpt


def _join(values: list[str], limit: int = 12) -> str:
    values = [str(v) for v in (values or []) if v]
    if not values:
        return "none"
    shown = values[:limit]
    suffix = f" (+{len(values) - limit} more)" if len(values) > limit else ""
    return ", ".join(shown) + suffix


def format_ingestion_summary_markdown(metadata: dict) -> str:
    repo = metadata.get("repo_map", {}) if isinstance(metadata, dict) else {}
    important = metadata.get("important_files", []) if isinstance(metadata, dict) else []
    redactions = metadata.get("redaction_summary", {}).get("total_redactions", 0)
    lines = [
        "## Source / Ingestion Summary",
        "",
        f"- source_count: {metadata.get('source_count', 0)}",
        f"- text_file_count: {metadata.get('text_file_count', 0)}",
        f"- image_file_count: {metadata.get('image_file_count', 0)}",
        f"- skipped_file_count: {metadata.get('skipped_file_count', 0)}",
        f"- truncated_file_count: {len(metadata.get('truncated_files', []))}",
        f"- redaction_count: {redactions}",
        f"- nested_archive_count: {len(metadata.get('nested_archives_processed', []))}",
        f"- detected_languages: {_join(metadata.get('detected_languages', []))}",
        f"- detected_frameworks: {_join(metadata.get('detected_frameworks', []))}",
        f"- repo_map: {repo.get('root_summary', 'not available')}",
    ]
    if important:
        lines.append("- important_files:")
        for item in important[:12]:
            lines.append(f"  - {item.get('path')} ({item.get('reason', 'important')})")
    warnings = metadata.get("ingestion_warnings") or []
    if warnings:
        lines.append("- ingestion_warnings:")
        for warning in warnings[:8]:
            lines.append(f"  - {warning}")
    return "\n".join(lines)


def evidence_pack_ingestion_block(items: list, metadata: dict | None) -> str:
    metadata = metadata or {}
    repo = metadata.get("repo_map", {})
    important = metadata.get("important_files", [])
    excerpts = []
    for item in sorted(
        [i for i in items if not i.is_img()],
        key=lambda i: -(getattr(i, "metadata", {}) or {}).get("importance_score", 0),
    )[:12]:
        meta = getattr(item, "metadata", {}) or {}
        excerpts.append(
            f"- {item.name} | {meta.get('language', 'text')} | "
            f"lines {meta.get('start_line', '?')}-{meta.get('end_line', '?')}: "
            f"{safe_excerpt(item.text, 280)}"
        )
    lines = [
        "INGESTION_CONTEXT:",
        f"- Source count: {metadata.get('source_count', 0)}",
        f"- Text files: {metadata.get('text_file_count', 0)}; images: {metadata.get('image_file_count', 0)}",
        f"- Skipped files: {metadata.get('skipped_file_count', 0)}; truncated files: {len(metadata.get('truncated_files', []))}",
        f"- Secret redactions: {metadata.get('redaction_summary', {}).get('total_redactions', 0)}",
        f"- Nested archives processed: {len(metadata.get('nested_archives_processed', []))}",
        f"- Detected languages: {_join(metadata.get('detected_languages', []))}",
        f"- Detected frameworks/tools: {_join(metadata.get('detected_frameworks', []))}",
        "",
        "REPO_MAP_SUMMARY:",
        repo.get("root_summary", "not available"),
        f"- Package/config files: {_join(repo.get('package_config_files', []), 18)}",
        f"- Likely entrypoints: {_join(repo.get('likely_entrypoints', []), 18)}",
        f"- Docs/readme files: {_join(repo.get('docs_readme_files', []), 18)}",
        f"- Test files: {_join(repo.get('test_files', []), 18)}",
        "",
        "IMPORTANT_FILES:",
    ]
    if important:
        lines.extend(f"- {i.get('path')} ({i.get('reason')})" for i in important[:20])
    else:
        lines.append("- none detected")
    lines += ["", "CONTENT_EXCERPTS:"]
    lines.extend(excerpts or ["- none"])
    return "\n".join(lines)


def estimate_cost(*args, **kwargs):
    """Compatibility wrapper for the legacy public import path."""
    from ai2ai.cli import estimate_cost as _estimate_cost

    return _estimate_cost(*args, **kwargs)
