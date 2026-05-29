"""Evidence Pack helpers."""

from __future__ import annotations

from ai2ai.ingestion.source_summary import evidence_pack_ingestion_block


def build_evidence_ingestion_context(items: list, ingestion_metadata: dict | None) -> str:
    return evidence_pack_ingestion_block(items, ingestion_metadata)


def phase_evidence(*args, **kwargs):
    """Compatibility wrapper for callers that import the phase from this module."""
    from ai2ai.cli import phase_evidence as _phase_evidence

    return _phase_evidence(*args, **kwargs)
