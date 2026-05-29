"""Configuration compatibility exports.

Sprint 3.5 establishes package boundaries without changing runtime behavior.
Constants are still sourced from ``ai2ai.cli`` during this safe split.
"""

from ai2ai.cli import (
    DEFAULT_OUTPUT_CONTRACT,
    MIN_VALID_PARTICIPANTS,
    RUN_METADATA,
    SYNTHESIS_MAX_OUTPUT_TOKENS,
)
