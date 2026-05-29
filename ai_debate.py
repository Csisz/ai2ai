"""Public CLI entry point for the AI Debate Pipeline.

The implementation lives in :mod:`ai2ai.cli`; this wrapper preserves the
historical ``python ai_debate.py ...`` command and import compatibility.
"""

from ai2ai.cli import *  # noqa: F401,F403


if __name__ == "__main__":
    main()
