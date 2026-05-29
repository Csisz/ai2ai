"""Scenario and role configuration loading.

This module keeps Sprint 4 scenario configuration outside the thin
``ai_debate.py`` wrapper and away from synthesis/provider logic.
"""

from __future__ import annotations

import json
from copy import deepcopy
from pathlib import Path
from typing import Any


class ScenarioConfigError(ValueError):
    """Raised when a scenario or role config cannot be used safely."""


QUALITY_ALIASES = {
    "deep": "best",
}

REQUIRED_SCENARIO_FIELDS = {
    "scenario_id",
    "description",
    "phases",
    "roles",
    "default_role_models",
    "judge_role",
    "moderator_role",
    "debate_participant_roles",
    "minimum_required_debate_participants",
    "synthesis_mode",
}

REQUIRED_ROLE_FIELDS = {
    "role_id",
    "title",
    "role_type",
    "description",
    "default_model_key",
    "fallback_model_keys",
    "prompt_focus",
    "responsibilities",
}


def project_root() -> Path:
    return Path(__file__).resolve().parents[2]


def default_config_root() -> Path:
    return project_root() / "config"


def normalize_quality(quality: str) -> str:
    return QUALITY_ALIASES.get((quality or "").strip(), quality)


def _read_json(path: Path) -> dict[str, Any]:
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except json.JSONDecodeError as exc:
        raise ScenarioConfigError(f"Invalid JSON in {path}: {exc}") from exc
    except OSError as exc:
        raise ScenarioConfigError(f"Cannot read {path}: {exc}") from exc


def load_role_config(config_root: Path | None = None) -> dict[str, dict[str, Any]]:
    root = config_root or default_config_root()
    path = root / "roles" / "roles.json"
    if not path.exists():
        return {}
    data = _read_json(path)
    roles = data.get("roles")
    if not isinstance(roles, list):
        raise ScenarioConfigError(f"Invalid role config {path}: expected roles list.")
    resolved = {}
    for role in roles:
        if not isinstance(role, dict):
            raise ScenarioConfigError(f"Invalid role config {path}: role entries must be objects.")
        missing = sorted(REQUIRED_ROLE_FIELDS - set(role))
        if missing:
            rid = role.get("role_id", "<unknown>")
            raise ScenarioConfigError(f"Invalid role config {path}: role {rid} missing {', '.join(missing)}.")
        rid = str(role["role_id"]).strip()
        if not rid:
            raise ScenarioConfigError(f"Invalid role config {path}: empty role_id.")
        if role.get("role_type") not in {"infrastructure", "debate_participant"}:
            raise ScenarioConfigError(f"Invalid role config {path}: role {rid} has invalid role_type.")
        resolved[rid] = dict(role)
    return resolved


def available_scenario_ids(builtin_scenarios: dict[str, Any] | None = None,
                           config_root: Path | None = None) -> list[str]:
    root = config_root or default_config_root()
    ids = set((builtin_scenarios or {}).keys())
    scenario_dir = root / "scenarios"
    if scenario_dir.exists():
        ids.update(p.stem for p in scenario_dir.glob("*.json"))
    return sorted(ids)


def load_scenario_config(scenario_id: str, config_root: Path | None = None) -> dict[str, Any] | None:
    root = config_root or default_config_root()
    path = root / "scenarios" / f"{scenario_id}.json"
    if not path.exists():
        return None
    data = _read_json(path)
    _validate_scenario_shape(data, path)
    if data["scenario_id"] != scenario_id:
        raise ScenarioConfigError(
            f"Invalid scenario config {path}: scenario_id must be {scenario_id!r}."
        )
    data["_config_path"] = str(path)
    return data


def _validate_scenario_shape(data: dict[str, Any], path: Path) -> None:
    missing = sorted(REQUIRED_SCENARIO_FIELDS - set(data))
    if missing:
        raise ScenarioConfigError(f"Invalid scenario config {path}: missing {', '.join(missing)}.")
    for key in ("phases", "roles", "debate_participant_roles"):
        if not isinstance(data.get(key), list) or not data.get(key):
            raise ScenarioConfigError(f"Invalid scenario config {path}: {key} must be a non-empty list.")
    if not isinstance(data.get("default_role_models"), dict) or not data["default_role_models"]:
        raise ScenarioConfigError(f"Invalid scenario config {path}: default_role_models must be an object.")
    if data.get("judge_role") not in data["roles"]:
        raise ScenarioConfigError(f"Invalid scenario config {path}: judge_role is not in roles.")
    if data.get("moderator_role") not in data["roles"]:
        raise ScenarioConfigError(f"Invalid scenario config {path}: moderator_role is not in roles.")
    missing_participants = [r for r in data["debate_participant_roles"] if r not in data["roles"]]
    if missing_participants:
        raise ScenarioConfigError(
            f"Invalid scenario config {path}: participant roles not in roles: {', '.join(missing_participants)}."
        )
    if int(data.get("minimum_required_debate_participants", 0)) < 1:
        raise ScenarioConfigError(
            f"Invalid scenario config {path}: minimum_required_debate_participants must be at least 1."
        )


def _builtin_to_config(scenario_id: str, builtin: dict[str, Any]) -> dict[str, Any]:
    quality_map = deepcopy(builtin.get("quality_map", {}))
    roles = sorted({role for mapping in quality_map.values() for role in mapping})
    participants = [r for r in roles if r not in {"moderator", "judge"}]
    return {
        "scenario_id": scenario_id,
        "description": builtin.get("description") or builtin.get("desc") or scenario_id,
        "phases": list(builtin.get("phases", [])),
        "roles": roles,
        "default_role_models": quality_map,
        "judge_role": "judge",
        "moderator_role": "moderator",
        "debate_participant_roles": participants,
        "minimum_required_debate_participants": 2 if scenario_id == "quick" else 2,
        "synthesis_mode": "default",
        "notes": "Resolved from built-in Python fallback.",
        "_config_path": "",
    }


def resolve_scenario(
    scenario_id: str,
    quality: str,
    role_overrides: dict[str, str] | None,
    catalog_keys: set[str],
    builtin_scenarios: dict[str, Any] | None = None,
    config_root: Path | None = None,
) -> dict[str, Any]:
    quality = normalize_quality(quality)
    root = config_root or default_config_root()
    scenario = load_scenario_config(scenario_id, root)
    source = "config"
    fallback_used = False
    if scenario is None:
        builtin = (builtin_scenarios or {}).get(scenario_id)
        if not builtin:
            raise ScenarioConfigError(f"Unknown scenario {scenario_id!r}; no config or built-in fallback found.")
        scenario = _builtin_to_config(scenario_id, builtin)
        source = "builtin_fallback"
        fallback_used = True

    roles_config = load_role_config(root)
    default_models = scenario["default_role_models"]
    if quality not in default_models:
        raise ScenarioConfigError(
            f"Scenario {scenario_id!r} has no default_role_models for quality {quality!r}."
        )
    base_mapping = dict(default_models[quality])
    overrides = dict(role_overrides or {})
    unknown_override_roles = [r for r in overrides if r not in scenario["roles"]]
    if unknown_override_roles:
        raise ScenarioConfigError(
            f"Scenario {scenario_id!r} does not define role(s): {', '.join(sorted(unknown_override_roles))}."
        )
    resolved_mapping = dict(base_mapping)
    resolved_mapping.update(overrides)

    missing_roles = [r for r in scenario["roles"] if r not in resolved_mapping]
    if missing_roles:
        raise ScenarioConfigError(
            f"Scenario {scenario_id!r} missing model mapping for role(s): {', '.join(missing_roles)}."
        )
    unknown_models = [
        f"{role}={model_key}"
        for role, model_key in resolved_mapping.items()
        if model_key not in catalog_keys
    ]
    if unknown_models:
        raise ScenarioConfigError(
            f"Scenario {scenario_id!r} references unknown model key(s): {', '.join(unknown_models)}."
        )

    role_definitions = {}
    for role in scenario["roles"]:
        base = roles_config.get(role, {
            "role_id": role,
            "title": role.replace("_", " ").title(),
            "role_type": "infrastructure" if role in {scenario["judge_role"], scenario["moderator_role"]} else "debate_participant",
            "description": "",
            "default_model_key": base_mapping.get(role, ""),
            "fallback_model_keys": [],
            "prompt_focus": "",
            "responsibilities": [],
        })
        merged = dict(base)
        merged["active_model_key"] = resolved_mapping.get(role)
        role_definitions[role] = merged

    return {
        "scenario_id": scenario["scenario_id"],
        "description": scenario["description"],
        "phases": list(scenario["phases"]),
        "roles": list(scenario["roles"]),
        "default_role_models": deepcopy(scenario["default_role_models"]),
        "active_role_models": resolved_mapping,
        "base_role_models": base_mapping,
        "role_overrides": overrides,
        "role_definitions": role_definitions,
        "judge_role": scenario["judge_role"],
        "moderator_role": scenario["moderator_role"],
        "debate_participant_roles": list(scenario["debate_participant_roles"]),
        "minimum_required_debate_participants": int(scenario["minimum_required_debate_participants"]),
        "synthesis_mode": scenario.get("synthesis_mode", "default"),
        "default_contract_id": scenario.get("default_contract_id"),
        "notes": scenario.get("notes", ""),
        "quality": quality,
        "config_source": source,
        "config_path": scenario.get("_config_path", ""),
        "fallback_used": fallback_used,
    }


def scenario_to_legacy(resolved: dict[str, Any]) -> dict[str, Any]:
    return {
        "desc": resolved["description"],
        "description": resolved["description"],
        "phases": list(resolved["phases"]),
        "quality_map": deepcopy(resolved["default_role_models"]),
        "minimum_required_debate_participants": resolved["minimum_required_debate_participants"],
        "judge_role": resolved["judge_role"],
        "moderator_role": resolved["moderator_role"],
        "debate_participant_roles": list(resolved["debate_participant_roles"]),
        "synthesis_mode": resolved.get("synthesis_mode", "default"),
        "default_contract_id": resolved.get("default_contract_id"),
    }


def scenario_summary(resolved: dict[str, Any]) -> dict[str, Any]:
    roles = resolved.get("role_definitions", {})
    return {
        "scenario_id": resolved.get("scenario_id"),
        "description": resolved.get("description"),
        "phases": resolved.get("phases", []),
        "roles": resolved.get("roles", []),
        "active_role_models": resolved.get("active_role_models", {}),
        "judge_role": resolved.get("judge_role"),
        "moderator_role": resolved.get("moderator_role"),
        "debate_participant_roles": resolved.get("debate_participant_roles", []),
        "minimum_required_debate_participants": resolved.get("minimum_required_debate_participants"),
        "synthesis_mode": resolved.get("synthesis_mode"),
        "default_contract_id": resolved.get("default_contract_id"),
        "config_source": resolved.get("config_source"),
        "fallback_used": resolved.get("fallback_used", False),
        "role_fallbacks": {
            role_id: role.get("fallback_model_keys", [])
            for role_id, role in roles.items()
        },
    }


def format_scenario_summary(resolved: dict[str, Any], full: bool = False) -> str:
    lines = [
        f"Scenario: {resolved.get('scenario_id')}",
        f"Description: {resolved.get('description')}",
        f"Quality: {resolved.get('quality')}",
        f"Source: {resolved.get('config_source')}",
        f"Phases: {' -> '.join(resolved.get('phases', []))}",
        f"Judge role: {resolved.get('judge_role')}",
        f"Moderator role: {resolved.get('moderator_role')}",
        "Participant roles: " + ", ".join(resolved.get("debate_participant_roles", [])),
        f"Minimum participants: {resolved.get('minimum_required_debate_participants')}",
        "Active model mapping:",
    ]
    for role, model_key in resolved.get("active_role_models", {}).items():
        lines.append(f"  - {role}: {model_key}")
    if full:
        lines.append("Fallback order:")
        for role, info in resolved.get("role_definitions", {}).items():
            fallbacks = info.get("fallback_model_keys") or []
            lines.append(f"  - {role}: {', '.join(fallbacks) if fallbacks else '(provider default)'}")
        lines.append("Roles:")
        for role, info in resolved.get("role_definitions", {}).items():
            lines.append(f"  - {role}: {info.get('title')} [{info.get('role_type')}]")
            if info.get("prompt_focus"):
                lines.append(f"    focus: {info.get('prompt_focus')}")
    return "\n".join(lines)
