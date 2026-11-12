from __future__ import annotations

import copy
import json
from pathlib import Path
from typing import Any

from proposal_engine import DEFAULT_CONFIG_DIR

REQUIRED_CONFIG_KEYS = [
    "carrier_name",
    "carrier_key",
    "aliases",
    "file_name_keywords",
    "section_keywords",
    "field_mappings",
    "benefit_mappings",
    "tier_mappings",
    "confidence_rules",
]


class CarrierConfigError(ValueError):
    """Raised when a carrier config is missing required structure."""


def load_carrier_config(carrier_key: str | None, config_dir: str | Path | None = None) -> dict[str, Any]:
    """Load a carrier config, falling back to default_generic when needed."""

    base_dir = Path(config_dir) if config_dir else DEFAULT_CONFIG_DIR
    requested_key = (carrier_key or "default_generic").strip().lower() or "default_generic"
    config_path = base_dir / f"{requested_key}.json"

    if not config_path.exists():
        config_path = base_dir / "default_generic.json"

    with config_path.open("r", encoding="utf-8") as file:
        config = json.load(file)

    default_config = config
    default_path = base_dir / "default_generic.json"
    if config_path.name != "default_generic.json" and default_path.exists():
        with default_path.open("r", encoding="utf-8") as file:
            default_config = json.load(file)
        config = _deep_merge(default_config, config)

    validate_carrier_config(config)
    config["_config_path"] = str(config_path)
    return config


def validate_carrier_config(config: dict[str, Any]) -> None:
    missing = [key for key in REQUIRED_CONFIG_KEYS if key not in config]
    if missing:
        raise CarrierConfigError(f"Carrier config is missing required key(s): {', '.join(missing)}")

    if not isinstance(config.get("section_keywords"), dict):
        raise CarrierConfigError("Carrier config section_keywords must be an object.")
    if not isinstance(config.get("field_mappings"), dict):
        raise CarrierConfigError("Carrier config field_mappings must be an object.")
    if not isinstance(config.get("benefit_mappings"), dict):
        raise CarrierConfigError("Carrier config benefit_mappings must be an object.")
    if not isinstance(config.get("tier_mappings"), dict):
        raise CarrierConfigError("Carrier config tier_mappings must be an object.")


def load_all_carrier_configs(config_dir: str | Path | None = None) -> list[dict[str, Any]]:
    base_dir = Path(config_dir) if config_dir else DEFAULT_CONFIG_DIR
    configs: list[dict[str, Any]] = []
    for path in sorted(base_dir.glob("*.json")):
        with path.open("r", encoding="utf-8") as file:
            config = json.load(file)
        validate_carrier_config(config)
        config["_config_path"] = str(path)
        configs.append(config)
    return configs


def _deep_merge(base: dict[str, Any], override: dict[str, Any]) -> dict[str, Any]:
    result = copy.deepcopy(base)
    for key, value in override.items():
        if isinstance(value, dict) and isinstance(result.get(key), dict):
            result[key] = _deep_merge(result[key], value)
        else:
            result[key] = copy.deepcopy(value)
    return result
