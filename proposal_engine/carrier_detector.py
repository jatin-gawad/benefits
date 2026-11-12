from __future__ import annotations

import re
from pathlib import Path
from typing import Any

from proposal_engine import DEFAULT_CONFIG_DIR
from proposal_engine.carrier_config_loader import load_all_carrier_configs


def detect_carrier(
    document: dict[str, Any],
    file_path: str | Path,
    config_dir: str | Path | None = None
) -> dict[str, Any]:
    """Score every carrier config and return the best carrier key plus confidence."""

    base_dir = Path(config_dir) if config_dir else DEFAULT_CONFIG_DIR
    configs = load_all_carrier_configs(base_dir)
    file_name = Path(file_path).name.lower()
    all_text = _normalize(document.get("all_text", ""))
    first_pages_text = _normalize(_first_pages_text(document))
    proposal_title = _normalize(_proposal_title(document))

    best: dict[str, Any] | None = None
    scored: list[dict[str, Any]] = []
    for config in configs:
        key = config.get("carrier_key", "default_generic")
        if key == "default_generic":
            continue

        names = _candidate_terms(config)
        score = 0
        details: list[str] = []

        for term in config.get("file_name_keywords", []):
            if _term_present(term, file_name):
                score += 12
                details.append(f"file name matched {term}")

        for term in names:
            if _term_present(term, proposal_title):
                score += 10
                details.append(f"title matched {term}")
            if _term_present(term, first_pages_text):
                score += 8
                details.append(f"first pages matched {term}")
            if _term_present(term, all_text):
                score += 5
                details.append(f"document text matched {term}")

        scored_item = {
            "carrier_key": key,
            "carrier_name": config.get("carrier_name", key),
            "score": score,
            "details": details,
            "config_path": config.get("_config_path")
        }
        scored.append(scored_item)
        if best is None or score > best["score"]:
            best = scored_item

    if not best or best["score"] < 5:
        return {
            "carrier_key": "default_generic",
            "carrier_name": "Default Generic",
            "confidence": "LOW",
            "score": best["score"] if best else 0,
            "details": ["No carrier config passed the detection threshold."],
            "scored_carriers": scored
        }

    confidence = "HIGH" if best["score"] >= 20 else "MEDIUM" if best["score"] >= 10 else "LOW"
    return {
        "carrier_key": best["carrier_key"],
        "carrier_name": best["carrier_name"],
        "confidence": confidence,
        "score": best["score"],
        "details": best["details"],
        "scored_carriers": scored
    }


def _candidate_terms(config: dict[str, Any]) -> list[str]:
    terms = [config.get("carrier_name", ""), *config.get("aliases", [])]
    return [term for term in terms if term and term.strip()]


def _first_pages_text(document: dict[str, Any]) -> str:
    pages = document.get("pages", [])
    return "\n".join(page.get("text", "") for page in pages[:3])


def _proposal_title(document: dict[str, Any]) -> str:
    lines = [line.strip() for line in document.get("all_text", "").splitlines() if line.strip()]
    return "\n".join(lines[:15])


def _normalize(value: str) -> str:
    return re.sub(r"\s+", " ", value or "").lower()


def _term_present(term: str, text: str) -> bool:
    normalized = _normalize(term)
    if not normalized:
        return False
    return normalized in text
