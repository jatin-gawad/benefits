"""Dynamic carrier proposal comparison engine."""

from pathlib import Path

PACKAGE_ROOT = Path(__file__).resolve().parent
PROJECT_ROOT = PACKAGE_ROOT.parent
DEFAULT_CONFIG_DIR = PROJECT_ROOT / "config" / "carriers"
DEFAULT_BENEFIT_LIBRARY_PATH = PROJECT_ROOT / "config" / "benefits" / "standard_benefits.json"
DEFAULT_TEMPLATE_PATH = PROJECT_ROOT / "config" / "templates" / "Total Cost ComparisonMktSvy.xlsx"
DEFAULT_OUTPUT_DIR = PROJECT_ROOT / "output"
DEFAULT_DATABASE_PATH = PROJECT_ROOT / "database" / "benefits_intelligence.db"

SUPPORTED_EXTENSIONS = {".pdf", ".xlsx", ".xlsm", ".xls", ".csv"}

__all__ = [
    "PACKAGE_ROOT",
    "PROJECT_ROOT",
    "DEFAULT_CONFIG_DIR",
    "DEFAULT_BENEFIT_LIBRARY_PATH",
    "DEFAULT_TEMPLATE_PATH",
    "DEFAULT_OUTPUT_DIR",
    "DEFAULT_DATABASE_PATH",
    "SUPPORTED_EXTENSIONS"
]
