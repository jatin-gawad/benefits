from __future__ import annotations

import sqlite3
import tempfile
import unittest
from contextlib import closing
from pathlib import Path

from proposal_engine.focused_comparison_mapper import build_focused_comparison_envelope
from proposal_engine.sqlite_loader import load_to_sqlite


class FocusedComparisonPersistenceTests(unittest.TestCase):
    def test_mapper_splits_carriers_into_existing_proposal_contract(self) -> None:
        with tempfile.TemporaryDirectory() as folder:
            source = Path(folder) / "comparison.pdf"
            source.write_bytes(b"same source package")
            envelope = build_focused_comparison_envelope(
                _focused_model(source),
                client_id="42",
                client_name="Canonical Client Ltd.",
            )

            self.assertEqual("42", envelope["client_id"])
            self.assertEqual(2, len(envelope["proposals"]))
            by_carrier = {
                proposal["proposal_summary"]["carrier_name"]: proposal
                for proposal in envelope["proposals"]
            }
            self.assertEqual({"Manulife", "RBC"}, set(by_carrier))
            self.assertEqual("80%", by_carrier["Manulife"]["plan_design"][0]["provision_value"])
            self.assertEqual("90%", by_carrier["RBC"]["plan_design"][0]["provision_value"])
            self.assertEqual([], by_carrier["Manulife"]["booklet_validation"])
            self.assertEqual(
                by_carrier["Manulife"]["proposal_summary"]["source_file_hash"],
                by_carrier["RBC"]["proposal_summary"]["source_file_hash"],
            )

    def test_shared_desktop_database_save_is_idempotent_and_does_not_alter_clients(self) -> None:
        with tempfile.TemporaryDirectory() as folder:
            root = Path(folder)
            source = root / "comparison.pdf"
            source.write_bytes(b"same source package")
            database_path = root / "Benefits_Intelligence.db"
            _create_core_desktop_clients(database_path)
            envelope = build_focused_comparison_envelope(
                _focused_model(source),
                client_id="42",
                client_name="Canonical Client Ltd.",
            )

            load_to_sqlite(envelope, database_path, client_id="42")
            first = _proposal_counts(database_path)
            load_to_sqlite(envelope, database_path, client_id="42")
            second = _proposal_counts(database_path)

            self.assertEqual(first, second)
            self.assertEqual(2, second["proposal_files"])
            self.assertEqual(2, second["proposal_plan_design"])
            self.assertEqual(2, second["proposal_rates"])
            self.assertEqual(2, second["proposal_premiums"])
            self.assertEqual(0, second["proposal_booklet_validation"])

            with closing(sqlite3.connect(database_path)) as connection:
                client_columns = {
                    row[1]
                    for row in connection.execute("PRAGMA table_info(Clients)").fetchall()
                }
                self.assertEqual(
                    {"client_id", "client_name", "status", "date_created", "date_updated"},
                    client_columns,
                )
                self.assertEqual(
                    (42, "Canonical Client Ltd."),
                    connection.execute("SELECT client_id, client_name FROM Clients").fetchone(),
                )
                self.assertEqual(
                    [("Manulife",), ("RBC",)],
                    connection.execute("SELECT carrier_name FROM carriers ORDER BY carrier_name").fetchall(),
                )
                self.assertEqual(
                    [(1,), (1,)],
                    connection.execute("SELECT reprocess_count FROM proposal_files ORDER BY id").fetchall(),
                )
                self.assertEqual(
                    [("42",)],
                    connection.execute("SELECT DISTINCT client_id FROM proposal_plan_design ORDER BY proposal_id").fetchall(),
                )
                self.assertEqual(
                    [(source.name,), (source.name,)],
                    connection.execute("SELECT source_file_name FROM proposal_plan_design ORDER BY proposal_id").fetchall(),
                )

    def test_booklet_validation_is_only_saved_when_validation_evidence_exists(self) -> None:
        with tempfile.TemporaryDirectory() as folder:
            root = Path(folder)
            database_path = root / "Benefits_Intelligence.db"
            _create_core_desktop_clients(database_path)
            proposal = {
                "proposal_summary": {
                    "client_name": "Canonical Client Ltd.",
                    "carrier_name": "Manulife",
                    "detected_carrier": "manulife",
                    "source_file_name": "proposal.pdf",
                    "source_file_hash": "validation-source",
                },
                "plan_design": [
                    {
                        "benefit_line": "Extended Health Care",
                        "provision_name": "Coinsurance",
                        "provision_value": "80%",
                        "booklet_value": "",
                        "variance": "",
                        "validation_status": "",
                        "status": "VALID",
                    }
                ],
            }
            envelope = {"client_id": "42", "proposals": [proposal]}

            load_to_sqlite(envelope, database_path, client_id="42")
            self.assertEqual(0, _proposal_counts(database_path)["proposal_booklet_validation"])

            proposal["plan_design"][0]["booklet_value"] = "80%"
            proposal["plan_design"][0]["validation_status"] = "match"
            load_to_sqlite(envelope, database_path, client_id="42")
            self.assertEqual(1, _proposal_counts(database_path)["proposal_booklet_validation"])

    def test_core_client_name_is_resolved_to_the_desktop_numeric_id(self) -> None:
        with tempfile.TemporaryDirectory() as folder:
            root = Path(folder)
            source = root / "comparison.pdf"
            source.write_bytes(b"same source package")
            database_path = root / "Benefits_Intelligence.db"
            _create_core_desktop_clients(database_path)
            envelope = build_focused_comparison_envelope(
                _focused_model(source),
                client_id="Canonical Client Ltd.",
                client_name="Canonical Client Ltd.",
            )

            load_to_sqlite(envelope, database_path, client_id=envelope["client_id"])

            with closing(sqlite3.connect(database_path)) as connection:
                self.assertEqual(
                    [("42",)],
                    connection.execute("SELECT DISTINCT client_id FROM proposal_plan_design").fetchall(),
                )


def _focused_model(source: Path) -> dict:
    return {
        "client_name": "Detected Client Name",
        "effective_date": "2026-08-01",
        "proposal_date": "2026-07-18",
        "source_files": [str(source)],
        "carrier_names": ["Manulife", "RBC"],
        "benefit_rows": [
            {
                "carrier_name": "Manulife",
                "carrier_index": 1,
                "benefit_line": "Extended Health Care",
                "class_name": "All Classes",
                "lives": 10,
                "volume": 100000,
                "rate": 1.25,
                "monthly_premium": 1250,
                "annual_premium": 15000,
                "source_file": str(source),
                "source_type": "PDF",
                "source_location": "Page 3",
                "confidence": "HIGH",
                "status": "VALID",
            },
            {
                "carrier_name": "RBC",
                "carrier_index": 2,
                "benefit_line": "Extended Health Care",
                "class_name": "All Classes",
                "lives": 10,
                "volume": 100000,
                "rate": 1.35,
                "monthly_premium": 1350,
                "annual_premium": 16200,
                "source_file": str(source),
                "source_type": "PDF",
                "source_location": "Page 8",
                "confidence": "HIGH",
                "status": "VALID",
            },
        ],
        "plan_design": [
            {
                "carrier_name": "Manulife",
                "carrier_index": 1,
                "benefit_line": "Extended Health Care",
                "provision_name": "Coinsurance",
                "provision_value": "80%",
                "source_file": str(source),
                "source_type": "PDF",
                "source_location": "Page 4",
                "confidence": "HIGH",
                "status": "VALID",
            },
            {
                "carrier_name": "RBC",
                "carrier_index": 2,
                "benefit_line": "Extended Health Care",
                "provision_name": "Coinsurance",
                "provision_value": "90%",
                "source_file": str(source),
                "source_type": "PDF",
                "source_location": "Page 9",
                "confidence": "HIGH",
                "status": "VALID",
            },
        ],
        "included_services": [],
        "rate_guarantee_by_carrier": {
            "Manulife": {"rate_guarantee": "24 months"},
            "RBC": {"rate_guarantee": "12 months"},
        },
        "audit": [],
    }


def _create_core_desktop_clients(database_path: Path) -> None:
    with closing(sqlite3.connect(database_path)) as connection:
        connection.execute(
            """
            CREATE TABLE Clients (
                client_id INTEGER PRIMARY KEY AUTOINCREMENT,
                client_name TEXT NOT NULL,
                status TEXT NOT NULL,
                date_created TEXT NOT NULL,
                date_updated TEXT NOT NULL
            )
            """
        )
        connection.execute(
            """
            INSERT INTO Clients (client_id, client_name, status, date_created, date_updated)
            VALUES (42, 'Canonical Client Ltd.', 'Active', '2026-01-01', '2026-01-01')
            """
        )
        connection.commit()


def _proposal_counts(database_path: Path) -> dict[str, int]:
    table_names = [
        "proposal_files",
        "proposal_plan_design",
        "proposal_rates",
        "proposal_premiums",
        "proposal_booklet_validation",
    ]
    with closing(sqlite3.connect(database_path)) as connection:
        return {
            table_name: connection.execute(f"SELECT COUNT(*) FROM {table_name}").fetchone()[0]
            for table_name in table_names
        }


if __name__ == "__main__":
    unittest.main()
