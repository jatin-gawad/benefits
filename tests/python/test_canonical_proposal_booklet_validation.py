from __future__ import annotations

import sqlite3
import tempfile
import unittest
from contextlib import closing
from pathlib import Path

from proposal_engine.booklet_validator import (
    validate_focused_proposals_against_booklet,
    validate_focused_proposals_against_booklet_facts,
)
from proposal_engine.sqlite_loader import load_to_sqlite


class CanonicalProposalBookletValidationTests(unittest.TestCase):
    def test_structured_match_and_difference_are_explicit(self) -> None:
        envelope = _envelope(("Manulife", "80 percent"), ("RBC", "90%"))

        result = validate_focused_proposals_against_booklet_facts(
            envelope,
            [_booklet_fact("80%")],
            booklet_path="current-booklet.pdf",
        )

        manulife, rbc = result["proposals"]
        self.assertEqual("MATCH", manulife["booklet_validation"][0]["validation_status"])
        self.assertEqual("DIFFERENT", rbc["booklet_validation"][0]["validation_status"])
        self.assertEqual("current-booklet.pdf", manulife["booklet_validation"][0]["booklet_source_file"])
        self.assertEqual(12, manulife["booklet_validation"][0]["booklet_source_page"])
        self.assertEqual("COMPLETED WITH REVIEW", result["_booklet_validation_summary"]["status"])
        self.assertEqual(2, result["_booklet_validation_summary"]["validation_rows"])

    def test_multiple_booklet_class_values_are_review_only(self) -> None:
        envelope = _envelope(("Manulife", "80%"))
        facts = [
            _booklet_fact("80%", class_name="Class A"),
            _booklet_fact("90%", class_name="Class B"),
        ]

        result = validate_focused_proposals_against_booklet_facts(envelope, facts)

        proposal = result["proposals"][0]
        self.assertEqual([], proposal["booklet_validation"])
        self.assertEqual(1, len(proposal["exceptions"]))
        self.assertIn("ambiguous", proposal["exceptions"][0]["message"].lower())
        self.assertEqual("NEEDS REVIEW", result["_booklet_validation_summary"]["status"])

    def test_basic_and_major_dental_rates_match_only_their_labeled_booklet_facts(self) -> None:
        envelope = {
            "client_id": "42",
            "client_name": "Canonical Client Ltd.",
            "proposals": [
                {
                    "proposal_summary": {
                        "client_name": "Canonical Client Ltd.",
                        "carrier_name": "Manulife",
                        "source_file": "manulife-proposal.pdf",
                    },
                    "plan_design": [
                        {
                            "benefit_line": "Dental Care",
                            "provision_name": "Basic Services",
                            "provision_value": "80%",
                            "confidence_score": "HIGH",
                        },
                        {
                            "benefit_line": "Dental Care",
                            "provision_name": "Major Restorative Services",
                            "provision_value": "50%",
                            "confidence_score": "HIGH",
                        },
                    ],
                    "booklet_validation": [],
                    "exceptions": [],
                }
            ],
        }
        facts = [
            _dental_booklet_fact("basic", "Basic Services Coinsurance", "80%"),
            _dental_booklet_fact("major", "Major Restorative Services Coinsurance", "50%"),
        ]

        result = validate_focused_proposals_against_booklet_facts(envelope, facts)

        validation = result["proposals"][0]["booklet_validation"]
        self.assertEqual(2, len(validation))
        self.assertEqual(["MATCH", "MATCH"], [row["validation_status"] for row in validation])
        self.assertEqual(["80%", "50%"], [row["booklet_value"] for row in validation])

    def test_unsupported_provision_does_not_create_validation_evidence(self) -> None:
        envelope = _envelope(("Manulife", "$300"), provision_name="Orthopaedic Shoes")

        result = validate_focused_proposals_against_booklet_facts(
            envelope,
            [_booklet_fact("$300", provision_code="maximum", provision_name="Maximum")],
        )

        proposal = result["proposals"][0]
        self.assertEqual([], proposal["booklet_validation"])
        self.assertIn("no canonical provision mapping", proposal["exceptions"][0]["message"].lower())

    def test_read_failure_keeps_proposals_and_records_review_issue(self) -> None:
        envelope = _envelope(("Manulife", "80%"))

        result = validate_focused_proposals_against_booklet(
            envelope,
            "missing-current-booklet.pdf",
        )

        self.assertEqual(1, len(result["proposals"]))
        self.assertEqual(1, len(result["proposals"][0]["plan_design"]))
        self.assertEqual([], result["proposals"][0]["booklet_validation"])
        self.assertEqual("NEEDS REVIEW", result["_booklet_validation_summary"]["status"])
        self.assertIn("could not run", result["proposals"][0]["exceptions"][0]["message"].lower())

    def test_validation_reuses_existing_table_and_reprocesses_idempotently(self) -> None:
        with tempfile.TemporaryDirectory() as folder:
            database_path = Path(folder) / "Benefits_Intelligence.db"
            _create_core_client(database_path)
            envelope = validate_focused_proposals_against_booklet_facts(
                _envelope(("Manulife", "80%"), ("RBC", "90%")),
                [_booklet_fact("80%")],
            )

            load_to_sqlite(envelope, database_path, client_id="42")
            load_to_sqlite(envelope, database_path, client_id="42")

            with closing(sqlite3.connect(database_path)) as connection:
                rows = connection.execute(
                    """
                    SELECT validation_status, provision_value, booklet_value
                    FROM proposal_booklet_validation
                    ORDER BY validation_status, provision_value
                    """
                ).fetchall()
                self.assertEqual(
                    [("DIFFERENT", "90%", "80%"), ("MATCH", "80%", "80%")],
                    rows,
                )
                self.assertEqual(
                    2,
                    connection.execute("SELECT COUNT(*) FROM proposal_files").fetchone()[0],
                )


def _envelope(*carrier_values: tuple[str, str], provision_name: str = "Coinsurance") -> dict:
    proposals = []
    for index, (carrier, value) in enumerate(carrier_values, start=1):
        proposals.append(
            {
                "proposal_summary": {
                    "client_name": "Canonical Client Ltd.",
                    "carrier_name": carrier,
                    "detected_carrier": carrier.lower(),
                    "source_file": f"{carrier.lower()}-proposal.pdf",
                    "source_file_name": f"{carrier.lower()}-proposal.pdf",
                    "source_file_hash": f"proposal-{index}",
                },
                "plan_design": [
                    {
                        "benefit_line": "Extended Health Care",
                        "provision_name": provision_name,
                        "provision_value": value,
                        "source_file": f"{carrier.lower()}-proposal.pdf",
                        "source_page": index + 2,
                        "confidence_score": "HIGH",
                        "status": "VALID",
                    }
                ],
                "booklet_validation": [],
                "exceptions": [],
            }
        )
    return {
        "client_id": "42",
        "client_name": "Canonical Client Ltd.",
        "proposals": proposals,
    }


def _booklet_fact(
    value: str,
    class_name: str = "All Employees",
    provision_code: str = "coinsurance",
    provision_name: str = "Coinsurance",
) -> dict:
    return {
        "benefit_line": "Extended Health Care",
        "benefit_code": "health",
        "provision_code": provision_code,
        "provision_name": provision_name,
        "provision_value": value,
        "class_name": class_name,
        "carrier": "Canada Life",
        "source_file": "current-booklet.pdf",
        "source_page": 12,
        "source_section": "Extended Health Care",
        "confidence_score": "HIGH",
    }


def _dental_booklet_fact(code: str, provision_name: str, value: str) -> dict:
    return {
        "benefit_line": f"Dental {provision_name}",
        "benefit_code": "dental",
        "provision_code": f"coinsurance.dental_{code}",
        "provision_name": provision_name,
        "provision_value": value,
        "class_name": "All Employees",
        "carrier": "Canada Life",
        "source_file": "current-booklet.pdf",
        "source_page": 22,
        "source_section": f"Dental Care - {provision_name}",
        "confidence_score": "HIGH",
    }


def _create_core_client(database_path: Path) -> None:
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


if __name__ == "__main__":
    unittest.main()
