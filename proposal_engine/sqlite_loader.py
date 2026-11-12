from __future__ import annotations

import hashlib
import json
import sqlite3
from contextlib import closing
from datetime import datetime, timezone
from pathlib import Path
from typing import Any

from proposal_engine import DEFAULT_DATABASE_PATH

COMMON_COLUMNS = [
    "client_id TEXT",
    "carrier_id INTEGER",
    "proposal_id INTEGER",
    "source_file_name TEXT",
    "source_file_hash TEXT",
    "source_page_number INTEGER",
    "source_sheet_name TEXT",
    "extraction_confidence TEXT",
    "status TEXT",
    "created_at TEXT",
    "updated_at TEXT",
]

TABLE_SPECS = {
    "clients": ["client_name TEXT", "external_client_id TEXT"],
    "carriers": ["carrier_key TEXT", "carrier_name TEXT"],
    "proposal_files": ["file_path TEXT", "detected_carrier TEXT", "detection_confidence TEXT", "reprocess_count INTEGER DEFAULT 0"],
    "benefits": ["benefit_key TEXT", "benefit_name TEXT", "normalized_benefit_name TEXT", "display_name TEXT", "benefit_category TEXT", "sort_order INTEGER", "is_standard INTEGER", "is_known INTEGER", "payload_json TEXT"],
    "proposal_summary": ["client_name TEXT", "carrier_name TEXT", "proposal_date TEXT", "effective_date TEXT", "detected_carrier TEXT", "detection_confidence TEXT", "payload_json TEXT"],
    "proposal_benefit_lines": ["benefit_line TEXT", "class_name TEXT", "tier TEXT", "lives REAL", "volume REAL", "rate REAL", "monthly_premium REAL", "annual_premium REAL", "payload_json TEXT"],
    "proposal_rates": ["benefit_line TEXT", "class_name TEXT", "tier TEXT", "age_band TEXT", "gender TEXT", "province TEXT", "lives REAL", "volume REAL", "rate REAL", "rate_basis TEXT", "monthly_premium REAL", "annual_premium REAL", "payload_json TEXT"],
    "proposal_premiums": ["benefit_line TEXT", "class_name TEXT", "tier TEXT", "monthly_premium REAL", "annual_premium REAL", "employer_cost REAL", "employee_cost REAL", "total_cost REAL", "payload_json TEXT"],
    "proposal_plan_design": ["benefit_line TEXT", "provision_name TEXT", "provision_value TEXT", "booklet_value TEXT", "variance TEXT", "validation_status TEXT", "payload_json TEXT"],
    "proposal_demographics": ["benefit_line TEXT", "class_name TEXT", "tier TEXT", "lives REAL", "volume REAL", "payload_json TEXT"],
    "proposal_underwriting_assumptions": ["assumption_name TEXT", "assumption_value TEXT", "payload_json TEXT"],
    "proposal_pooling_retention": ["benefit_line TEXT", "pooling_level TEXT", "pooling_charge REAL", "commission REAL", "retention REAL", "target_loss_ratio REAL", "rate_guarantee TEXT", "pooled_benefits TEXT", "experience_rated_benefits TEXT", "ibnr TEXT", "rate_caps TEXT", "first_renewal_cap TEXT", "second_renewal_cap TEXT", "payload_json TEXT"],
    "proposal_booklet_validation": ["benefit_line TEXT", "provision_name TEXT", "provision_value TEXT", "booklet_value TEXT", "variance TEXT", "validation_status TEXT", "payload_json TEXT"],
    "proposal_exceptions": ["exception_type TEXT", "field_name TEXT", "message TEXT", "payload_json TEXT"],
    "proposal_export_audit": ["action TEXT", "table_name TEXT", "rows_loaded INTEGER", "message TEXT", "payload_json TEXT"],
}

DETAIL_TABLE_MAP = {
    "benefits": "benefits",
    "benefit_lines": "proposal_benefit_lines",
    "rates": "proposal_rates",
    "premiums": "proposal_premiums",
    "plan_design": "proposal_plan_design",
    "demographics": "proposal_demographics",
    "underwriting_assumptions": "proposal_underwriting_assumptions",
    "pooling_retention": "proposal_pooling_retention",
    "exceptions": "proposal_exceptions",
}


def load_to_sqlite(
    validated_data: dict[str, Any],
    database_path: str | Path | None = None,
    client_id: str | None = None
) -> list[dict[str, Any]]:
    """Create/update proposal comparison tables and load normalized proposal data."""

    db_path = Path(database_path) if database_path else DEFAULT_DATABASE_PATH
    db_path.parent.mkdir(parents=True, exist_ok=True)
    audit: list[dict[str, Any]] = []

    with closing(sqlite3.connect(db_path)) as connection:
        connection.row_factory = sqlite3.Row
        _create_tables(connection)

        proposals = validated_data.get("proposals") or [validated_data]
        for proposal in proposals:
            proposal_audit = _load_one_proposal(connection, proposal, client_id or validated_data.get("client_id"))
            audit.extend(proposal_audit)

        _insert_audit_rows(connection, audit, client_id)
        connection.commit()

    validated_data["sqlite_audit"] = audit
    return audit


def sqlite_preview(database_path: str | Path | None = None) -> dict[str, int]:
    db_path = Path(database_path) if database_path else DEFAULT_DATABASE_PATH
    if not db_path.exists():
        return {}
    preview: dict[str, int] = {}
    with closing(sqlite3.connect(db_path)) as connection:
        for table_name in TABLE_SPECS:
            try:
                preview[table_name] = connection.execute(f"SELECT COUNT(*) FROM {table_name}").fetchone()[0]
            except sqlite3.Error:
                preview[table_name] = 0
    return preview


def _create_tables(connection: sqlite3.Connection) -> None:
    for table_name, columns in TABLE_SPECS.items():
        if table_name == "clients" and _is_core_clients_table(connection):
            continue
        column_sql = ["id INTEGER PRIMARY KEY AUTOINCREMENT", *COMMON_COLUMNS, *columns]
        connection.execute(f"CREATE TABLE IF NOT EXISTS {table_name} ({', '.join(column_sql)})")
        _ensure_columns(connection, table_name, [*COMMON_COLUMNS, *columns])
    connection.execute("CREATE INDEX IF NOT EXISTS idx_proposal_files_hash ON proposal_files(source_file_hash)")
    connection.execute("CREATE INDEX IF NOT EXISTS idx_proposal_files_client_hash ON proposal_files(client_id, carrier_id, source_file_hash)")
    connection.execute("CREATE INDEX IF NOT EXISTS idx_proposal_summary_client ON proposal_summary(client_id)")
    connection.execute("CREATE INDEX IF NOT EXISTS idx_proposal_rates_proposal ON proposal_rates(proposal_id)")
    connection.execute("CREATE INDEX IF NOT EXISTS idx_proposal_premiums_proposal ON proposal_premiums(proposal_id)")
    connection.execute("CREATE INDEX IF NOT EXISTS idx_proposal_plan_design_client ON proposal_plan_design(client_id)")
    connection.execute("CREATE INDEX IF NOT EXISTS idx_proposal_plan_design_proposal ON proposal_plan_design(proposal_id)")
    connection.execute("CREATE INDEX IF NOT EXISTS idx_proposal_validation_proposal ON proposal_booklet_validation(proposal_id)")


def _is_core_clients_table(connection: sqlite3.Connection) -> bool:
    columns = _table_columns(connection, "clients")
    return bool(columns) and "client_id" in columns and "id" not in columns


def _table_columns(connection: sqlite3.Connection, table_name: str) -> set[str]:
    return {
        str(row[1]).casefold()
        for row in connection.execute(f"PRAGMA table_info({table_name})").fetchall()
    }


def _ensure_columns(connection: sqlite3.Connection, table_name: str, column_specs: list[str]) -> None:
    existing = {
        row[1]
        for row in connection.execute(f"PRAGMA table_info({table_name})").fetchall()
    }
    for column_spec in column_specs:
        column_name = column_spec.split()[0]
        if column_name not in existing:
            connection.execute(f"ALTER TABLE {table_name} ADD COLUMN {column_spec}")


def _load_one_proposal(connection: sqlite3.Connection, proposal: dict[str, Any], client_id: str | None) -> list[dict[str, Any]]:
    audit: list[dict[str, Any]] = []
    now = _now()
    summary = proposal.get("proposal_summary", {})
    source_file = summary.get("source_file")
    source_path = Path(source_file) if source_file else None
    source_file_name = source_path.name if source_path else summary.get("source_file_name") or "unknown"
    source_hash = summary.get("source_file_hash") or (
        _file_hash(source_path)
        if source_path and source_path.exists() and source_path.is_file()
        else _text_hash(source_file_name)
    )
    effective_client_id = str(client_id or summary.get("client_name") or "default_client")

    uses_core_clients = _is_core_clients_table(connection)
    client_row_id = _upsert_client(connection, effective_client_id, summary.get("client_name"), now)
    if uses_core_clients and client_row_id > 0:
        effective_client_id = str(client_row_id)
    carrier_row_id = _upsert_carrier(connection, summary.get("detected_carrier"), summary.get("carrier_name"), now)
    proposal_id, action = _upsert_proposal_file(
        connection,
        effective_client_id,
        carrier_row_id,
        source_file_name,
        source_file or "",
        source_hash,
        summary.get("detected_carrier"),
        summary.get("detection_confidence"),
        now,
    )

    _delete_existing_detail_rows(connection, proposal_id)
    _insert_summary(connection, proposal, effective_client_id, carrier_row_id, proposal_id, source_file_name, source_hash, now)
    audit.append(_audit(action, "proposal_files", 1, proposal_id, source_file_name, source_hash, "VALID", "Proposal file registered."))

    for data_key, table_name in DETAIL_TABLE_MAP.items():
        rows = proposal.get(data_key, [])
        for row in rows:
            _insert_detail_row(connection, table_name, row, effective_client_id, carrier_row_id, proposal_id, source_file_name, source_hash, now)
        audit.append(_audit("load", table_name, len(rows), proposal_id, source_file_name, source_hash, "VALID", f"Loaded {len(rows)} row(s)."))

    booklet_rows = proposal.get("booklet_validation")
    if booklet_rows is None:
        booklet_rows = [
            row
            for row in proposal.get("plan_design", [])
            if any(row.get(field) not in (None, "") for field in ("booklet_value", "variance", "validation_status"))
        ]
    for row in booklet_rows:
        _insert_detail_row(connection, "proposal_booklet_validation", row, effective_client_id, carrier_row_id, proposal_id, source_file_name, source_hash, now)
    audit.append(_audit("load", "proposal_booklet_validation", len(booklet_rows), proposal_id, source_file_name, source_hash, "VALID", f"Loaded {len(booklet_rows)} row(s)."))

    _ = client_row_id
    return audit


def _upsert_client(connection: sqlite3.Connection, client_id: str, client_name: str | None, now: str) -> int:
    columns = _table_columns(connection, "clients")
    if "id" not in columns:
        row = connection.execute(
            "SELECT client_id FROM clients WHERE CAST(client_id AS TEXT) = ? LIMIT 1",
            (client_id,),
        ).fetchone()
        if row is None and client_name:
            row = connection.execute(
                "SELECT client_id FROM clients WHERE LOWER(client_name) = LOWER(?) LIMIT 1",
                (client_name,),
            ).fetchone()
        return int(row[0]) if row and str(row[0]).isdigit() else 0

    row = connection.execute("SELECT id FROM clients WHERE client_id = ?", (client_id,)).fetchone()
    if row:
        connection.execute(
            "UPDATE clients SET client_name = ?, updated_at = ? WHERE id = ?",
            (client_name, now, row["id"]),
        )
        return int(row["id"])

    cursor = connection.execute(
        _insert_sql("clients", ["client_id", "client_name", "external_client_id", "status", "created_at", "updated_at"]),
        (client_id, client_name, client_id, "VALID", now, now),
    )
    return int(cursor.lastrowid)


def _upsert_carrier(connection: sqlite3.Connection, carrier_key: str | None, carrier_name: str | None, now: str) -> int:
    key = carrier_key or "default_generic"
    name = carrier_name or "Unknown Carrier"
    row = connection.execute("SELECT id FROM carriers WHERE carrier_key = ?", (key,)).fetchone()
    if row:
        connection.execute("UPDATE carriers SET carrier_name = ?, updated_at = ? WHERE id = ?", (name, now, row["id"]))
        return int(row["id"])
    cursor = connection.execute(
        _insert_sql("carriers", ["carrier_key", "carrier_name", "status", "created_at", "updated_at"]),
        (key, name, "VALID", now, now),
    )
    return int(cursor.lastrowid)


def _upsert_proposal_file(
    connection: sqlite3.Connection,
    client_id: str,
    carrier_id: int,
    source_file_name: str,
    file_path: str,
    source_hash: str,
    detected_carrier: str | None,
    detection_confidence: str | None,
    now: str,
) -> tuple[int, str]:
    rows = connection.execute(
        """
        SELECT id, reprocess_count
        FROM proposal_files
        WHERE source_file_hash = ? AND client_id = ? AND carrier_id = ?
        ORDER BY updated_at DESC, id DESC
        """,
        (source_hash, client_id, carrier_id),
    ).fetchall()
    row = rows[0] if rows else None
    for duplicate in rows[1:]:
        duplicate_id = int(duplicate["id"])
        _delete_existing_detail_rows(connection, duplicate_id)
        connection.execute("DELETE FROM proposal_files WHERE id = ?", (duplicate_id,))

    if row:
        connection.execute(
            """
            UPDATE proposal_files
            SET client_id = ?, carrier_id = ?, proposal_id = ?, source_file_name = ?, file_path = ?,
                detected_carrier = ?, detection_confidence = ?, reprocess_count = ?, updated_at = ?
            WHERE id = ?
            """,
            (
                client_id,
                carrier_id,
                row["id"],
                source_file_name,
                file_path,
                detected_carrier,
                detection_confidence,
                int(row["reprocess_count"] or 0) + 1,
                now,
                row["id"],
            ),
        )
        return int(row["id"]), "reprocess"

    cursor = connection.execute(
        _insert_sql(
            "proposal_files",
            [
                "client_id",
                "carrier_id",
                "proposal_id",
                "source_file_name",
                "source_file_hash",
                "file_path",
                "detected_carrier",
                "detection_confidence",
                "extraction_confidence",
                "status",
                "created_at",
                "updated_at",
            ],
        ),
        (
            client_id,
            carrier_id,
            None,
            source_file_name,
            source_hash,
            file_path,
            detected_carrier,
            detection_confidence,
            detection_confidence,
            "VALID",
            now,
            now,
        ),
    )
    proposal_id = int(cursor.lastrowid)
    connection.execute("UPDATE proposal_files SET proposal_id = ? WHERE id = ?", (proposal_id, proposal_id))
    return proposal_id, "insert"


def _delete_existing_detail_rows(connection: sqlite3.Connection, proposal_id: int) -> None:
    for table_name in [
        "proposal_summary",
        *DETAIL_TABLE_MAP.values(),
        "proposal_booklet_validation",
    ]:
        connection.execute(f"DELETE FROM {table_name} WHERE proposal_id = ?", (proposal_id,))


def _insert_summary(
    connection: sqlite3.Connection,
    proposal: dict[str, Any],
    client_id: str,
    carrier_id: int,
    proposal_id: int,
    source_file_name: str,
    source_hash: str,
    now: str,
) -> None:
    summary = proposal.get("proposal_summary", {})
    row = {
        **_common_values(summary, client_id, carrier_id, proposal_id, source_file_name, source_hash, now),
        "client_name": summary.get("client_name"),
        "carrier_name": summary.get("carrier_name"),
        "proposal_date": summary.get("proposal_date"),
        "effective_date": summary.get("effective_date"),
        "detected_carrier": summary.get("detected_carrier"),
        "detection_confidence": summary.get("detection_confidence"),
        "payload_json": json.dumps(summary, default=str),
    }
    _insert_dict(connection, "proposal_summary", row)


def _insert_detail_row(
    connection: sqlite3.Connection,
    table_name: str,
    source: dict[str, Any],
    client_id: str,
    carrier_id: int,
    proposal_id: int,
    source_file_name: str,
    source_hash: str,
    now: str,
) -> None:
    detail_source_file = source.get("source_file")
    detail_source_path = Path(str(detail_source_file)) if detail_source_file else None
    detail_source_name = detail_source_path.name if detail_source_path else source_file_name
    detail_source_hash = source.get("source_file_hash") or (
        _file_hash(detail_source_path)
        if detail_source_path and detail_source_path.exists() and detail_source_path.is_file()
        else source_hash
    )
    row = _common_values(source, client_id, carrier_id, proposal_id, detail_source_name, detail_source_hash, now)
    for column in _table_specific_columns(table_name):
        row[column] = source.get(column)
    row["payload_json"] = json.dumps(source, default=str)
    _insert_dict(connection, table_name, row)


def _insert_audit_rows(connection: sqlite3.Connection, audit: list[dict[str, Any]], client_id: str | None) -> None:
    now = _now()
    for item in audit:
        row = {
            "client_id": client_id,
            "carrier_id": None,
            "proposal_id": item.get("proposal_id"),
            "source_file_name": item.get("source_file_name"),
            "source_file_hash": item.get("source_file_hash"),
            "source_page_number": None,
            "source_sheet_name": None,
            "extraction_confidence": None,
            "status": item.get("status"),
            "created_at": now,
            "updated_at": now,
            "action": item.get("action"),
            "table_name": item.get("table_name"),
            "rows_loaded": item.get("rows_loaded"),
            "message": item.get("message"),
            "payload_json": json.dumps(item, default=str),
        }
        _insert_dict(connection, "proposal_export_audit", row)


def _common_values(
    source: dict[str, Any],
    client_id: str,
    carrier_id: int,
    proposal_id: int,
    source_file_name: str,
    source_hash: str,
    now: str,
) -> dict[str, Any]:
    return {
        "client_id": client_id,
        "carrier_id": carrier_id,
        "proposal_id": proposal_id,
        "source_file_name": source_file_name,
        "source_file_hash": source_hash,
        "source_page_number": source.get("source_page"),
        "source_sheet_name": source.get("source_sheet"),
        "extraction_confidence": source.get("confidence_score"),
        "status": source.get("status"),
        "created_at": now,
        "updated_at": now,
    }


def _table_specific_columns(table_name: str) -> list[str]:
    columns = []
    for column_spec in TABLE_SPECS[table_name]:
        column = column_spec.split()[0]
        if column != "payload_json":
            columns.append(column)
    return columns


def _insert_dict(connection: sqlite3.Connection, table_name: str, values: dict[str, Any]) -> None:
    columns = list(values.keys())
    placeholders = ", ".join("?" for _ in columns)
    quoted_columns = ", ".join(columns)
    connection.execute(
        f"INSERT INTO {table_name} ({quoted_columns}) VALUES ({placeholders})",
        [values[column] for column in columns],
    )


def _insert_sql(table_name: str, columns: list[str]) -> str:
    return f"INSERT INTO {table_name} ({', '.join(columns)}) VALUES ({', '.join('?' for _ in columns)})"


def _file_hash(path: Path) -> str:
    hasher = hashlib.sha256()
    with path.open("rb") as file:
        for chunk in iter(lambda: file.read(1024 * 1024), b""):
            hasher.update(chunk)
    return hasher.hexdigest()


def _text_hash(value: str) -> str:
    return hashlib.sha256(value.encode("utf-8")).hexdigest()


def _now() -> str:
    return datetime.now(timezone.utc).isoformat()


def _audit(
    action: str,
    table_name: str,
    rows_loaded: int,
    proposal_id: int,
    source_file_name: str,
    source_hash: str,
    status: str,
    message: str,
) -> dict[str, Any]:
    return {
        "action": action,
        "table_name": table_name,
        "rows_loaded": rows_loaded,
        "proposal_id": proposal_id,
        "source_file_name": source_file_name,
        "source_file_hash": source_hash,
        "status": status,
        "message": message,
        "created_at": _now(),
    }
