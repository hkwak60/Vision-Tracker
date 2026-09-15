from __future__ import annotations

import sqlite3
import sys
import re
from contextlib import closing
from dataclasses import dataclass, replace
from datetime import datetime
from pathlib import Path
from typing import Any

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter


APP_TITLE = "Vision Issue Tracker"
APP_DIR = Path(sys.executable).parent if getattr(sys, "frozen", False) else Path.cwd()
DB_PATH = APP_DIR / "data" / "vision_issues.db"

LINES = ["1-1", "1-2", "2-1", "2-2"]
INSTRUMENTS = [
    "Pinhole",
    "Pouch Align",
    "Lead",
    "Sealing",
    "Lead Align",
    "Welding(+)",
    "Welding(-)",
]
INSTRUMENT_SEPARATOR = " / "
WORKERS = ["Hojun Kwak", "Kijung Kim", "Jihoon Yun", "Jisub Yun", "Seongwon Park"]
ACTIVE_STATUS_OPTIONS = ["Action Required", "Monitoring"]
STATUS_OPTIONS = ACTIVE_STATUS_OPTIONS + ["Resolved"]
CATEGORY_MAP = {
    "Hardware": ["Camera", "Lighting"],
    "Software": ["Program Crash", "Program Update", "Mavin Model Update", "UI", "PLC", "Other"],
    "Recipe": ["Overkill", "Underkill", "Add Measure", "Bypass/Unbypass"],
    "Deep Learning": [
        "Overkill Training",
        "Leakage Training",
        "Overkill&Leakage",
        "Dataset Major Change",
        "Model Revert",
        "Model Update",
    ],
    "Camera Grab Fail": [""],
    "Production": [""],
    "Other": [""],
}
CATEGORIES = list(CATEGORY_MAP.keys())
VERSION_GROUPS = {
    "Welding": ["Welding(+)", "Welding(-)"],
    "Common": ["Pinhole", "Pouch Align", "Lead Align"],
    "New Lead": ["Lead"],
    "Sealing": ["Sealing"],
}
INSTRUMENT_GROUP = {
    instrument: group_name
    for group_name, instruments in VERSION_GROUPS.items()
    for instrument in instruments
}
NO_ALGO_INSTRUMENTS = {"Sealing"}
DL_MODEL_FAMILIES = [
    "Crop_A",
    "Crop_B",
    "Crop_micro",
    "Crop_micro_tabside",
    "Gap_DL",
    "HORNMARK",
    "LEADEDGE",
    "SEGMENTATION",
    "SEPA",
    "SEPA_SHOULDER",
]
DL_CHANGE_TYPES = CATEGORY_MAP["Deep Learning"]
DL_TRAINING_CHANGE_TYPES = [
    "Overkill Training",
    "Leakage Training",
    "Overkill&Leakage",
    "Dataset Major Change",
]
DL_APPLICATION_CHANGE_TYPES = ["Model Update", "Model Revert"]
DL_TRAINED_STATUS_OPTIONS = ["Action Required", "Applied"]
DL_INSTRUMENTS = ["Welding(-)", "Welding(+)"]
DL_MACHINE_SEPARATOR = " / "
DL_MACHINE_TARGETS = [
    {
        "key": f"{line} {instrument}",
        "line": line,
        "instrument": instrument,
        "polarity": "Anode" if instrument == "Welding(-)" else "Cathode",
    }
    for line in LINES
    for instrument in DL_INSTRUMENTS
]
DL_MACHINE_KEYS = [target["key"] for target in DL_MACHINE_TARGETS]
DL_MACHINE_BY_KEY = {target["key"]: target for target in DL_MACHINE_TARGETS}


def instrument_uses_algo(instrument: str) -> bool:
    return instrument not in NO_ALGO_INSTRUMENTS


def version_group_uses_algo(group_name: str) -> bool:
    return any(instrument_uses_algo(instrument) for instrument in VERSION_GROUPS.get(group_name, []))


SW_VERSION_PATTERN = re.compile(r"^\s*(\d{6})\.(\d{4})\s*$")
ALGO_VERSION_PATTERN = re.compile(r"^\s*(\d+)\.(\d+)\.(\d+)\.(\d+)\s*$")


def sw_version_sort_key(value: str) -> tuple[int, int] | None:
    match = SW_VERSION_PATTERN.match(value or "")
    if not match:
        return None
    return int(match.group(1)), int(match.group(2))


def algo_version_sort_key(value: str) -> tuple[int, int, int, int] | None:
    match = ALGO_VERSION_PATTERN.match(value or "")
    if not match:
        return None
    return tuple(int(part) for part in match.groups())


def version_sort_key(value: str, component: str | None = None) -> tuple[int, ...] | None:
    if component == "sw":
        key = sw_version_sort_key(value)
        if key is not None:
            return key
    elif component == "algo":
        key = algo_version_sort_key(value)
        if key is not None:
            return key
    else:
        key = sw_version_sort_key(value) or algo_version_sort_key(value)
        if key is not None:
            return key

    digits: list[str] = []
    current = ""
    for char in value:
        if char.isdigit():
            current += char
        elif current:
            digits.append(current)
            current = ""
    if current:
        digits.append(current)
    if not digits:
        return None
    return tuple(int(part) for part in digits)


def version_component_list_sort_key(item: dict[str, str], component: str) -> tuple[int, tuple[int, ...], str, int]:
    key = version_sort_key(item.get("version", ""), component)
    try:
        row_id = int(item.get("id", "0") or 0)
    except ValueError:
        row_id = 0
    return (1 if key is not None else 0, key or tuple(), item.get("updated_at", ""), row_id)


def split_instruments(value: str) -> list[str]:
    instruments = [item.strip() for item in value.split("/") if item.strip()]
    return instruments


def format_instruments(values: list[str] | tuple[str, ...] | set[str]) -> str:
    ordered = [instrument for instrument in INSTRUMENTS if instrument in values]
    return INSTRUMENT_SEPARATOR.join(ordered)


def split_dl_targets(value: str) -> list[str]:
    return [item.strip() for item in value.split(DL_MACHINE_SEPARATOR) if item.strip()]


def serialize_dl_targets(values: list[str] | tuple[str, ...] | set[str]) -> str:
    ordered = [target for target in DL_MACHINE_KEYS if target in values]
    return DL_MACHINE_SEPARATOR.join(ordered)


def dl_targets_by_line(values: list[str] | tuple[str, ...] | set[str]) -> dict[str, list[str]]:
    grouped: dict[str, list[str]] = {}
    for key in [target for target in DL_MACHINE_KEYS if target in values]:
        target = DL_MACHINE_BY_KEY[key]
        grouped.setdefault(target["line"], []).append(target["instrument"])
    return grouped


def infer_dl_scope(values: list[str] | tuple[str, ...] | set[str]) -> str:
    selected = {target for target in values if target in DL_MACHINE_BY_KEY}
    if selected == set(DL_MACHINE_KEYS):
        return "Universal"
    lines = {DL_MACHINE_BY_KEY[target]["line"] for target in selected}
    instruments = {DL_MACHINE_BY_KEY[target]["instrument"] for target in selected}
    if len(lines) == 1 and instruments == set(DL_INSTRUMENTS):
        return "Line-specific"
    if len(instruments) == 1 and len(lines) > 1:
        return "Polarity-specific"
    return "Custom"


@dataclass(frozen=True)
class IssueInput:
    issue_time: str
    line: str
    instrument: str
    worker: str
    category: str
    subcategory: str
    title: str
    description: str
    status: str = ACTIVE_STATUS_OPTIONS[0]
    resolved_time: str = ""
    resolution_notes: str = ""


@dataclass(frozen=True)
class VersionInput:
    update_time: str
    group_name: str
    line: str
    instrument: str
    sw_version: str
    algo_version: str
    description: str
    worker: str
    sw_description: str = ""
    algo_description: str = ""


@dataclass(frozen=True)
class DeepLearningTrainedInput:
    trained_time: str
    model_family: str
    model_version: str
    change_type: str
    target_machines: tuple[str, ...]
    worker: str
    description: str
    status: str = ACTIVE_STATUS_OPTIONS[0]


@dataclass(frozen=True)
class DeepLearningApplicationInput:
    applied_time: str
    model_family: str
    model_version: str
    change_type: str
    target_machines: tuple[str, ...]
    worker: str
    description: str


def now_text() -> str:
    return datetime.now().strftime("%Y-%m-%d %H:%M")


def downtime_duration(issue_time: str, end_time: datetime | None = None) -> str:
    end_time = end_time or datetime.now()
    try:
        start_time = datetime.strptime(issue_time, "%Y-%m-%d %H:%M")
    except ValueError:
        return ""
    minutes = max(0, int((end_time - start_time).total_seconds() // 60))
    hours, remaining_minutes = divmod(minutes, 60)
    return f"{hours:02d}:{remaining_minutes:02d}"


def clean_source_metadata(value: str | None) -> str:
    if not value:
        return ""
    metadata_patterns = [
        re.compile(r"^\s*Source\s+No\.?\s*:.*$", re.IGNORECASE),
        re.compile(r"^\s*Original\s+Vision\s*:.*$", re.IGNORECASE),
    ]
    kept_lines = [
        line
        for line in str(value).splitlines()
        if not any(pattern.match(line) for pattern in metadata_patterns)
    ]
    cleaned = "\n".join(kept_lines).strip()
    return re.sub(r"\n{3,}", "\n\n", cleaned)


def connect(db_path: Path = DB_PATH) -> sqlite3.Connection:
    db_path.parent.mkdir(parents=True, exist_ok=True)
    conn = sqlite3.connect(db_path)
    conn.row_factory = sqlite3.Row
    return conn


def ensure_column(conn: sqlite3.Connection, table_name: str, column_name: str, definition: str) -> None:
    columns = {row["name"] for row in conn.execute(f"PRAGMA table_info({table_name})")}
    if column_name not in columns:
        conn.execute(f"ALTER TABLE {table_name} ADD COLUMN {column_name} {definition}")


def clean_issue_source_metadata(conn: sqlite3.Connection) -> None:
    rows = conn.execute(
        """
        SELECT id, description, resolution_notes
        FROM issues
        WHERE COALESCE(description, '') LIKE '%Source No%'
           OR COALESCE(description, '') LIKE '%Original Vision%'
           OR COALESCE(resolution_notes, '') LIKE '%Source No%'
           OR COALESCE(resolution_notes, '') LIKE '%Original Vision%'
        """
    ).fetchall()
    for row in rows:
        description = clean_source_metadata(row["description"])
        resolution_notes = clean_source_metadata(row["resolution_notes"])
        if description != (row["description"] or "") or resolution_notes != (row["resolution_notes"] or ""):
            conn.execute(
                """
                UPDATE issues
                SET description = ?, resolution_notes = ?
                WHERE id = ?
                """,
                (description, resolution_notes, row["id"]),
            )


def initialize_database(db_path: Path = DB_PATH) -> None:
    with closing(connect(db_path)) as conn:
        conn.execute(
            """
            CREATE TABLE IF NOT EXISTS issues (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                created_at TEXT NOT NULL,
                issue_time TEXT NOT NULL,
                resolved_time TEXT,
                line TEXT NOT NULL,
                instrument TEXT NOT NULL,
                worker TEXT NOT NULL,
                category TEXT NOT NULL,
                subcategory TEXT,
                title TEXT NOT NULL,
                description TEXT,
                status TEXT NOT NULL,
                resolution_notes TEXT
            )
            """
        )
        conn.execute(
            """
            CREATE INDEX IF NOT EXISTS idx_issues_lookup
            ON issues(status, category, subcategory, line, instrument, issue_time)
            """
        )
        conn.execute(
            """
            CREATE TABLE IF NOT EXISTS version_templates (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                group_name TEXT NOT NULL,
                sw_version TEXT NOT NULL,
                algo_version TEXT NOT NULL,
                description TEXT,
                worker TEXT NOT NULL,
                created_at TEXT NOT NULL,
                updated_at TEXT NOT NULL
            )
            """
        )
        conn.execute(
            """
            CREATE INDEX IF NOT EXISTS idx_version_templates_lookup
            ON version_templates(group_name, updated_at)
            """
        )
        conn.execute(
            """
            CREATE TABLE IF NOT EXISTS version_history (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                created_at TEXT NOT NULL,
                update_time TEXT NOT NULL,
                group_name TEXT NOT NULL,
                line TEXT NOT NULL,
                instrument TEXT NOT NULL,
                sw_version TEXT NOT NULL,
                algo_version TEXT NOT NULL,
                description TEXT,
                worker TEXT NOT NULL,
                created_issue_id INTEGER,
                FOREIGN KEY(created_issue_id) REFERENCES issues(id)
            )
            """
        )
        conn.execute(
            """
            CREATE INDEX IF NOT EXISTS idx_version_history_lookup
            ON version_history(line, instrument, group_name, update_time)
            """
        )
        conn.execute(
            """
            CREATE TABLE IF NOT EXISTS dl_trained_models (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                created_at TEXT NOT NULL,
                trained_time TEXT NOT NULL,
                model_family TEXT NOT NULL,
                model_version TEXT NOT NULL,
                change_type TEXT NOT NULL,
                target_machines TEXT NOT NULL,
                scope TEXT NOT NULL,
                description TEXT,
                worker TEXT NOT NULL,
                status TEXT NOT NULL,
                created_issue_id INTEGER,
                FOREIGN KEY(created_issue_id) REFERENCES issues(id)
            )
            """
        )
        conn.execute(
            """
            CREATE INDEX IF NOT EXISTS idx_dl_trained_models_lookup
            ON dl_trained_models(status, model_family, trained_time)
            """
        )
        conn.execute(
            """
            CREATE TABLE IF NOT EXISTS dl_model_applications (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                created_at TEXT NOT NULL,
                applied_time TEXT NOT NULL,
                model_family TEXT NOT NULL,
                model_version TEXT NOT NULL,
                change_type TEXT NOT NULL,
                line TEXT NOT NULL,
                polarity TEXT NOT NULL,
                instrument TEXT NOT NULL,
                machine TEXT NOT NULL,
                scope TEXT NOT NULL,
                description TEXT,
                worker TEXT NOT NULL,
                created_issue_id INTEGER,
                FOREIGN KEY(created_issue_id) REFERENCES issues(id)
            )
            """
        )
        conn.execute(
            """
            CREATE INDEX IF NOT EXISTS idx_dl_model_applications_lookup
            ON dl_model_applications(model_family, line, instrument, applied_time)
            """
        )
        ensure_column(conn, "version_templates", "sw_description", "TEXT")
        ensure_column(conn, "version_templates", "algo_description", "TEXT")
        ensure_column(conn, "version_history", "sw_description", "TEXT")
        ensure_column(conn, "version_history", "algo_description", "TEXT")
        ensure_column(conn, "version_history", "sw_touched", "INTEGER")
        ensure_column(conn, "version_history", "algo_touched", "INTEGER")
        backfill_version_history_component_flags(conn)
        backfill_version_history_descriptions(conn)
        conn.execute("UPDATE issues SET status = 'Action Required' WHERE status = 'Open'")
        conn.execute("UPDATE issues SET status = 'Monitoring' WHERE status = 'In Progress'")
        conn.execute("UPDATE issues SET subcategory = 'Overkill' WHERE subcategory = 'Overkill(False Reject)'")
        conn.execute("UPDATE issues SET subcategory = 'Underkill' WHERE subcategory = 'Underkill(False Accept)'")
        clean_issue_source_metadata(conn)
        conn.commit()


def validate_issue(issue: IssueInput) -> list[str]:
    errors: list[str] = []
    required = {
        "Issue time": issue.issue_time,
        "Line": issue.line,
        "Instrument": issue.instrument,
        "Worker": issue.worker,
        "Category": issue.category,
        "Title": issue.title,
        "Status": issue.status,
    }
    for label, value in required.items():
        if not value.strip():
            errors.append(f"{label} is required.")
    try:
        datetime.strptime(issue.issue_time, "%Y-%m-%d %H:%M")
    except ValueError:
        errors.append("Issue time must use YYYY-MM-DD HH:MM format.")
    if issue.line not in LINES:
        errors.append("Line is not valid.")
    issue_instruments = split_instruments(issue.instrument)
    if not issue_instruments:
        errors.append("Instrument is required.")
    invalid_instruments = [instrument for instrument in issue_instruments if instrument not in INSTRUMENTS]
    if invalid_instruments:
        errors.append("Instrument is not valid.")
    if issue.category not in CATEGORY_MAP:
        errors.append("Category is not valid.")
    if issue.status not in STATUS_OPTIONS:
        errors.append("Status is not valid.")
    allowed_subcategories = CATEGORY_MAP.get(issue.category, [])
    if issue.subcategory and issue.subcategory not in allowed_subcategories:
        errors.append("Subcategory is not valid for the selected category.")
    return errors


def create_issue(issue: IssueInput, db_path: Path = DB_PATH) -> int:
    errors = validate_issue(issue)
    if errors:
        raise ValueError("\n".join(errors))

    with closing(connect(db_path)) as conn:
        cursor = conn.execute(
            """
            INSERT INTO issues (
                created_at, issue_time, resolved_time, line, instrument, worker,
                category, subcategory, title, description, status, resolution_notes
            )
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """,
            (
                now_text(),
                issue.issue_time,
                issue.resolved_time,
                issue.line,
                issue.instrument,
                issue.worker,
                issue.category,
                issue.subcategory,
                issue.title,
                issue.description,
                issue.status,
                issue.resolution_notes,
            ),
        )
        conn.commit()
        return int(cursor.lastrowid)


def create_issues_for_lines(issue: IssueInput, lines: list[str] | tuple[str, ...] | set[str], db_path: Path = DB_PATH) -> list[int]:
    requested_lines = [line.strip() for line in lines if line.strip()]
    if not requested_lines:
        raise ValueError("Line is required.")
    invalid_lines = [line for line in requested_lines if line not in LINES]
    if invalid_lines:
        raise ValueError("Line is not valid.")
    selected_lines = [line for line in LINES if line in requested_lines]
    ids: list[int] = []
    for line in selected_lines:
        ids.append(create_issue(replace(issue, line=line), db_path))
    return ids


def update_issue(issue_id: int, issue: IssueInput, db_path: Path = DB_PATH) -> None:
    errors = validate_issue(issue)
    if errors:
        raise ValueError("\n".join(errors))

    with closing(connect(db_path)) as conn:
        conn.execute(
            """
            UPDATE issues
            SET issue_time = ?, resolved_time = ?, line = ?, instrument = ?,
                worker = ?, category = ?, subcategory = ?, title = ?,
                description = ?, status = ?, resolution_notes = ?
            WHERE id = ?
            """,
            (
                issue.issue_time,
                issue.resolved_time,
                issue.line,
                issue.instrument,
                issue.worker,
                issue.category,
                issue.subcategory,
                issue.title,
                issue.description,
                issue.status,
                issue.resolution_notes,
                issue_id,
            ),
        )
        conn.commit()


def resolve_issue(issue_id: int, notes: str = "", db_path: Path = DB_PATH) -> None:
    with closing(connect(db_path)) as conn:
        conn.execute(
            """
            UPDATE issues
            SET status = 'Resolved',
                resolution_notes = CASE WHEN ? = '' THEN resolution_notes ELSE ? END
            WHERE id = ?
            """,
            (notes, notes, issue_id),
        )
        conn.commit()


def set_issue_status(issue_id: int, status: str, db_path: Path = DB_PATH) -> None:
    if status not in STATUS_OPTIONS:
        raise ValueError("Status is not valid.")
    with closing(connect(db_path)) as conn:
        conn.execute("UPDATE issues SET status = ? WHERE id = ?", (status, issue_id))
        conn.commit()


def delete_issue(issue_id: int, db_path: Path = DB_PATH) -> None:
    with closing(connect(db_path)) as conn:
        conn.execute("DELETE FROM issues WHERE id = ?", (issue_id,))
        conn.commit()


def validate_version_update(version: VersionInput) -> list[str]:
    errors: list[str] = []
    required = {
        "Update time": version.update_time,
        "Group": version.group_name,
        "Line": version.line,
        "Instrument": version.instrument,
        "SW Version": version.sw_version,
        "Worker": version.worker,
    }
    if instrument_uses_algo(version.instrument):
        required["Algo Version"] = version.algo_version
    for label, value in required.items():
        if not value.strip():
            errors.append(f"{label} is required.")
    try:
        datetime.strptime(version.update_time, "%Y-%m-%d %H:%M")
    except ValueError:
        errors.append("Update time must use YYYY-MM-DD HH:MM format.")
    if version.group_name not in VERSION_GROUPS:
        errors.append("Version group is not valid.")
    if version.line not in LINES:
        errors.append("Line is not valid.")
    if version.instrument not in INSTRUMENTS:
        errors.append("Instrument is not valid.")
    if version.instrument and version.group_name != INSTRUMENT_GROUP.get(version.instrument):
        errors.append("Instrument is not part of the selected version group.")
    return errors


def split_version_description(description: str | None) -> tuple[str, str]:
    value = (description or "").strip()
    sw_marker = "[SW Description]"
    algo_marker = "[Algo Description]"
    if sw_marker in value or algo_marker in value:
        sw_text = value
        algo_text = ""
        if sw_marker in value:
            sw_text = value.split(sw_marker, 1)[1]
        if algo_marker in sw_text:
            sw_text, algo_text = sw_text.split(algo_marker, 1)
        elif algo_marker in value:
            _, algo_text = value.split(algo_marker, 1)
        return sw_text.strip(), algo_text.strip()
    return value, ""


def combine_version_description(sw_description: str, algo_description: str, uses_algo: bool) -> str:
    sw_text = sw_description.strip()
    algo_text = algo_description.strip()
    if not uses_algo:
        return sw_text
    parts = []
    if sw_text:
        parts.append(f"[SW Description]\n{sw_text}")
    if algo_text:
        parts.append(f"[Algo Description]\n{algo_text}")
    return "\n\n".join(parts)


def version_description_parts(row: sqlite3.Row) -> tuple[str, str]:
    keys = set(row.keys())
    fallback_sw, fallback_algo = split_version_description(row["description"] if "description" in keys else "")
    sw_description = (row["sw_description"] or "").strip() if "sw_description" in keys else ""
    algo_description = (row["algo_description"] or "").strip() if "algo_description" in keys else ""
    if not sw_description:
        sw_description = fallback_sw
    if not algo_description:
        algo_description = fallback_algo
    if sw_description or algo_description:
        return sw_description, algo_description
    return "", ""


def infer_version_history_component_flags(row: sqlite3.Row) -> tuple[bool, bool]:
    sw_description, algo_description = version_description_parts(row)
    uses_algo = True
    keys = set(row.keys())
    if "instrument" in keys:
        uses_algo = instrument_uses_algo(row["instrument"])
    elif "group_name" in keys:
        uses_algo = version_group_uses_algo(row["group_name"])

    if not uses_algo:
        return True, False
    if sw_description and not algo_description:
        return True, False
    if algo_description and not sw_description:
        return False, True
    return True, True


def backfill_version_history_component_flags(conn: sqlite3.Connection) -> None:
    rows = conn.execute(
        """
        SELECT id, group_name, instrument, description, sw_description, algo_description,
               sw_touched, algo_touched
        FROM version_history
        WHERE sw_touched IS NULL OR algo_touched IS NULL
        """
    ).fetchall()
    for row in rows:
        sw_touched, algo_touched = infer_version_history_component_flags(row)
        conn.execute(
            """
            UPDATE version_history
            SET sw_touched = ?, algo_touched = ?
            WHERE id = ?
            """,
            (1 if sw_touched else 0, 1 if algo_touched else 0, row["id"]),
        )


def version_component_description_from_templates(
    conn: sqlite3.Connection,
    group_name: str,
    component: str,
    version: str,
) -> str:
    if not version.strip():
        return ""
    version_column = "sw_version" if component == "sw" else "algo_version"
    rows = conn.execute(
        f"""
        SELECT description, sw_description, algo_description
        FROM version_templates
        WHERE group_name = ? AND {version_column} = ?
        ORDER BY updated_at DESC, id DESC
        """,
        (group_name, version.strip()),
    ).fetchall()
    for row in rows:
        sw_description, algo_description = version_description_parts(row)
        description = sw_description if component == "sw" else algo_description
        if description:
            return description
    return ""


def backfill_version_history_descriptions(conn: sqlite3.Connection) -> None:
    rows = conn.execute(
        """
        SELECT id, group_name, instrument, sw_version, algo_version, description,
               sw_description, algo_description, sw_touched, algo_touched
        FROM version_history
        WHERE COALESCE(description, '') = ''
           OR COALESCE(sw_description, '') = ''
           OR COALESCE(algo_description, '') = ''
        """
    ).fetchall()
    for row in rows:
        sw_description, algo_description = version_description_parts(row)
        sw_touched, algo_touched = version_history_component_flags(row)
        if sw_touched and not sw_description:
            sw_description = version_component_description_from_templates(
                conn, row["group_name"], "sw", row["sw_version"]
            )
        if algo_touched and not algo_description:
            algo_description = version_component_description_from_templates(
                conn, row["group_name"], "algo", row["algo_version"]
            )
        if not version_group_uses_algo(row["group_name"]):
            algo_description = ""
        description = combine_version_description(
            sw_description,
            algo_description,
            version_group_uses_algo(row["group_name"]),
        )
        if (
            description != (row["description"] or "")
            or sw_description != (row["sw_description"] or "")
            or algo_description != (row["algo_description"] or "")
        ):
            conn.execute(
                """
                UPDATE version_history
                SET description = ?, sw_description = ?, algo_description = ?
                WHERE id = ?
                """,
                (description, sw_description, algo_description, row["id"]),
            )


def refresh_combined_version_descriptions(conn: sqlite3.Connection, table_name: str, group_name: str) -> None:
    rows = conn.execute(
        f"""
        SELECT id, description, sw_description, algo_description
        FROM {table_name}
        WHERE group_name = ?
        """,
        (group_name,),
    ).fetchall()
    uses_algo = version_group_uses_algo(group_name)
    for row in rows:
        sw_description, algo_description = version_description_parts(row)
        combined = combine_version_description(sw_description, algo_description, uses_algo)
        conn.execute(
            f"UPDATE {table_name} SET description = ? WHERE id = ?",
            (combined, row["id"]),
        )


def save_version_template(
    group_name: str,
    sw_version: str,
    algo_version: str,
    description: str,
    worker: str,
    db_path: Path = DB_PATH,
    sw_description: str | None = None,
    algo_description: str | None = None,
) -> int:
    if group_name not in VERSION_GROUPS:
        raise ValueError("Version group is not valid.")
    if not sw_version.strip():
        raise ValueError("SW Version is required.")
    if version_group_uses_algo(group_name) and not algo_version.strip():
        raise ValueError("Algo Version is required.")
    if sw_description is None and algo_description is None:
        sw_description_value, algo_description_value = split_version_description(description)
    else:
        sw_description_value = (sw_description or "").strip()
        algo_description_value = (algo_description or "").strip()
    if not version_group_uses_algo(group_name):
        algo_description_value = ""
    timestamp = now_text()
    with closing(connect(db_path)) as conn:
        if not sw_description_value:
            sw_description_value = version_component_description_from_templates(
                conn, group_name, "sw", sw_version
            )
        if version_group_uses_algo(group_name) and not algo_description_value:
            algo_description_value = version_component_description_from_templates(
                conn, group_name, "algo", algo_version
            )
        combined_description = (
            combine_version_description(
                sw_description_value,
                algo_description_value,
                version_group_uses_algo(group_name),
            )
            or description
        )
        existing = conn.execute(
            """
            SELECT id, description, sw_description, algo_description FROM version_templates
            WHERE group_name = ? AND sw_version = ? AND algo_version = ?
            ORDER BY id DESC LIMIT 1
            """,
            (group_name, sw_version.strip(), algo_version.strip()),
        ).fetchone()
        if existing:
            existing_sw_description, existing_algo_description = version_description_parts(existing)
            if not sw_description_value:
                sw_description_value = existing_sw_description
            if version_group_uses_algo(group_name) and not algo_description_value:
                algo_description_value = existing_algo_description
            combined_description = (
                combine_version_description(
                    sw_description_value,
                    algo_description_value,
                    version_group_uses_algo(group_name),
                )
                or existing["description"]
                or description
            )
            conn.execute(
                """
                UPDATE version_templates
                SET description = ?, sw_description = ?, algo_description = ?, worker = ?, updated_at = ?
                WHERE id = ?
                """,
                (
                    combined_description,
                    sw_description_value,
                    algo_description_value,
                    worker,
                    timestamp,
                    existing["id"],
                ),
            )
            conn.commit()
            return int(existing["id"])
        cursor = conn.execute(
            """
            INSERT INTO version_templates (
                group_name, sw_version, algo_version, description, sw_description, algo_description,
                worker, created_at, updated_at
            )
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
            """,
            (
                group_name,
                sw_version.strip(),
                algo_version.strip(),
                combined_description,
                sw_description_value,
                algo_description_value,
                worker,
                timestamp,
                timestamp,
            ),
        )
        conn.commit()
        return int(cursor.lastrowid)


def recent_version_templates(group_name: str, limit: int = 3, db_path: Path = DB_PATH) -> list[sqlite3.Row]:
    with closing(connect(db_path)) as conn:
        return list(
            conn.execute(
                """
                SELECT id, group_name, sw_version, algo_version, description, sw_description,
                       algo_description, worker, created_at, updated_at
                FROM version_templates
                WHERE group_name = ?
                ORDER BY updated_at DESC, id DESC
                LIMIT ?
                """,
                (group_name, limit),
            )
        )


def get_version_template(template_id: int, db_path: Path = DB_PATH) -> sqlite3.Row | None:
    with closing(connect(db_path)) as conn:
        return conn.execute(
            """
            SELECT id, group_name, sw_version, algo_version, description, sw_description,
                   algo_description, worker, created_at, updated_at
            FROM version_templates
            WHERE id = ?
            """,
            (template_id,),
        ).fetchone()


def version_component_templates(
    group_name: str,
    component: str,
    limit: int = 50,
    db_path: Path = DB_PATH,
) -> list[dict[str, str]]:
    if group_name not in VERSION_GROUPS:
        raise ValueError("Version group is not valid.")
    if component not in {"sw", "algo"}:
        raise ValueError("Version component is not valid.")
    if component == "algo" and not version_group_uses_algo(group_name):
        return []

    version_column = "sw_version" if component == "sw" else "algo_version"
    with closing(connect(db_path)) as conn:
        rows = conn.execute(
            """
            SELECT id, group_name, sw_version, algo_version, description, sw_description,
                   algo_description, worker, created_at, updated_at
            FROM version_templates
            WHERE group_name = ?
            ORDER BY updated_at DESC, id DESC
            """,
            (group_name,),
        ).fetchall()

    templates: list[dict[str, str]] = []
    seen: set[str] = set()
    for row in rows:
        version = (row[version_column] or "").strip()
        if not version or version in seen:
            continue
        sw_description, algo_description = version_description_parts(row)
        templates.append(
            {
                "id": str(row["id"]),
                "group_name": row["group_name"],
                "component": component,
                "version": version,
                "description": sw_description if component == "sw" else algo_description,
                "worker": row["worker"],
                "created_at": row["created_at"],
                "updated_at": row["updated_at"],
            }
        )
        seen.add(version)
    templates.sort(key=lambda item: version_component_list_sort_key(item, component), reverse=True)
    return templates[:limit]


def save_version_component_template(
    group_name: str,
    component: str,
    version: str,
    description: str,
    worker: str,
    db_path: Path = DB_PATH,
) -> None:
    if group_name not in VERSION_GROUPS:
        raise ValueError("Version group is not valid.")
    if component not in {"sw", "algo"}:
        raise ValueError("Version component is not valid.")
    if component == "algo" and not version_group_uses_algo(group_name):
        raise ValueError("This version group does not use Algo versions.")
    version_value = version.strip()
    if not version_value:
        raise ValueError("Version is required.")

    version_column = "sw_version" if component == "sw" else "algo_version"
    description_column = "sw_description" if component == "sw" else "algo_description"
    timestamp = now_text()
    with closing(connect(db_path)) as conn:
        template_cursor = conn.execute(
            f"""
            UPDATE version_templates
            SET {description_column} = ?, worker = ?, updated_at = ?
            WHERE group_name = ? AND {version_column} = ?
            """,
            (description, worker, timestamp, group_name, version_value),
        )
        conn.execute(
            f"""
            UPDATE version_history
            SET {description_column} = ?, worker = ?
            WHERE group_name = ? AND {version_column} = ?
            """,
            (description, worker, group_name, version_value),
        )
        if template_cursor.rowcount == 0:
            sw_version = version_value if component == "sw" else ""
            algo_version = version_value if component == "algo" else ""
            sw_description = description if component == "sw" else ""
            algo_description = description if component == "algo" else ""
            combined = combine_version_description(sw_description, algo_description, version_group_uses_algo(group_name))
            conn.execute(
                """
                INSERT INTO version_templates (
                    group_name, sw_version, algo_version, description, sw_description, algo_description,
                    worker, created_at, updated_at
                )
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
                """,
                (
                    group_name,
                    sw_version,
                    algo_version,
                    combined,
                    sw_description,
                    algo_description,
                    worker,
                    timestamp,
                    timestamp,
                ),
            )
        refresh_combined_version_descriptions(conn, "version_templates", group_name)
        refresh_combined_version_descriptions(conn, "version_history", group_name)
        conn.commit()


def update_version_component_template(
    group_name: str,
    component: str,
    old_version: str,
    new_version: str,
    description: str,
    worker: str,
    db_path: Path = DB_PATH,
) -> None:
    if group_name not in VERSION_GROUPS:
        raise ValueError("Version group is not valid.")
    if component not in {"sw", "algo"}:
        raise ValueError("Version component is not valid.")
    if component == "algo" and not version_group_uses_algo(group_name):
        raise ValueError("This version group does not use Algo versions.")
    if not old_version.strip():
        raise ValueError("Select a version first.")
    if not new_version.strip():
        raise ValueError("Version is required.")

    version_column = "sw_version" if component == "sw" else "algo_version"
    description_column = "sw_description" if component == "sw" else "algo_description"
    timestamp = now_text()
    with closing(connect(db_path)) as conn:
        conn.execute(
            f"""
            UPDATE version_templates
            SET {version_column} = ?, {description_column} = ?, worker = ?, updated_at = ?
            WHERE group_name = ? AND {version_column} = ?
            """,
            (new_version.strip(), description, worker, timestamp, group_name, old_version.strip()),
        )
        conn.execute(
            f"""
            UPDATE version_history
            SET {version_column} = ?, {description_column} = ?, worker = ?
            WHERE group_name = ? AND {version_column} = ?
            """,
            (new_version.strip(), description, worker, group_name, old_version.strip()),
        )
        refresh_combined_version_descriptions(conn, "version_templates", group_name)
        refresh_combined_version_descriptions(conn, "version_history", group_name)
        conn.commit()


def delete_version_component_template(
    group_name: str,
    component: str,
    version: str,
    db_path: Path = DB_PATH,
) -> None:
    if group_name not in VERSION_GROUPS:
        raise ValueError("Version group is not valid.")
    if component not in {"sw", "algo"}:
        raise ValueError("Version component is not valid.")
    if not version.strip():
        return
    version_column = "sw_version" if component == "sw" else "algo_version"
    with closing(connect(db_path)) as conn:
        conn.execute(
            f"DELETE FROM version_templates WHERE group_name = ? AND {version_column} = ?",
            (group_name, version.strip()),
        )
        conn.execute(
            f"DELETE FROM version_history WHERE group_name = ? AND {version_column} = ?",
            (group_name, version.strip()),
        )
        conn.commit()


def update_version_template(
    template_id: int,
    sw_version: str,
    algo_version: str,
    description: str,
    worker: str,
    db_path: Path = DB_PATH,
) -> None:
    if not sw_version.strip():
        raise ValueError("SW Version is required.")
    timestamp = now_text()
    with closing(connect(db_path)) as conn:
        row = conn.execute(
            """
            SELECT group_name, sw_version, algo_version
            FROM version_templates
            WHERE id = ?
            """,
            (template_id,),
        ).fetchone()
        if row is None:
            raise ValueError("Version template was not found.")
        if version_group_uses_algo(row["group_name"]) and not algo_version.strip():
            raise ValueError("Algo Version is required.")
        sw_description, algo_description = split_version_description(description)
        if not version_group_uses_algo(row["group_name"]):
            algo_description = ""
        conn.execute(
            """
            UPDATE version_templates
            SET sw_version = ?, algo_version = ?, description = ?, sw_description = ?,
                algo_description = ?, worker = ?, updated_at = ?
            WHERE id = ?
            """,
            (
                sw_version.strip(),
                algo_version.strip(),
                description,
                sw_description,
                algo_description,
                worker,
                timestamp,
                template_id,
            ),
        )
        conn.execute(
            """
            UPDATE version_history
            SET sw_version = ?, algo_version = ?, description = ?, sw_description = ?,
                algo_description = ?, worker = ?
            WHERE group_name = ? AND sw_version = ? AND algo_version = ?
            """,
            (
                sw_version.strip(),
                algo_version.strip(),
                description,
                sw_description,
                algo_description,
                worker,
                row["group_name"],
                row["sw_version"],
                row["algo_version"],
            ),
        )
        conn.commit()


def delete_version_template(template_id: int, db_path: Path = DB_PATH) -> None:
    with closing(connect(db_path)) as conn:
        row = conn.execute(
            """
            SELECT group_name, sw_version, algo_version
            FROM version_templates
            WHERE id = ?
            """,
            (template_id,),
        ).fetchone()
        if row is None:
            return
        conn.execute("DELETE FROM version_templates WHERE id = ?", (template_id,))
        conn.execute(
            """
            DELETE FROM version_history
            WHERE group_name = ? AND sw_version = ? AND algo_version = ?
            """,
            (row["group_name"], row["sw_version"], row["algo_version"]),
        )
        conn.commit()


def latest_version_by_instrument(db_path: Path = DB_PATH) -> dict[tuple[str, str], sqlite3.Row]:
    with closing(connect(db_path)) as conn:
        rows = conn.execute(
            """
            SELECT vh.*
            FROM version_history vh
            JOIN (
                SELECT line, instrument, MAX(update_time || printf('%012d', id)) AS latest_key
                FROM version_history
                GROUP BY line, instrument
            ) latest
              ON latest.line = vh.line
             AND latest.instrument = vh.instrument
             AND latest.latest_key = vh.update_time || printf('%012d', vh.id)
            """
        ).fetchall()
        return {(row["line"], row["instrument"]): row for row in rows}


def version_history_component_flags(row: sqlite3.Row) -> tuple[bool, bool]:
    keys = set(row.keys())
    if (
        "sw_touched" in keys
        and "algo_touched" in keys
        and row["sw_touched"] is not None
        and row["algo_touched"] is not None
    ):
        sw_touched = bool(row["sw_touched"])
        algo_touched = bool(row["algo_touched"])
        if "instrument" in keys and not instrument_uses_algo(row["instrument"]):
            algo_touched = False
        elif "group_name" in keys and not version_group_uses_algo(row["group_name"]):
            algo_touched = False
        return sw_touched, algo_touched
    return infer_version_history_component_flags(row)


def latest_dashboard_versions(db_path: Path = DB_PATH) -> dict[tuple[str, str], dict[str, str]]:
    states: dict[tuple[str, str], dict[str, str]] = {}
    with closing(connect(db_path)) as conn:
        rows = conn.execute(
            """
            SELECT id, created_at, update_time, group_name, line, instrument, sw_version,
                   algo_version, description, sw_description, algo_description, worker,
                   created_issue_id, sw_touched, algo_touched
            FROM version_history
            ORDER BY id ASC
            """
        ).fetchall()

    for row in rows:
        key = (row["line"], row["instrument"])
        state = states.setdefault(
            key,
            {
                "line": row["line"],
                "instrument": row["instrument"],
                "group_name": row["group_name"],
                "sw_version": "",
                "algo_version": "",
                "update_time": "",
                "sw_update_time": "",
                "algo_update_time": "",
            },
        )
        sw_touched, algo_touched = version_history_component_flags(row)
        if sw_touched:
            state["sw_version"] = row["sw_version"] or ""
            state["sw_update_time"] = row["update_time"] or ""
            state["group_name"] = row["group_name"]
            state["update_time"] = row["update_time"] or state["update_time"]
        if instrument_uses_algo(row["instrument"]) and algo_touched:
            state["algo_version"] = row["algo_version"] or ""
            state["algo_update_time"] = row["update_time"] or ""
            state["group_name"] = row["group_name"]
            state["update_time"] = row["update_time"] or state["update_time"]
    return states


def version_history_rows(db_path: Path = DB_PATH) -> list[sqlite3.Row]:
    with closing(connect(db_path)) as conn:
        return list(
            conn.execute(
                """
                SELECT id, update_time, group_name, line, instrument, sw_version,
                       algo_version, description, sw_description, algo_description,
                       sw_touched, algo_touched, worker, created_issue_id
                FROM version_history
                ORDER BY update_time DESC, id DESC
                """
            )
        )


def export_version_dashboard_to_excel(output_path: Path, db_path: Path = DB_PATH) -> None:
    output_path.parent.mkdir(parents=True, exist_ok=True)
    latest = latest_dashboard_versions(db_path)
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Version Dashboard"

    headers = [
        "Line",
        "Vision",
        "Group",
        "SW Version",
        "Algo Version",
        "Last Updated",
    ]
    sheet.append(headers)

    for cell in sheet[1]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="1F4E78")

    for line in LINES:
        for instrument in INSTRUMENTS:
            row = latest.get((line, instrument))
            sheet.append(
                [
                    line,
                    instrument,
                    INSTRUMENT_GROUP[instrument],
                    row["sw_version"] if row else "",
                    row["algo_version"] if row else "",
                    row["update_time"] if row else "",
                ]
            )

    for column_index, header in enumerate(headers, start=1):
        max_length = len(header)
        for cell in sheet[get_column_letter(column_index)]:
            max_length = max(max_length, len(str(cell.value or "")))
        sheet.column_dimensions[get_column_letter(column_index)].width = min(max_length + 2, 64)

    sheet.freeze_panes = "A2"
    workbook.save(output_path)


def create_version_update(
    version: VersionInput,
    create_program_update_issue: bool = True,
    db_path: Path = DB_PATH,
) -> int:
    errors = validate_version_update(version)
    if errors:
        raise ValueError("\n".join(errors))

    uses_algo = version_group_uses_algo(version.group_name)
    sw_description = version.sw_description.strip()
    algo_description = version.algo_description.strip()
    if not sw_description and not algo_description:
        sw_description, algo_description = split_version_description(version.description)
    if not uses_algo:
        algo_description = ""

    has_entered_description = bool(sw_description or algo_description)
    sw_touched = bool(sw_description) or not has_entered_description or not uses_algo
    algo_touched = uses_algo and (bool(algo_description) or not has_entered_description)

    with closing(connect(db_path)) as conn:
        if sw_touched and not sw_description:
            sw_description = version_component_description_from_templates(
                conn, version.group_name, "sw", version.sw_version
            )
        if algo_touched and not algo_description:
            algo_description = version_component_description_from_templates(
                conn, version.group_name, "algo", version.algo_version
            )

    description = (
        combine_version_description(
            sw_description,
            algo_description,
            uses_algo,
        )
        or version.description
    )

    save_version_template(
        version.group_name,
        version.sw_version,
        version.algo_version,
        description,
        version.worker,
        db_path,
        sw_description=sw_description,
        algo_description=algo_description,
    )
    created_issue_id: int | None = None
    if create_program_update_issue:
        version_text = f"SW {version.sw_version}"
        if instrument_uses_algo(version.instrument):
            version_text = f"{version_text} / Algo {version.algo_version}"
        issue = IssueInput(
            issue_time=version.update_time,
            resolved_time="00:00",
            line=version.line,
            instrument=version.instrument,
            worker=version.worker,
            category="Software",
            subcategory="Program Update",
            title=f"Program Update - {version.line} {version.instrument} {version_text}",
            description=description,
            status="Monitoring",
        )
        created_issue_id = create_issue(issue, db_path)

    with closing(connect(db_path)) as conn:
        cursor = conn.execute(
            """
            INSERT INTO version_history (
                created_at, update_time, group_name, line, instrument, sw_version,
                algo_version, description, sw_description, algo_description,
                sw_touched, algo_touched, worker, created_issue_id
            )
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """,
            (
                now_text(),
                version.update_time,
                version.group_name,
                version.line,
                version.instrument,
                version.sw_version,
                version.algo_version,
                description,
                sw_description,
                algo_description,
                1 if sw_touched else 0,
                1 if algo_touched else 0,
                version.worker,
                created_issue_id,
            ),
        )
        conn.commit()
        return int(cursor.lastrowid)


def validate_dl_targets(target_machines: tuple[str, ...]) -> list[str]:
    requested = [target.strip() for target in target_machines if target.strip()]
    if not requested:
        return ["Select at least one Welding machine."]
    invalid = [target for target in requested if target not in DL_MACHINE_BY_KEY]
    if invalid:
        return ["Welding machine selection is not valid."]
    return []


def validate_dl_trained_model(model: DeepLearningTrainedInput) -> list[str]:
    errors: list[str] = []
    required = {
        "Trained time": model.trained_time,
        "Model family": model.model_family,
        "Model version": model.model_version,
        "Change type": model.change_type,
        "Worker": model.worker,
    }
    for label, value in required.items():
        if not value.strip():
            errors.append(f"{label} is required.")
    try:
        datetime.strptime(model.trained_time, "%Y-%m-%d %H:%M")
    except ValueError:
        errors.append("Trained time must use YYYY-MM-DD HH:MM format.")
    if model.model_family not in DL_MODEL_FAMILIES:
        errors.append("Model family is not valid.")
    if model.change_type not in DL_TRAINING_CHANGE_TYPES:
        errors.append("Deep Learning change type is not valid.")
    if model.status not in DL_TRAINED_STATUS_OPTIONS:
        errors.append("Trained model status is not valid.")
    errors.extend(validate_dl_targets(model.target_machines))
    return errors


def validate_dl_application(model: DeepLearningApplicationInput) -> list[str]:
    errors: list[str] = []
    required = {
        "Applied time": model.applied_time,
        "Model family": model.model_family,
        "Model version": model.model_version,
        "Change type": model.change_type,
        "Worker": model.worker,
    }
    for label, value in required.items():
        if not value.strip():
            errors.append(f"{label} is required.")
    try:
        datetime.strptime(model.applied_time, "%Y-%m-%d %H:%M")
    except ValueError:
        errors.append("Applied time must use YYYY-MM-DD HH:MM format.")
    if model.model_family not in DL_MODEL_FAMILIES:
        errors.append("Model family is not valid.")
    if model.change_type not in DL_APPLICATION_CHANGE_TYPES:
        errors.append("Deep Learning change type is not valid.")
    errors.extend(validate_dl_targets(model.target_machines))
    return errors


def dl_issue_description(
    model_family: str,
    model_version: str,
    change_type: str,
    scope: str,
    target_machines: str,
    description: str,
) -> str:
    lines = [
        f"Model Family: {model_family}",
        f"Model Version: {model_version}",
        f"Change Type: {change_type}",
        f"Scope: {scope}",
        f"Targets: {target_machines}",
    ]
    if description.strip():
        lines.extend(["", description.strip()])
    return "\n".join(lines)


def create_dl_issues_for_targets(
    issue_time: str,
    model_family: str,
    model_version: str,
    change_type: str,
    target_machines: tuple[str, ...],
    scope: str,
    description: str,
    worker: str,
    status: str,
    title_prefix: str,
    db_path: Path = DB_PATH,
) -> list[int]:
    target_text = serialize_dl_targets(target_machines)
    issue_description = dl_issue_description(
        model_family,
        model_version,
        change_type,
        scope,
        target_text,
        description,
    )
    issue_ids: list[int] = []
    for line, instruments in dl_targets_by_line(target_machines).items():
        instrument_text = format_instruments(instruments)
        issue_ids.append(
            create_issue(
                IssueInput(
                    issue_time=issue_time,
                    resolved_time="00:00",
                    line=line,
                    instrument=instrument_text,
                    worker=worker,
                    category="Deep Learning",
                    subcategory="Model Update",
                    title=f"{title_prefix} - {line} {instrument_text} {model_family} {model_version}",
                    description=issue_description,
                    status=status,
                ),
                db_path,
            )
        )
    return issue_ids


def create_dl_trained_model(
    model: DeepLearningTrainedInput,
    create_action_issue: bool = True,
    db_path: Path = DB_PATH,
) -> int:
    errors = validate_dl_trained_model(model)
    if errors:
        raise ValueError("\n".join(errors))
    targets = tuple(target for target in DL_MACHINE_KEYS if target in set(model.target_machines))
    target_text = serialize_dl_targets(targets)
    scope = infer_dl_scope(targets)
    created_issue_id = 0
    if create_action_issue:
        issue_ids = create_dl_issues_for_targets(
            model.trained_time,
            model.model_family,
            model.model_version,
            model.change_type,
            targets,
            scope,
            model.description,
            model.worker,
            "Action Required",
            "Deep Learning Model Ready",
            db_path,
        )
        created_issue_id = issue_ids[0] if issue_ids else 0
    with closing(connect(db_path)) as conn:
        cursor = conn.execute(
            """
            INSERT INTO dl_trained_models (
                created_at, trained_time, model_family, model_version, change_type,
                target_machines, scope, description, worker, status, created_issue_id
            )
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """,
            (
                now_text(),
                model.trained_time,
                model.model_family,
                model.model_version,
                model.change_type,
                target_text,
                scope,
                model.description,
                model.worker,
                model.status,
                created_issue_id or None,
            ),
        )
        conn.commit()
        return int(cursor.lastrowid)


def create_dl_model_application(
    model: DeepLearningApplicationInput,
    create_monitoring_issue: bool = True,
    trained_model_id: int | None = None,
    db_path: Path = DB_PATH,
) -> list[int]:
    errors = validate_dl_application(model)
    if errors:
        raise ValueError("\n".join(errors))
    targets = tuple(target for target in DL_MACHINE_KEYS if target in set(model.target_machines))
    scope = infer_dl_scope(targets)
    created_issue_ids: list[int] = []
    if create_monitoring_issue:
        created_issue_ids = create_dl_issues_for_targets(
            model.applied_time,
            model.model_family,
            model.model_version,
            model.change_type,
            targets,
            scope,
            model.description,
            model.worker,
            "Monitoring",
            "Deep Learning Model Applied",
            db_path,
        )
    rows: list[int] = []
    with closing(connect(db_path)) as conn:
        for target_key in targets:
            target = DL_MACHINE_BY_KEY[target_key]
            cursor = conn.execute(
                """
                INSERT INTO dl_model_applications (
                    created_at, applied_time, model_family, model_version, change_type,
                    line, polarity, instrument, machine, scope, description, worker, created_issue_id
                )
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                """,
                (
                    now_text(),
                    model.applied_time,
                    model.model_family,
                    model.model_version,
                    model.change_type,
                    target["line"],
                    target["polarity"],
                    target["instrument"],
                    target_key,
                    scope,
                    model.description,
                    model.worker,
                    created_issue_ids[0] if created_issue_ids else None,
                ),
            )
            rows.append(int(cursor.lastrowid))
        if trained_model_id:
            conn.execute(
                "UPDATE dl_trained_models SET status = 'Applied' WHERE id = ?",
                (trained_model_id,),
            )
        conn.commit()
    return rows


def list_dl_trained_models(
    model_family: str | None = None,
    status: str | None = None,
    db_path: Path = DB_PATH,
) -> list[sqlite3.Row]:
    clauses: list[str] = []
    params: list[Any] = []
    if model_family:
        clauses.append("model_family = ?")
        params.append(model_family)
    if status:
        clauses.append("status = ?")
        params.append(status)
    where = f"WHERE {' AND '.join(clauses)}" if clauses else ""
    with closing(connect(db_path)) as conn:
        return list(
            conn.execute(
                f"""
                SELECT id, created_at, trained_time, model_family, model_version,
                       change_type, target_machines, scope, description, worker,
                       status, created_issue_id
                FROM dl_trained_models
                {where}
                ORDER BY trained_time DESC, id DESC
                """,
                params,
            )
        )


def dl_application_rows(model_family: str | None = None, db_path: Path = DB_PATH) -> list[sqlite3.Row]:
    clauses: list[str] = []
    params: list[Any] = []
    if model_family:
        clauses.append("model_family = ?")
        params.append(model_family)
    where = f"WHERE {' AND '.join(clauses)}" if clauses else ""
    with closing(connect(db_path)) as conn:
        return list(
            conn.execute(
                f"""
                SELECT id, created_at, applied_time, model_family, model_version,
                       change_type, line, polarity, instrument, machine, scope,
                       description, worker, created_issue_id
                FROM dl_model_applications
                {where}
                ORDER BY applied_time DESC, id DESC
                """,
                params,
            )
        )


def latest_dl_applied_models(
    model_family: str | None = None,
    db_path: Path = DB_PATH,
) -> dict[tuple[str, str, str], sqlite3.Row]:
    latest: dict[tuple[str, str, str], sqlite3.Row] = {}
    rows = sorted(
        dl_application_rows(model_family, db_path),
        key=lambda row: (row["applied_time"], int(row["id"])),
    )
    for row in rows:
        latest[(row["model_family"], row["line"], row["instrument"])] = row
    return latest


def mark_dl_trained_model_status(
    trained_model_id: int,
    status: str,
    db_path: Path = DB_PATH,
) -> None:
    if status not in DL_TRAINED_STATUS_OPTIONS:
        raise ValueError("Trained model status is not valid.")
    with closing(connect(db_path)) as conn:
        conn.execute(
            "UPDATE dl_trained_models SET status = ? WHERE id = ?",
            (status, trained_model_id),
        )
        conn.commit()


def export_deep_learning_dashboard_to_excel(output_path: Path, db_path: Path = DB_PATH) -> None:
    output_path.parent.mkdir(parents=True, exist_ok=True)
    latest = latest_dl_applied_models(db_path=db_path)
    trained_rows = list_dl_trained_models(db_path=db_path)
    workbook = Workbook()
    applied_sheet = workbook.active
    applied_sheet.title = "Applied Models"
    trained_sheet = workbook.create_sheet("Trained Models")

    applied_headers = [
        "Model Family",
        "Line",
        "Polarity",
        "Instrument",
        "Model Version",
        "Scope",
        "Time",
        "Logged By",
        "Description",
    ]
    trained_headers = [
        "Status",
        "Model Family",
        "Model Version",
        "Change Type",
        "Scope",
        "Target Machines",
        "Time",
        "Logged By",
        "Description",
    ]
    applied_sheet.append(applied_headers)
    trained_sheet.append(trained_headers)

    for family in DL_MODEL_FAMILIES:
        for target in DL_MACHINE_TARGETS:
            row = latest.get((family, target["line"], target["instrument"]))
            applied_sheet.append(
                [
                    family,
                    target["line"],
                    target["polarity"],
                    target["instrument"],
                    row["model_version"] if row else "",
                    row["scope"] if row else "",
                    row["applied_time"] if row else "",
                    row["worker"] if row else "",
                    row["description"] if row else "",
                ]
            )

    for row in trained_rows:
        trained_sheet.append(
            [
                row["status"],
                row["model_family"],
                row["model_version"],
                row["change_type"],
                row["scope"],
                row["target_machines"],
                row["trained_time"],
                row["worker"],
                row["description"],
            ]
        )

    for sheet in [applied_sheet, trained_sheet]:
        for cell in sheet[1]:
            cell.font = Font(bold=True, color="FFFFFF")
            cell.fill = PatternFill("solid", fgColor="1F4E78")
            cell.alignment = Alignment(vertical="top")
        for column_index, header in enumerate([cell.value for cell in sheet[1]], start=1):
            column_letter = get_column_letter(column_index)
            max_length = len(str(header or ""))
            for cell in sheet[column_letter]:
                max_length = max(max_length, len(str(cell.value or "")))
                cell.alignment = Alignment(wrap_text=header == "Description", vertical="top")
            if header == "Description":
                sheet.column_dimensions[column_letter].width = 55
            else:
                sheet.column_dimensions[column_letter].width = min(max_length + 2, 28)
        sheet.freeze_panes = "A2"

    workbook.save(output_path)


def build_search_query(filters: dict[str, str]) -> tuple[str, list[Any]]:
    clauses: list[str] = []
    params: list[Any] = []

    exact_fields = ["status", "line", "category", "subcategory", "worker"]
    for field in exact_fields:
        value = filters.get(field, "").strip()
        if value:
            clauses.append(f"{field} = ?")
            params.append(value)

    selected_instruments = split_instruments(filters.get("instrument", "").strip())
    if selected_instruments:
        instrument_clauses: list[str] = []
        for instrument in selected_instruments:
            instrument_clauses.append(
                "(instrument = ? OR instrument LIKE ? OR instrument LIKE ? OR instrument LIKE ?)"
            )
            params.extend(
                [
                    instrument,
                    f"{instrument}{INSTRUMENT_SEPARATOR}%",
                    f"%{INSTRUMENT_SEPARATOR}{instrument}{INSTRUMENT_SEPARATOR}%",
                    f"%{INSTRUMENT_SEPARATOR}{instrument}",
                ]
            )
        clauses.append(f"({' OR '.join(instrument_clauses)})")

    date_from = filters.get("date_from", "").strip()
    date_to = filters.get("date_to", "").strip()
    if date_from:
        clauses.append("issue_time >= ?")
        params.append(date_from)
    if date_to:
        clauses.append("issue_time <= ?")
        params.append(date_to)

    keyword = filters.get("keyword", "").strip()
    if keyword:
        clauses.append("(title LIKE ? OR description LIKE ? OR resolution_notes LIKE ?)")
        like_value = f"%{keyword}%"
        params.extend([like_value, like_value, like_value])

    where = f"WHERE {' AND '.join(clauses)}" if clauses else ""
    query = f"""
        SELECT id, issue_time, resolved_time, line, instrument, worker, category,
               subcategory, title, description, status, resolution_notes
        FROM issues
        {where}
        ORDER BY issue_time ASC, id ASC
    """
    return query, params


def active_issues(db_path: Path = DB_PATH) -> list[sqlite3.Row]:
    with closing(connect(db_path)) as conn:
        return list(
            conn.execute(
                """
                SELECT id, issue_time, resolved_time, line, instrument, worker, category,
                       subcategory, title, description, status, resolution_notes
                FROM issues
                WHERE status IN (?, ?)
                ORDER BY issue_time ASC, id ASC
                """,
                tuple(ACTIVE_STATUS_OPTIONS),
            )
        )


def dashboard_counts(db_path: Path = DB_PATH) -> dict[str, int]:
    today = datetime.now().strftime("%Y-%m-%d")
    with closing(connect(db_path)) as conn:
        rows = conn.execute(
            """
            SELECT status, COUNT(*) AS count
            FROM issues
            GROUP BY status
            """
        ).fetchall()
        counts = {row["status"]: int(row["count"]) for row in rows}
        resolved_today = conn.execute(
            """
            SELECT COUNT(*) AS count
            FROM issues
            WHERE status = 'Resolved' AND issue_time >= ? AND issue_time < ?
            """,
            (f"{today} 00:00", f"{today} 23:59"),
        ).fetchone()
        counts["Resolved Today"] = int(resolved_today["count"]) if resolved_today else 0
        counts["Active"] = sum(counts.get(status, 0) for status in ACTIVE_STATUS_OPTIONS)
        return counts


def issue_time_bounds(db_path: Path = DB_PATH) -> tuple[str, str]:
    today = datetime.now().strftime("%Y-%m-%d")
    with closing(connect(db_path)) as conn:
        row = conn.execute(
            """
            SELECT MIN(issue_time) AS first_time,
                   MAX(issue_time) AS latest_time
            FROM issues
            """
        ).fetchone()
        first_time = row["first_time"] if row and row["first_time"] else f"{today} 00:00"
        latest_time = row["latest_time"] if row and row["latest_time"] else f"{today} 23:59"
        return first_time, latest_time


def search_issues(filters: dict[str, str] | None = None, db_path: Path = DB_PATH) -> list[sqlite3.Row]:
    filters = filters or {}
    query, params = build_search_query(filters)
    with closing(connect(db_path)) as conn:
        return list(conn.execute(query, params))


def get_issue(issue_id: int, db_path: Path = DB_PATH) -> sqlite3.Row | None:
    with closing(connect(db_path)) as conn:
        return conn.execute("SELECT * FROM issues WHERE id = ?", (issue_id,)).fetchone()


def export_issues_to_excel(rows: list[sqlite3.Row], output_path: Path) -> None:
    output_path.parent.mkdir(parents=True, exist_ok=True)
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Issue Report"

    headers = [
        "ID",
        "Line",
        "Instrument",
        "Issue Time",
        "Downtime",
        "Category",
        "Title",
        "Status",
        "Description",
        "Resolution Notes",
    ]
    sheet.append(headers)
    hidden_headers = {"Downtime"}
    wrapped_headers = {"Description", "Resolution Notes"}
    fixed_widths = {
        "Title": 48,
        "Description": 48,
        "Resolution Notes": 113.57,  # Excel column width equivalent for roughly 800 px.
    }

    for cell in sheet[1]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="1F4E78")
        cell.alignment = Alignment(vertical="top")

    for row_number, row in enumerate(rows, start=1):
        values = []
        for header in headers:
            if header == "ID":
                values.append(row_number)
                continue
            value = row[header_key(header)]
            if header in wrapped_headers:
                value = clean_source_metadata(value)
            values.append(value)
        sheet.append(values)

    for column_index, header in enumerate(headers, start=1):
        column_letter = get_column_letter(column_index)
        max_length = len(header)
        for cell in sheet[column_letter]:
            max_length = max(max_length, len(str(cell.value or "")))
            if header in wrapped_headers:
                cell.alignment = Alignment(wrap_text=True, vertical="top")
            else:
                cell.alignment = Alignment(vertical="top")
        if header in fixed_widths:
            sheet.column_dimensions[column_letter].width = fixed_widths[header]
        else:
            sheet.column_dimensions[column_letter].width = max_length + 2
        if header in hidden_headers:
            sheet.column_dimensions[column_letter].hidden = True

    sheet.freeze_panes = "A2"
    workbook.save(output_path)


def header_key(header: str) -> str:
    return {
        "ID": "id",
        "Line": "line",
        "Instrument": "instrument",
        "Issue Time": "issue_time",
        "Downtime": "resolved_time",
        "Category": "category",
        "Title": "title",
        "Status": "status",
        "Description": "description",
        "Resolution Notes": "resolution_notes",
    }[header]


# Online Apps Script / Google Sheets adapter. These wrappers preserve the local
# SQLite implementation when no online config is present.
import json as _json
import os as _os
import urllib.error as _urllib_error
import urllib.request as _urllib_request

_LOCAL_INITIALIZE_DATABASE = initialize_database
_LOCAL_CREATE_ISSUE = create_issue
_LOCAL_CREATE_ISSUES_FOR_LINES = create_issues_for_lines
_LOCAL_UPDATE_ISSUE = update_issue
_LOCAL_RESOLVE_ISSUE = resolve_issue
_LOCAL_SET_ISSUE_STATUS = set_issue_status
_LOCAL_DELETE_ISSUE = delete_issue
_LOCAL_ACTIVE_ISSUES = active_issues
_LOCAL_DASHBOARD_COUNTS = dashboard_counts
_LOCAL_ISSUE_TIME_BOUNDS = issue_time_bounds
_LOCAL_SEARCH_ISSUES = search_issues
_LOCAL_GET_ISSUE = get_issue
_LOCAL_SAVE_VERSION_TEMPLATE = save_version_template
_LOCAL_RECENT_VERSION_TEMPLATES = recent_version_templates
_LOCAL_GET_VERSION_TEMPLATE = get_version_template
_LOCAL_VERSION_COMPONENT_TEMPLATES = version_component_templates
_LOCAL_SAVE_VERSION_COMPONENT_TEMPLATE = save_version_component_template
_LOCAL_UPDATE_VERSION_COMPONENT_TEMPLATE = update_version_component_template
_LOCAL_DELETE_VERSION_COMPONENT_TEMPLATE = delete_version_component_template
_LOCAL_UPDATE_VERSION_TEMPLATE = update_version_template
_LOCAL_DELETE_VERSION_TEMPLATE = delete_version_template
_LOCAL_LATEST_VERSION_BY_INSTRUMENT = latest_version_by_instrument
_LOCAL_LATEST_DASHBOARD_VERSIONS = latest_dashboard_versions
_LOCAL_VERSION_HISTORY_ROWS = version_history_rows
_LOCAL_CREATE_VERSION_UPDATE = create_version_update
_LOCAL_CREATE_DL_TRAINED_MODEL = create_dl_trained_model
_LOCAL_CREATE_DL_MODEL_APPLICATION = create_dl_model_application
_LOCAL_LIST_DL_TRAINED_MODELS = list_dl_trained_models
_LOCAL_DL_APPLICATION_ROWS = dl_application_rows
_LOCAL_LATEST_DL_APPLIED_MODELS = latest_dl_applied_models
_LOCAL_MARK_DL_TRAINED_MODEL_STATUS = mark_dl_trained_model_status
_LOCAL_EXPORT_DEEP_LEARNING_DASHBOARD_TO_EXCEL = export_deep_learning_dashboard_to_excel

_ISSUE_HEADERS = [
    "id", "created_at", "issue_time", "resolved_time", "line", "instrument", "worker",
    "category", "subcategory", "title", "description", "status", "resolution_notes",
]
_TEMPLATE_HEADERS = [
    "id", "group_name", "sw_version", "algo_version", "description", "sw_description",
    "algo_description", "worker", "created_at", "updated_at",
]
_HISTORY_HEADERS = [
    "id", "created_at", "update_time", "group_name", "line", "instrument", "sw_version",
    "algo_version", "description", "sw_description", "algo_description", "sw_touched",
    "algo_touched", "worker", "created_issue_id",
]
_DL_TRAINED_HEADERS = [
    "id", "created_at", "trained_time", "model_family", "model_version", "change_type",
    "target_machines", "scope", "description", "worker", "status", "created_issue_id",
]
_DL_APPLICATION_HEADERS = [
    "id", "created_at", "applied_time", "model_family", "model_version", "change_type",
    "line", "polarity", "instrument", "machine", "scope", "description", "worker",
    "created_issue_id",
]
_NUMERIC_FIELDS = {"id", "created_issue_id", "sw_touched", "algo_touched"}


def _config_path() -> Path:
    override = _os.environ.get("VISION_TRACKER_CONFIG", "").strip()
    if override:
        return Path(override)
    return APP_DIR / "config.json"


def _load_online_config() -> dict[str, Any]:
    path = _config_path()
    if not path.exists():
        return {}
    try:
        return _json.loads(path.read_text(encoding="utf-8-sig"))
    except Exception:
        return {}


def _online_enabled(db_path: Path = DB_PATH) -> bool:
    if Path(db_path) != DB_PATH:
        return False
    config = _load_online_config()
    url = str(config.get("apps_script_url", ""))
    token = str(config.get("api_token", ""))
    return (
        str(config.get("mode", "")).lower() == "online"
        and url.startswith("https://script.google.com/")
        and "PASTE_" not in url
        and bool(token)
        and "PASTE_" not in token
    )


def _online_request(action: str, **payload: Any) -> dict[str, Any]:
    config = _load_online_config()
    url = str(config.get("apps_script_url", ""))
    token = str(config.get("api_token", ""))
    timeout = int(config.get("request_timeout_seconds", 20) or 20)
    body = dict(payload)
    body["action"] = action
    body["token"] = token
    data = _json.dumps(body).encode("utf-8")
    request = _urllib_request.Request(
        url,
        data=data,
        headers={"Content-Type": "application/json"},
        method="POST",
    )
    try:
        with _urllib_request.urlopen(request, timeout=timeout) as response:
            raw = response.read().decode("utf-8")
    except _urllib_error.HTTPError as exc:
        exc.close()
        hint = "Check the deployed Apps Script URL and access settings in config.json."
        raise ValueError(
            f"Online API HTTP {exc.code}. {hint} "
            "The save result is unconfirmed; refresh history before retrying."
        ) from exc
    except _urllib_error.URLError as exc:
        raise ValueError(f"Online API connection failed: {exc.reason}") from exc
    try:
        result = _json.loads(raw)
    except ValueError as exc:
        raise ValueError("Online API returned a non-JSON response. Check the deployment URL and access settings.") from exc
    if not isinstance(result, dict):
        raise ValueError("Online API returned an invalid response.")
    if not result.get("ok"):
        raise ValueError(str(result.get("error") or "Online API request failed."))
    return result


def _normalize_online_row(table: str, row: dict[str, Any] | None) -> dict[str, Any] | None:
    if row is None:
        return None
    headers = {
        "issues": _ISSUE_HEADERS,
        "version_templates": _TEMPLATE_HEADERS,
        "version_history": _HISTORY_HEADERS,
        "dl_trained_models": _DL_TRAINED_HEADERS,
        "dl_model_applications": _DL_APPLICATION_HEADERS,
    }[table]
    normalized: dict[str, Any] = {}
    for header in headers:
        value = row.get(header, "")
        if header in _NUMERIC_FIELDS:
            normalized[header] = int(value) if str(value).strip() not in {"", "None"} else 0
        else:
            normalized[header] = "" if value is None else str(value)
    return normalized


def _online_rows(table: str) -> list[dict[str, Any]]:
    rows = _online_request("list", table=table).get("rows", [])
    return [_normalize_online_row(table, row) for row in rows]


def _online_get(table: str, row_id: int) -> dict[str, Any] | None:
    row = _online_request("get", table=table, id=row_id).get("row")
    return _normalize_online_row(table, row)


def _online_append(table: str, row: dict[str, Any]) -> dict[str, Any]:
    return _normalize_online_row(table, _online_request("append", table=table, row=row).get("row"))


def _online_bulk_append(table: str, rows: list[dict[str, Any]]) -> list[dict[str, Any]]:
    created = _online_request("bulkAppend", table=table, rows=rows).get("rows", [])
    return [_normalize_online_row(table, row) for row in created]


def _online_update(table: str, row_id: int, row: dict[str, Any]) -> dict[str, Any]:
    return _normalize_online_row(table, _online_request("update", table=table, id=row_id, row=row).get("row"))


def _online_delete(table: str, row_id: int) -> None:
    _online_request("delete", table=table, id=row_id)


def _issue_to_online_row(issue: IssueInput, created_at: str | None = None) -> dict[str, Any]:
    return {
        "created_at": created_at or now_text(),
        "issue_time": issue.issue_time,
        "resolved_time": issue.resolved_time,
        "line": issue.line,
        "instrument": issue.instrument,
        "worker": issue.worker,
        "category": issue.category,
        "subcategory": issue.subcategory,
        "title": issue.title,
        "description": issue.description,
        "status": issue.status,
        "resolution_notes": issue.resolution_notes,
    }


def _sorted_issues(rows: list[dict[str, Any]]) -> list[dict[str, Any]]:
    return sorted(rows, key=lambda row: (row.get("issue_time", ""), int(row.get("id", 0))))


def _issue_matches_filters(row: dict[str, Any], filters: dict[str, str]) -> bool:
    for field in ["status", "line", "category", "subcategory", "worker"]:
        value = filters.get(field, "").strip()
        if value and row.get(field, "") != value:
            return False
    selected_instruments = split_instruments(filters.get("instrument", "").strip())
    if selected_instruments:
        row_instruments = set(split_instruments(row.get("instrument", "")))
        if not any(instrument in row_instruments for instrument in selected_instruments):
            return False
    date_from = filters.get("date_from", "").strip()
    date_to = filters.get("date_to", "").strip()
    issue_time = row.get("issue_time", "")
    if date_from and issue_time < date_from:
        return False
    if date_to and issue_time > date_to:
        return False
    keyword = filters.get("keyword", "").strip().lower()
    if keyword:
        text = " ".join(str(row.get(field, "")) for field in ["title", "description", "resolution_notes"]).lower()
        if keyword not in text:
            return False
    return True


def _sort_by_updated(rows: list[dict[str, Any]]) -> list[dict[str, Any]]:
    return sorted(rows, key=lambda row: (row.get("updated_at", ""), int(row.get("id", 0))), reverse=True)


def _online_version_component_description(group_name: str, component: str, version: str) -> str:
    if not version.strip():
        return ""
    version_column = "sw_version" if component == "sw" else "algo_version"
    for row in _sort_by_updated(_online_rows("version_templates")):
        if row["group_name"] == group_name and row[version_column] == version.strip():
            sw_description, algo_description = version_description_parts(row)
            description = sw_description if component == "sw" else algo_description
            if description:
                return description
    return ""


def _combined_for_online_row(row: dict[str, Any]) -> str:
    return combine_version_description(
        row.get("sw_description", ""),
        row.get("algo_description", ""),
        version_group_uses_algo(row.get("group_name", "")),
    )


def initialize_database(db_path: Path = DB_PATH) -> None:
    if not _online_enabled(db_path):
        return _LOCAL_INITIALIZE_DATABASE(db_path)
    _online_request("init")


def create_issue(issue: IssueInput, db_path: Path = DB_PATH) -> int:
    if not _online_enabled(db_path):
        return _LOCAL_CREATE_ISSUE(issue, db_path)
    errors = validate_issue(issue)
    if errors:
        raise ValueError("\n".join(errors))
    row = _online_append("issues", _issue_to_online_row(issue))
    return int(row["id"])


def create_issues_for_lines(issue: IssueInput, lines: list[str] | tuple[str, ...] | set[str], db_path: Path = DB_PATH) -> list[int]:
    if not _online_enabled(db_path):
        return _LOCAL_CREATE_ISSUES_FOR_LINES(issue, lines, db_path)
    requested_lines = [line.strip() for line in lines if line.strip()]
    if not requested_lines:
        raise ValueError("Line is required.")
    invalid_lines = [line for line in requested_lines if line not in LINES]
    if invalid_lines:
        raise ValueError("Line is not valid.")
    selected_lines = [line for line in LINES if line in requested_lines]
    created_at = now_text()
    rows = []
    for line in selected_lines:
        line_issue = IssueInput(**{**issue.__dict__, "line": line})
        errors = validate_issue(line_issue)
        if errors:
            raise ValueError("\n".join(errors))
        rows.append(_issue_to_online_row(line_issue, created_at=created_at))
    return [int(row["id"]) for row in _online_bulk_append("issues", rows)]


def update_issue(issue_id: int, issue: IssueInput, db_path: Path = DB_PATH) -> None:
    if not _online_enabled(db_path):
        return _LOCAL_UPDATE_ISSUE(issue_id, issue, db_path)
    errors = validate_issue(issue)
    if errors:
        raise ValueError("\n".join(errors))
    _online_update("issues", issue_id, _issue_to_online_row(issue))


def resolve_issue(issue_id: int, notes: str = "", db_path: Path = DB_PATH) -> None:
    if not _online_enabled(db_path):
        return _LOCAL_RESOLVE_ISSUE(issue_id, notes, db_path)
    updates: dict[str, Any] = {"status": "Resolved"}
    if notes:
        updates["resolution_notes"] = notes
    _online_update("issues", issue_id, updates)


def set_issue_status(issue_id: int, status: str, db_path: Path = DB_PATH) -> None:
    if not _online_enabled(db_path):
        return _LOCAL_SET_ISSUE_STATUS(issue_id, status, db_path)
    if status not in STATUS_OPTIONS:
        raise ValueError("Status is not valid.")
    _online_update("issues", issue_id, {"status": status})


def delete_issue(issue_id: int, db_path: Path = DB_PATH) -> None:
    if not _online_enabled(db_path):
        return _LOCAL_DELETE_ISSUE(issue_id, db_path)
    _online_delete("issues", issue_id)


def active_issues(db_path: Path = DB_PATH) -> list[dict[str, Any]]:
    if not _online_enabled(db_path):
        return _LOCAL_ACTIVE_ISSUES(db_path)
    return _sorted_issues([row for row in _online_rows("issues") if row.get("status") in ACTIVE_STATUS_OPTIONS])


def dashboard_counts(db_path: Path = DB_PATH) -> dict[str, int]:
    if not _online_enabled(db_path):
        return _LOCAL_DASHBOARD_COUNTS(db_path)
    today = datetime.now().strftime("%Y-%m-%d")
    counts: dict[str, int] = {}
    for row in _online_rows("issues"):
        status = row.get("status", "")
        counts[status] = counts.get(status, 0) + 1
    counts["Resolved Today"] = sum(
        1
        for row in _online_rows("issues")
        if row.get("status") == "Resolved" and f"{today} 00:00" <= row.get("issue_time", "") < f"{today} 23:59"
    )
    counts["Active"] = sum(counts.get(status, 0) for status in ACTIVE_STATUS_OPTIONS)
    return counts


def issue_time_bounds(db_path: Path = DB_PATH) -> tuple[str, str]:
    if not _online_enabled(db_path):
        return _LOCAL_ISSUE_TIME_BOUNDS(db_path)
    today = datetime.now().strftime("%Y-%m-%d")
    times = [row.get("issue_time", "") for row in _online_rows("issues") if row.get("issue_time")]
    if not times:
        return f"{today} 00:00", f"{today} 23:59"
    return min(times), max(times)


def search_issues(filters: dict[str, str] | None = None, db_path: Path = DB_PATH) -> list[dict[str, Any]]:
    if not _online_enabled(db_path):
        return _LOCAL_SEARCH_ISSUES(filters, db_path)
    filters = filters or {}
    return _sorted_issues([row for row in _online_rows("issues") if _issue_matches_filters(row, filters)])


def get_issue(issue_id: int, db_path: Path = DB_PATH) -> dict[str, Any] | None:
    if not _online_enabled(db_path):
        return _LOCAL_GET_ISSUE(issue_id, db_path)
    return _online_get("issues", issue_id)


def save_version_template(
    group_name: str,
    sw_version: str,
    algo_version: str,
    description: str,
    worker: str,
    db_path: Path = DB_PATH,
    sw_description: str | None = None,
    algo_description: str | None = None,
) -> int:
    if not _online_enabled(db_path):
        return _LOCAL_SAVE_VERSION_TEMPLATE(group_name, sw_version, algo_version, description, worker, db_path, sw_description, algo_description)
    if group_name not in VERSION_GROUPS:
        raise ValueError("Version group is not valid.")
    if not sw_version.strip():
        raise ValueError("SW Version is required.")
    if version_group_uses_algo(group_name) and not algo_version.strip():
        raise ValueError("Algo Version is required.")
    if sw_description is None and algo_description is None:
        sw_description_value, algo_description_value = split_version_description(description)
    else:
        sw_description_value = (sw_description or "").strip()
        algo_description_value = (algo_description or "").strip()
    if not version_group_uses_algo(group_name):
        algo_description_value = ""
    if not sw_description_value:
        sw_description_value = _online_version_component_description(group_name, "sw", sw_version)
    if version_group_uses_algo(group_name) and not algo_description_value:
        algo_description_value = _online_version_component_description(group_name, "algo", algo_version)
    combined_description = combine_version_description(sw_description_value, algo_description_value, version_group_uses_algo(group_name)) or description
    timestamp = now_text()
    templates = _online_rows("version_templates")
    existing = None
    for row in sorted(templates, key=lambda item: int(item.get("id", 0)), reverse=True):
        if row["group_name"] == group_name and row["sw_version"] == sw_version.strip() and row["algo_version"] == algo_version.strip():
            existing = row
            break
    row_data = {
        "group_name": group_name,
        "sw_version": sw_version.strip(),
        "algo_version": algo_version.strip(),
        "description": combined_description,
        "sw_description": sw_description_value,
        "algo_description": algo_description_value,
        "worker": worker,
        "updated_at": timestamp,
    }
    if existing:
        updated = _online_update("version_templates", int(existing["id"]), row_data)
        return int(updated["id"])
    row_data["created_at"] = timestamp
    created = _online_append("version_templates", row_data)
    return int(created["id"])


def recent_version_templates(group_name: str, limit: int = 3, db_path: Path = DB_PATH) -> list[dict[str, Any]]:
    if not _online_enabled(db_path):
        return _LOCAL_RECENT_VERSION_TEMPLATES(group_name, limit, db_path)
    return _sort_by_updated([row for row in _online_rows("version_templates") if row["group_name"] == group_name])[:limit]


def get_version_template(template_id: int, db_path: Path = DB_PATH) -> dict[str, Any] | None:
    if not _online_enabled(db_path):
        return _LOCAL_GET_VERSION_TEMPLATE(template_id, db_path)
    return _online_get("version_templates", template_id)


def version_component_templates(group_name: str, component: str, limit: int = 50, db_path: Path = DB_PATH) -> list[dict[str, str]]:
    if not _online_enabled(db_path):
        return _LOCAL_VERSION_COMPONENT_TEMPLATES(group_name, component, limit, db_path)
    if group_name not in VERSION_GROUPS:
        raise ValueError("Version group is not valid.")
    if component not in {"sw", "algo"}:
        raise ValueError("Version component is not valid.")
    if component == "algo" and not version_group_uses_algo(group_name):
        return []
    version_column = "sw_version" if component == "sw" else "algo_version"
    templates: list[dict[str, str]] = []
    seen: set[str] = set()
    for row in _sort_by_updated([row for row in _online_rows("version_templates") if row["group_name"] == group_name]):
        version = row.get(version_column, "").strip()
        if not version or version in seen:
            continue
        sw_description, algo_description = version_description_parts(row)
        templates.append({
            "id": str(row["id"]),
            "group_name": row["group_name"],
            "component": component,
            "version": version,
            "description": sw_description if component == "sw" else algo_description,
            "worker": row["worker"],
            "created_at": row["created_at"],
            "updated_at": row["updated_at"],
        })
        seen.add(version)
    templates.sort(key=lambda item: version_component_list_sort_key(item, component), reverse=True)
    return templates[:limit]


def save_version_component_template(group_name: str, component: str, version: str, description: str, worker: str, db_path: Path = DB_PATH) -> None:
    if not _online_enabled(db_path):
        return _LOCAL_SAVE_VERSION_COMPONENT_TEMPLATE(group_name, component, version, description, worker, db_path)
    if group_name not in VERSION_GROUPS:
        raise ValueError("Version group is not valid.")
    if component not in {"sw", "algo"}:
        raise ValueError("Version component is not valid.")
    if component == "algo" and not version_group_uses_algo(group_name):
        raise ValueError("This version group does not use Algo versions.")
    version_value = version.strip()
    if not version_value:
        raise ValueError("Version is required.")

    version_column = "sw_version" if component == "sw" else "algo_version"
    description_column = "sw_description" if component == "sw" else "algo_description"
    timestamp = now_text()
    matching_templates = [
        row for row in _sort_by_updated(_online_rows("version_templates"))
        if row["group_name"] == group_name and row.get(version_column) == version_value
    ]
    if matching_templates:
        for row in matching_templates:
            updates = {description_column: description, "worker": worker, "updated_at": timestamp}
            merged = {**row, **updates}
            updates["description"] = _combined_for_online_row(merged)
            _online_update("version_templates", int(row["id"]), updates)
    else:
        sw_description = description if component == "sw" else ""
        algo_description = description if component == "algo" else ""
        row_data = {
            "group_name": group_name,
            "sw_version": version_value if component == "sw" else "",
            "algo_version": version_value if component == "algo" else "",
            "description": combine_version_description(sw_description, algo_description, version_group_uses_algo(group_name)),
            "sw_description": sw_description,
            "algo_description": algo_description,
            "worker": worker,
            "created_at": timestamp,
            "updated_at": timestamp,
        }
        _online_append("version_templates", row_data)

    for row in list(_online_rows("version_history")):
        if row["group_name"] == group_name and row.get(version_column) == version_value:
            updates = {description_column: description, "worker": worker}
            merged = {**row, **updates}
            updates["description"] = _combined_for_online_row(merged)
            _online_update("version_history", int(row["id"]), updates)


def update_version_component_template(group_name: str, component: str, old_version: str, new_version: str, description: str, worker: str, db_path: Path = DB_PATH) -> None:
    if not _online_enabled(db_path):
        return _LOCAL_UPDATE_VERSION_COMPONENT_TEMPLATE(group_name, component, old_version, new_version, description, worker, db_path)
    if group_name not in VERSION_GROUPS:
        raise ValueError("Version group is not valid.")
    if component not in {"sw", "algo"}:
        raise ValueError("Version component is not valid.")
    if component == "algo" and not version_group_uses_algo(group_name):
        raise ValueError("This version group does not use Algo versions.")
    if not old_version.strip():
        raise ValueError("Select a version first.")
    if not new_version.strip():
        raise ValueError("Version is required.")
    version_column = "sw_version" if component == "sw" else "algo_version"
    description_column = "sw_description" if component == "sw" else "algo_description"
    timestamp = now_text()
    for table in ["version_templates", "version_history"]:
        for row in _online_rows(table):
            if row["group_name"] == group_name and row.get(version_column) == old_version.strip():
                updates = {version_column: new_version.strip(), description_column: description, "worker": worker}
                if table == "version_templates":
                    updates["updated_at"] = timestamp
                merged = {**row, **updates}
                updates["description"] = _combined_for_online_row(merged)
                _online_update(table, int(row["id"]), updates)


def delete_version_component_template(group_name: str, component: str, version: str, db_path: Path = DB_PATH) -> None:
    if not _online_enabled(db_path):
        return _LOCAL_DELETE_VERSION_COMPONENT_TEMPLATE(group_name, component, version, db_path)
    if group_name not in VERSION_GROUPS:
        raise ValueError("Version group is not valid.")
    if component not in {"sw", "algo"}:
        raise ValueError("Version component is not valid.")
    if not version.strip():
        return
    version_column = "sw_version" if component == "sw" else "algo_version"
    for table in ["version_templates", "version_history"]:
        for row in _online_rows(table):
            if row["group_name"] == group_name and row.get(version_column) == version.strip():
                _online_delete(table, int(row["id"]))


def update_version_template(template_id: int, sw_version: str, algo_version: str, description: str, worker: str, db_path: Path = DB_PATH) -> None:
    if not _online_enabled(db_path):
        return _LOCAL_UPDATE_VERSION_TEMPLATE(template_id, sw_version, algo_version, description, worker, db_path)
    template = _online_get("version_templates", template_id)
    if template is None:
        raise ValueError("Version template was not found.")
    if not sw_version.strip():
        raise ValueError("SW Version is required.")
    if version_group_uses_algo(template["group_name"]) and not algo_version.strip():
        raise ValueError("Algo Version is required.")
    sw_description, algo_description = split_version_description(description)
    if not version_group_uses_algo(template["group_name"]):
        algo_description = ""
    updates = {
        "sw_version": sw_version.strip(),
        "algo_version": algo_version.strip(),
        "description": description,
        "sw_description": sw_description,
        "algo_description": algo_description,
        "worker": worker,
        "updated_at": now_text(),
    }
    _online_update("version_templates", template_id, updates)
    for row in _online_rows("version_history"):
        if row["group_name"] == template["group_name"] and row["sw_version"] == template["sw_version"] and row["algo_version"] == template["algo_version"]:
            history_updates = dict(updates)
            history_updates.pop("updated_at", None)
            _online_update("version_history", int(row["id"]), history_updates)


def delete_version_template(template_id: int, db_path: Path = DB_PATH) -> None:
    if not _online_enabled(db_path):
        return _LOCAL_DELETE_VERSION_TEMPLATE(template_id, db_path)
    template = _online_get("version_templates", template_id)
    if template is None:
        return
    _online_delete("version_templates", template_id)
    for row in _online_rows("version_history"):
        if row["group_name"] == template["group_name"] and row["sw_version"] == template["sw_version"] and row["algo_version"] == template["algo_version"]:
            _online_delete("version_history", int(row["id"]))


def latest_version_by_instrument(db_path: Path = DB_PATH) -> dict[tuple[str, str], dict[str, Any]]:
    if not _online_enabled(db_path):
        return _LOCAL_LATEST_VERSION_BY_INSTRUMENT(db_path)
    latest: dict[tuple[str, str], dict[str, Any]] = {}
    for row in sorted(_online_rows("version_history"), key=lambda item: (item.get("update_time", ""), int(item.get("id", 0)))):
        latest[(row["line"], row["instrument"])] = row
    return latest


def latest_dashboard_versions(db_path: Path = DB_PATH) -> dict[tuple[str, str], dict[str, str]]:
    if not _online_enabled(db_path):
        return _LOCAL_LATEST_DASHBOARD_VERSIONS(db_path)
    states: dict[tuple[str, str], dict[str, str]] = {}
    for row in sorted(_online_rows("version_history"), key=lambda item: int(item.get("id", 0))):
        key = (row["line"], row["instrument"])
        state = states.setdefault(key, {
            "line": row["line"],
            "instrument": row["instrument"],
            "group_name": row["group_name"],
            "sw_version": "",
            "algo_version": "",
            "update_time": "",
            "sw_update_time": "",
            "algo_update_time": "",
        })
        sw_touched, algo_touched = version_history_component_flags(row)
        if sw_touched:
            state["sw_version"] = row["sw_version"] or ""
            state["sw_update_time"] = row["update_time"] or ""
            state["group_name"] = row["group_name"]
            state["update_time"] = row["update_time"] or state["update_time"]
        if instrument_uses_algo(row["instrument"]) and algo_touched:
            state["algo_version"] = row["algo_version"] or ""
            state["algo_update_time"] = row["update_time"] or ""
            state["group_name"] = row["group_name"]
            state["update_time"] = row["update_time"] or state["update_time"]
    return states


def version_history_rows(db_path: Path = DB_PATH) -> list[dict[str, Any]]:
    if not _online_enabled(db_path):
        return _LOCAL_VERSION_HISTORY_ROWS(db_path)
    return sorted(_online_rows("version_history"), key=lambda row: (row.get("update_time", ""), int(row.get("id", 0))), reverse=True)


def create_version_update(version: VersionInput, create_program_update_issue: bool = True, db_path: Path = DB_PATH) -> int:
    if not _online_enabled(db_path):
        return _LOCAL_CREATE_VERSION_UPDATE(version, create_program_update_issue, db_path)
    errors = validate_version_update(version)
    if errors:
        raise ValueError("\n".join(errors))
    uses_algo = version_group_uses_algo(version.group_name)
    sw_description = version.sw_description.strip()
    algo_description = version.algo_description.strip()
    if not sw_description and not algo_description:
        sw_description, algo_description = split_version_description(version.description)
    if not uses_algo:
        algo_description = ""
    has_entered_description = bool(sw_description or algo_description)
    sw_touched = bool(sw_description) or not has_entered_description or not uses_algo
    algo_touched = uses_algo and (bool(algo_description) or not has_entered_description)
    if sw_touched and not sw_description:
        sw_description = _online_version_component_description(version.group_name, "sw", version.sw_version)
    if algo_touched and not algo_description:
        algo_description = _online_version_component_description(version.group_name, "algo", version.algo_version)
    description = combine_version_description(sw_description, algo_description, uses_algo) or version.description
    save_version_template(
        version.group_name,
        version.sw_version,
        version.algo_version,
        description,
        version.worker,
        db_path,
        sw_description=sw_description,
        algo_description=algo_description,
    )
    created_issue_id: int | None = None
    if create_program_update_issue:
        version_text = f"SW {version.sw_version}"
        if instrument_uses_algo(version.instrument):
            version_text = f"{version_text} / Algo {version.algo_version}"
        issue = IssueInput(
            issue_time=version.update_time,
            resolved_time="00:00",
            line=version.line,
            instrument=version.instrument,
            worker=version.worker,
            category="Software",
            subcategory="Program Update",
            title=f"Program Update - {version.line} {version.instrument} {version_text}",
            description=description,
            status="Monitoring",
        )
        created_issue_id = create_issue(issue, db_path)
    created = _online_append("version_history", {
        "created_at": now_text(),
        "update_time": version.update_time,
        "group_name": version.group_name,
        "line": version.line,
        "instrument": version.instrument,
        "sw_version": version.sw_version,
        "algo_version": version.algo_version,
        "description": description,
        "sw_description": sw_description,
        "algo_description": algo_description,
        "sw_touched": 1 if sw_touched else 0,
        "algo_touched": 1 if algo_touched else 0,
        "worker": version.worker,
        "created_issue_id": created_issue_id or "",
    })
    return int(created["id"])

# Online 1.1 cache overlay. This intentionally sits after the first online
# adapter so existing high-level functions reuse these optimized low-level calls.
import uuid as _uuid

def _online_cache_path() -> Path:
    base = _os.environ.get("LOCALAPPDATA") or _os.environ.get("APPDATA")
    if base:
        return Path(base) / "VisionIssueTracker" / "online_cache.db"
    return APP_DIR / "data" / "online_cache.db"


_ONLINE_CACHE_PATH = _online_cache_path()
_ONLINE_SCHEMA_VERSION = "online_1_2"

for _header in ["updated_at", "deleted_at", "client_request_id"]:
    if _header not in _ISSUE_HEADERS:
        _ISSUE_HEADERS.append(_header)
for _header in ["deleted_at", "client_request_id"]:
    if _header not in _TEMPLATE_HEADERS:
        _TEMPLATE_HEADERS.append(_header)
for _header in ["updated_at", "deleted_at", "client_request_id"]:
    if _header not in _HISTORY_HEADERS:
        _HISTORY_HEADERS.append(_header)
for _header in ["updated_at", "deleted_at", "client_request_id"]:
    if _header not in _DL_TRAINED_HEADERS:
        _DL_TRAINED_HEADERS.append(_header)
for _header in ["updated_at", "deleted_at", "client_request_id"]:
    if _header not in _DL_APPLICATION_HEADERS:
        _DL_APPLICATION_HEADERS.append(_header)

_TABLE_HEADERS = {
    "issues": _ISSUE_HEADERS,
    "version_templates": _TEMPLATE_HEADERS,
    "version_history": _HISTORY_HEADERS,
    "dl_trained_models": _DL_TRAINED_HEADERS,
    "dl_model_applications": _DL_APPLICATION_HEADERS,
}


def online_mode(db_path: Path = DB_PATH) -> bool:
    return _online_enabled(db_path)


def _to_online_int(value: Any) -> int:
    text = str(value).strip()
    if text in {"", "None", "none", "null"}:
        return 0
    try:
        return int(text)
    except ValueError:
        return int(float(text))


def _normalize_online_row(table: str, row: dict[str, Any] | None) -> dict[str, Any] | None:
    if row is None:
        return None
    normalized: dict[str, Any] = {}
    for header in _TABLE_HEADERS[table]:
        value = row.get(header, "")
        if header in _NUMERIC_FIELDS:
            normalized[header] = _to_online_int(value)
        else:
            normalized[header] = "" if value is None else str(value)
    return normalized


def _online_cache_connect() -> sqlite3.Connection:
    _ONLINE_CACHE_PATH.parent.mkdir(parents=True, exist_ok=True)
    conn = sqlite3.connect(_ONLINE_CACHE_PATH)
    conn.row_factory = sqlite3.Row
    conn.execute(
        """
        CREATE TABLE IF NOT EXISTS online_rows (
            table_name TEXT NOT NULL,
            row_id INTEGER NOT NULL,
            row_json TEXT NOT NULL,
            updated_at TEXT NOT NULL DEFAULT '',
            deleted_at TEXT NOT NULL DEFAULT '',
            PRIMARY KEY (table_name, row_id)
        )
        """
    )
    conn.execute(
        """
        CREATE TABLE IF NOT EXISTS online_state (
            key TEXT PRIMARY KEY,
            value TEXT NOT NULL
        )
        """
    )
    conn.execute(
        "INSERT OR REPLACE INTO online_state(key, value) VALUES('schema_version', ?)",
        (_ONLINE_SCHEMA_VERSION,),
    )
    conn.commit()
    return conn


def _online_cache_init() -> None:
    with closing(_online_cache_connect()):
        pass


def _online_cache_get_state(key: str) -> str:
    with closing(_online_cache_connect()) as conn:
        row = conn.execute("SELECT value FROM online_state WHERE key = ?", (key,)).fetchone()
        return str(row["value"]) if row else ""


def _online_cache_set_state(key: str, value: str) -> None:
    with closing(_online_cache_connect()) as conn:
        conn.execute("INSERT OR REPLACE INTO online_state(key, value) VALUES(?, ?)", (key, value))
        conn.commit()


def _online_cache_rows(table: str, include_deleted: bool = False) -> list[dict[str, Any]]:
    with closing(_online_cache_connect()) as conn:
        query = "SELECT row_json FROM online_rows WHERE table_name = ?"
        if not include_deleted:
            query += " AND deleted_at = ''"
        rows = [
            _normalize_online_row(table, _json.loads(row["row_json"]))
            for row in conn.execute(query, (table,))
        ]
    return sorted([row for row in rows if row is not None], key=lambda row: int(row.get("id", 0)))


def _online_cache_get(table: str, row_id: int) -> dict[str, Any] | None:
    with closing(_online_cache_connect()) as conn:
        row = conn.execute(
            "SELECT row_json FROM online_rows WHERE table_name = ? AND row_id = ? AND deleted_at = ''",
            (table, row_id),
        ).fetchone()
        if not row:
            return None
        return _normalize_online_row(table, _json.loads(row["row_json"]))


def _online_cache_apply_rows(table: str, rows: list[dict[str, Any]]) -> None:
    if not rows:
        return
    with closing(_online_cache_connect()) as conn:
        for raw in rows:
            row = _normalize_online_row(table, raw)
            if row is None:
                continue
            row_id = int(row.get("id", 0))
            if row_id <= 0:
                continue
            if row.get("deleted_at", ""):
                conn.execute("DELETE FROM online_rows WHERE table_name = ? AND row_id = ?", (table, row_id))
                continue
            conn.execute(
                """
                INSERT OR REPLACE INTO online_rows(table_name, row_id, row_json, updated_at, deleted_at)
                VALUES(?, ?, ?, ?, ?)
                """,
                (
                    table,
                    row_id,
                    _json.dumps(row, ensure_ascii=False),
                    row.get("updated_at", ""),
                    row.get("deleted_at", ""),
                ),
            )
        conn.commit()


def _online_cache_replace_rows(table: str, rows: list[dict[str, Any]]) -> None:
    with closing(_online_cache_connect()) as conn:
        conn.execute("DELETE FROM online_rows WHERE table_name = ?", (table,))
        conn.commit()
    _online_cache_apply_rows(table, rows)


def _online_cache_delete(table: str, row_id: int) -> None:
    with closing(_online_cache_connect()) as conn:
        conn.execute("DELETE FROM online_rows WHERE table_name = ? AND row_id = ?", (table, row_id))
        conn.commit()


def _safe_online_cache_apply_rows(table: str, rows: list[dict[str, Any]]) -> None:
    try:
        _online_cache_apply_rows(table, rows)
    except Exception:
        # The online write already succeeded. Cache failures should not make a saved issue look failed.
        pass


def _safe_online_cache_delete(table: str, row_id: int) -> None:
    try:
        _online_cache_delete(table, row_id)
    except Exception:
        pass


def _normalize_row_list(table: str, rows: list[dict[str, Any]]) -> list[dict[str, Any]]:
    normalized = []
    for row in rows or []:
        item = _normalize_online_row(table, row)
        if item is not None:
            normalized.append(item)
    return normalized


def _apply_sync_result(result: dict[str, Any], replace: bool) -> None:
    tables = result.get("tables", {})
    for table in _TABLE_HEADERS:
        rows = _normalize_row_list(table, tables.get(table, []))
        if replace:
            _online_cache_replace_rows(table, rows)
        else:
            _online_cache_apply_rows(table, rows)
    server_time = str(result.get("server_time") or now_text())
    _online_cache_set_state("last_sync_at", server_time)
    if replace:
        _online_cache_set_state("last_full_sync_at", server_time)


def _online_legacy_full_refresh() -> None:
    for table in _TABLE_HEADERS:
        try:
            rows = _normalize_row_list(table, _online_request("list", table=table).get("rows", []))
        except ValueError as exc:
            if "Invalid table" in str(exc):
                continue
            raise
        _online_cache_replace_rows(table, rows)
    timestamp = now_text()
    _online_cache_set_state("last_sync_at", timestamp)
    _online_cache_set_state("last_full_sync_at", timestamp)


def refresh_online_cache(force_full: bool = False) -> None:
    if not _online_enabled():
        return
    _online_cache_init()
    last_sync = _online_cache_get_state("last_sync_at")
    try:
        if force_full or not last_sync:
            _apply_sync_result(_online_request("bootstrap"), replace=True)
        else:
            _apply_sync_result(_online_request("changesSince", since=last_sync), replace=False)
    except ValueError as exc:
        if "Unknown action" in str(exc):
            _online_legacy_full_refresh()
            return
        raise


def initialize_database(db_path: Path = DB_PATH) -> None:
    if not _online_enabled(db_path):
        return _LOCAL_INITIALIZE_DATABASE(db_path)
    _online_cache_init()


def _online_rows(table: str) -> list[dict[str, Any]]:
    return _online_cache_rows(table)


def _online_get(table: str, row_id: int) -> dict[str, Any] | None:
    row = _online_cache_get(table, row_id)
    if row is not None:
        return row
    row = _normalize_online_row(table, _online_request("get", table=table, id=row_id).get("row"))
    if row is not None:
        _safe_online_cache_apply_rows(table, [row])
    return row


def _with_client_request_id(row: dict[str, Any]) -> dict[str, Any]:
    prepared = dict(row)
    prepared.setdefault("client_request_id", str(_uuid.uuid4()))
    return prepared


def _online_append(table: str, row: dict[str, Any]) -> dict[str, Any]:
    created = _normalize_online_row(
        table,
        _online_request("append", table=table, row=_with_client_request_id(row)).get("row"),
    )
    if created is None:
        raise ValueError("Online API did not return a created row.")
    _safe_online_cache_apply_rows(table, [created])
    return created


def _online_bulk_append(table: str, rows: list[dict[str, Any]]) -> list[dict[str, Any]]:
    created = _normalize_row_list(
        table,
        _online_request("bulkAppend", table=table, rows=[_with_client_request_id(row) for row in rows]).get("rows", []),
    )
    _safe_online_cache_apply_rows(table, created)
    return created


def _online_update(table: str, row_id: int, row: dict[str, Any]) -> dict[str, Any]:
    updated = _normalize_online_row(
        table,
        _online_request("update", table=table, id=row_id, row=row).get("row"),
    )
    if updated is None:
        raise ValueError("Online API did not return an updated row.")
    _safe_online_cache_apply_rows(table, [updated])
    return updated


def _online_delete(table: str, row_id: int) -> None:
    deleted = _normalize_online_row(table, _online_request("delete", table=table, id=row_id).get("row"))
    if deleted is not None:
        _safe_online_cache_apply_rows(table, [deleted])
    else:
        _safe_online_cache_delete(table, row_id)


def search_issues(filters: dict[str, str] | None = None, db_path: Path = DB_PATH) -> list[dict[str, Any]]:
    if not _online_enabled(db_path):
        return _LOCAL_SEARCH_ISSUES(filters, db_path)
    filters = filters or {}
    rows = _online_rows("issues")
    return _sorted_issues([row for row in rows if _issue_matches_filters(row, filters)])


def create_dl_trained_model(
    model: DeepLearningTrainedInput,
    create_action_issue: bool = True,
    db_path: Path = DB_PATH,
) -> int:
    if not _online_enabled(db_path):
        return _LOCAL_CREATE_DL_TRAINED_MODEL(model, create_action_issue, db_path)
    errors = validate_dl_trained_model(model)
    if errors:
        raise ValueError("\n".join(errors))
    targets = tuple(target for target in DL_MACHINE_KEYS if target in set(model.target_machines))
    target_text = serialize_dl_targets(targets)
    scope = infer_dl_scope(targets)
    created_issue_id = 0
    if create_action_issue:
        issue_ids = create_dl_issues_for_targets(
            model.trained_time,
            model.model_family,
            model.model_version,
            model.change_type,
            targets,
            scope,
            model.description,
            model.worker,
            "Action Required",
            "Deep Learning Model Ready",
            db_path,
        )
        created_issue_id = issue_ids[0] if issue_ids else 0
    created = _online_append(
        "dl_trained_models",
        {
            "created_at": now_text(),
            "trained_time": model.trained_time,
            "model_family": model.model_family,
            "model_version": model.model_version,
            "change_type": model.change_type,
            "target_machines": target_text,
            "scope": scope,
            "description": model.description,
            "worker": model.worker,
            "status": model.status,
            "created_issue_id": created_issue_id or "",
        },
    )
    return int(created["id"])


def create_dl_model_application(
    model: DeepLearningApplicationInput,
    create_monitoring_issue: bool = True,
    trained_model_id: int | None = None,
    db_path: Path = DB_PATH,
) -> list[int]:
    if not _online_enabled(db_path):
        return _LOCAL_CREATE_DL_MODEL_APPLICATION(model, create_monitoring_issue, trained_model_id, db_path)
    errors = validate_dl_application(model)
    if errors:
        raise ValueError("\n".join(errors))
    targets = tuple(target for target in DL_MACHINE_KEYS if target in set(model.target_machines))
    scope = infer_dl_scope(targets)
    created_issue_ids: list[int] = []
    if create_monitoring_issue:
        created_issue_ids = create_dl_issues_for_targets(
            model.applied_time,
            model.model_family,
            model.model_version,
            model.change_type,
            targets,
            scope,
            model.description,
            model.worker,
            "Monitoring",
            "Deep Learning Model Applied",
            db_path,
        )
    created_rows = _online_bulk_append(
        "dl_model_applications",
        [
            {
                "created_at": now_text(),
                "applied_time": model.applied_time,
                "model_family": model.model_family,
                "model_version": model.model_version,
                "change_type": model.change_type,
                "line": DL_MACHINE_BY_KEY[target_key]["line"],
                "polarity": DL_MACHINE_BY_KEY[target_key]["polarity"],
                "instrument": DL_MACHINE_BY_KEY[target_key]["instrument"],
                "machine": target_key,
                "scope": scope,
                "description": model.description,
                "worker": model.worker,
                "created_issue_id": created_issue_ids[0] if created_issue_ids else "",
            }
            for target_key in targets
        ],
    )
    if trained_model_id:
        mark_dl_trained_model_status(trained_model_id, "Applied", db_path)
    return [int(row["id"]) for row in created_rows]


def list_dl_trained_models(
    model_family: str | None = None,
    status: str | None = None,
    db_path: Path = DB_PATH,
) -> list[dict[str, Any]]:
    if not _online_enabled(db_path):
        return _LOCAL_LIST_DL_TRAINED_MODELS(model_family, status, db_path)
    rows = _online_rows("dl_trained_models")
    if model_family:
        rows = [row for row in rows if row.get("model_family") == model_family]
    if status:
        rows = [row for row in rows if row.get("status") == status]
    return sorted(rows, key=lambda row: (row.get("trained_time", ""), int(row.get("id", 0))), reverse=True)


def dl_application_rows(model_family: str | None = None, db_path: Path = DB_PATH) -> list[dict[str, Any]]:
    if not _online_enabled(db_path):
        return _LOCAL_DL_APPLICATION_ROWS(model_family, db_path)
    rows = _online_rows("dl_model_applications")
    if model_family:
        rows = [row for row in rows if row.get("model_family") == model_family]
    return sorted(rows, key=lambda row: (row.get("applied_time", ""), int(row.get("id", 0))), reverse=True)


def latest_dl_applied_models(
    model_family: str | None = None,
    db_path: Path = DB_PATH,
) -> dict[tuple[str, str, str], dict[str, Any]]:
    if not _online_enabled(db_path):
        return _LOCAL_LATEST_DL_APPLIED_MODELS(model_family, db_path)
    latest: dict[tuple[str, str, str], dict[str, Any]] = {}
    rows = sorted(
        dl_application_rows(model_family, db_path),
        key=lambda row: (row.get("applied_time", ""), int(row.get("id", 0))),
    )
    for row in rows:
        latest[(row["model_family"], row["line"], row["instrument"])] = row
    return latest


def mark_dl_trained_model_status(
    trained_model_id: int,
    status: str,
    db_path: Path = DB_PATH,
) -> None:
    if not _online_enabled(db_path):
        return _LOCAL_MARK_DL_TRAINED_MODEL_STATUS(trained_model_id, status, db_path)
    if status not in DL_TRAINED_STATUS_OPTIONS:
        raise ValueError("Trained model status is not valid.")
    _online_update("dl_trained_models", trained_model_id, {"status": status})


def export_deep_learning_dashboard_to_excel(output_path: Path, db_path: Path = DB_PATH) -> None:
    if not _online_enabled(db_path):
        return _LOCAL_EXPORT_DEEP_LEARNING_DASHBOARD_TO_EXCEL(output_path, db_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    latest = latest_dl_applied_models(db_path=db_path)
    trained_rows = list_dl_trained_models(db_path=db_path)
    workbook = Workbook()
    applied_sheet = workbook.active
    applied_sheet.title = "Applied Models"
    trained_sheet = workbook.create_sheet("Trained Models")
    applied_headers = [
        "Model Family",
        "Line",
        "Polarity",
        "Instrument",
        "Model Version",
        "Scope",
        "Time",
        "Logged By",
        "Description",
    ]
    trained_headers = [
        "Status",
        "Model Family",
        "Model Version",
        "Change Type",
        "Scope",
        "Target Machines",
        "Time",
        "Logged By",
        "Description",
    ]
    applied_sheet.append(applied_headers)
    trained_sheet.append(trained_headers)
    for family in DL_MODEL_FAMILIES:
        for target in DL_MACHINE_TARGETS:
            row = latest.get((family, target["line"], target["instrument"]))
            applied_sheet.append(
                [
                    family,
                    target["line"],
                    target["polarity"],
                    target["instrument"],
                    row["model_version"] if row else "",
                    row["scope"] if row else "",
                    row["applied_time"] if row else "",
                    row["worker"] if row else "",
                    row["description"] if row else "",
                ]
            )
    for row in trained_rows:
        trained_sheet.append(
            [
                row["status"],
                row["model_family"],
                row["model_version"],
                row["change_type"],
                row["scope"],
                row["target_machines"],
                row["trained_time"],
                row["worker"],
                row["description"],
            ]
        )
    for sheet in [applied_sheet, trained_sheet]:
        for cell in sheet[1]:
            cell.font = Font(bold=True, color="FFFFFF")
            cell.fill = PatternFill("solid", fgColor="1F4E78")
            cell.alignment = Alignment(vertical="top")
        for column_index, header in enumerate([cell.value for cell in sheet[1]], start=1):
            column_letter = get_column_letter(column_index)
            max_length = len(str(header or ""))
            for cell in sheet[column_letter]:
                max_length = max(max_length, len(str(cell.value or "")))
                cell.alignment = Alignment(wrap_text=header == "Description", vertical="top")
            if header == "Description":
                sheet.column_dimensions[column_letter].width = 55
            else:
                sheet.column_dimensions[column_letter].width = min(max_length + 2, 28)
        sheet.freeze_panes = "A2"
    workbook.save(output_path)
