from pathlib import Path
from tempfile import TemporaryDirectory

from openpyxl import load_workbook

from vision_tracker import (
    IssueInput,
    VersionInput,
    active_issues,
    create_version_update,
    create_issue,
    dashboard_counts,
    delete_issue,
    export_issues_to_excel,
    export_version_dashboard_to_excel,
    initialize_database,
    issue_time_bounds,
    latest_version_by_instrument,
    recent_version_templates,
    resolve_issue,
    search_issues,
    set_issue_status,
    update_issue,
    version_history_rows,
)


def run_tests() -> None:
    with TemporaryDirectory() as temp_dir:
        db_path = Path(temp_dir) / "test.db"
        initialize_database(db_path)

        issue_id = create_issue(
            IssueInput(
                issue_time="2026-06-17 08:10",
                line="1-1",
                instrument="Pinhole",
                worker="Hojun Kwak",
                category="Hardware",
                subcategory="Camera",
                title="Camera disconnect during inspection",
                description="Camera stopped responding during production.",
            ),
            db_path,
        )
        assert issue_id == 1

        rows = search_issues({"status": "Action Required", "category": "Hardware"}, db_path)
        assert len(rows) == 1
        assert rows[0]["title"] == "Camera disconnect during inspection"
        active = active_issues(db_path)
        assert len(active) == 1
        counts = dashboard_counts(db_path)
        assert counts["Action Required"] == 1
        assert counts["Active"] == 1

        update_issue(
            issue_id,
            IssueInput(
                issue_time="2026-06-17 08:10",
                resolved_time="2026-06-17 08:42",
                line="1-1",
                instrument="Pinhole",
                worker="Hojun Kwak",
                category="Hardware",
                subcategory="Camera",
                title="Camera disconnect during inspection",
                description="Camera stopped responding during production.",
                status="Resolved",
                resolution_notes="Reconnected camera cable and restarted program.",
            ),
            db_path,
        )

        resolved = search_issues({"status": "Resolved"}, db_path)
        assert len(resolved) == 1

        export_path = Path(temp_dir) / "report.xlsx"
        export_issues_to_excel(resolved, export_path)
        workbook = load_workbook(export_path)
        sheet = workbook.active
        assert sheet["B2"].value == "2026-06-17 08:10"
        assert sheet["A2"].value == 1
        assert sheet["I2"].value == "Camera disconnect during inspection"
        assert sheet["C1"].value == "Downtime Duration"
        assert sheet["F1"].value == "Logged By"

        issue_id_2 = create_issue(
            IssueInput(
                issue_time="2026-06-17 09:00",
                line="1-2",
                instrument="Lead",
                worker="Kijung Kim",
                category="Camera Grab Fail",
                subcategory="",
                title="Grab timeout",
                description="Camera failed to grab during cycle.",
                status="Monitoring",
            ),
            db_path,
        )
        resolve_issue(issue_id_2, db_path=db_path)
        grab_fail = search_issues({"category": "Camera Grab Fail"}, db_path)
        assert len(grab_fail) == 1
        assert grab_fail[0]["status"] == "Resolved"

        set_issue_status(issue_id, "Monitoring", db_path)
        monitoring = search_issues({"status": "Monitoring"}, db_path)
        assert len(monitoring) == 1

        issue_id_3 = create_issue(
            IssueInput(
                issue_time="2026-06-17 10:00",
                line="2-1",
                instrument="Sealing",
                worker="Jihoon Yun",
                category="Recipe",
                subcategory="Overkill",
                title="Overkill trend",
                description="Reject rate increased after recipe change.",
            ),
            db_path,
        )
        delete_issue(issue_id_3, db_path)
        deleted = search_issues({"keyword": "Overkill trend"}, db_path)
        assert len(deleted) == 0

        early_id = create_issue(
            IssueInput(
                issue_time="2026-06-17 07:30",
                line="2-2",
                instrument="Welding(+)",
                worker="Jisub Yun",
                category="Production",
                subcategory="",
                title="Early production note",
                description="Created after later records but should sort first by issue time.",
            ),
            db_path,
        )
        ordered = search_issues({}, db_path)
        assert ordered[0]["id"] == early_id
        first_time, latest_time = issue_time_bounds(db_path)
        assert first_time == "2026-06-17 07:30"
        assert latest_time == "2026-06-17 09:00"

        plc_id = create_issue(
            IssueInput(
                issue_time="2026-06-17 11:00",
                line="2-2",
                instrument="Lead",
                worker="Yun Jihoon",
                category="Software",
                subcategory="PLC",
                title="PLC communication check",
                description="PLC communication issue classification test.",
            ),
            db_path,
        )
        plc_rows = search_issues({"category": "Software", "subcategory": "PLC"}, db_path)
        assert len(plc_rows) == 1
        assert plc_rows[0]["id"] == plc_id

        bypass_id = create_issue(
            IssueInput(
                issue_time="2026-06-17 12:00",
                line="2-2",
                instrument="Welding(-)",
                worker="Yun Jihoon",
                category="Recipe",
                subcategory="Bypass/Unbypass",
                title="Bypass setting check",
                description="Bypass/Unbypass classification test.",
            ),
            db_path,
        )
        bypass_rows = search_issues({"category": "Recipe", "subcategory": "Bypass/Unbypass"}, db_path)
        assert len(bypass_rows) == 1
        assert bypass_rows[0]["id"] == bypass_id

        multi_id = create_issue(
            IssueInput(
                issue_time="2026-06-17 13:00",
                line="1-1",
                instrument="Welding(+) / Welding(-)",
                worker="Yun Jihoon",
                category="Recipe",
                subcategory="Add Measure",
                title="Both welding visions updated",
                description="Multiple vision selection test.",
            ),
            db_path,
        )
        plus_rows = search_issues({"instrument": "Welding(+)"}, db_path)
        assert any(row["id"] == multi_id for row in plus_rows)
        minus_rows = search_issues({"instrument": "Welding(-)"}, db_path)
        assert any(row["id"] == multi_id for row in minus_rows)
        multi_filter_rows = search_issues({"instrument": "Lead / Welding(-)"}, db_path)
        assert any(row["id"] == multi_id for row in multi_filter_rows)

        create_version_update(
            VersionInput(
                update_time="2026-06-17 14:00",
                group_name="Welding",
                line="1-1",
                instrument="Welding(+)",
                sw_version="SW-1.0.0",
                algo_version="ALG-2.0.0",
                description="Initial welding plus version record.",
                worker="Jihoon Yun",
            ),
            True,
            db_path,
        )
        create_version_update(
            VersionInput(
                update_time="2026-06-17 14:05",
                group_name="Welding",
                line="1-1",
                instrument="Welding(-)",
                sw_version="SW-1.0.0",
                algo_version="ALG-2.0.0",
                description="Initial welding minus version record.",
                worker="Jihoon Yun",
            ),
            True,
            db_path,
        )
        create_version_update(
            VersionInput(
                update_time="2026-06-17 15:00",
                group_name="Welding",
                line="1-1",
                instrument="Welding(+)",
                sw_version="SW-1.1.0",
                algo_version="ALG-2.1.0",
                description="Updated plus vision only.",
                worker="Jihoon Yun",
            ),
            True,
            db_path,
        )
        latest_versions = latest_version_by_instrument(db_path)
        assert latest_versions[("1-1", "Welding(+)")]["sw_version"] == "SW-1.1.0"
        assert latest_versions[("1-1", "Welding(-)")]["sw_version"] == "SW-1.0.0"
        templates = recent_version_templates("Welding", 3, db_path)
        assert len(templates) == 2
        assert templates[0]["sw_version"] == "SW-1.1.0"
        history = version_history_rows(db_path)
        assert history[0]["instrument"] == "Welding(+)"
        update_issues = search_issues({"category": "Software", "subcategory": "Program Update"}, db_path)
        assert len(update_issues) == 3
        assert all(row["status"] == "Monitoring" for row in update_issues)

        version_export_path = Path(temp_dir) / "version_dashboard.xlsx"
        export_version_dashboard_to_excel(version_export_path, db_path)
        version_workbook = load_workbook(version_export_path)
        version_sheet = version_workbook.active
        assert version_sheet.title == "Version Dashboard"
        assert version_sheet["A1"].value == "Line"
        assert version_sheet.max_row == 1 + 4 * 7
        exported_rows = list(version_sheet.iter_rows(min_row=2, values_only=True))
        welding_plus = [row for row in exported_rows if row[0] == "1-1" and row[1] == "Welding(+)"][0]
        welding_minus = [row for row in exported_rows if row[0] == "1-1" and row[1] == "Welding(-)"][0]
        assert welding_plus[3] == "SW-1.1.0"
        assert welding_minus[3] == "SW-1.0.0"


if __name__ == "__main__":
    run_tests()
    print("All tests passed.")
