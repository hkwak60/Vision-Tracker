from __future__ import annotations

import calendar
import os
import sys
from datetime import datetime, timedelta
from pathlib import Path


if getattr(sys, "frozen", False):
    tcl_root = Path(getattr(sys, "_MEIPASS", Path(sys.executable).parent)) / "tcl"
    os.environ.setdefault("TCL_LIBRARY", str(tcl_root / "tcl8.6"))
    os.environ.setdefault("TK_LIBRARY", str(tcl_root / "tk8.6"))

import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from vision_tracker import (
    ACTIVE_STATUS_OPTIONS,
    APP_TITLE,
    CATEGORIES,
    CATEGORY_MAP,
    INSTRUMENTS,
    INSTRUMENT_GROUP,
    LINES,
    STATUS_OPTIONS,
    VERSION_GROUPS,
    WORKERS,
    IssueInput,
    VersionInput,
    active_issues,
    create_version_update,
    create_issue,
    delete_issue,
    delete_version_component_template,
    export_issues_to_excel,
    export_version_dashboard_to_excel,
    format_instruments,
    get_issue,
    initialize_database,
    instrument_uses_algo,
    issue_time_bounds,
    latest_dashboard_versions,
    latest_version_by_instrument,
    now_text,
    resolve_issue,
    search_issues,
    set_issue_status,
    split_instruments,
    update_issue,
    update_version_component_template,
    version_component_templates,
    version_group_uses_algo,
    version_sort_key,
)


STATUS_TAGS = {
    "Action Required": ("#fff2f0", "#9f1239"),
    "Monitoring": ("#fff8db", "#854d0e"),
    "Resolved": ("#ecfdf3", "#166534"),
}

LANGUAGES = ["English", "한국어"]
VERSION_TAB_SPACER = " " * 80
TRANSLATIONS = {
    "한국어": {
        "Current Worker": "작업자",
        "Language": "언어",
        "Issue Board": "이슈 보드",
        "Search / Report": "검색 / 보고서",
        "Version History": "버전 기록",
        "Action Required": "조치 필요",
        "Monitoring": "모니터링",
        "Resolved": "해결됨",
        "Resolved Today": "오늘 해결",
        "Active": "진행 중",
        "Refresh": "새로고침",
        "Edit": "수정",
        "Delete": "삭제",
        "Create Issue": "이슈 등록",
        "Edit Issue": "이슈 수정",
        "Save Changes": "변경 저장",
        "Cancel": "취소",
        "Move to Action Required": "조치 필요로 이동",
        "Move to Monitoring": "모니터링으로 이동",
        "No Issue Selected": "선택된 이슈 없음",
        "Select a card to view details, or create a new issue.": "카드를 선택해 상세를 확인하거나 새 이슈를 등록하세요.",
        "Selected Issue": "선택된 이슈",
        "Title": "제목",
        "Status": "상태",
        "Line / Instrument": "라인 / 비전",
        "Category": "분류",
        "Issue Time": "발생 시간",
        "Logged By": "작성자",
        "Downtime Duration": "다운타임",
        "Description": "설명",
        "Issue Record": "이슈 기록",
        "Line": "라인",
        "Vision": "비전",
        "Subcategory": "세부 분류",
        "Resolution Notes": "조치 내용",
        "New": "새 기록",
        "Save": "저장",
        "ID": "번호",
        "Instrument": "비전",
        "Worker": "작업자",
        "Keyword": "키워드",
        "From": "시작",
        "To": "종료",
        "Today": "오늘",
        "This Week": "이번 주",
        "Camera Grab Fail": "카메라 Grab 실패",
        "Recipe Issues": "레시피 이슈",
        "Clear": "초기화",
        "Search": "검색",
        "Excel": "엑셀",
        "Export Dashboard": "대시보드 추출",
        "SW behind": "SW 낮음",
        "Algo behind": "Algo 낮음",
        "Updated 7d": "최근 7일 업데이트",
        "Vision Filter": "비전 필터",
        "Version Dashboard": "버전 대시보드",
        "Version Update": "버전 업데이트",
        "Version Group": "버전 그룹",
        "Version Template": "버전 템플릿",
        "SW Version": "SW 버전",
        "Algo Version": "Algo 버전",
        "Update Time": "업데이트 시간",
        "Save Version Update": "버전 업데이트 저장",
        "Save Version": "버전 저장",
        "Delete Version": "버전 삭제",
        "Create Monitoring Issue": "모니터링 이슈 등록",
        "Version Description": "버전 설명",
        "SW Description": "SW 설명",
        "Algo Description": "Algo 설명",
        "Save SW": "SW 저장",
        "Save Algo": "Algo 저장",
        "Delete SW": "SW 삭제",
        "Delete Algo": "Algo 삭제",
        "Group": "그룹",
    }
}


class VisionIssueApp(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        initialize_database()
        self.title(APP_TITLE)
        self.geometry("1180x740")
        self.minsize(980, 620)
        self.selected_issue_id: int | None = None
        self.loaded_issue_worker = ""
        self.search_rows = []
        self.active_issue_rows = []
        self.language_var = tk.StringVar(value="한국어")
        self.current_worker_var = tk.StringVar(value=WORKERS[0])
        self.translated_widgets: list[tuple[tk.Widget, str, str]] = []

        self.configure(bg="#f4f6f8")
        self.style = ttk.Style(self)
        self.style.theme_use("clam")
        self.configure_styles()

        self.build_layout()
        self.initialize_issue_form_state()
        self.refresh_open_issues()
        self.search_records()

    def configure_styles(self) -> None:
        self.style.configure("TNotebook", background="#f4f6f8", borderwidth=0)
        self.style.configure("TNotebook.Tab", padding=(18, 10), font=("Segoe UI", 10, "bold"))
        self.style.map("TNotebook.Tab", background=[("disabled", "#ffffff")], foreground=[("disabled", "#ffffff")])
        self.style.configure("TFrame", background="#f4f6f8")
        self.style.configure("Panel.TFrame", background="#ffffff", relief="solid", borderwidth=1)
        self.style.configure("TLabel", background="#f4f6f8", font=("Segoe UI", 10))
        self.style.configure("Panel.TLabel", background="#ffffff", font=("Segoe UI", 10))
        self.style.configure("Header.TLabel", background="#f4f6f8", font=("Segoe UI", 18, "bold"))
        self.style.configure("Subheader.TLabel", background="#ffffff", font=("Segoe UI", 12, "bold"))
        self.style.configure("TButton", font=("Segoe UI", 10), padding=(12, 7))
        self.style.configure("Accent.TButton", font=("Segoe UI", 10, "bold"), padding=(12, 7))
        self.style.configure("Treeview", font=("Segoe UI", 9), rowheight=28)
        self.style.configure("Treeview.Heading", font=("Segoe UI", 9, "bold"))
        self.style.configure("Card.TFrame", background="#ffffff", relief="solid", borderwidth=1)
        self.style.configure("CardTitle.TLabel", background="#ffffff", font=("Segoe UI", 9))
        self.style.configure("CardValue.TLabel", background="#ffffff", font=("Segoe UI", 20, "bold"))

    def initialize_issue_form_state(self) -> None:
        self.issue_date_var = tk.StringVar()
        self.issue_hour_var = tk.StringVar()
        self.issue_minute_var = tk.StringVar()
        self.resolved_time_var = tk.StringVar(value="00:00")
        self.line_var = tk.StringVar(value=LINES[0])
        self.instrument_var = tk.StringVar(value=INSTRUMENTS[0])
        self.selected_instruments = {INSTRUMENTS[0]}
        self.category_var = tk.StringVar(value=CATEGORIES[0])
        self.subcategory_var = tk.StringVar(value=CATEGORY_MAP[CATEGORIES[0]][0])
        self.status_var = tk.StringVar(value=ACTIVE_STATUS_OPTIONS[0])
        self.title_var = tk.StringVar()
        self.form_description_value = ""
        self.form_resolution_value = ""
        self.line_instrument_traces_added = False
        self.set_issue_datetime(now_text())

    def build_layout(self) -> None:
        header = ttk.Frame(self, padding=(18, 14, 18, 8))
        header.pack(fill="x")
        ttk.Label(header, text=APP_TITLE, style="Header.TLabel").pack(side="left")
        profile = ttk.Frame(header)
        profile.pack(side="right")
        self.tr_label(profile, "Language").pack(side="left", padx=(0, 8))
        language_combo = ttk.Combobox(
            profile,
            textvariable=self.language_var,
            values=LANGUAGES,
            state="readonly",
            width=10,
        )
        language_combo.pack(side="left", padx=(0, 18))
        language_combo.bind("<<ComboboxSelected>>", self.apply_language)

        self.notebook = ttk.Notebook(self)
        self.notebook.pack(fill="both", expand=True, padx=18, pady=(0, 18))

        self.open_tab = ttk.Frame(self.notebook, padding=14)
        self.search_tab = ttk.Frame(self.notebook, padding=14)
        self.version_spacer_tab = ttk.Frame(self.notebook)
        self.version_tab = ttk.Frame(self.notebook, padding=14)
        self.notebook.add(self.open_tab, text="Issue Board")
        self.notebook.add(self.search_tab, text="Search / Report")
        self.notebook.add(self.version_spacer_tab, text=VERSION_TAB_SPACER, state="disabled")
        self.notebook.add(self.version_tab, text="Version History")

        self.build_open_tab()
        self.build_search_tab()
        self.build_version_tab()
        self.apply_language()

    def text(self, key: str) -> str:
        return TRANSLATIONS.get(self.language_var.get(), {}).get(key, key)

    def register_text(self, widget: tk.Widget, key: str, option: str = "text") -> tk.Widget:
        self.translated_widgets.append((widget, key, option))
        widget.configure(**{option: self.text(key)})
        return widget

    def tr_label(self, parent: tk.Widget, key: str, **kwargs) -> ttk.Label:
        label = ttk.Label(parent, **kwargs)
        return self.register_text(label, key)

    def tr_button(self, parent: tk.Widget, key: str, command, prefix: str = "", **kwargs) -> ttk.Button:
        button = ttk.Button(parent, command=command, **kwargs)
        self.translated_widgets.append((button, key, "text"))
        button.configure(text=f"{prefix}{self.text(key)}")
        button.translation_prefix = prefix
        return button

    def apply_language(self, _event: tk.Event | None = None) -> None:
        active_widgets: list[tuple[tk.Widget, str, str]] = []
        for widget, key, option in self.translated_widgets:
            try:
                if not widget.winfo_exists():
                    continue
                prefix = getattr(widget, "translation_prefix", "")
                widget.configure(**{option: f"{prefix}{self.text(key)}"})
                active_widgets.append((widget, key, option))
            except tk.TclError:
                continue
        self.translated_widgets = active_widgets
        if hasattr(self, "notebook"):
            self.notebook.tab(self.open_tab, text=self.text("Issue Board"))
            self.notebook.tab(self.search_tab, text=self.text("Search / Report"))
            self.notebook.tab(self.version_spacer_tab, text=VERSION_TAB_SPACER)
            self.notebook.tab(self.version_tab, text=self.text("Version History"))
        for tree_name in ["search_tree"]:
            if hasattr(self, tree_name):
                self.update_tree_headings(getattr(self, tree_name))
        if hasattr(self, "version_create_issue_button"):
            self.refresh_create_issue_button()

    def update_tree_headings(self, tree: ttk.Treeview) -> None:
        headings = {
            "id": "ID",
            "issue_time": "Issue Time",
            "line": "Line",
            "instrument": "Instrument",
            "category": "Category",
            "subcategory": "Subcategory",
            "title": "Title",
            "status": "Status",
            "worker": "Logged By",
        }
        for column, key in headings.items():
            tree.heading(column, text=self.text(key))

    def build_open_tab(self) -> None:
        toolbar = ttk.Frame(self.open_tab)
        toolbar.pack(fill="x", pady=(0, 10))
        self.tr_button(toolbar, "Refresh", self.refresh_open_issues, prefix="↻ ").pack(side="left")
        self.tr_button(toolbar, "Create Issue", self.show_create_issue_form, prefix="+ ", style="Accent.TButton").pack(side="right")

        content = ttk.Frame(self.open_tab)
        content.pack(fill="both", expand=True)
        content.columnconfigure(0, weight=3)
        content.columnconfigure(1, weight=2)
        content.rowconfigure(0, weight=1)

        board = ttk.Frame(content)
        board.grid(row=0, column=0, sticky="nsew", padx=(0, 12))
        board.columnconfigure(0, weight=1)
        board.columnconfigure(1, weight=1)
        board.rowconfigure(0, weight=1)

        self.board_columns: dict[str, ttk.Frame] = {}
        self.board_column_canvases: dict[str, tk.Canvas] = {}
        self.board_count_vars: dict[str, tk.StringVar] = {}
        self.issue_card_widgets: dict[int, tuple[tk.Frame, list[tk.Widget]]] = {}
        for column_index, status in enumerate(ACTIVE_STATUS_OPTIONS):
            column = ttk.Frame(board, style="Panel.TFrame", padding=10)
            column.grid(row=0, column=column_index, sticky="nsew", padx=(0, 10) if column_index == 0 else (0, 0))
            column.columnconfigure(0, weight=1)
            column.rowconfigure(1, weight=1)

            header = ttk.Frame(column, style="Panel.TFrame")
            header.grid(row=0, column=0, sticky="ew", pady=(0, 8))
            self.tr_label(header, status, style="Subheader.TLabel").pack(side="left")
            count_var = tk.StringVar(value="0")
            ttk.Label(header, textvariable=count_var, style="Panel.TLabel").pack(side="right")
            self.board_count_vars[status] = count_var

            canvas = tk.Canvas(column, background="#ffffff", highlightthickness=0)
            canvas.grid(row=1, column=0, sticky="nsew")
            scrollbar = ttk.Scrollbar(column, orient="vertical", command=canvas.yview)
            scrollbar.grid(row=1, column=1, sticky="ns")
            canvas.configure(yscrollcommand=scrollbar.set)
            card_frame = ttk.Frame(canvas, style="Panel.TFrame")
            window = canvas.create_window((0, 0), window=card_frame, anchor="nw")
            card_frame.bind("<Configure>", lambda _event, target=canvas: target.configure(scrollregion=target.bbox("all")))
            canvas.bind("<Configure>", lambda event, target=canvas, item=window: target.itemconfigure(item, width=event.width))
            canvas.bind("<MouseWheel>", lambda event, target=canvas: target.yview_scroll(int(-event.delta / 60), "units"))
            self.board_columns[status] = card_frame
            self.board_column_canvases[status] = canvas

        side_panel = ttk.Frame(content, style="Panel.TFrame")
        side_panel.grid(row=0, column=1, sticky="nsew")
        side_panel.columnconfigure(0, weight=1)
        side_panel.rowconfigure(0, weight=1)
        self.board_side_canvas = tk.Canvas(side_panel, background="#ffffff", highlightthickness=0)
        self.board_side_canvas.grid(row=0, column=0, sticky="nsew")
        side_scroll = ttk.Scrollbar(side_panel, orient="vertical", command=self.board_side_canvas.yview)
        side_scroll.grid(row=0, column=1, sticky="ns")
        self.board_side_canvas.configure(yscrollcommand=side_scroll.set)
        self.board_side_frame = ttk.Frame(self.board_side_canvas, style="Panel.TFrame", padding=14)
        side_window = self.board_side_canvas.create_window((0, 0), window=self.board_side_frame, anchor="nw")
        self.board_side_frame.bind(
            "<Configure>",
            lambda _event: self.board_side_canvas.configure(scrollregion=self.board_side_canvas.bbox("all")),
        )
        self.board_side_canvas.bind(
            "<Configure>",
            lambda event: self.board_side_canvas.itemconfigure(side_window, width=event.width),
        )
        self.board_side_canvas.bind(
            "<MouseWheel>",
            lambda event: self.board_side_canvas.yview_scroll(int(-event.delta / 60), "units"),
        )
        self.show_empty_issue_detail()

    def bind_mousewheel_recursive(self, widget: tk.Widget, canvas: tk.Canvas) -> None:
        def on_mousewheel(event: tk.Event) -> str:
            canvas.yview_scroll(int(-event.delta / 60), "units")
            return "break"

        widget.bind("<MouseWheel>", on_mousewheel)
        for child in widget.winfo_children():
            self.bind_mousewheel_recursive(child, canvas)

    def clear_board_side_panel(self) -> None:
        for child in self.board_side_frame.winfo_children():
            child.destroy()
        self.board_side_canvas.yview_moveto(0)

    def show_empty_issue_detail(self) -> None:
        self.clear_board_side_panel()
        self.selected_issue_id = None
        self.tr_label(self.board_side_frame, "No Issue Selected", style="Subheader.TLabel").pack(anchor="w", pady=(0, 10))
        ttk.Label(
            self.board_side_frame,
            text=self.text("Select a card to view details, or create a new issue."),
            style="Panel.TLabel",
            wraplength=420,
            justify="left",
        ).pack(anchor="w", fill="x")
        self.bind_mousewheel_recursive(self.board_side_frame, self.board_side_canvas)
        self.refresh_board_card_styles()

    def show_issue_detail(self, issue_id: int) -> None:
        row = get_issue(issue_id)
        if row is None:
            messagebox.showerror(APP_TITLE, "Issue was not found.")
            self.show_empty_issue_detail()
            return
        self.selected_issue_id = issue_id
        self.clear_board_side_panel()
        self.board_side_frame.columnconfigure(0, weight=1)

        header = ttk.Frame(self.board_side_frame, style="Panel.TFrame")
        header.pack(fill="x", pady=(0, 10))
        ttk.Label(
            header,
            text=row["title"] or "-",
            style="Subheader.TLabel",
            wraplength=420,
            justify="left",
        ).pack(side="left", fill="x", expand=True, anchor="w")

        status_bg, status_fg = STATUS_TAGS.get(row["status"], ("#eef2ff", "#312e81"))
        tk.Label(
            header,
            text=self.text(row["status"]),
            bg=status_bg,
            fg=status_fg,
            padx=8,
            pady=3,
            font=("Segoe UI", 9, "bold"),
        ).pack(side="right", padx=(8, 0))

        category_text = row["category"]
        if row["subcategory"]:
            category_text = f"{category_text} / {row['subcategory']}"
        detail_rows = [
            ("Line / Instrument", f"{row['line']} / {row['instrument']}"),
            ("Category", category_text),
            ("Issue Time", row["issue_time"] or "-"),
            ("Downtime Duration", row["resolved_time"] or "-"),
            ("Logged By", row["worker"] or "-"),
        ]
        for label, value in detail_rows:
            self.add_detail_row(self.board_side_frame, label, value)

        self.tr_label(self.board_side_frame, "Description", style="Panel.TLabel").pack(anchor="w", pady=(14, 2))
        description = tk.Text(self.board_side_frame, height=8, wrap="word", font=("Segoe UI", 9), relief="solid", bd=1)
        description.pack(fill="x")
        description.insert("1.0", row["description"] or "")
        description.configure(state="disabled")

        self.tr_label(self.board_side_frame, "Resolution Notes", style="Panel.TLabel").pack(anchor="w", pady=(14, 2))
        resolution = tk.Text(self.board_side_frame, height=5, wrap="word", font=("Segoe UI", 9), relief="solid", bd=1)
        resolution.pack(fill="x")
        resolution.insert("1.0", row["resolution_notes"] or "")
        resolution.configure(state="disabled")

        actions = ttk.Frame(self.board_side_frame, style="Panel.TFrame")
        actions.pack(fill="x", pady=(14, 0))
        self.tr_button(actions, "Edit", lambda issue=issue_id: self.show_edit_issue_form(issue), prefix="✎ ").pack(fill="x", pady=(0, 6))
        if row["status"] == "Action Required":
            self.tr_button(
                actions,
                "Move to Monitoring",
                lambda issue=issue_id: self.move_issue_status(issue, "Monitoring"),
                prefix="→ ",
            ).pack(fill="x", pady=(0, 6))
        elif row["status"] == "Monitoring":
            self.tr_button(
                actions,
                "Move to Action Required",
                lambda issue=issue_id: self.move_issue_status(issue, "Action Required"),
                prefix="→ ",
            ).pack(fill="x", pady=(0, 6))
        self.tr_button(actions, "Resolved", lambda issue=issue_id: self.resolve_issue_from_board(issue), prefix="✓ ").pack(fill="x", pady=(0, 6))
        self.tr_button(actions, "Delete", lambda issue=issue_id: self.delete_issue_by_id(issue), prefix="✕ ").pack(fill="x")

        self.bind_mousewheel_recursive(self.board_side_frame, self.board_side_canvas)
        self.refresh_board_card_styles()

    def add_detail_row(self, parent: ttk.Frame, label: str, value: str) -> None:
        self.tr_label(parent, label, style="Panel.TLabel").pack(anchor="w", pady=(8, 0))
        ttk.Label(parent, text=value or "-", style="Panel.TLabel", wraplength=420, justify="left").pack(anchor="w", fill="x")

    def show_create_issue_form(self) -> None:
        self.clear_form()
        self.render_issue_form("create")

    def show_edit_issue_form(self, issue_id: int | None = None) -> None:
        issue_id = issue_id or self.selected_issue_id
        if issue_id is None:
            messagebox.showwarning(APP_TITLE, "Select an issue first.")
            return
        if not self.load_issue_state(issue_id):
            return
        self.render_issue_form("edit")

    def render_issue_form(self, mode: str) -> None:
        self.clear_board_side_panel()
        form_title = "Create Issue" if mode == "create" else "Edit Issue"
        save_title = "Create Issue" if mode == "create" else "Save Changes"
        panel = self.board_side_frame
        panel.columnconfigure(1, weight=1)
        self.tr_label(panel, form_title, style="Subheader.TLabel").grid(row=0, column=0, columnspan=2, sticky="w", pady=(0, 12))

        self.add_line_instrument_grid(panel, 1)
        self.add_datetime_picker(panel, "Issue Time", 2, 0)
        self.add_labeled_combo(panel, "Logged By", self.current_worker_var, WORKERS, 3, 0)
        category_combo = self.add_labeled_combo(panel, "Category", self.category_var, CATEGORIES, 4, 0)
        category_combo.bind("<<ComboboxSelected>>", lambda _event: self.update_subcategories())
        self.subcategory_combo = self.add_labeled_combo(panel, "Subcategory", self.subcategory_var, CATEGORY_MAP[self.category_var.get()], 5, 0)
        self.add_labeled_combo(panel, "Status", self.status_var, STATUS_OPTIONS, 6, 0)
        self.add_labeled_entry(panel, "Downtime Duration", self.resolved_time_var, 7, 0)
        self.add_labeled_entry(panel, "Title", self.title_var, 8, 0)

        self.tr_label(panel, "Description", style="Panel.TLabel").grid(row=9, column=0, columnspan=2, sticky="w", pady=(10, 2))
        self.description_text = tk.Text(panel, height=7, wrap="word", font=("Segoe UI", 10))
        self.description_text.grid(row=10, column=0, columnspan=2, sticky="nsew")
        self.description_text.insert("1.0", self.form_description_value)

        self.tr_label(panel, "Resolution Notes", style="Panel.TLabel").grid(row=11, column=0, columnspan=2, sticky="w", pady=(10, 2))
        self.resolution_text = tk.Text(panel, height=5, wrap="word", font=("Segoe UI", 10))
        self.resolution_text.grid(row=12, column=0, columnspan=2, sticky="nsew")
        self.resolution_text.insert("1.0", self.form_resolution_value)

        actions = ttk.Frame(panel, style="Panel.TFrame")
        actions.grid(row=13, column=0, columnspan=2, sticky="ew", pady=(14, 0))
        self.tr_button(actions, "Cancel", self.cancel_issue_form).pack(side="left")
        self.tr_button(actions, save_title, self.save_issue, prefix="✓ ", style="Accent.TButton").pack(side="right")
        self.bind_mousewheel_recursive(self.board_side_frame, self.board_side_canvas)

    def cancel_issue_form(self) -> None:
        if self.selected_issue_id is None:
            self.show_empty_issue_detail()
        else:
            self.show_issue_detail(self.selected_issue_id)

    def add_line_instrument_grid(self, parent: ttk.Frame, row: int) -> None:
        grid = ttk.Frame(parent, style="Panel.TFrame")
        grid.grid(row=row, column=0, columnspan=2, sticky="ew", pady=(0, 12))
        grid.columnconfigure(0, weight=1)
        self.line_buttons: dict[str, tk.Button] = {}
        self.instrument_buttons: dict[str, tk.Button] = {}

        self.tr_label(grid, "Line", style="Panel.TLabel").grid(row=0, column=0, sticky="w", pady=(0, 3))
        line_frame = ttk.Frame(grid, style="Panel.TFrame")
        line_frame.grid(row=1, column=0, sticky="ew", pady=(0, 8))
        for column_index, line in enumerate(LINES):
            button = tk.Button(
                line_frame,
                text=line,
                width=8,
                relief="raised",
                command=lambda selected_line=line: self.select_line(selected_line),
            )
            button.grid(row=0, column=column_index, sticky="ew", padx=2, pady=3)
            line_frame.columnconfigure(column_index, weight=1)
            self.line_buttons[line] = button

        self.tr_label(grid, "Vision", style="Panel.TLabel").grid(row=2, column=0, sticky="w", pady=(0, 3))
        vision_frame = ttk.Frame(grid, style="Panel.TFrame")
        vision_frame.grid(row=3, column=0, sticky="ew")
        for index, instrument in enumerate(INSTRUMENTS):
            button = tk.Button(
                vision_frame,
                text=instrument,
                width=13,
                relief="raised",
                command=lambda selected_instrument=instrument: self.select_instrument(selected_instrument),
            )
            button.grid(row=index // 2, column=index % 2, sticky="ew", padx=2, pady=3)
            vision_frame.columnconfigure(index % 2, weight=1)
            self.instrument_buttons[instrument] = button
        if not getattr(self, "line_instrument_traces_added", False):
            self.line_var.trace_add("write", lambda *_args: self.refresh_line_instrument_buttons())
            self.instrument_var.trace_add("write", lambda *_args: self.refresh_line_instrument_buttons())
            self.line_instrument_traces_added = True
        self.refresh_line_instrument_buttons()

    def select_line(self, line: str) -> None:
        self.line_var.set(line)

    def select_instrument(self, instrument: str) -> None:
        if instrument in self.selected_instruments and len(self.selected_instruments) > 1:
            self.selected_instruments.remove(instrument)
        else:
            self.selected_instruments.add(instrument)
        self.instrument_var.set(format_instruments(self.selected_instruments))

    def refresh_line_instrument_buttons(self) -> None:
        selected_bg = "#1f6feb"
        selected_fg = "#ffffff"
        default_bg = self.cget("bg")
        default_fg = "#111827"
        for line, button in getattr(self, "line_buttons", {}).items():
            is_selected = line == self.line_var.get()
            button.configure(
                background=selected_bg if is_selected else default_bg,
                foreground=selected_fg if is_selected else default_fg,
                relief="sunken" if is_selected else "raised",
            )
        current_instruments = set(split_instruments(self.instrument_var.get()))
        self.selected_instruments = current_instruments or {INSTRUMENTS[0]}
        if not current_instruments:
            self.instrument_var.set(format_instruments(self.selected_instruments))
            return
        for instrument, button in getattr(self, "instrument_buttons", {}).items():
            is_selected = instrument in current_instruments
            button.configure(
                background=selected_bg if is_selected else default_bg,
                foreground=selected_fg if is_selected else default_fg,
                relief="sunken" if is_selected else "raised",
            )

    def add_filter_instrument_buttons(self, parent: ttk.Frame, row: int) -> None:
        frame = ttk.Frame(parent, style="Panel.TFrame")
        frame.grid(row=row, column=0, columnspan=6, sticky="w", pady=(8, 0))
        self.tr_label(frame, "Vision Filter", style="Panel.TLabel").pack(side="left", padx=(4, 8))
        self.filter_instrument_buttons: dict[str, tk.Button] = {}
        for instrument in INSTRUMENTS:
            button = tk.Button(
                frame,
                text=instrument,
                width=12,
                relief="raised",
                command=lambda selected_instrument=instrument: self.toggle_filter_instrument(selected_instrument),
            )
            button.pack(side="left", padx=(0, 4))
            self.filter_instrument_buttons[instrument] = button

    def toggle_filter_instrument(self, instrument: str) -> None:
        if instrument in self.filter_instruments:
            self.filter_instruments.remove(instrument)
        else:
            self.filter_instruments.add(instrument)
        self.filter_instrument.set(format_instruments(self.filter_instruments))
        self.refresh_filter_instrument_buttons()
        self.search_records()

    def refresh_filter_instrument_buttons(self) -> None:
        selected_bg = "#1f6feb"
        selected_fg = "#ffffff"
        default_bg = self.cget("bg")
        default_fg = "#111827"
        for instrument, button in getattr(self, "filter_instrument_buttons", {}).items():
            is_selected = instrument in self.filter_instruments
            button.configure(
                background=selected_bg if is_selected else default_bg,
                foreground=selected_fg if is_selected else default_fg,
                relief="sunken" if is_selected else "raised",
            )

    def build_search_tab(self) -> None:
        filters = ttk.Frame(self.search_tab, style="Panel.TFrame", padding=14)
        filters.pack(fill="x", pady=(0, 10))

        self.filter_status = tk.StringVar()
        self.filter_line = tk.StringVar()
        self.filter_instrument = tk.StringVar()
        self.filter_instruments: set[str] = set()
        self.filter_category = tk.StringVar()
        self.filter_subcategory = tk.StringVar()
        self.filter_keyword = tk.StringVar()
        self.filter_from = tk.StringVar()
        self.filter_to = tk.StringVar()
        self.reset_search_date_bounds()

        self.add_filter_combo(filters, "Status", self.filter_status, [""] + STATUS_OPTIONS, 0, 0)
        self.add_filter_combo(filters, "Line", self.filter_line, [""] + LINES, 0, 2)
        category_filter = self.add_filter_combo(filters, "Category", self.filter_category, [""] + CATEGORIES, 1, 0)
        category_filter.bind("<<ComboboxSelected>>", lambda _event: self.update_filter_subcategories())
        self.filter_subcategory_combo = self.add_filter_combo(filters, "Subcategory", self.filter_subcategory, [""], 1, 2)
        self.add_filter_entry(filters, "Keyword", self.filter_keyword, 1, 4)
        self.add_filter_entry(filters, "From", self.filter_from, 2, 0)
        self.add_filter_entry(filters, "To", self.filter_to, 2, 2)

        self.add_filter_instrument_buttons(filters, 3)

        quick_filters = ttk.Frame(filters, style="Panel.TFrame")
        quick_filters.grid(row=4, column=0, columnspan=6, sticky="w", pady=(8, 0))
        self.tr_button(quick_filters, "Today", lambda: self.apply_quick_filter("today")).pack(side="left", padx=(0, 6))
        self.tr_button(quick_filters, "This Week", lambda: self.apply_quick_filter("week")).pack(side="left", padx=(0, 6))
        self.tr_button(quick_filters, "Action Required", lambda: self.apply_quick_filter("action")).pack(side="left", padx=(0, 6))
        self.tr_button(quick_filters, "Monitoring", lambda: self.apply_quick_filter("monitoring")).pack(side="left", padx=(0, 6))
        self.tr_button(quick_filters, "Camera Grab Fail", lambda: self.apply_quick_filter("camera_grab")).pack(side="left", padx=(0, 6))
        self.tr_button(quick_filters, "Recipe Issues", lambda: self.apply_quick_filter("recipe")).pack(side="left", padx=(0, 6))
        self.tr_button(quick_filters, "Clear", self.clear_search_filters).pack(side="left")

        buttons = ttk.Frame(filters, style="Panel.TFrame")
        buttons.grid(row=2, column=4, columnspan=2, sticky="e", padx=6, pady=6)
        self.tr_button(buttons, "Search", self.search_records, prefix="⌕ ", style="Accent.TButton").pack(side="left", padx=(0, 8))
        self.tr_button(buttons, "Excel", self.export_search_results, prefix="⇩ ").pack(side="left")
        self.tr_button(buttons, "Delete", lambda: self.delete_selected_issue(self.search_tree), prefix="✕ ").pack(side="left", padx=(8, 0))

        search_table = ttk.Frame(self.search_tab)
        search_table.pack(fill="both", expand=True)
        search_table.columnconfigure(0, weight=1)
        search_table.rowconfigure(0, weight=1)
        self.search_tree = self.make_issue_tree(search_table)
        self.search_tree.grid(row=0, column=0, sticky="nsew")
        search_scroll = ttk.Scrollbar(search_table, orient="vertical", command=self.search_tree.yview)
        search_scroll.grid(row=0, column=1, sticky="ns")
        search_xscroll = ttk.Scrollbar(search_table, orient="horizontal", command=self.search_tree.xview)
        search_xscroll.grid(row=1, column=0, sticky="ew")
        self.search_tree.configure(yscrollcommand=search_scroll.set, xscrollcommand=search_xscroll.set)
        self.search_tree.bind("<Double-1>", lambda _event: self.load_selected_search_issue())

    def add_version_legend_item(self, parent: ttk.Frame, color: str, label_key: str) -> None:
        dot = tk.Canvas(parent, width=10, height=10, bg="#ffffff", highlightthickness=0)
        dot.create_oval(2, 2, 8, 8, fill=color, outline=color)
        dot.pack(side="left", padx=(10, 4))
        self.tr_label(parent, label_key, style="Panel.TLabel").pack(side="left")

    def create_version_dashboard_cell(
        self,
        parent: ttk.Frame,
        row: int,
        column: int,
        line: str,
        instrument: str,
    ) -> dict[str, tk.Widget]:
        shell = tk.Frame(parent, bg="#d8dee8", padx=1, pady=1)
        shell.grid(row=row, column=column, sticky="nsew", padx=2, pady=2)
        shell.columnconfigure(0, weight=1)

        cell = tk.Frame(shell, bg="#ffffff")
        cell.pack(fill="both", expand=True)
        recent_bar = tk.Frame(cell, width=3, bg="#ffffff")
        recent_bar.pack(side="left", fill="y")
        body = tk.Frame(cell, bg="#ffffff", padx=5, pady=4)
        body.pack(side="left", fill="both", expand=True)
        body.columnconfigure(0, weight=1)

        sw_label = tk.Label(
            body,
            text="-",
            bg="#ffffff",
            fg="#111827",
            font=("Segoe UI", 8),
            anchor="w",
        )
        sw_label.grid(row=0, column=0, sticky="ew")
        sw_dot = tk.Canvas(body, width=9, height=9, bg="#ffffff", highlightthickness=0)
        sw_dot.grid(row=0, column=1, sticky="e", padx=(4, 0))

        algo_label = tk.Label(
            body,
            text="",
            bg="#ffffff",
            fg="#111827",
            font=("Segoe UI", 8),
            anchor="w",
        )
        algo_label.grid(row=1, column=0, sticky="ew", pady=(1, 0))
        algo_dot = tk.Canvas(body, width=9, height=9, bg="#ffffff", highlightthickness=0)
        algo_dot.grid(row=1, column=1, sticky="e", padx=(4, 0), pady=(1, 0))

        widgets: dict[str, tk.Widget] = {
            "shell": shell,
            "cell": cell,
            "recent_bar": recent_bar,
            "body": body,
            "sw_label": sw_label,
            "sw_dot": sw_dot,
            "algo_label": algo_label,
            "algo_dot": algo_dot,
        }
        self.bind_version_cell_click(widgets, line, instrument)
        return widgets

    def bind_version_cell_click(self, widgets: dict[str, tk.Widget], line: str, instrument: str) -> None:
        for widget in widgets.values():
            widget.bind(
                "<Button-1>",
                lambda _event, selected_line=line, selected_instrument=instrument: self.select_version_target(
                    selected_line,
                    selected_instrument,
                ),
            )

    def draw_status_dot(self, canvas: tk.Canvas, color: str | None) -> None:
        canvas.delete("all")
        canvas.configure(bg="#ffffff")
        if color:
            canvas.create_oval(2, 2, 8, 8, fill=color, outline=color)

    def normalized_version_key(self, value: str, width: int) -> tuple[int, ...] | None:
        key = version_sort_key(value)
        if key is None:
            return None
        return key + (0,) * max(0, width - len(key))

    def latest_current_version_keys(self, rows: dict[tuple[str, str], object], instrument: str) -> tuple[tuple[int, ...] | None, tuple[int, ...] | None]:
        sw_keys = []
        algo_keys = []
        for line in LINES:
            row = rows.get((line, instrument))
            if not row:
                continue
            sw_key = version_sort_key(row["sw_version"])
            if sw_key is not None:
                sw_keys.append(sw_key)
            if instrument_uses_algo(instrument):
                algo_key = version_sort_key(row["algo_version"])
                if algo_key is not None:
                    algo_keys.append(algo_key)

        sw_width = max((len(key) for key in sw_keys), default=0)
        algo_width = max((len(key) for key in algo_keys), default=0)
        max_sw = max((key + (0,) * (sw_width - len(key)) for key in sw_keys), default=None)
        max_algo = max((key + (0,) * (algo_width - len(key)) for key in algo_keys), default=None)
        return max_sw, max_algo

    def build_version_tab(self) -> None:
        content = ttk.Frame(self.version_tab)
        content.pack(fill="both", expand=True)
        content.columnconfigure(0, weight=3)
        content.columnconfigure(1, weight=2)
        content.rowconfigure(1, weight=1)

        dashboard_panel = ttk.Frame(content, style="Panel.TFrame", padding=12)
        dashboard_panel.grid(row=0, column=0, columnspan=2, sticky="ew", pady=(0, 10))
        dashboard_header = ttk.Frame(dashboard_panel, style="Panel.TFrame")
        dashboard_header.pack(fill="x", pady=(0, 8))
        self.tr_label(dashboard_header, "Version Dashboard", style="Subheader.TLabel").pack(side="left")
        self.tr_button(dashboard_header, "Export Dashboard", self.export_version_dashboard, prefix="⇩ ").pack(side="right")
        self.tr_button(dashboard_header, "Refresh", self.refresh_version_history, prefix="? ", width=10).pack(side="right", padx=(0, 6))

        legend = ttk.Frame(dashboard_header, style="Panel.TFrame")
        legend.pack(side="right", padx=(0, 14))
        self.add_version_legend_item(legend, "#f59e0b", "SW behind")
        self.add_version_legend_item(legend, "#2563eb", "Algo behind")
        tk.Frame(legend, width=16, height=4, bg="#22c55e").pack(side="left", padx=(10, 4))
        self.tr_label(legend, "Updated 7d", style="Panel.TLabel").pack(side="left")

        self.version_dashboard = ttk.Frame(dashboard_panel, style="Panel.TFrame")
        self.version_dashboard.pack(fill="x")
        self.version_cells: dict[tuple[str, str], dict[str, tk.Widget]] = {}
        ttk.Label(self.version_dashboard, text="", style="Panel.TLabel", width=7).grid(row=0, column=0, sticky="nsew")
        for column_index, instrument in enumerate(INSTRUMENTS, start=1):
            ttk.Label(
                self.version_dashboard,
                text=instrument,
                style="Panel.TLabel",
                anchor="center",
                font=("Segoe UI", 8, "bold"),
            ).grid(row=0, column=column_index, sticky="ew", padx=2, pady=(0, 4))
            self.version_dashboard.columnconfigure(column_index, weight=1, uniform="vision_version")
        for row_index, line in enumerate(LINES, start=1):
            ttk.Label(
                self.version_dashboard,
                text=line,
                style="Panel.TLabel",
                anchor="center",
                font=("Segoe UI", 9, "bold"),
            ).grid(row=row_index, column=0, sticky="nsew", padx=(0, 4), pady=2)
            for column_index, instrument in enumerate(INSTRUMENTS, start=1):
                self.version_cells[(line, instrument)] = self.create_version_dashboard_cell(
                    self.version_dashboard,
                    row_index,
                    column_index,
                    line,
                    instrument,
                )

        editor_panel = ttk.Frame(content, style="Panel.TFrame")
        editor_panel.grid(row=1, column=0, sticky="nsew", padx=(0, 10))
        editor_panel.columnconfigure(0, weight=1)
        editor_panel.rowconfigure(0, weight=1)
        editor_canvas = tk.Canvas(editor_panel, background="#ffffff", highlightthickness=0)
        editor_canvas.grid(row=0, column=0, sticky="nsew")
        editor = ttk.Frame(editor_canvas, style="Panel.TFrame", padding=14)
        editor_window = editor_canvas.create_window((0, 0), window=editor, anchor="nw")
        editor.bind("<Configure>", lambda _event: editor_canvas.configure(scrollregion=editor_canvas.bbox("all")))
        editor_canvas.bind("<Configure>", lambda event: editor_canvas.itemconfigure(editor_window, width=event.width))
        editor_canvas.bind("<MouseWheel>", lambda event: editor_canvas.yview_scroll(int(-event.delta / 60), "units"))
        editor.columnconfigure(1, weight=1)
        editor.columnconfigure(3, weight=1)
        editor.rowconfigure(5, weight=1)
        self.tr_label(editor, "Version Update", style="Subheader.TLabel").grid(row=0, column=0, columnspan=4, sticky="w", pady=(0, 10))

        self.version_group_var = tk.StringVar(value=list(VERSION_GROUPS.keys())[0])
        self.version_update_time_var = tk.StringVar(value=now_text())
        self.version_sw_var = tk.StringVar()
        self.version_algo_var = tk.StringVar()
        self.version_create_issue_var = tk.BooleanVar(value=True)
        self.version_selected_lines = {LINES[0]}
        self.version_selected_instruments = {VERSION_GROUPS[self.version_group_var.get()][0]}

        target_frame = ttk.Frame(editor, style="Panel.TFrame")
        target_frame.grid(row=1, column=0, columnspan=4, sticky="ew", pady=(0, 10))
        self.version_line_buttons: dict[str, tk.Button] = {}
        self.version_instrument_buttons: dict[str, tk.Button] = {}
        self.tr_label(target_frame, "Line", style="Panel.TLabel").grid(row=0, column=0, sticky="w", padx=(0, 8), pady=3)
        for column_index, line in enumerate(LINES, start=1):
            button = tk.Button(target_frame, text=line, width=10, command=lambda selected_line=line: self.toggle_version_line(selected_line))
            button.grid(row=0, column=column_index, padx=2, pady=3)
            self.version_line_buttons[line] = button
        self.tr_label(target_frame, "Vision", style="Panel.TLabel").grid(row=1, column=0, sticky="w", padx=(0, 8), pady=3)
        for index, instrument in enumerate(INSTRUMENTS):
            button = tk.Button(target_frame, text=instrument, width=15, command=lambda selected_instrument=instrument: self.toggle_version_instrument(selected_instrument))
            button.grid(row=1 + index // 4, column=1 + index % 4, sticky="ew", padx=2, pady=3)
            target_frame.columnconfigure(1 + index % 4, weight=1)
            self.version_instrument_buttons[instrument] = button

        self.add_labeled_entry(editor, "SW Version", self.version_sw_var, 2, 0)
        self.version_algo_entry = self.add_labeled_entry(editor, "Algo Version", self.version_algo_var, 2, 2)
        self.add_labeled_entry(editor, "Update Time", self.version_update_time_var, 3, 0)
        self.add_labeled_combo(editor, "Logged By", self.current_worker_var, WORKERS, 3, 2)
        self.version_create_issue_button = tk.Button(
            editor,
            anchor="w",
            relief="raised",
            command=self.toggle_create_issue_option,
        )
        self.version_create_issue_button.grid(row=4, column=0, columnspan=4, sticky="w", pady=7)
        self.refresh_create_issue_button()

        self.version_update_description_label = self.tr_label(editor, "Description", style="Panel.TLabel")
        self.version_update_description_label.grid(row=5, column=0, sticky="nw", pady=7)
        self.version_single_description_text = tk.Text(editor, height=5, wrap="word", font=("Segoe UI", 10))
        self.version_single_description_text.grid(row=5, column=1, columnspan=3, sticky="nsew", pady=7)
        self.version_single_description_text.bind("<MouseWheel>", lambda event: editor_canvas.yview_scroll(int(-event.delta / 60), "units"))

        self.version_split_description_frame = ttk.Frame(editor, style="Panel.TFrame")
        self.version_split_description_frame.grid(row=5, column=1, columnspan=3, sticky="nsew", pady=7)
        self.version_split_description_frame.columnconfigure(0, weight=1)
        self.version_split_description_frame.columnconfigure(1, weight=0)
        self.version_split_description_frame.columnconfigure(2, weight=1)
        self.version_split_description_frame.rowconfigure(1, weight=1)
        ttk.Label(self.version_split_description_frame, text="SW", style="Panel.TLabel", font=("Segoe UI", 9, "bold")).grid(row=0, column=0, sticky="w", pady=(0, 3))
        ttk.Label(self.version_split_description_frame, text="Algo", style="Panel.TLabel", font=("Segoe UI", 9, "bold")).grid(row=0, column=2, sticky="w", padx=(10, 0), pady=(0, 3))
        self.version_sw_description_text = tk.Text(self.version_split_description_frame, height=5, wrap="word", font=("Segoe UI", 10))
        self.version_sw_description_text.grid(row=1, column=0, sticky="nsew", padx=(0, 5))
        divider = tk.Frame(self.version_split_description_frame, width=1, bg="#d1d5db")
        divider.grid(row=1, column=1, sticky="ns")
        self.version_algo_description_text = tk.Text(self.version_split_description_frame, height=5, wrap="word", font=("Segoe UI", 10))
        self.version_algo_description_text.grid(row=1, column=2, sticky="nsew", padx=(10, 0))
        self.version_sw_description_text.bind("<MouseWheel>", lambda event: editor_canvas.yview_scroll(int(-event.delta / 60), "units"))
        self.version_algo_description_text.bind("<MouseWheel>", lambda event: editor_canvas.yview_scroll(int(-event.delta / 60), "units"))

        actions = ttk.Frame(editor, style="Panel.TFrame")
        actions.grid(row=6, column=0, columnspan=4, sticky="ew", pady=(10, 0))
        self.tr_button(actions, "Refresh", self.refresh_version_history, prefix="↻ ").pack(side="left")
        self.tr_button(actions, "Save Version Update", self.save_version_updates, prefix="✓ ", style="Accent.TButton").pack(side="right")

        description_container = ttk.Frame(content, style="Panel.TFrame")
        description_container.grid(row=1, column=1, sticky="nsew")
        description_container.columnconfigure(0, weight=1)
        description_container.rowconfigure(0, weight=1)
        description_canvas = tk.Canvas(description_container, background="#ffffff", highlightthickness=0)
        description_canvas.grid(row=0, column=0, sticky="nsew")
        description_panel = ttk.Frame(description_canvas, style="Panel.TFrame", padding=12)
        description_window = description_canvas.create_window((0, 0), window=description_panel, anchor="nw")
        description_panel.bind("<Configure>", lambda _event: description_canvas.configure(scrollregion=description_canvas.bbox("all")))
        description_canvas.bind("<Configure>", lambda event: description_canvas.itemconfigure(description_window, width=event.width))
        description_canvas.bind("<MouseWheel>", lambda event: description_canvas.yview_scroll(int(-event.delta / 60), "units"))
        description_panel.columnconfigure(0, weight=1)
        description_panel.columnconfigure(1, weight=1)
        description_panel.rowconfigure(2, weight=1)
        description_panel.rowconfigure(7, weight=2)
        self.tr_label(description_panel, "Version Description", style="Subheader.TLabel").grid(row=0, column=0, sticky="w", pady=(0, 8))
        group_buttons = ttk.Frame(description_panel, style="Panel.TFrame")
        group_buttons.grid(row=1, column=0, columnspan=2, sticky="ew", pady=(0, 8))
        self.version_description_group_var = tk.StringVar(value=list(VERSION_GROUPS.keys())[0])
        self.version_description_buttons: dict[str, tk.Button] = {}
        for index, group_name in enumerate(VERSION_GROUPS):
            button = tk.Button(
                group_buttons,
                text=group_name,
                width=10,
                command=lambda selected_group=group_name: self.select_description_group(selected_group),
            )
            button.grid(row=index // 2, column=index % 2, sticky="ew", padx=2, pady=2)
            self.version_description_buttons[group_name] = button
        group_buttons.columnconfigure(0, weight=1)
        group_buttons.columnconfigure(1, weight=1)

        self.version_description_sw_var = tk.StringVar()
        self.version_description_algo_var = tk.StringVar()
        self.version_description_sw_selected_version: str | None = None
        self.version_description_algo_selected_version: str | None = None

        self.version_sw_description_panel = ttk.Frame(description_panel, style="Panel.TFrame")
        self.version_sw_description_panel.grid(row=2, column=0, sticky="nsew", padx=(0, 6))
        self.version_sw_description_panel.columnconfigure(0, weight=1)
        self.version_sw_description_panel.rowconfigure(1, weight=1)
        self.version_sw_description_panel.rowconfigure(3, weight=2)
        self.version_description_sw_entry = self.add_inline_labeled_entry(
            self.version_sw_description_panel,
            "SW Version",
            self.version_description_sw_var,
            0,
            width=18,
        )
        self.version_description_sw_list = tk.Listbox(self.version_sw_description_panel, height=5, exportselection=False, font=("Segoe UI", 9))
        self.version_description_sw_list.grid(row=1, column=0, sticky="nsew")
        self.version_description_sw_list.bind("<<ListboxSelect>>", lambda _event: self.show_selected_sw_description())
        self.tr_label(self.version_sw_description_panel, "SW Description", style="Panel.TLabel").grid(row=2, column=0, sticky="w", pady=(6, 2))
        self.version_description_sw_text = tk.Text(self.version_sw_description_panel, height=7, wrap="word", font=("Segoe UI", 9))
        self.version_description_sw_text.grid(row=3, column=0, sticky="nsew", pady=(4, 0))
        sw_actions = ttk.Frame(self.version_sw_description_panel, style="Panel.TFrame")
        sw_actions.grid(row=4, column=0, sticky="ew", pady=(8, 0))
        self.tr_button(sw_actions, "Delete SW", self.delete_selected_sw_version_template, prefix="✕ ").pack(side="left")
        self.tr_button(sw_actions, "Save SW", self.save_selected_sw_version_template, prefix="✓ ", style="Accent.TButton").pack(side="right")

        self.version_algo_description_panel = ttk.Frame(description_panel, style="Panel.TFrame")
        self.version_algo_description_panel.grid(row=2, column=1, sticky="nsew", padx=(6, 0))
        self.version_algo_description_panel.columnconfigure(0, weight=1)
        self.version_algo_description_panel.rowconfigure(1, weight=1)
        self.version_algo_description_panel.rowconfigure(3, weight=2)
        self.version_description_algo_entry = self.add_inline_labeled_entry(
            self.version_algo_description_panel,
            "Algo Version",
            self.version_description_algo_var,
            0,
            width=18,
        )
        self.version_description_algo_list = tk.Listbox(self.version_algo_description_panel, height=5, exportselection=False, font=("Segoe UI", 9))
        self.version_description_algo_list.grid(row=1, column=0, sticky="nsew")
        self.version_description_algo_list.bind("<<ListboxSelect>>", lambda _event: self.show_selected_algo_description())
        self.tr_label(self.version_algo_description_panel, "Algo Description", style="Panel.TLabel").grid(row=2, column=0, sticky="w", pady=(6, 2))
        self.version_description_algo_text = tk.Text(self.version_algo_description_panel, height=7, wrap="word", font=("Segoe UI", 9))
        self.version_description_algo_text.grid(row=3, column=0, sticky="nsew", pady=(4, 0))
        algo_actions = ttk.Frame(self.version_algo_description_panel, style="Panel.TFrame")
        algo_actions.grid(row=4, column=0, sticky="ew", pady=(8, 0))
        self.tr_button(algo_actions, "Delete Algo", self.delete_selected_algo_version_template, prefix="✕ ").pack(side="left")
        self.tr_button(algo_actions, "Save Algo", self.save_selected_algo_version_template, prefix="✓ ", style="Accent.TButton").pack(side="right")

        self.bind_mousewheel_recursive(editor, editor_canvas)
        self.bind_mousewheel_recursive(description_panel, description_canvas)
        self.on_version_group_changed()
        self.refresh_version_history()

    def toggle_create_issue_option(self) -> None:
        self.version_create_issue_var.set(not self.version_create_issue_var.get())
        self.refresh_create_issue_button()

    def refresh_create_issue_button(self) -> None:
        marker = "☑" if self.version_create_issue_var.get() else "☐"
        self.version_create_issue_button.configure(text=f"{marker} {self.text('Create Monitoring Issue')}")

    def format_version_template_label(self, row) -> str:
        if not version_group_uses_algo(row["group_name"]):
            return f"SW {row['sw_version']}"
        return f"SW {row['sw_version']} / A {row['algo_version']}"

    def refresh_version_algo_input(self) -> None:
        uses_algo = version_group_uses_algo(self.version_group_var.get())
        state = "normal" if uses_algo else "disabled"
        self.version_algo_entry.configure(state=state)
        self.refresh_version_update_description_layout()
        if not uses_algo:
            self.version_algo_var.set("")

    def refresh_version_update_description_layout(self) -> None:
        if not hasattr(self, "version_single_description_text"):
            return
        uses_algo = version_group_uses_algo(self.version_group_var.get())
        if uses_algo:
            self.version_single_description_text.grid_remove()
            self.version_split_description_frame.grid()
        else:
            self.version_split_description_frame.grid_remove()
            self.version_single_description_text.grid()

    def split_version_update_description(self, value: str) -> tuple[str, str]:
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
        return value.strip(), ""

    def set_version_update_description(self, value: str) -> None:
        sw_text, algo_text = self.split_version_update_description(value or "")
        for widget, text in [
            (self.version_single_description_text, value or ""),
            (self.version_sw_description_text, sw_text),
            (self.version_algo_description_text, algo_text),
        ]:
            widget.delete("1.0", "end")
            widget.insert("1.0", text)

    def version_update_description_value(self) -> str:
        if not version_group_uses_algo(self.version_group_var.get()):
            return self.version_single_description_text.get("1.0", "end").strip()
        sw_text = self.version_sw_description_text.get("1.0", "end").strip()
        algo_text = self.version_algo_description_text.get("1.0", "end").strip()
        parts = []
        if sw_text:
            parts.append(f"[SW Description]\n{sw_text}")
        if algo_text:
            parts.append(f"[Algo Description]\n{algo_text}")
        return "\n\n".join(parts)

    def version_update_description_parts(self) -> tuple[str, str]:
        if not version_group_uses_algo(self.version_group_var.get()):
            return self.version_single_description_text.get("1.0", "end").strip(), ""
        return (
            self.version_sw_description_text.get("1.0", "end").strip(),
            self.version_algo_description_text.get("1.0", "end").strip(),
        )

    def refresh_description_algo_input(self) -> None:
        uses_algo = version_group_uses_algo(self.version_description_group_var.get())
        if uses_algo:
            self.version_sw_description_panel.grid_configure(column=0, columnspan=1, padx=(0, 6))
            self.version_algo_description_panel.grid(row=2, column=1, sticky="nsew", padx=(6, 0))
        else:
            self.version_algo_description_panel.grid_remove()
            self.version_sw_description_panel.grid_configure(column=0, columnspan=2, padx=(0, 0))
            self.version_description_algo_selected_version = None
            self.version_description_algo_var.set("")
            self.version_description_algo_list.delete(0, "end")
            self.set_version_description_text(self.version_description_algo_text, "")

    def on_version_group_changed(self) -> None:
        group_name = self.version_group_var.get()
        if not any(INSTRUMENT_GROUP[instrument] == group_name for instrument in self.version_selected_instruments):
            self.version_selected_instruments = {VERSION_GROUPS[group_name][0]}
        self.refresh_version_algo_input()
        self.refresh_version_target_buttons()

    def select_description_group(self, group_name: str) -> None:
        self.version_description_group_var.set(group_name)
        self.populate_version_description_dashboard()

    def refresh_description_group_buttons(self) -> None:
        selected_bg = "#1f6feb"
        selected_fg = "#ffffff"
        default_bg = self.cget("bg")
        default_fg = "#111827"
        for group_name, button in self.version_description_buttons.items():
            is_selected = group_name == self.version_description_group_var.get()
            button.configure(
                background=selected_bg if is_selected else default_bg,
                foreground=selected_fg if is_selected else default_fg,
                relief="sunken" if is_selected else "raised",
            )

    def populate_version_description_dashboard(self) -> None:
        self.refresh_description_group_buttons()
        self.refresh_description_algo_input()
        group_name = self.version_description_group_var.get()
        self.version_description_sw_rows = version_component_templates(group_name, "sw")
        self.version_description_sw_list.delete(0, "end")
        for row in self.version_description_sw_rows:
            self.version_description_sw_list.insert("end", row["version"])
        if self.version_description_sw_rows:
            self.version_description_sw_list.selection_set(0)
            self.show_selected_sw_description()
        else:
            self.version_description_sw_var.set("")
            self.version_description_sw_selected_version = None
            self.set_version_description_text(self.version_description_sw_text, "")

        if version_group_uses_algo(group_name):
            self.version_description_algo_rows = version_component_templates(group_name, "algo")
            self.version_description_algo_list.delete(0, "end")
            for row in self.version_description_algo_rows:
                self.version_description_algo_list.insert("end", row["version"])
            if self.version_description_algo_rows:
                self.version_description_algo_list.selection_set(0)
                self.show_selected_algo_description()
            else:
                self.version_description_algo_var.set("")
                self.version_description_algo_selected_version = None
                self.set_version_description_text(self.version_description_algo_text, "")

    def show_selected_sw_description(self) -> None:
        selection = self.version_description_sw_list.curselection()
        if not selection:
            self.version_description_sw_selected_version = None
            self.set_version_description_text(self.version_description_sw_text, "")
            return
        row = self.version_description_sw_rows[selection[0]]
        self.version_description_sw_selected_version = row["version"]
        self.version_description_sw_var.set(row["version"])
        self.set_version_description_text(self.version_description_sw_text, row["description"] or "")

    def show_selected_algo_description(self) -> None:
        selection = self.version_description_algo_list.curselection()
        if not selection:
            self.version_description_algo_selected_version = None
            self.set_version_description_text(self.version_description_algo_text, "")
            return
        row = self.version_description_algo_rows[selection[0]]
        self.version_description_algo_selected_version = row["version"]
        self.version_description_algo_var.set(row["version"])
        self.set_version_description_text(self.version_description_algo_text, row["description"] or "")

    def set_version_description_text(self, widget: tk.Text, value: str) -> None:
        widget.delete("1.0", "end")
        widget.insert("1.0", value)

    def save_selected_sw_version_template(self) -> None:
        self.save_selected_version_component_template("sw")

    def save_selected_algo_version_template(self) -> None:
        self.save_selected_version_component_template("algo")

    def save_selected_version_component_template(self, component: str) -> None:
        if component == "sw":
            selected_version = self.version_description_sw_selected_version
            new_version = self.version_description_sw_var.get().strip()
            description = self.version_description_sw_text.get("1.0", "end").strip()
        else:
            selected_version = self.version_description_algo_selected_version
            new_version = self.version_description_algo_var.get().strip()
            description = self.version_description_algo_text.get("1.0", "end").strip()
        if not selected_version:
            messagebox.showwarning(APP_TITLE, "Select a version first.")
            return
        try:
            update_version_component_template(
                self.version_description_group_var.get(),
                component,
                selected_version,
                new_version,
                description,
                self.current_worker_var.get().strip(),
            )
        except ValueError as exc:
            messagebox.showerror(APP_TITLE, str(exc))
            return
        self.refresh_version_history()
        messagebox.showinfo(APP_TITLE, "Version updated.")

    def delete_selected_sw_version_template(self) -> None:
        self.delete_selected_version_component_template("sw")

    def delete_selected_algo_version_template(self) -> None:
        self.delete_selected_version_component_template("algo")

    def delete_selected_version_component_template(self, component: str) -> None:
        if component == "sw":
            selected_version = self.version_description_sw_selected_version
            label = f"SW {selected_version}" if selected_version else ""
        else:
            selected_version = self.version_description_algo_selected_version
            label = f"Algo {selected_version}" if selected_version else ""
        if not selected_version:
            messagebox.showwarning(APP_TITLE, "Select a version first.")
            return
        if not messagebox.askyesno(
            APP_TITLE,
            f"Delete this version?\n\n{self.version_description_group_var.get()} / {label}\n\nApplied version dashboard records for this version will also be removed.",
        ):
            return
        delete_version_component_template(
            self.version_description_group_var.get(),
            component,
            selected_version,
        )
        self.refresh_version_history()
        messagebox.showinfo(APP_TITLE, "Version deleted.")

    def select_version_target(self, line: str, instrument: str) -> None:
        self.version_group_var.set(INSTRUMENT_GROUP[instrument])
        self.version_selected_lines = {line}
        self.version_selected_instruments = {instrument}
        self.refresh_version_algo_input()
        self.refresh_version_target_buttons()
        self.select_version_description_for_target(line, instrument)

    def select_version_description_for_target(self, line: str, instrument: str) -> None:
        group_name = INSTRUMENT_GROUP[instrument]
        self.version_description_group_var.set(group_name)
        self.populate_version_description_dashboard()
        row = getattr(self, "version_dashboard_latest", {}).get((line, instrument))
        if not row:
            return
        self.select_version_description_component("sw", row.get("sw_version", ""))
        if instrument_uses_algo(instrument):
            self.select_version_description_component("algo", row.get("algo_version", ""))

    def select_version_description_component(self, component: str, version: str) -> None:
        if not version:
            return
        if component == "sw":
            rows = getattr(self, "version_description_sw_rows", [])
            listbox = self.version_description_sw_list
            show = self.show_selected_sw_description
        else:
            rows = getattr(self, "version_description_algo_rows", [])
            listbox = self.version_description_algo_list
            show = self.show_selected_algo_description
        listbox.selection_clear(0, "end")
        for index, row in enumerate(rows):
            if row["version"] == version:
                listbox.selection_set(index)
                listbox.see(index)
                show()
                return

    def toggle_version_line(self, line: str) -> None:
        if line in self.version_selected_lines and len(self.version_selected_lines) > 1:
            self.version_selected_lines.remove(line)
        else:
            self.version_selected_lines.add(line)
        self.refresh_version_target_buttons()

    def toggle_version_instrument(self, instrument: str) -> None:
        instrument_group = INSTRUMENT_GROUP[instrument]
        if instrument_group != self.version_group_var.get():
            self.version_group_var.set(instrument_group)
            self.version_selected_instruments = {instrument}
            self.refresh_version_algo_input()
            self.refresh_version_target_buttons()
            return
        if instrument in self.version_selected_instruments and len(self.version_selected_instruments) > 1:
            self.version_selected_instruments.remove(instrument)
        else:
            self.version_selected_instruments.add(instrument)
        self.refresh_version_target_buttons()

    def refresh_version_target_buttons(self) -> None:
        selected_bg = "#1f6feb"
        selected_fg = "#ffffff"
        disabled_bg = "#e5e7eb"
        default_bg = self.cget("bg")
        default_fg = "#111827"
        for line, button in self.version_line_buttons.items():
            is_selected = line in self.version_selected_lines
            button.configure(
                background=selected_bg if is_selected else default_bg,
                foreground=selected_fg if is_selected else default_fg,
                relief="sunken" if is_selected else "raised",
            )
        group_name = self.version_group_var.get()
        for instrument, button in self.version_instrument_buttons.items():
            is_allowed = INSTRUMENT_GROUP[instrument] == group_name
            is_selected = instrument in self.version_selected_instruments
            button.configure(
                state="normal",
                background=selected_bg if is_selected else (default_bg if is_allowed else disabled_bg),
                foreground=selected_fg if is_selected else (default_fg if is_allowed else "#6b7280"),
                relief="sunken" if is_selected else "raised",
            )

    def save_version_updates(self) -> None:
        lines = sorted(self.version_selected_lines, key=LINES.index)
        instruments = sorted(self.version_selected_instruments, key=INSTRUMENTS.index)
        if not lines or not instruments:
            messagebox.showwarning(APP_TITLE, "Select at least one line and one vision.")
            return
        description = self.version_update_description_value()
        sw_description, algo_description = self.version_update_description_parts()
        algo_version = self.version_algo_var.get().strip() if version_group_uses_algo(self.version_group_var.get()) else ""
        saved_count = 0
        try:
            for line in lines:
                for instrument in instruments:
                    create_version_update(
                        VersionInput(
                            update_time=self.version_update_time_var.get().strip(),
                            group_name=self.version_group_var.get(),
                            line=line,
                            instrument=instrument,
                            sw_version=self.version_sw_var.get().strip(),
                            algo_version=algo_version,
                            description=description,
                            worker=self.current_worker_var.get().strip(),
                            sw_description=sw_description,
                            algo_description=algo_description,
                        ),
                        self.version_create_issue_var.get(),
                    )
                    saved_count += 1
        except ValueError as exc:
            messagebox.showerror(APP_TITLE, str(exc))
            return
        self.refresh_version_history()
        self.refresh_open_issues()
        self.search_records()
        messagebox.showinfo(APP_TITLE, f"{saved_count} version update record(s) saved.")

    def refresh_version_history(self) -> None:
        self.populate_version_dashboard()
        self.populate_version_description_dashboard()

    def export_version_dashboard(self) -> None:
        default_name = f"vision_version_dashboard_{now_text().replace(':', '').replace(' ', '_')}.xlsx"
        output = filedialog.asksaveasfilename(
            title="Save Version Dashboard",
            defaultextension=".xlsx",
            initialfile=default_name,
            filetypes=[("Excel Workbook", "*.xlsx")],
        )
        if not output:
            return
        export_version_dashboard_to_excel(Path(output))
        messagebox.showinfo(APP_TITLE, f"Version dashboard saved:\n{output}")

    def populate_version_dashboard(self) -> None:
        latest = latest_dashboard_versions()
        self.version_dashboard_latest = latest
        recent_threshold = datetime.now() - timedelta(days=7)
        latest_keys = {
            instrument: self.latest_current_version_keys(latest, instrument)
            for instrument in INSTRUMENTS
        }
        for line in LINES:
            for instrument in INSTRUMENTS:
                row = latest.get((line, instrument))
                cell = self.version_cells[(line, instrument)]
                shell = cell["shell"]
                recent_bar = cell["recent_bar"]
                sw_label = cell["sw_label"]
                algo_label = cell["algo_label"]
                sw_dot = cell["sw_dot"]
                algo_dot = cell["algo_dot"]

                if row is None:
                    shell.configure(bg="#e5e7eb")
                    recent_bar.configure(bg="#ffffff")
                    sw_label.configure(text="-", fg="#6b7280")
                    algo_label.configure(text="")
                    self.draw_status_dot(sw_dot, None)
                    self.draw_status_dot(algo_dot, None)
                    continue

                max_sw, max_algo = latest_keys[instrument]
                sw_key = self.normalized_version_key(row["sw_version"], len(max_sw or ()))
                algo_key = self.normalized_version_key(row["algo_version"], len(max_algo or ()))
                sw_outdated = max_sw is not None and sw_key is not None and sw_key < max_sw
                algo_outdated = (
                    instrument_uses_algo(instrument)
                    and max_algo is not None
                    and algo_key is not None
                    and algo_key < max_algo
                )
                try:
                    updated = datetime.strptime(row["update_time"], "%Y-%m-%d %H:%M")
                except ValueError:
                    updated = datetime.min

                shell.configure(bg="#d8dee8")
                recent_bar.configure(bg="#22c55e" if updated >= recent_threshold else "#ffffff")
                sw_label.configure(text=f"SW {row['sw_version']}", fg="#111827")
                self.draw_status_dot(sw_dot, "#f59e0b" if sw_outdated else None)
                if instrument_uses_algo(instrument):
                    algo_label.grid()
                    algo_dot.grid()
                    algo_label.configure(text=f"A {row['algo_version']}", fg="#111827")
                    self.draw_status_dot(algo_dot, "#2563eb" if algo_outdated else None)
                else:
                    algo_label.grid_remove()
                    algo_dot.grid_remove()
                    self.draw_status_dot(algo_dot, None)

    def make_issue_tree(self, parent: ttk.Frame) -> ttk.Treeview:
        columns = ("id", "issue_time", "line", "instrument", "category", "subcategory", "title", "status", "worker")
        tree = ttk.Treeview(parent, columns=columns, show="headings", selectmode="browse")
        headings = {
            "id": "ID",
            "issue_time": "Issue Time",
            "line": "Line",
            "instrument": "Instrument",
            "category": "Category",
            "subcategory": "Subcategory",
            "title": "Title",
            "status": "Status",
            "worker": "Logged By",
        }
        widths = {
            "id": 58,
            "issue_time": 140,
            "line": 72,
            "instrument": 190,
            "category": 110,
            "subcategory": 170,
            "title": 290,
            "status": 100,
            "worker": 130,
        }
        for column in columns:
            tree.heading(column, text=self.text(headings[column]))
            tree.column(column, width=widths[column], minwidth=widths[column], stretch=False, anchor="w")
        for status, (background, foreground) in STATUS_TAGS.items():
            tree.tag_configure(status, background=background, foreground=foreground)
        tree.bind("<MouseWheel>", lambda event: tree.yview_scroll(int(-event.delta / 60), "units"))
        return tree

    def add_labeled_entry(self, parent: ttk.Frame, label: str, variable: tk.StringVar, row: int, column: int, columnspan: int = 1) -> ttk.Entry:
        self.tr_label(parent, label, style="Panel.TLabel").grid(row=row, column=column, sticky="w", padx=(0, 8), pady=7)
        entry = ttk.Entry(parent, textvariable=variable)
        entry.grid(row=row, column=column + 1, columnspan=columnspan, sticky="ew", pady=7)
        return entry

    def add_inline_labeled_entry(
        self,
        parent: ttk.Frame,
        label: str,
        variable: tk.StringVar,
        row: int,
        width: int = 17,
    ) -> ttk.Entry:
        frame = ttk.Frame(parent, style="Panel.TFrame")
        frame.grid(row=row, column=0, columnspan=2, sticky="w", pady=5)
        self.tr_label(frame, label, style="Panel.TLabel", width=12).pack(side="left", padx=(0, 8))
        entry = ttk.Entry(frame, textvariable=variable, width=width)
        entry.pack(side="left")
        return entry

    def add_datetime_picker(self, parent: ttk.Frame, label: str, row: int, column: int) -> None:
        self.tr_label(parent, label, style="Panel.TLabel").grid(row=row, column=column, sticky="w", padx=(0, 8), pady=7)
        frame = ttk.Frame(parent, style="Panel.TFrame")
        frame.grid(row=row, column=column + 1, sticky="w", pady=7)
        self.issue_date_entry = ttk.Entry(frame, textvariable=self.issue_date_var, width=12)
        self.issue_date_entry.pack(side="left", padx=(0, 10))
        self.issue_date_entry.bind("<Button-1>", self.open_calendar_popup)
        tk.Spinbox(frame, from_=0, to=23, textvariable=self.issue_hour_var, width=3, wrap=True, format="%02.0f").pack(side="left")
        ttk.Label(frame, text=":", style="Panel.TLabel").pack(side="left", padx=3)
        tk.Spinbox(frame, from_=0, to=59, textvariable=self.issue_minute_var, width=3, wrap=True, format="%02.0f").pack(side="left")

    def open_calendar_popup(self, _event: tk.Event | None = None) -> None:
        if hasattr(self, "calendar_popup") and self.calendar_popup.winfo_exists():
            self.position_calendar_popup(self.calendar_popup)
            return
        popup = tk.Toplevel(self)
        self.calendar_popup = popup
        popup.title("Select Issue Date")
        popup.resizable(False, False)
        popup.transient(self)
        popup.overrideredirect(True)
        popup.bind("<FocusOut>", lambda _event: popup.destroy())

        selected = self.parse_issue_datetime()
        year_var = tk.IntVar(value=selected.year)
        month_var = tk.IntVar(value=selected.month)

        header = ttk.Frame(popup, padding=8)
        header.pack(fill="x")

        body = ttk.Frame(popup, padding=(8, 0, 8, 8))
        body.pack()

        def draw_calendar() -> None:
            for child in body.winfo_children():
                child.destroy()
            for index, day_name in enumerate(["Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun"]):
                ttk.Label(body, text=day_name, width=5, anchor="center").grid(row=0, column=index, padx=1, pady=1)
            month_days = calendar.monthcalendar(year_var.get(), month_var.get())
            for week_index, week in enumerate(month_days, start=1):
                for day_index, day in enumerate(week):
                    if day == 0:
                        ttk.Label(body, text="", width=5).grid(row=week_index, column=day_index, padx=1, pady=1)
                        continue
                    date_text = f"{year_var.get():04d}-{month_var.get():02d}-{day:02d}"
                    ttk.Button(body, text=str(day), width=4, command=lambda value=date_text: select_date(value)).grid(
                        row=week_index, column=day_index, padx=1, pady=1
                    )

        def change_month(delta: int) -> None:
            month = month_var.get() + delta
            year = year_var.get()
            if month < 1:
                month = 12
                year -= 1
            elif month > 12:
                month = 1
                year += 1
            month_var.set(month)
            year_var.set(year)
            title_var.set(f"{calendar.month_name[month]} {year}")
            draw_calendar()

        def select_date(date_text: str) -> None:
            self.issue_date_var.set(date_text)
            popup.destroy()

        title_var = tk.StringVar(value=f"{calendar.month_name[month_var.get()]} {year_var.get()}")
        ttk.Button(header, text="<", width=3, command=lambda: change_month(-1)).pack(side="left")
        ttk.Label(header, textvariable=title_var, width=18, anchor="center").pack(side="left", padx=6)
        ttk.Button(header, text=">", width=3, command=lambda: change_month(1)).pack(side="left")
        draw_calendar()
        popup.update_idletasks()
        self.position_calendar_popup(popup)
        popup.lift()
        popup.focus_force()

    def position_calendar_popup(self, popup: tk.Toplevel) -> None:
        self.issue_date_entry.update_idletasks()
        x = self.issue_date_entry.winfo_rootx()
        y = self.issue_date_entry.winfo_rooty() + self.issue_date_entry.winfo_height()
        popup_width = max(popup.winfo_reqwidth(), 220)
        popup_height = max(popup.winfo_reqheight(), 180)
        screen_width = self.winfo_screenwidth()
        screen_height = self.winfo_screenheight()
        x = max(0, min(x, screen_width - popup_width - 8))
        if y + popup_height > screen_height:
            y = max(0, self.issue_date_entry.winfo_rooty() - popup_height)
        popup.geometry(f"+{x}+{y}")

    def add_labeled_combo(self, parent: ttk.Frame, label: str, variable: tk.StringVar, values: list[str], row: int, column: int) -> ttk.Combobox:
        self.tr_label(parent, label, style="Panel.TLabel").grid(row=row, column=column, sticky="w", padx=(0, 8), pady=7)
        combo = ttk.Combobox(parent, textvariable=variable, values=values, state="readonly")
        combo.grid(row=row, column=column + 1, sticky="ew", pady=7)
        return combo

    def add_filter_combo(self, parent: ttk.Frame, label: str, variable: tk.StringVar, values: list[str], row: int, column: int) -> ttk.Combobox:
        self.tr_label(parent, label, style="Panel.TLabel").grid(row=row, column=column, sticky="w", padx=(4, 8), pady=6)
        combo = ttk.Combobox(parent, textvariable=variable, values=values, state="readonly", width=20)
        combo.grid(row=row, column=column + 1, sticky="ew", padx=(0, 12), pady=6)
        return combo

    def add_filter_entry(self, parent: ttk.Frame, label: str, variable: tk.StringVar, row: int, column: int) -> ttk.Entry:
        self.tr_label(parent, label, style="Panel.TLabel").grid(row=row, column=column, sticky="w", padx=(4, 8), pady=6)
        entry = ttk.Entry(parent, textvariable=variable, width=22)
        entry.grid(row=row, column=column + 1, sticky="ew", padx=(0, 12), pady=6)
        return entry

    def update_subcategories(self) -> None:
        values = CATEGORY_MAP.get(self.category_var.get(), [""])
        if hasattr(self, "subcategory_combo") and self.subcategory_combo.winfo_exists():
            self.subcategory_combo.configure(values=values)
        self.subcategory_var.set(values[0])

    def update_filter_subcategories(self) -> None:
        category = self.filter_category.get()
        values = [""] + CATEGORY_MAP.get(category, [])
        self.filter_subcategory_combo.configure(values=values)
        self.filter_subcategory.set("")

    def form_issue(self) -> IssueInput:
        return IssueInput(
            issue_time=self.issue_datetime_text(),
            resolved_time=self.resolved_time_var.get().strip() or "00:00",
            line=self.line_var.get().strip(),
            instrument=self.instrument_var.get().strip(),
            worker=self.current_worker_var.get().strip(),
            category=self.category_var.get().strip(),
            subcategory=self.subcategory_var.get().strip(),
            title=self.title_var.get().strip(),
            description=self.description_text.get("1.0", "end").strip(),
            status=self.status_var.get().strip(),
            resolution_notes=self.resolution_text.get("1.0", "end").strip(),
        )

    def save_issue(self) -> None:
        try:
            issue = self.form_issue()
            if self.selected_issue_id is None:
                saved_id = create_issue(issue)
                messagebox.showinfo(APP_TITLE, "Issue saved.")
            else:
                update_issue(self.selected_issue_id, issue)
                saved_id = self.selected_issue_id
                messagebox.showinfo(APP_TITLE, "Issue updated.")
            self.refresh_open_issues()
            self.search_records()
            self.show_issue_detail(saved_id)
            self.select_issue_in_tree(self.search_tree, saved_id)
        except ValueError as exc:
            messagebox.showerror(APP_TITLE, str(exc))

    def clear_form(self) -> None:
        self.selected_issue_id = None
        self.loaded_issue_worker = ""
        self.current_worker_var.set(WORKERS[0])
        self.set_issue_datetime(now_text())
        self.resolved_time_var.set("00:00")
        self.line_var.set(LINES[0])
        self.selected_instruments = {INSTRUMENTS[0]}
        self.instrument_var.set(INSTRUMENTS[0])
        self.category_var.set(CATEGORIES[0])
        self.update_subcategories()
        self.status_var.set(ACTIVE_STATUS_OPTIONS[0])
        self.title_var.set("")
        self.form_description_value = ""
        self.form_resolution_value = ""
        for widget_name in ["description_text", "resolution_text"]:
            widget = getattr(self, widget_name, None)
            if widget is not None and widget.winfo_exists():
                widget.delete("1.0", "end")

    def refresh_open_issues(self) -> None:
        self.active_issue_rows = active_issues()
        self.render_issue_board()

    def render_issue_board(self) -> None:
        rows_by_status = {status: [] for status in ACTIVE_STATUS_OPTIONS}
        for row in self.active_issue_rows:
            if row["status"] in rows_by_status:
                rows_by_status[row["status"]].append(row)

        self.issue_card_widgets = {}
        for status, frame in self.board_columns.items():
            for child in frame.winfo_children():
                child.destroy()
            self.board_count_vars[status].set(str(len(rows_by_status[status])))
            if not rows_by_status[status]:
                ttk.Label(frame, text="-", style="Panel.TLabel").pack(anchor="center", pady=18)
            for row in rows_by_status[status]:
                self.add_issue_card(frame, row, self.board_column_canvases[status])
            self.bind_mousewheel_recursive(frame, self.board_column_canvases[status])
            self.board_column_canvases[status].yview_moveto(0)
        self.refresh_board_card_styles()

    def add_issue_card(self, parent: ttk.Frame, row, canvas: tk.Canvas) -> None:
        issue_id = int(row["id"])
        card = tk.Frame(
            parent,
            bg="#ffffff",
            highlightbackground="#d8dee8",
            highlightthickness=1,
            padx=10,
            pady=8,
        )
        card.pack(fill="x", padx=(0, 2), pady=(0, 8))
        widgets: list[tk.Widget] = []

        title = tk.Label(
            card,
            text=row["title"] or "-",
            bg="#ffffff",
            fg="#111827",
            font=("Segoe UI", 10, "bold"),
            anchor="w",
            justify="left",
            wraplength=330,
        )
        title.pack(fill="x", anchor="w")
        widgets.append(title)

        category_text = row["category"]
        if row["subcategory"]:
            category_text = f"{category_text} / {row['subcategory']}"
        lines = [
            f"{row['line']} / {row['instrument']}",
            category_text,
            row["issue_time"] or "-",
            f"{self.text('Downtime Duration')}: {row['resolved_time'] or '00:00'}",
        ]
        for value in lines:
            label = tk.Label(
                card,
                text=value,
                bg="#ffffff",
                fg="#4b5563",
                font=("Segoe UI", 9),
                anchor="w",
                justify="left",
                wraplength=330,
            )
            label.pack(fill="x", anchor="w", pady=(4, 0))
            widgets.append(label)

        self.issue_card_widgets[issue_id] = (card, widgets)
        self.bind_issue_card_click(card, issue_id)
        self.bind_mousewheel_recursive(card, canvas)

    def bind_issue_card_click(self, widget: tk.Widget, issue_id: int) -> None:
        widget.bind("<Button-1>", lambda _event, selected_id=issue_id: self.select_issue_card(selected_id))
        for child in widget.winfo_children():
            self.bind_issue_card_click(child, issue_id)

    def select_issue_card(self, issue_id: int) -> None:
        self.show_issue_detail(issue_id)

    def refresh_board_card_styles(self) -> None:
        for issue_id, (card, widgets) in getattr(self, "issue_card_widgets", {}).items():
            is_selected = issue_id == self.selected_issue_id
            bg = "#eef6ff" if is_selected else "#ffffff"
            border = "#1f6feb" if is_selected else "#d8dee8"
            card.configure(bg=bg, highlightbackground=border, highlightthickness=2 if is_selected else 1)
            for widget in widgets:
                widget.configure(bg=bg)

    def move_issue_status(self, issue_id: int, status: str) -> None:
        try:
            set_issue_status(issue_id, status)
        except ValueError as exc:
            messagebox.showerror(APP_TITLE, str(exc))
            return
        self.refresh_open_issues()
        self.search_records()
        self.show_issue_detail(issue_id)

    def resolve_issue_from_board(self, issue_id: int) -> None:
        resolve_issue(issue_id)
        self.refresh_open_issues()
        self.search_records()
        self.show_empty_issue_detail()

    def search_records(self) -> None:
        filters = {
            "status": self.filter_status.get(),
            "line": self.filter_line.get(),
            "instrument": self.filter_instrument.get(),
            "category": self.filter_category.get(),
            "subcategory": self.filter_subcategory.get(),
            "keyword": self.filter_keyword.get(),
            "date_from": self.filter_from.get(),
            "date_to": self.filter_to.get(),
        }
        self.search_rows = search_issues(filters)
        self.populate_tree(self.search_tree, self.search_rows)
        if self.search_tree.get_children():
            self.search_tree.yview_moveto(1.0)

    def reset_search_date_bounds(self) -> None:
        first_time, latest_time = issue_time_bounds()
        self.filter_from.set(first_time)
        self.filter_to.set(latest_time)

    def apply_quick_filter(self, filter_name: str) -> None:
        if filter_name == "today":
            today = datetime.now().strftime("%Y-%m-%d")
            self.filter_from.set(f"{today} 00:00")
            self.filter_to.set(f"{today} 23:59")
        elif filter_name == "week":
            today_dt = datetime.now()
            start = today_dt - timedelta(days=today_dt.weekday())
            self.filter_from.set(start.strftime("%Y-%m-%d 00:00"))
            self.filter_to.set(today_dt.strftime("%Y-%m-%d 23:59"))
        elif filter_name == "action":
            self.filter_status.set("Action Required")
        elif filter_name == "monitoring":
            self.filter_status.set("Monitoring")
        elif filter_name == "camera_grab":
            self.filter_category.set("Camera Grab Fail")
            self.update_filter_subcategories()
        elif filter_name == "recipe":
            self.filter_category.set("Recipe")
            self.update_filter_subcategories()
        self.search_records()

    def clear_search_filters(self) -> None:
        for variable in [
            self.filter_status,
            self.filter_line,
            self.filter_category,
            self.filter_subcategory,
            self.filter_keyword,
        ]:
            variable.set("")
        self.filter_instruments.clear()
        self.filter_instrument.set("")
        self.refresh_filter_instrument_buttons()
        self.reset_search_date_bounds()
        self.update_filter_subcategories()
        self.search_records()

    def populate_tree(self, tree: ttk.Treeview, rows: list) -> None:
        for item in tree.get_children():
            tree.delete(item)
        for row_number, row in enumerate(rows, start=1):
            tree.insert(
                "",
                "end",
                iid=str(row["id"]),
                values=(
                    row_number,
                    row["issue_time"],
                    row["line"],
                    row["instrument"],
                    row["category"],
                    row["subcategory"],
                    row["title"],
                    row["status"],
                    row["worker"],
                ),
                tags=(row["status"],),
            )

    def selected_tree_id(self, tree: ttk.Treeview) -> int | None:
        selection = tree.selection()
        if not selection:
            return None
        return int(selection[0])

    def load_selected_open_issue(self) -> None:
        if self.selected_issue_id is None:
            messagebox.showwarning(APP_TITLE, "Select an issue first.")
            return
        self.show_edit_issue_form(self.selected_issue_id)

    def load_issue_state(self, issue_id: int) -> bool:
        row = get_issue(issue_id)
        if row is None:
            messagebox.showerror(APP_TITLE, "Issue was not found.")
            return False
        self.selected_issue_id = issue_id
        self.loaded_issue_worker = row["worker"] or ""
        self.current_worker_var.set(row["worker"] or WORKERS[0])
        self.set_issue_datetime(row["issue_time"])
        self.resolved_time_var.set(row["resolved_time"] or "00:00")
        self.line_var.set(row["line"])
        self.instrument_var.set(row["instrument"])
        self.category_var.set(row["category"])
        self.update_subcategories()
        self.subcategory_var.set(row["subcategory"] or "")
        self.status_var.set(row["status"])
        self.title_var.set(row["title"])
        self.form_description_value = row["description"] or ""
        self.form_resolution_value = row["resolution_notes"] or ""
        return True

    def load_issue_into_form(self, issue_id: int) -> None:
        self.show_edit_issue_form(issue_id)

    def load_selected_search_issue(self) -> None:
        issue_id = self.selected_tree_id(self.search_tree)
        if issue_id is None:
            messagebox.showwarning(APP_TITLE, "Select a search result first.")
            return
        self.notebook.select(self.open_tab)
        self.show_issue_detail(issue_id)

    def select_issue_in_tree(self, tree: ttk.Treeview, issue_id: int) -> None:
        for item in tree.get_children():
            if int(item) == issue_id:
                tree.selection_set(item)
                tree.focus(item)
                tree.see(item)
                return

    def parse_issue_datetime(self) -> datetime:
        try:
            return datetime.strptime(self.issue_datetime_text(), "%Y-%m-%d %H:%M")
        except ValueError:
            return datetime.now()

    def set_issue_datetime(self, value: str) -> None:
        try:
            parsed = datetime.strptime(value, "%Y-%m-%d %H:%M")
        except ValueError:
            parsed = datetime.now()
        self.issue_date_var.set(parsed.strftime("%Y-%m-%d"))
        self.issue_hour_var.set(parsed.strftime("%H"))
        self.issue_minute_var.set(parsed.strftime("%M"))

    def issue_datetime_text(self) -> str:
        try:
            hour = max(0, min(23, int(self.issue_hour_var.get() or 0)))
        except ValueError:
            hour = 0
        try:
            minute = max(0, min(59, int(self.issue_minute_var.get() or 0)))
        except ValueError:
            minute = 0
        self.issue_hour_var.set(f"{hour:02d}")
        self.issue_minute_var.set(f"{minute:02d}")
        return f"{self.issue_date_var.get().strip()} {hour:02d}:{minute:02d}"

    def resolve_selected_open_issue(self) -> None:
        if self.selected_issue_id is None:
            messagebox.showwarning(APP_TITLE, "Select an issue first.")
            return
        self.resolve_issue_from_board(self.selected_issue_id)

    def quick_status_selected(self, tree: ttk.Treeview, status: str) -> None:
        issue_id = self.selected_tree_id(tree)
        if issue_id is None:
            messagebox.showwarning(APP_TITLE, "Select an issue first.")
            return
        try:
            set_issue_status(issue_id, status)
        except ValueError as exc:
            messagebox.showerror(APP_TITLE, str(exc))
            return
        self.refresh_open_issues()
        self.search_records()
        self.show_issue_detail(issue_id)

    def delete_loaded_issue(self) -> None:
        if self.selected_issue_id is None:
            messagebox.showwarning(APP_TITLE, "Load or select an issue first.")
            return
        self.delete_issue_by_id(self.selected_issue_id)

    def delete_selected_issue(self, tree: ttk.Treeview) -> None:
        issue_id = self.selected_tree_id(tree)
        if issue_id is None:
            messagebox.showwarning(APP_TITLE, "Select an issue first.")
            return
        self.delete_issue_by_id(issue_id)

    def delete_issue_by_id(self, issue_id: int) -> None:
        row = get_issue(issue_id)
        title = row["title"] if row else "selected issue"
        if not messagebox.askyesno(
            APP_TITLE,
            f"정말 삭제하시겠습니까?\n\n{title}\n\n삭제 후 되돌릴 수 없습니다.",
        ):
            return
        delete_issue(issue_id)
        if self.selected_issue_id == issue_id:
            self.clear_form()
            self.show_empty_issue_detail()
        self.refresh_open_issues()
        self.search_records()

    def export_search_results(self) -> None:
        if not self.search_rows:
            messagebox.showwarning(APP_TITLE, "No search results to export.")
            return
        default_name = f"PKG Inspection Daily Issue List_{datetime.now().strftime('%y%m%d')}.xlsx"
        output = filedialog.asksaveasfilename(
            title="Save Excel Report",
            defaultextension=".xlsx",
            initialfile=default_name,
            filetypes=[("Excel Workbook", "*.xlsx")],
        )
        if not output:
            return
        export_issues_to_excel(self.search_rows, Path(output))
        messagebox.showinfo(APP_TITLE, f"Excel report saved:\n{output}")


if __name__ == "__main__":
    app = VisionIssueApp()
    app.mainloop()
