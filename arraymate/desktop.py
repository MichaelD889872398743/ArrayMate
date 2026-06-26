"""Tkinter desktop UI for ArrayMate."""

from __future__ import annotations

from datetime import datetime
import json
import os
import platform
import subprocess
import tkinter as tk
import webbrowser
from tkinter import filedialog, messagebox, ttk
from typing import Any, Optional

from arraymate.core import (
    ArrayCandidate,
    ArrayMateCoreError,
    ColumnTransform,
    TablePreview,
    TableTransformOptions,
    build_table_preview,
    get_output_format,
    infer_column_transform_types,
)
from arraymate.service import ArrayMateService, LoadResult


class ArrayMate:
    """Desktop UI for converting JSON arrays to table files."""

    NO_UNFOLD_LABEL = "Keep selected table"
    REPOSITORY_URL = "https://github.com/MichaelD889872398743/ArrayMate"
    WINDOW_WIDTH = 1000
    WINDOW_HEIGHT = 660
    DEFAULT_FONT = ("Arial", 10)
    TITLE_FONT = ("Arial", 16, "bold")
    CODE_FONT = ("Consolas", 10)
    UI_FONT = ("Segoe UI", 10)
    UI_FONT_SMALL = ("Segoe UI", 9)
    PANE_TITLE_FONT = ("Segoe UI", 8, "bold")
    MAX_PREVIEW_COLUMNS = 10
    COLOR_BG = "#1e1e1e"
    COLOR_PANEL = "#252526"
    COLOR_PANEL_2 = "#2d2d30"
    COLOR_BORDER = "#3c3c3c"
    COLOR_TEXT = "#d4d4d4"
    COLOR_MUTED = "#9da5b4"
    COLOR_ACCENT = "#007acc"
    COLOR_GREEN = "#4ec9b0"
    COLOR_YELLOW = "#ffd166"
    COLOR_RED = "#f48771"

    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title("ArrayMate")
        self.root.geometry(f"{self.WINDOW_WIDTH}x{self.WINDOW_HEIGHT}")
        self.root.resizable(True, True)

        self.service = ArrayMateService()
        self.json_data: Optional[dict[str, Any] | list[Any]] = None
        self.array_keys: list[str] = []
        self.candidate_by_path: dict[str, ArrayCandidate] = {}
        self.effective_candidate_key = ""
        self.auto_filename = True
        self.column_transforms: dict[str, ColumnTransform] = {}
        self.current_preview_columns: list[str] = []
        self.advanced_column_combo: Optional[ttk.Combobox] = None
        self.advanced_type_combo: Optional[ttk.Combobox] = None
        self.advanced_section_frame: Optional[ttk.Frame] = None

        self.json_file_path = tk.StringVar()
        self.selected_array_key = tk.StringVar()
        self.output_folder = tk.StringVar()
        self.output_filename = tk.StringVar()
        self.output_format = tk.StringVar(value="Excel (.xlsx)")
        self.stringify_all = tk.BooleanVar(value=False)
        self.stringify_formulas = tk.BooleanVar(value=False)
        self.include_parent_metadata = tk.BooleanVar(value=False)
        self.selected_nested_candidate = tk.StringVar()
        self.advanced_options_visible = tk.BooleanVar(value=False)
        self.json_input_visible = tk.BooleanVar(value=True)
        self.advanced_column = tk.StringVar()
        self.advanced_type = tk.StringVar(value="Keep")
        self.advanced_find = tk.StringVar()
        self.advanced_replace = tk.StringVar()
        self.advanced_status = tk.StringVar(value="No column action selected")

        self.setup_ui()
        self._schedule_auto_filename_refresh()

    def setup_ui(self) -> None:
        """Set up the application window."""
        self._configure_styles()
        self.root.configure(bg=self.COLOR_BG)
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)

        shell = ttk.Frame(self.root, style="App.TFrame")
        shell.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        shell.columnconfigure(0, weight=1)
        shell.rowconfigure(1, weight=1)

        self._create_source_bar(shell)
        self._create_workspace(shell)
        self._create_status_section(shell)

    def _configure_styles(self) -> None:
        style = ttk.Style(self.root)
        style.theme_use("clam")
        style.configure(".", font=self.UI_FONT)
        style.configure("App.TFrame", background=self.COLOR_BG)
        style.configure("Header.TFrame", background=self.COLOR_PANEL)
        style.configure("Panel.TFrame", background=self.COLOR_PANEL)
        style.configure("Center.TFrame", background=self.COLOR_BG)
        style.configure("Card.TFrame", background=self.COLOR_PANEL, bordercolor=self.COLOR_BORDER, relief="solid")
        style.configure("Section.TLabelframe", background=self.COLOR_PANEL, bordercolor=self.COLOR_BORDER)
        style.configure("Section.TLabelframe.Label", background=self.COLOR_PANEL, foreground=self.COLOR_TEXT, font=self.UI_FONT_SMALL)
        style.configure("TLabel", background=self.COLOR_BG, foreground=self.COLOR_TEXT)
        style.configure("Muted.TLabel", background=self.COLOR_BG, foreground=self.COLOR_MUTED, font=self.UI_FONT_SMALL)
        style.configure("Panel.TLabel", background=self.COLOR_PANEL, foreground=self.COLOR_TEXT)
        style.configure("PanelMuted.TLabel", background=self.COLOR_PANEL, foreground=self.COLOR_MUTED, font=self.UI_FONT_SMALL)
        style.configure("PaneTitle.TLabel", background=self.COLOR_PANEL, foreground=self.COLOR_MUTED, font=self.PANE_TITLE_FONT)
        style.configure("HeaderTitle.TLabel", background=self.COLOR_PANEL, foreground=self.COLOR_TEXT, font=("Segoe UI", 14, "bold"))
        style.configure("Status.TLabel", background=self.COLOR_ACCENT, foreground="white", font=self.UI_FONT_SMALL)
        style.configure("TEntry", fieldbackground="#1b1b1b", foreground=self.COLOR_TEXT, insertcolor=self.COLOR_TEXT, bordercolor=self.COLOR_BORDER)
        style.configure("TCombobox", fieldbackground="#1b1b1b", foreground=self.COLOR_TEXT, arrowcolor=self.COLOR_TEXT, bordercolor=self.COLOR_BORDER)
        style.map(
            "TCombobox",
            fieldbackground=[("readonly", "#1b1b1b")],
            selectbackground=[("readonly", "#1b1b1b")],
            selectforeground=[("readonly", self.COLOR_TEXT)],
            foreground=[("readonly", self.COLOR_TEXT)],
        )
        style.configure("TButton", background=self.COLOR_PANEL_2, foreground=self.COLOR_TEXT, bordercolor=self.COLOR_BORDER, focusthickness=0)
        style.map("TButton", background=[("active", "#3a3d41")])
        style.configure("Accent.TButton", background=self.COLOR_ACCENT, foreground="white", bordercolor=self.COLOR_ACCENT)
        style.map("Accent.TButton", background=[("active", "#1688d1"), ("disabled", "#3c3c3c")])
        style.configure("TCheckbutton", background=self.COLOR_BG, foreground=self.COLOR_TEXT)
        style.map("TCheckbutton", background=[("active", self.COLOR_BG)], foreground=[("disabled", self.COLOR_MUTED)])
        style.configure("Panel.TCheckbutton", background=self.COLOR_PANEL, foreground=self.COLOR_TEXT)
        style.map("Panel.TCheckbutton", background=[("active", self.COLOR_PANEL)], foreground=[("disabled", self.COLOR_MUTED)])
        style.configure(
            "Treeview",
            background=self.COLOR_PANEL,
            fieldbackground=self.COLOR_PANEL,
            foreground=self.COLOR_TEXT,
            bordercolor=self.COLOR_BORDER,
            rowheight=24,
        )
        style.configure("Treeview.Heading", background="#202020", foreground=self.COLOR_MUTED, relief="flat", font=self.UI_FONT_SMALL)
        style.map("Treeview", background=[("selected", "#094771")], foreground=[("selected", self.COLOR_TEXT)])

    def _create_source_bar(self, parent: ttk.Frame) -> None:
        header = ttk.Frame(parent, style="Header.TFrame", padding=(14, 10))
        header.grid(row=0, column=0, sticky=(tk.W, tk.E))
        header.columnconfigure(2, weight=1)

        ttk.Label(header, text="ArrayMate", style="HeaderTitle.TLabel").grid(row=0, column=0, sticky=tk.W, padx=(0, 18))
        ttk.Label(header, text="Step 1", style="PanelMuted.TLabel").grid(row=0, column=1, sticky=tk.W, padx=(0, 8))
        ttk.Entry(header, textvariable=self.json_file_path).grid(row=0, column=2, sticky=(tk.W, tk.E), padx=(0, 8))
        ttk.Button(header, text="Browse", command=self.browse_json_file).grid(row=0, column=3, padx=(0, 6))

    def _create_workspace(self, parent: ttk.Frame) -> None:
        workspace = ttk.Frame(parent, style="App.TFrame")
        workspace.grid(row=1, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        workspace.columnconfigure(0, minsize=300, weight=0)
        workspace.columnconfigure(1, weight=1)
        workspace.columnconfigure(2, minsize=330, weight=0)
        workspace.rowconfigure(0, weight=1)

        self._create_structure_pane(workspace)
        self._create_preview_pane(workspace)
        self._create_right_pane(workspace)

    def _create_structure_pane(self, parent: ttk.Frame) -> None:
        left = ttk.Frame(parent, style="Panel.TFrame", padding=(0, 8, 0, 8))
        left.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        left.rowconfigure(1, weight=1)
        left.columnconfigure(0, weight=1)

        ttk.Label(left, text="STEP 2 - PARSED STRUCTURE", style="PaneTitle.TLabel").grid(
            row=0, column=0, sticky=(tk.W, tk.E), padx=12, pady=(0, 8)
        )
        self.array_tree = ttk.Treeview(
            left,
            columns=("items", "columns", "status"),
            show="tree headings",
            selectmode="browse",
        )
        self.array_tree.heading("#0", text="Path")
        self.array_tree.heading("items", text="Rows")
        self.array_tree.heading("columns", text="Columns")
        self.array_tree.heading("status", text="Status")
        self.array_tree.column("#0", width=165, minwidth=120, stretch=True)
        self.array_tree.column("items", width=48, minwidth=42, anchor=tk.E, stretch=False)
        self.array_tree.column("columns", width=42, minwidth=38, anchor=tk.E, stretch=False)
        self.array_tree.column("status", width=72, minwidth=56, stretch=False)
        self.array_tree.grid(row=1, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), padx=(8, 0))
        self.array_tree.bind("<<TreeviewSelect>>", self.on_array_selected)

    def _create_preview_pane(self, parent: ttk.Frame) -> None:
        center = ttk.Frame(parent, style="Center.TFrame", padding=14)
        center.grid(row=0, column=1, sticky=(tk.W, tk.E, tk.N, tk.S))
        center.columnconfigure(0, weight=1)
        center.rowconfigure(2, weight=1)

        info_frame = ttk.Frame(center, style="Center.TFrame")
        info_frame.grid(row=0, column=0, sticky=(tk.W, tk.E), pady=(0, 10))
        info_frame.columnconfigure(0, weight=1)
        self.array_info_label = ttk.Label(info_frame, text="No JSON loaded", style="Muted.TLabel")
        self.array_info_label.grid(row=0, column=0, sticky=(tk.W, tk.E), padx=(0, 10))
        self.paste_json_button = ttk.Button(
            info_frame,
            text="Hide JSON",
            command=self.open_json_input_window,
            style="Accent.TButton",
        )
        self.paste_json_button.grid(row=0, column=1, sticky=tk.E)

        self._create_inline_json_input(center)

        preview_frame = ttk.Frame(center, style="Card.TFrame", padding=0)
        preview_frame.grid(row=2, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        preview_frame.columnconfigure(0, weight=1)
        preview_frame.rowconfigure(1, weight=1)

        preview_header = ttk.Frame(preview_frame, style="Panel.TFrame", padding=(12, 9))
        preview_header.grid(row=0, column=0, sticky=(tk.W, tk.E))
        preview_header.columnconfigure(1, weight=1)
        ttk.Label(preview_header, text="Preview", style="Panel.TLabel").grid(row=0, column=0, sticky=tk.W)
        self.preview_label = ttk.Label(preview_header, text="Select an exportable array to preview rows.", style="PanelMuted.TLabel")
        self.preview_label.grid(row=0, column=1, sticky=tk.E)

        self.preview_tree = ttk.Treeview(preview_frame, show="headings")
        self.preview_tree.grid(row=1, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

    def _create_inline_json_input(self, parent: ttk.Frame) -> None:
        self.json_input_frame = ttk.Frame(parent, style="Card.TFrame", padding=0)
        self.json_input_frame.columnconfigure(0, weight=1)
        self.json_input_frame.rowconfigure(1, weight=1)

        json_header = ttk.Frame(self.json_input_frame, style="Panel.TFrame", padding=(12, 9))
        json_header.grid(row=0, column=0, sticky=(tk.W, tk.E))
        json_header.columnconfigure(0, weight=1)
        ttk.Label(json_header, text="JSON Input", style="Panel.TLabel").grid(row=0, column=0, sticky=tk.W)
        ttk.Button(json_header, text="Load JSON", command=self.load_json_from_text, style="Accent.TButton").grid(
            row=0, column=1, padx=(0, 6)
        )
        ttk.Button(json_header, text="Clear", command=self.clear_json_text).grid(row=0, column=2, padx=(0, 6))

        input_body = ttk.Frame(self.json_input_frame, style="Panel.TFrame", padding=(12, 0, 12, 12))
        input_body.grid(row=1, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        input_body.columnconfigure(0, weight=1)
        input_body.rowconfigure(0, weight=1)

        self.json_text = tk.Text(
            input_body,
            wrap=tk.WORD,
            font=self.CODE_FONT,
            height=8,
            background="#1b1b1b",
            foreground=self.COLOR_TEXT,
            insertbackground=self.COLOR_TEXT,
            relief=tk.FLAT,
            borderwidth=1,
            highlightthickness=1,
            highlightbackground=self.COLOR_BORDER,
            highlightcolor=self.COLOR_ACCENT,
        )
        self.json_text.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        scrollbar = ttk.Scrollbar(input_body, orient=tk.VERTICAL, command=self.json_text.yview)
        scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))
        self.json_text.configure(yscrollcommand=scrollbar.set)
        self._set_json_input_visible(self.json_input_visible.get())

    def _create_right_pane(self, parent: ttk.Frame) -> None:
        right = ttk.Frame(parent, style="Panel.TFrame", padding=14)
        right.grid(row=0, column=2, sticky=(tk.W, tk.E, tk.N, tk.S))
        right.columnconfigure(0, weight=1)

        ttk.Label(right, text="STEP 3 - TRANSFORM & EXPORT", style="PaneTitle.TLabel").grid(
            row=0, column=0, sticky=(tk.W, tk.E), pady=(0, 10)
        )
        self._create_output_settings_section(right, row=1, column=0)
        self._create_transform_section(right, row=2, column=0)
        self._create_warning_section(right, row=3, column=0)
        self._create_project_section(right, row=4, column=0)

    def _create_transform_section(self, parent: ttk.Frame, row: int, column: int) -> None:
        options_frame = ttk.LabelFrame(parent, text="Transform Options", style="Section.TLabelframe", padding="10")
        options_frame.grid(row=row, column=column, sticky=(tk.W, tk.E, tk.N), pady=(0, 12))
        options_frame.columnconfigure(0, weight=1)

        ttk.Label(options_frame, text="Unfold level", style="PanelMuted.TLabel").grid(row=0, column=0, sticky=tk.W)
        self.nested_candidate_combo = ttk.Combobox(
            options_frame,
            textvariable=self.selected_nested_candidate,
            state="disabled",
        )
        self.nested_candidate_combo.grid(row=1, column=0, sticky=(tk.W, tk.E), pady=(4, 10))
        self.nested_candidate_combo.bind("<<ComboboxSelected>>", self.refresh_selected_candidate)

        ttk.Checkbutton(
            options_frame,
            text="Stringify everything",
            variable=self.stringify_all,
            command=self.refresh_selected_candidate,
            style="Panel.TCheckbutton",
        ).grid(row=2, column=0, sticky=tk.W, pady=(0, 4))
        ttk.Checkbutton(
            options_frame,
            text="Stringify formulas",
            variable=self.stringify_formulas,
            command=self.refresh_selected_candidate,
            style="Panel.TCheckbutton",
        ).grid(row=3, column=0, sticky=tk.W, pady=(0, 4))
        self.parent_metadata_check = ttk.Checkbutton(
            options_frame,
            text="Include parent metadata",
            variable=self.include_parent_metadata,
            command=self.refresh_selected_candidate,
            state="disabled",
            style="Panel.TCheckbutton",
        )
        self.parent_metadata_check.grid(row=4, column=0, sticky=tk.W, pady=(0, 10))

        self.advanced_toggle_button = ttk.Button(
            options_frame,
            text="Show Advanced Options",
            command=self.open_advanced_options,
        )
        self.advanced_toggle_button.grid(row=5, column=0, sticky=(tk.W, tk.E))

        self.advanced_section_frame = ttk.Frame(options_frame, style="Panel.TFrame", padding=(0, 10, 0, 0))
        self.advanced_section_frame.columnconfigure(0, weight=1)
        self._create_column_actions_tab(self.advanced_section_frame)

    def _create_output_settings_section(self, parent: ttk.Frame, row: int = 3, column: int = 0) -> None:
        output_frame = ttk.LabelFrame(parent, text="Export", style="Section.TLabelframe", padding="10")
        output_frame.grid(row=row, column=column, sticky=(tk.W, tk.E, tk.N), pady=(0, 12))
        output_frame.columnconfigure(0, weight=1)

        ttk.Label(output_frame, text="Output format", style="PanelMuted.TLabel").grid(row=0, column=0, sticky=tk.W)
        format_combobox = ttk.Combobox(
            output_frame,
            textvariable=self.output_format,
            values=["Excel (.xlsx)", "CSV (.csv)", "JSON (.json)"],
            state="readonly",
        )
        format_combobox.set(self.output_format.get())
        format_combobox.grid(row=1, column=0, sticky=(tk.W, tk.E), pady=(4, 10))
        format_combobox.bind("<<ComboboxSelected>>", self.on_format_selected)

        ttk.Label(output_frame, text="File name", style="PanelMuted.TLabel").grid(row=2, column=0, sticky=tk.W)
        filename_entry = ttk.Entry(output_frame, textvariable=self.output_filename)
        filename_entry.grid(row=3, column=0, sticky=(tk.W, tk.E), pady=(4, 10))
        filename_entry.bind("<KeyRelease>", self.on_filename_edited)
        self.extension_label = ttk.Label(output_frame, text=".xlsx")
        self.extension_label.grid(row=3, column=1, sticky=tk.W, padx=(6, 0), pady=(4, 10))

        ttk.Label(output_frame, text="Save folder", style="PanelMuted.TLabel").grid(row=4, column=0, sticky=tk.W)
        folder_frame = ttk.Frame(output_frame, style="Panel.TFrame")
        folder_frame.grid(row=5, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(4, 10))
        folder_frame.columnconfigure(0, weight=1)
        ttk.Entry(folder_frame, textvariable=self.output_folder).grid(row=0, column=0, sticky=(tk.W, tk.E), padx=(0, 6))
        ttk.Button(folder_frame, text="Browse", command=self.browse_output_folder).grid(row=0, column=1)

        button_frame = ttk.Frame(output_frame)
        button_frame.grid(row=6, column=0, columnspan=2, sticky=(tk.W, tk.E))
        button_frame.columnconfigure(0, weight=1)
        self.process_button = ttk.Button(
            button_frame,
            text="Convert to File",
            command=self.convert_to_file,
            state="disabled",
            style="Accent.TButton",
        )
        self.process_button.grid(row=0, column=0, sticky=(tk.W, tk.E), padx=(0, 8))
        ttk.Button(button_frame, text="Clear", command=self.clear_all).grid(row=0, column=1)

    def _create_process_buttons(self, parent: ttk.Frame) -> None:
        button_frame = ttk.Frame(parent, style="Panel.TFrame")
        button_frame.grid(row=4, column=0, pady=12)

        self.process_button = ttk.Button(
            button_frame,
            text="Convert to File",
            command=self.convert_to_file,
            state="disabled",
        )
        self.process_button.pack(side=tk.LEFT, padx=(0, 10))
        ttk.Button(button_frame, text="Clear All", command=self.clear_all).pack(side=tk.LEFT)

    def _create_status_section(self, parent: ttk.Frame) -> None:
        status_frame = ttk.Frame(parent, style="Header.TFrame")
        status_frame.grid(row=2, column=0, sticky=(tk.W, tk.E))
        status_frame.columnconfigure(0, weight=1)
        self.status_label = ttk.Label(
            status_frame,
            text="Ready to convert JSON arrays",
            style="Status.TLabel",
            padding=(10, 4),
        )
        self.status_label.grid(row=0, column=0, sticky=(tk.W, tk.E))

    def _create_warning_section(self, parent: ttk.Frame, row: int, column: int) -> None:
        warning_frame = ttk.LabelFrame(parent, text="Detected Warnings", style="Section.TLabelframe", padding="10")
        warning_frame.grid(row=row, column=column, sticky=(tk.W, tk.E, tk.N))
        warning_frame.columnconfigure(0, weight=1)
        self.warning_label = ttk.Label(warning_frame, text="JSON not parsed yet", style="PanelMuted.TLabel", wraplength=280)
        self.warning_label.grid(row=0, column=0, sticky=(tk.W, tk.E))

    def _create_project_section(self, parent: ttk.Frame, row: int, column: int) -> None:
        project_frame = ttk.LabelFrame(parent, text="Project", style="Section.TLabelframe", padding="10")
        project_frame.grid(row=row, column=column, sticky=(tk.W, tk.E, tk.S), pady=(12, 0))
        project_frame.columnconfigure(0, weight=1)
        ttk.Label(project_frame, text="ArrayMate on GitHub", style="PanelMuted.TLabel").grid(
            row=0, column=0, sticky=tk.W, pady=(0, 6)
        )
        ttk.Button(project_frame, text="Open Repository", command=self.open_repository).grid(
            row=1, column=0, sticky=(tk.W, tk.E)
        )

    def clear_all(self) -> None:
        """Clear inputs and reset application state."""
        self.json_file_path.set("")
        self.selected_array_key.set("")
        self.output_folder.set("")
        self.output_filename.set("")
        self.auto_filename = True
        self.output_format.set("Excel (.xlsx)")
        self.extension_label["text"] = ".xlsx"
        self.process_button["text"] = "Convert to File"
        self.process_button["state"] = "disabled"

        self.json_data = None
        self.array_keys = []
        self.candidate_by_path = {}
        self.effective_candidate_key = ""
        self.column_transforms = {}
        self.current_preview_columns = []
        self.service.clear()
        self.array_tree.delete(*self.array_tree.get_children())
        self._clear_preview()
        self.array_info_label["text"] = "No JSON loaded"
        self.warning_label["text"] = "JSON not parsed yet"
        self.status_label["text"] = "Ready to convert JSON arrays"
        self.status_label["foreground"] = "green"
        self.clear_json_text()
        self._set_json_input_visible(True)
        self.include_parent_metadata.set(False)
        self.stringify_all.set(False)
        self.stringify_formulas.set(False)
        self._reset_column_action_form()
        self.parent_metadata_check.state(["disabled"])
        self._clear_nested_candidate_action()

    def browse_json_file(self) -> None:
        file_path = filedialog.askopenfilename(
            title="Select JSON File",
            filetypes=[("JSON files", "*.json"), ("All files", "*.*")],
        )

        if file_path:
            self.json_file_path.set(file_path)
            self.load_json_file()

    def open_json_input_window(self) -> None:
        self._set_json_input_visible(not self.json_input_visible.get())

    def open_repository(self) -> None:
        webbrowser.open(self.REPOSITORY_URL)
        self.status_label["text"] = "Repository opened in browser"

    def _set_json_input_visible(self, visible: bool) -> None:
        self.json_input_visible.set(visible)
        if visible:
            self.json_input_frame.grid(row=1, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(0, 12))
            self.paste_json_button["text"] = "Hide JSON"
            self.json_text.focus()
        else:
            self.json_input_frame.grid_remove()
            self.paste_json_button["text"] = "Show JSON"

    def load_json_from_text(self) -> None:
        json_text = self.json_text.get("1.0", tk.END).strip()
        if not json_text:
            messagebox.showerror("Error", "Please enter JSON data")
            return

        try:
            load_result = self.service.load_text(json_text)
            self.json_data = self.service.json_data
            self.json_file_path.set("")
            self._apply_load_result(load_result, "JSON data")
            if load_result.array_candidates:
                self._set_json_input_visible(False)
            else:
                messagebox.showerror("Error", "No arrays found in the JSON data")
        except json.JSONDecodeError as e:
            messagebox.showerror("Error", f"Invalid JSON format: {str(e)}")
        except Exception as e:
            messagebox.showerror("Error", f"Error parsing JSON: {str(e)}")

    def clear_json_text(self) -> None:
        self.json_text.delete("1.0", tk.END)

    def load_json_file(self) -> None:
        try:
            load_result = self.service.load_file(self.json_file_path.get())
            self.json_data = self.service.json_data
            self._apply_load_result(load_result, "JSON file")
            if not load_result.array_candidates:
                messagebox.showerror("Error", "No arrays found in JSON file")
            else:
                self._set_json_input_visible(False)
        except json.JSONDecodeError as e:
            messagebox.showerror("Error", f"Invalid JSON file: {str(e)}")
            self.status_label["text"] = "Error: Invalid JSON file"
            self.status_label["foreground"] = "red"
        except Exception as e:
            messagebox.showerror("Error", f"Error loading file: {str(e)}")
            self.status_label["text"] = "Error loading file"
            self.status_label["foreground"] = "red"

    def _apply_load_result(self, load_result: LoadResult, source_label: str) -> None:
        self.array_keys = load_result.array_keys
        self.candidate_by_path = {candidate.display_path: candidate for candidate in load_result.array_candidates}
        self.column_transforms = {}
        self._reset_column_action_form()
        self.array_tree.delete(*self.array_tree.get_children())
        self._clear_preview()

        for candidate in load_result.array_candidates:
            status = self._candidate_status(candidate)
            self.array_tree.insert(
                "",
                tk.END,
                iid=candidate.display_path,
                text=candidate.display_path,
                values=(candidate.item_count, candidate.column_count or "", status),
            )

        selected_candidate = self._default_candidate(load_result.array_candidates)
        if selected_candidate:
            self.array_tree.selection_set(selected_candidate.display_path)
            self.array_tree.focus(selected_candidate.display_path)
            self._select_candidate(selected_candidate)
            self.status_label["text"] = f"Found {len(load_result.array_candidates)} array candidate(s) in {source_label}"
            self.status_label["foreground"] = "green"
        else:
            self.selected_array_key.set("")
            self.array_info_label["text"] = f"No arrays found in {source_label}"
            self.process_button["state"] = "disabled"
            self.status_label["text"] = f"No arrays found in {source_label}"
            self.status_label["foreground"] = "orange"

    def _default_candidate(self, candidates: list[ArrayCandidate]) -> Optional[ArrayCandidate]:
        return next((candidate for candidate in candidates if candidate.exportable), candidates[0] if candidates else None)

    def _candidate_status(self, candidate: ArrayCandidate) -> str:
        if candidate.exportable and candidate.warning:
            return f"Exportable, {candidate.warning.lower()}"
        if candidate.exportable:
            return "Exportable"
        return candidate.warning or "Not exportable"

    def on_format_selected(self, event: Optional[tk.Event] = None) -> None:
        output_format = get_output_format(self.output_format.get())
        self.extension_label["text"] = output_format.extension
        self.process_button["text"] = f"Convert to {output_format.label}"

    def browse_output_folder(self) -> None:
        folder_path = filedialog.askdirectory(title="Select Output Folder")
        if folder_path:
            self.output_folder.set(folder_path)
            self._refresh_auto_filename()

    def on_filename_edited(self, event: Optional[tk.Event] = None) -> None:
        self.auto_filename = False

    def on_array_selected(self, event: Optional[tk.Event] = None) -> None:
        selection = self.array_tree.selection()
        if not selection:
            return

        candidate = self.candidate_by_path.get(selection[0])
        if candidate:
            self._select_candidate(candidate)

    def refresh_selected_candidate(self, event: Optional[tk.Event] = None) -> None:
        selected_key = self.selected_array_key.get()
        candidate = self.candidate_by_path.get(selected_key)
        if candidate:
            self._select_candidate(candidate, reset_unfold=False)

    def _select_candidate(self, candidate: ArrayCandidate, reset_unfold: bool = True) -> None:
        previous_key = self.selected_array_key.get()
        if previous_key and previous_key != candidate.display_path:
            self.column_transforms = {}
            self._reset_column_action_form()
        self.selected_array_key.set(candidate.display_path)
        self._update_nested_candidate_action(candidate, reset_unfold=reset_unfold or previous_key != candidate.display_path)
        unfold_key = self._unfold_key()
        effective_candidate = self.candidate_by_path.get(unfold_key, candidate) if unfold_key else candidate
        if unfold_key:
            self.include_parent_metadata.set(False)
            self.parent_metadata_check.state(["disabled"])
        else:
            self._update_parent_metadata_option(
                candidate,
                is_new_selection=self.effective_candidate_key != candidate.display_path,
            )
        self.effective_candidate_key = effective_candidate.display_path
        if not self.output_filename.get():
            self.auto_filename = True
        self._refresh_auto_filename(effective_candidate.display_path)

        if effective_candidate.exportable:
            try:
                array_data = self.service.get_table_data(
                    candidate.display_path,
                    unfold_key=unfold_key,
                    include_parent_metadata=self._include_parent_metadata_for(candidate),
                    transform_options=self._table_transform_options(),
                )
                if array_data is None:
                    raise ArrayMateCoreError("Selected array is invalid")
                preview = build_table_preview(array_data, effective_candidate.display_path)
                self.array_info_label["text"] = self._candidate_detail_text(candidate, effective_candidate, preview)
                self.warning_label["text"] = self._warning_text(effective_candidate, preview)
                self._render_preview(preview)
                self.process_button["state"] = "normal"
            except ArrayMateCoreError as e:
                self.warning_label["text"] = str(e)
                self._clear_preview(str(e))
                self.process_button["state"] = "disabled"
        else:
            self.array_info_label["text"] = self._candidate_detail_text(candidate, effective_candidate)
            self.warning_label["text"] = effective_candidate.warning or "This array is not exportable as a table."
            self._clear_preview(effective_candidate.warning or "This array is not exportable as a table.")
            self.process_button["state"] = "disabled"

    def _candidate_detail_text(
        self,
        candidate: ArrayCandidate,
        effective_candidate: ArrayCandidate,
        preview: Optional[TablePreview] = None,
    ) -> str:
        details = [
            f"Selected: {candidate.display_path}",
            f"rows: {preview.rows if preview else effective_candidate.item_count}",
            f"columns: {len(preview.columns) if preview else effective_candidate.column_count}",
        ]
        if effective_candidate.display_path != candidate.display_path:
            details.append(f"unfolding: {effective_candidate.display_path}")
        if effective_candidate.source_count > 1:
            details.append(f"grouped from {effective_candidate.source_count} parent records")
        if effective_candidate.warning:
            details.append(effective_candidate.warning)
        return " | ".join(details)

    def _warning_text(self, candidate: ArrayCandidate, preview: TablePreview) -> str:
        warnings = ["JSON parsed successfully"]
        if candidate.warning:
            warnings.append(candidate.warning)
        warnings.extend(preview.warnings)
        if self.stringify_formulas.get():
            warnings.append("Formula protection enabled")
        return "\n".join(warnings)

    def _update_parent_metadata_option(self, candidate: ArrayCandidate, is_new_selection: bool) -> None:
        supports_parent_metadata = self._supports_parent_metadata(candidate)
        if supports_parent_metadata:
            self.parent_metadata_check.state(["!disabled"])
            if is_new_selection:
                self.include_parent_metadata.set(True)
        else:
            self.include_parent_metadata.set(False)
            self.parent_metadata_check.state(["disabled"])

    def _supports_parent_metadata(self, candidate: ArrayCandidate) -> bool:
        return any(segment is Ellipsis for segment in candidate.path)

    def _include_parent_metadata_for(self, candidate: ArrayCandidate) -> bool:
        return self._supports_parent_metadata(candidate) and self.include_parent_metadata.get()

    def _table_transform_options(self) -> TableTransformOptions:
        return TableTransformOptions(
            stringify_all=self.stringify_all.get(),
            stringify_formulas=self.stringify_formulas.get(),
            column_transforms=tuple(self.column_transforms.values()),
        )

    def _effective_candidate(self, candidate: ArrayCandidate) -> ArrayCandidate:
        nested_key = self._unfold_key()
        if nested_key:
            return self.candidate_by_path.get(nested_key, candidate)
        return candidate

    def _unfold_key(self) -> Optional[str]:
        nested_key = self.selected_nested_candidate.get()
        if nested_key and nested_key != self.NO_UNFOLD_LABEL:
            return nested_key
        return None

    def _update_nested_candidate_action(self, candidate: ArrayCandidate, reset_unfold: bool) -> None:
        nested_candidates = self.service.get_nested_array_candidates(candidate.display_path, max_nested_levels=3)
        if not nested_candidates:
            self._clear_nested_candidate_action()
            return

        nested_values = [self.NO_UNFOLD_LABEL] + [nested_candidate.display_path for nested_candidate in nested_candidates]
        self.nested_candidate_combo["values"] = nested_values
        if reset_unfold or self.selected_nested_candidate.get() not in nested_values:
            self.selected_nested_candidate.set(self.NO_UNFOLD_LABEL)
            self.nested_candidate_combo.state(["!disabled"])

    def _clear_nested_candidate_action(self) -> None:
        self.selected_nested_candidate.set("")
        self.nested_candidate_combo["values"] = []
        self.nested_candidate_combo.state(["disabled"])

    def _render_preview(self, preview: TablePreview) -> None:
        self.preview_tree.delete(*self.preview_tree.get_children())
        column_names = [column.name for column in preview.columns[: self.MAX_PREVIEW_COLUMNS]]
        self.current_preview_columns = [column.name for column in preview.columns]
        self._refresh_column_action_columns()
        self.preview_tree["columns"] = column_names

        for column_name in column_names:
            self.preview_tree.heading(column_name, text=column_name)
            self.preview_tree.column(column_name, width=120, minwidth=80, stretch=True)

        for row in preview.preview_rows[:6]:
            values = [self._format_preview_value(row.get(column_name)) for column_name in column_names]
            self.preview_tree.insert("", tk.END, values=values)

        column_summary = ", ".join(
            f"{column.name}: {column.inferred_type}" for column in preview.columns[: self.MAX_PREVIEW_COLUMNS]
        )
        warning_text = f" | {'; '.join(preview.warnings)}" if preview.warnings else ""
        self.preview_label["text"] = f"{preview.rows} rows | {len(preview.columns)} columns | {column_summary}{warning_text}"

    def _clear_preview(self, message: str = "Select an exportable array to preview rows.") -> None:
        self.preview_tree.delete(*self.preview_tree.get_children())
        self.preview_tree["columns"] = ()
        self.current_preview_columns = []
        self._refresh_column_action_columns()
        self.preview_label["text"] = message

    def open_advanced_options(self) -> None:
        self.advanced_options_visible.set(not self.advanced_options_visible.get())
        if self.advanced_section_frame is None:
            return

        if self.advanced_options_visible.get():
            self.advanced_section_frame.grid(row=6, column=0, sticky=(tk.W, tk.E))
            self.advanced_toggle_button["text"] = "Hide Advanced Options"
            self._refresh_column_action_columns()
        else:
            self.advanced_section_frame.grid_remove()
            self.advanced_toggle_button["text"] = "Show Advanced Options"

    def _create_column_actions_tab(self, parent: ttk.Frame) -> None:
        parent.columnconfigure(1, weight=1)

        ttk.Label(parent, text="Column Actions", style="Panel.TLabel").grid(row=0, column=0, columnspan=2, sticky=tk.W)
        ttk.Label(parent, text="Column", style="PanelMuted.TLabel").grid(row=1, column=0, sticky=tk.W, padx=(0, 8), pady=(8, 0))
        self.advanced_column_combo = ttk.Combobox(
            parent,
            textvariable=self.advanced_column,
            state="readonly",
            width=34,
        )
        self.advanced_column_combo.grid(row=1, column=1, sticky=(tk.W, tk.E), pady=(8, 0))
        self.advanced_column_combo.bind("<<ComboboxSelected>>", self._load_column_action)

        ttk.Label(parent, text="Data Type", style="PanelMuted.TLabel").grid(row=2, column=0, sticky=tk.W, padx=(0, 8), pady=(8, 0))
        self.advanced_type_combo = ttk.Combobox(
            parent,
            textvariable=self.advanced_type,
            values=["Keep", "Text", "Number", "Integer", "Boolean"],
            state="readonly",
            width=18,
        )
        self.advanced_type_combo.grid(row=2, column=1, sticky=tk.W, pady=(8, 0))

        ttk.Label(parent, text="Find", style="PanelMuted.TLabel").grid(row=3, column=0, sticky=tk.W, padx=(0, 8), pady=(8, 0))
        ttk.Entry(parent, textvariable=self.advanced_find).grid(row=3, column=1, sticky=(tk.W, tk.E), pady=(8, 0))

        ttk.Label(parent, text="Replace", style="PanelMuted.TLabel").grid(row=4, column=0, sticky=tk.W, padx=(0, 8), pady=(8, 0))
        ttk.Entry(parent, textvariable=self.advanced_replace).grid(row=4, column=1, sticky=(tk.W, tk.E), pady=(8, 0))

        button_frame = ttk.Frame(parent)
        button_frame.grid(row=5, column=0, columnspan=2, sticky=tk.W, pady=(12, 0))
        ttk.Button(button_frame, text="Apply Column Action", command=self._save_column_action).pack(side=tk.LEFT, padx=(0, 8))
        ttk.Button(button_frame, text="Clear Column Action", command=self._clear_column_action).pack(side=tk.LEFT)

        ttk.Label(parent, textvariable=self.advanced_status, style="PanelMuted.TLabel", wraplength=280).grid(
            row=6, column=0, columnspan=2, sticky=tk.W, pady=(10, 0)
        )
        ttk.Label(parent, text="Additional Actions", style="Panel.TLabel").grid(row=7, column=0, columnspan=2, sticky=tk.W, pady=(14, 0))
        ttk.Label(parent, text="Placeholder for future table-level actions.", style="PanelMuted.TLabel").grid(
            row=8, column=0, columnspan=2, sticky=tk.W, pady=(4, 0)
        )

    def _refresh_column_action_columns(self) -> None:
        combo = self.advanced_column_combo
        if combo is None:
            return

        try:
            combo.winfo_exists()
        except tk.TclError:
            self.advanced_column_combo = None
            return

        combo["values"] = self.current_preview_columns
        if not self.current_preview_columns:
            self._reset_column_action_form()
            combo.state(["disabled"])
            return

        combo.state(["!disabled"])
        if self.advanced_column.get() not in self.current_preview_columns:
            self.advanced_column.set(self.current_preview_columns[0])
        self._load_column_action()

    def _load_column_action(self, event: Optional[tk.Event] = None) -> None:
        column = self.advanced_column.get()
        possible_types = self._possible_types_for_column(column)
        if self.advanced_type_combo is not None:
            self.advanced_type_combo["values"] = possible_types

        transform = self.column_transforms.get(column)
        if transform is None:
            self.advanced_type.set("Keep")
            self.advanced_find.set("")
            self.advanced_replace.set("")
            self.advanced_status.set(self._column_type_hint(column, possible_types))
            return

        self.advanced_type.set(transform.data_type if transform.data_type in possible_types else "Keep")
        self.advanced_find.set(transform.find_text)
        self.advanced_replace.set(transform.replace_text)
        self.advanced_status.set(f"{self._column_action_status_text(transform)} | {self._column_type_hint(column, possible_types)}")

    def _save_column_action(self) -> None:
        column = self.advanced_column.get()
        if not column:
            messagebox.showerror("Error", "Please select a column")
            return

        transform = ColumnTransform(
            column=column,
            data_type=self.advanced_type.get(),
            find_text=self.advanced_find.get(),
            replace_text=self.advanced_replace.get(),
        )
        next_transforms = dict(self.column_transforms)
        if transform.data_type == "Keep" and not transform.find_text:
            next_transforms.pop(column, None)
        else:
            next_transforms[column] = transform

        try:
            self._validate_column_transforms(next_transforms)
        except ArrayMateCoreError as e:
            self.advanced_status.set(f"Cannot apply: {e}")
            self.status_label["text"] = str(e)
            self.status_label["foreground"] = "red"
            return

        self.column_transforms = next_transforms
        self.advanced_status.set(
            f"No action set for {column}" if transform.data_type == "Keep" and not transform.find_text else self._column_action_status_text(transform)
        )
        self.refresh_selected_candidate()

    def _clear_column_action(self) -> None:
        column = self.advanced_column.get()
        if column:
            self.column_transforms.pop(column, None)
        self.advanced_type.set("Keep")
        self.advanced_find.set("")
        self.advanced_replace.set("")
        self.advanced_status.set(f"No action set for {column}" if column else "No column selected")
        self.refresh_selected_candidate()

    def _reset_column_action_form(self) -> None:
        self.advanced_column.set("")
        self.advanced_type.set("Keep")
        self.advanced_find.set("")
        self.advanced_replace.set("")
        self.advanced_status.set("No column action selected")

    def _column_action_status_text(self, transform: ColumnTransform) -> str:
        parts = [f"{transform.column}: {transform.data_type}"]
        if transform.find_text:
            parts.append(f"replace '{transform.find_text}' with '{transform.replace_text}'")
        return " | ".join(parts)

    def _possible_types_for_column(self, column: str) -> tuple[str, ...]:
        if not column:
            return ("Keep", "Text")

        array_data = self._current_table_data_without_column_transforms()
        return infer_column_transform_types(array_data, column)

    def _column_type_hint(self, column: str, possible_types: tuple[str, ...]) -> str:
        if not column:
            return "No column selected"
        return f"Possible types for {column}: {', '.join(possible_types)}"

    def _current_table_data_without_column_transforms(self) -> Optional[list[Any]]:
        candidate = self.candidate_by_path.get(self.selected_array_key.get())
        if candidate is None:
            return None

        return self.service.get_table_data(
            candidate.display_path,
            unfold_key=self._unfold_key(),
            include_parent_metadata=self._include_parent_metadata_for(candidate),
            transform_options=TableTransformOptions(
                stringify_all=self.stringify_all.get(),
                stringify_formulas=self.stringify_formulas.get(),
            ),
        )

    def _validate_column_transforms(self, column_transforms: dict[str, ColumnTransform]) -> None:
        candidate = self.candidate_by_path.get(self.selected_array_key.get())
        if candidate is None:
            return

        array_data = self.service.get_table_data(
            candidate.display_path,
            unfold_key=self._unfold_key(),
            include_parent_metadata=self._include_parent_metadata_for(candidate),
            transform_options=TableTransformOptions(
                stringify_all=self.stringify_all.get(),
                stringify_formulas=self.stringify_formulas.get(),
                column_transforms=tuple(column_transforms.values()),
            ),
        )
        if array_data is None:
            raise ArrayMateCoreError("Selected array is invalid")
        build_table_preview(array_data, self.effective_candidate_key or candidate.display_path)

    def _format_preview_value(self, value: Any) -> str:
        if isinstance(value, dict):
            return "[record]"
        if isinstance(value, list):
            return "[table]"
        if value is None:
            return ""
        return str(value)

    def _suggest_filename(self, array_key: str) -> str:
        clean_name = "".join(char if char.isalnum() else "_" for char in array_key).strip("_")
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        return f"{clean_name or 'array'}_{timestamp}"

    def _refresh_auto_filename(self, array_key: Optional[str] = None) -> None:
        if not self.auto_filename:
            return

        selected_key = array_key or self.effective_candidate_key or self.selected_array_key.get()
        if selected_key:
            self.output_filename.set(self._suggest_filename(selected_key))

    def _schedule_auto_filename_refresh(self) -> None:
        if self.auto_filename and self.selected_array_key.get():
            self._refresh_auto_filename()
        self.root.after(1000, self._schedule_auto_filename_refresh)

    def convert_to_file(self) -> None:
        if not self.selected_array_key.get():
            messagebox.showerror("Error", "Please select an array to convert")
            return
        if not self.output_folder.get():
            messagebox.showerror("Error", "Please select an output folder")
            return
        if not self.output_filename.get():
            messagebox.showerror("Error", "Please enter a filename")
            return

        try:
            export_plan = self.service.create_export_plan(
                self.output_folder.get(),
                self.output_filename.get(),
                self.output_format.get(),
            )

            if os.path.exists(export_plan.file_path):
                result = messagebox.askyesno(
                    "File Exists",
                    f"File '{export_plan.filename}' already exists in the selected folder.\nDo you want to overwrite it?",
                )
                if not result:
                    return

            selected_candidate = self.candidate_by_path.get(self.selected_array_key.get())
            unfold_key = self._unfold_key() if selected_candidate else None
            export_result = self.service.export_array(
                self.selected_array_key.get(),
                export_plan,
                include_parent_metadata=bool(
                    selected_candidate and not unfold_key and self._include_parent_metadata_for(selected_candidate)
                ),
                unfold_key=unfold_key,
                transform_options=self._table_transform_options(),
            )
            messagebox.showinfo(
                "Success",
                f"{export_result.output_format.label} file saved successfully!\n"
                f"File: {export_result.file_path}\n"
                f"Rows: {export_result.rows}\n"
                f"Columns: {export_result.columns}",
            )

            self.status_label["text"] = f"{export_result.output_format.label} file saved: {export_result.filename}"
            self.status_label["foreground"] = "green"
            self._open_exported_file(export_result.output_format.label, export_result.file_path)
        except ArrayMateCoreError as e:
            messagebox.showerror("Error", str(e))
            self.status_label["text"] = str(e)
            self.status_label["foreground"] = "red"
        except Exception as e:
            output_format = get_output_format(self.output_format.get())
            messagebox.showerror("Error", f"Error creating {output_format.label} file: {str(e)}")
            self.status_label["text"] = f"Error creating {output_format.label} file"
            self.status_label["foreground"] = "red"

    def _open_exported_file(self, output_label: str, file_path: str) -> None:
        if output_label == "Excel":
            self.open_excel_file(file_path)
        elif output_label == "CSV":
            self.open_csv_file(file_path)
        elif output_label == "JSON":
            self.open_json_file(file_path)
        else:
            self.open_file_location(file_path)

    def open_excel_file(self, file_path: str) -> None:
        self._open_file(file_path, "Excel")

    def open_csv_file(self, file_path: str) -> None:
        self._open_file(file_path, "CSV")

    def open_json_file(self, file_path: str) -> None:
        self._open_file(file_path, "JSON")

    def _open_file(self, file_path: str, label: str) -> None:
        try:
            system = platform.system()
            if system == "Windows":
                os.startfile(file_path)
            elif system == "Darwin":
                subprocess.run(["open", file_path], check=True)
            else:
                subprocess.run(["xdg-open", file_path], check=True)
            self.status_label["text"] = f"{label} file opened: {os.path.basename(file_path)}"
            self.status_label["foreground"] = "green"
        except Exception as e:
            messagebox.showwarning(
                "Warning",
                f"{label} file saved successfully, but could not open automatically.\n"
                f"File location: {file_path}\n"
                f"Error: {str(e)}",
            )
            self.status_label["text"] = f"{label} file saved, but could not open: {os.path.basename(file_path)}"
            self.status_label["foreground"] = "orange"

    def open_file_location(self, file_path: str) -> None:
        try:
            system = platform.system()
            if system == "Windows":
                subprocess.run(["explorer", "/select,", file_path], check=True)
            elif system == "Darwin":
                subprocess.run(["open", "-R", file_path], check=True)
            else:
                subprocess.run(["xdg-open", os.path.dirname(file_path)], check=True)
            self.status_label["text"] = f"File location opened: {os.path.basename(file_path)}"
            self.status_label["foreground"] = "green"
        except Exception:
            messagebox.showinfo("File Saved", f"File saved successfully!\nLocation: {file_path}")
            self.status_label["text"] = f"File saved: {os.path.basename(file_path)}"
            self.status_label["foreground"] = "green"


def main() -> None:
    """Main entry point for the application."""
    root = tk.Tk()
    ArrayMate(root)
    root.mainloop()


if __name__ == "__main__":
    main()
