"""Tkinter desktop UI for ArrayMate."""

from __future__ import annotations

import json
import os
import platform
import subprocess
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from typing import Any, Optional

from arraymate.core import ArrayCandidate, ArrayMateCoreError, TablePreview, build_table_preview, get_output_format
from arraymate.service import ArrayMateService, LoadResult


class ArrayMate:
    """Desktop UI for converting JSON arrays to table files."""

    NO_UNFOLD_LABEL = "Keep selected table"
    WINDOW_WIDTH = 1100
    WINDOW_HEIGHT = 760
    DEFAULT_FONT = ("Arial", 10)
    TITLE_FONT = ("Arial", 16, "bold")
    CODE_FONT = ("Consolas", 10)
    MAX_PREVIEW_COLUMNS = 10

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

        self.json_file_path = tk.StringVar()
        self.selected_array_key = tk.StringVar()
        self.output_folder = tk.StringVar()
        self.output_filename = tk.StringVar()
        self.output_format = tk.StringVar(value="Excel (.xlsx)")
        self.stringify_all = tk.BooleanVar(value=False)
        self.stringify_formulas = tk.BooleanVar(value=False)
        self.include_parent_metadata = tk.BooleanVar(value=False)
        self.selected_nested_candidate = tk.StringVar()

        self.setup_ui()

    def setup_ui(self) -> None:
        """Set up the application window."""
        main_frame = ttk.Frame(self.root, padding="18")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(0, weight=1)

        ttk.Label(main_frame, text="ArrayMate", font=self.TITLE_FONT).grid(row=0, column=0, pady=(0, 14))
        self._create_file_selection_section(main_frame)
        self._create_candidate_section(main_frame)
        self._create_output_settings_section(main_frame)
        self._create_process_buttons(main_frame)
        self._create_status_section(main_frame)

    def _create_file_selection_section(self, parent: ttk.Frame) -> None:
        file_frame = ttk.LabelFrame(parent, text="Step 1: Select JSON Source", padding="10")
        file_frame.grid(row=1, column=0, sticky=(tk.W, tk.E), pady=(0, 10))
        file_frame.columnconfigure(1, weight=1)

        ttk.Label(file_frame, text="JSON File:").grid(row=0, column=0, sticky=tk.W, padx=(0, 10))
        ttk.Entry(file_frame, textvariable=self.json_file_path, width=60).grid(
            row=0,
            column=1,
            sticky=(tk.W, tk.E),
            padx=(0, 10),
        )
        ttk.Button(file_frame, text="Browse", command=self.browse_json_file).grid(row=0, column=2)

        ttk.Label(file_frame, text="OR Paste JSON:").grid(row=1, column=0, sticky=tk.W, padx=(0, 10), pady=(10, 0))
        ttk.Button(file_frame, text="Open JSON Input", command=self.open_json_input_window).grid(
            row=1,
            column=1,
            sticky=tk.W,
            pady=(10, 0),
        )

    def _create_candidate_section(self, parent: ttk.Frame) -> None:
        candidate_frame = ttk.LabelFrame(parent, text="Step 2: Pick a Table Candidate", padding="10")
        candidate_frame.grid(row=2, column=0, sticky=(tk.W, tk.E), pady=(0, 10))
        candidate_frame.columnconfigure(0, weight=1)
        candidate_frame.columnconfigure(1, weight=1)

        ttk.Label(candidate_frame, text="Discovered Arrays:").grid(row=0, column=0, columnspan=2, sticky=tk.W)
        self.array_tree = ttk.Treeview(
            candidate_frame,
            columns=("items", "columns", "status"),
            show="tree headings",
            height=8,
            selectmode="browse",
        )
        self.array_tree.heading("#0", text="Path")
        self.array_tree.heading("items", text="Rows")
        self.array_tree.heading("columns", text="Columns")
        self.array_tree.heading("status", text="Status")
        self.array_tree.column("#0", width=420, minwidth=220, stretch=True)
        self.array_tree.column("items", width=80, minwidth=60, anchor=tk.E, stretch=False)
        self.array_tree.column("columns", width=80, minwidth=60, anchor=tk.E, stretch=False)
        self.array_tree.column("status", width=260, minwidth=140, stretch=True)
        self.array_tree.grid(row=1, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(4, 8))
        self.array_tree.bind("<<TreeviewSelect>>", self.on_array_selected)

        self.array_info_label = ttk.Label(candidate_frame, text="No JSON loaded")
        self.array_info_label.grid(row=2, column=0, columnspan=2, sticky=tk.W, pady=(0, 8))

        options_frame = ttk.LabelFrame(candidate_frame, text="Quick Options", padding="8")
        options_frame.grid(row=3, column=0, sticky=(tk.W, tk.E, tk.N), padx=(0, 8))
        ttk.Checkbutton(options_frame, text="Stringify everything", variable=self.stringify_all, state="disabled").grid(
            row=0,
            column=0,
            sticky=tk.W,
        )
        ttk.Checkbutton(options_frame, text="Stringify formulas", variable=self.stringify_formulas, state="disabled").grid(
            row=1,
            column=0,
            sticky=tk.W,
            pady=(4, 0),
        )
        self.parent_metadata_check = ttk.Checkbutton(
            options_frame,
            text="Include parent metadata",
            variable=self.include_parent_metadata,
            command=self.refresh_selected_candidate,
            state="disabled",
        )
        self.parent_metadata_check.grid(row=2, column=0, sticky=tk.W, pady=(4, 0))
        ttk.Label(options_frame, text="Unfold level:").grid(row=3, column=0, sticky=tk.W, pady=(8, 0))
        nested_action_frame = ttk.Frame(options_frame)
        nested_action_frame.grid(row=4, column=0, sticky=(tk.W, tk.E), pady=(4, 0))
        nested_action_frame.columnconfigure(0, weight=1)
        self.nested_candidate_combo = ttk.Combobox(
            nested_action_frame,
            textvariable=self.selected_nested_candidate,
            state="disabled",
            width=34,
        )
        self.nested_candidate_combo.grid(row=0, column=0, sticky=(tk.W, tk.E))
        self.nested_candidate_combo.bind("<<ComboboxSelected>>", self.refresh_selected_candidate)
        self.nested_candidate_button = ttk.Button(
            nested_action_frame,
            text="Apply",
            command=self.refresh_selected_candidate,
            state="disabled",
        )
        self.nested_candidate_button.grid(row=0, column=1, padx=(6, 0))
        ttk.Label(options_frame, text="Placeholder: string parsing options will be wired later.").grid(
            row=5,
            column=0,
            sticky=tk.W,
            pady=(8, 0),
        )

        advanced_frame = ttk.LabelFrame(candidate_frame, text="Advanced Column Types", padding="8")
        advanced_frame.grid(row=3, column=1, sticky=(tk.W, tk.E, tk.N))
        ttk.Label(advanced_frame, text="Placeholder: per-column type controls will appear here.").grid(
            row=0,
            column=0,
            sticky=tk.W,
        )

        preview_frame = ttk.LabelFrame(candidate_frame, text="Preview", padding="8")
        preview_frame.grid(row=4, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(10, 0))
        preview_frame.columnconfigure(0, weight=1)

        self.preview_tree = ttk.Treeview(preview_frame, show="headings", height=6)
        self.preview_tree.grid(row=0, column=0, sticky=(tk.W, tk.E))
        self.preview_label = ttk.Label(preview_frame, text="Select an exportable array to preview rows.")
        self.preview_label.grid(row=1, column=0, sticky=tk.W, pady=(6, 0))

    def _create_output_settings_section(self, parent: ttk.Frame) -> None:
        output_frame = ttk.LabelFrame(parent, text="Step 3: Set Output Location", padding="10")
        output_frame.grid(row=3, column=0, sticky=(tk.W, tk.E), pady=(0, 10))
        output_frame.columnconfigure(1, weight=1)

        ttk.Label(output_frame, text="Output Format:").grid(row=0, column=0, sticky=tk.W, padx=(0, 10))
        format_combobox = ttk.Combobox(
            output_frame,
            textvariable=self.output_format,
            values=["Excel (.xlsx)", "CSV (.csv)", "JSON (.json)"],
            state="readonly",
            width=15,
        )
        format_combobox.grid(row=0, column=1, sticky=tk.W)
        format_combobox.bind("<<ComboboxSelected>>", self.on_format_selected)

        ttk.Label(output_frame, text="Save Folder:").grid(row=1, column=0, sticky=tk.W, padx=(0, 10), pady=(10, 0))
        ttk.Entry(output_frame, textvariable=self.output_folder, width=60).grid(
            row=1,
            column=1,
            sticky=(tk.W, tk.E),
            padx=(0, 10),
            pady=(10, 0),
        )
        ttk.Button(output_frame, text="Browse", command=self.browse_output_folder).grid(row=1, column=2, pady=(10, 0))

        ttk.Label(output_frame, text="File Name:").grid(row=2, column=0, sticky=tk.W, padx=(0, 10), pady=(10, 0))
        ttk.Entry(output_frame, textvariable=self.output_filename, width=60).grid(
            row=2,
            column=1,
            sticky=(tk.W, tk.E),
            padx=(0, 10),
            pady=(10, 0),
        )
        self.extension_label = ttk.Label(output_frame, text=".xlsx")
        self.extension_label.grid(row=2, column=2, sticky=tk.W, pady=(10, 0))

    def _create_process_buttons(self, parent: ttk.Frame) -> None:
        button_frame = ttk.Frame(parent)
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
        status_frame = ttk.LabelFrame(parent, text="Status & Events", padding="10")
        status_frame.grid(row=5, column=0, sticky=(tk.W, tk.E), pady=(6, 0))
        status_frame.columnconfigure(0, weight=1)
        self.status_label = ttk.Label(
            status_frame,
            text="Ready to convert JSON arrays",
            font=self.DEFAULT_FONT,
            foreground="green",
        )
        self.status_label.grid(row=0, column=0, sticky=tk.W)

    def clear_all(self) -> None:
        """Clear inputs and reset application state."""
        self.json_file_path.set("")
        self.selected_array_key.set("")
        self.output_folder.set("")
        self.output_filename.set("")
        self.output_format.set("Excel (.xlsx)")
        self.extension_label["text"] = ".xlsx"
        self.process_button["text"] = "Convert to File"
        self.process_button["state"] = "disabled"

        self.json_data = None
        self.array_keys = []
        self.candidate_by_path = {}
        self.effective_candidate_key = ""
        self.service.clear()
        self.array_tree.delete(*self.array_tree.get_children())
        self._clear_preview()
        self.array_info_label["text"] = "No JSON loaded"
        self.status_label["text"] = "Ready to convert JSON arrays"
        self.status_label["foreground"] = "green"
        self.include_parent_metadata.set(False)
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
        self.json_input_window = tk.Toplevel(self.root)
        self.json_input_window.title("Paste JSON Data")
        self.json_input_window.geometry("700x460")
        self.json_input_window.resizable(True, True)
        self.json_input_window.columnconfigure(0, weight=1)
        self.json_input_window.rowconfigure(1, weight=1)

        ttk.Label(
            self.json_input_window,
            text="Paste your JSON data below:",
            font=self.DEFAULT_FONT,
        ).grid(row=0, column=0, sticky=tk.W, padx=10, pady=(10, 5))

        self.json_text = tk.Text(self.json_input_window, wrap=tk.WORD, font=self.CODE_FONT)
        self.json_text.grid(row=1, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), padx=10, pady=(0, 10))

        scrollbar = ttk.Scrollbar(self.json_input_window, orient=tk.VERTICAL, command=self.json_text.yview)
        scrollbar.grid(row=1, column=1, sticky=(tk.N, tk.S))
        self.json_text.configure(yscrollcommand=scrollbar.set)

        button_frame = ttk.Frame(self.json_input_window)
        button_frame.grid(row=2, column=0, columnspan=2, pady=(0, 10))
        ttk.Button(button_frame, text="Load JSON", command=self.load_json_from_text).pack(side=tk.LEFT, padx=(0, 10))
        ttk.Button(button_frame, text="Clear", command=self.clear_json_text).pack(side=tk.LEFT, padx=(0, 10))
        ttk.Button(button_frame, text="Cancel", command=self.json_input_window.destroy).pack(side=tk.LEFT)
        self.json_text.focus()

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
                self.json_input_window.destroy()
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
            if self.selected_array_key.get():
                self.output_filename.set(self._suggest_filename(self.selected_array_key.get()))

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

    def use_selected_nested_candidate(self) -> None:
        self.refresh_selected_candidate()

    def _select_candidate(self, candidate: ArrayCandidate, reset_unfold: bool = True) -> None:
        previous_key = self.selected_array_key.get()
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
        if self.output_folder.get():
            self.output_filename.set(self._suggest_filename(effective_candidate.display_path))

        array_data = self.service.get_table_data(
            candidate.display_path,
            unfold_key=unfold_key,
            include_parent_metadata=self._include_parent_metadata_for(candidate),
        )
        if effective_candidate.exportable and array_data is not None:
            try:
                preview = build_table_preview(array_data, effective_candidate.display_path)
                self.array_info_label["text"] = self._candidate_detail_text(candidate, effective_candidate, preview)
                self._render_preview(preview)
                self.process_button["state"] = "normal"
            except ArrayMateCoreError as e:
                self._clear_preview(str(e))
                self.process_button["state"] = "disabled"
        else:
            self.array_info_label["text"] = self._candidate_detail_text(candidate, effective_candidate)
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
        self.nested_candidate_button.state(["!disabled"])

    def _clear_nested_candidate_action(self) -> None:
        self.selected_nested_candidate.set("")
        self.nested_candidate_combo["values"] = []
        self.nested_candidate_combo.state(["disabled"])
        self.nested_candidate_button.state(["disabled"])

    def _render_preview(self, preview: TablePreview) -> None:
        self.preview_tree.delete(*self.preview_tree.get_children())
        column_names = [column.name for column in preview.columns[: self.MAX_PREVIEW_COLUMNS]]
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
        self.preview_label["text"] = message

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
        return f"{clean_name or 'array'}_data"

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
