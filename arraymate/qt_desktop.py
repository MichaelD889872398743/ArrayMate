"""PySide6 desktop UI for ArrayMate."""

from __future__ import annotations

import json
import os
import platform
import subprocess
import sys
import webbrowser
from datetime import datetime
from typing import Any, Optional

from PySide6.QtCore import Qt, QTimer
from PySide6.QtWidgets import (
    QApplication,
    QAbstractScrollArea,
    QAbstractItemView,
    QCheckBox,
    QComboBox,
    QFileDialog,
    QFrame,
    QGridLayout,
    QGroupBox,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QLineEdit,
    QMainWindow,
    QMessageBox,
    QPushButton,
    QSizePolicy,
    QTableWidget,
    QTableWidgetItem,
    QTextEdit,
    QTreeWidget,
    QTreeWidgetItem,
    QVBoxLayout,
    QWidget,
)

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


class ArrayMateWindow(QMainWindow):
    """Modern Qt desktop UI for converting JSON arrays to table files."""

    NO_UNFOLD_LABEL = "Keep selected table"
    REPOSITORY_URL = "https://github.com/MichaelD889872398743/ArrayMate"
    MAX_PREVIEW_COLUMNS = 10

    def __init__(self) -> None:
        super().__init__()
        self.setWindowTitle("ArrayMate")
        self.resize(1180, 720)

        self.service = ArrayMateService()
        self.candidate_by_path: dict[str, ArrayCandidate] = {}
        self.selected_array_key = ""
        self.effective_candidate_key = ""
        self.column_transforms: dict[str, ColumnTransform] = {}
        self.current_preview_columns: list[str] = []
        self.auto_filename = True
        self.suppress_text_auto_parse = False

        self._build_ui()
        self.json_parse_timer = QTimer(self)
        self.json_parse_timer.setSingleShot(True)
        self.json_parse_timer.setInterval(700)
        self.json_parse_timer.timeout.connect(self._auto_load_json_from_text)
        self.json_text.textChanged.connect(self._schedule_json_auto_parse)
        self._apply_styles()

    def _build_ui(self) -> None:
        root = QWidget()
        root_layout = QVBoxLayout(root)
        root_layout.setContentsMargins(0, 0, 0, 0)
        root_layout.setSpacing(0)
        self.setCentralWidget(root)

        root_layout.addWidget(self._source_bar())

        workspace = QFrame()
        workspace.setObjectName("workspace")
        workspace_layout = QHBoxLayout(workspace)
        workspace_layout.setContentsMargins(0, 0, 0, 0)
        workspace_layout.setSpacing(0)
        root_layout.addWidget(workspace, 1)

        workspace_layout.addWidget(self._structure_pane(), 0)
        workspace_layout.addWidget(self._preview_pane(), 1)
        workspace_layout.addWidget(self._right_pane(), 0)

        self.status_label = QLabel("Ready to convert JSON arrays")
        self.status_label.setObjectName("statusBar")
        self.status_label.setMinimumHeight(26)
        root_layout.addWidget(self.status_label)

    def _source_bar(self) -> QWidget:
        header = QFrame()
        header.setObjectName("header")
        layout = QHBoxLayout(header)
        layout.setContentsMargins(16, 10, 16, 10)
        layout.setSpacing(8)

        title = QLabel("ArrayMate")
        title.setObjectName("appTitle")
        layout.addWidget(title)

        step = QLabel("Step 1")
        step.setObjectName("mutedOnPanel")
        layout.addWidget(step)

        self.file_path_edit = QLineEdit()
        self.file_path_edit.setPlaceholderText("Loaded file path")
        layout.addWidget(self.file_path_edit, 1)

        load_button = QPushButton("Load JSON File")
        load_button.setObjectName("primaryButton")
        load_button.clicked.connect(self.browse_json_file)
        layout.addWidget(load_button)
        return header

    def _structure_pane(self) -> QWidget:
        pane = QFrame()
        pane.setObjectName("panel")
        pane.setFixedWidth(310)
        layout = QVBoxLayout(pane)
        layout.setContentsMargins(10, 10, 0, 10)
        layout.setSpacing(8)

        title = QLabel("STEP 2 - PARSED STRUCTURE")
        title.setObjectName("paneTitle")
        layout.addWidget(title)

        self.array_tree = QTreeWidget()
        self.array_tree.setColumnCount(4)
        self.array_tree.setHeaderLabels(["Path", "Rows", "Cols", "Status"])
        self.array_tree.setRootIsDecorated(False)
        self.array_tree.setAlternatingRowColors(False)
        self.array_tree.itemSelectionChanged.connect(self.on_array_selected)
        self.array_tree.setColumnWidth(0, 150)
        self.array_tree.setColumnWidth(1, 48)
        self.array_tree.setColumnWidth(2, 44)
        layout.addWidget(self.array_tree, 1)
        return pane

    def _preview_pane(self) -> QWidget:
        pane = QFrame()
        pane.setObjectName("center")
        pane.setMinimumWidth(0)
        pane.setSizePolicy(QSizePolicy.Policy.Ignored, QSizePolicy.Policy.Expanding)
        layout = QVBoxLayout(pane)
        layout.setContentsMargins(14, 14, 14, 14)
        layout.setSpacing(10)

        info_row = QHBoxLayout()
        self.array_info_label = QLabel("No JSON loaded")
        self.array_info_label.setObjectName("muted")
        self.array_info_label.setMinimumWidth(0)
        self.array_info_label.setSizePolicy(QSizePolicy.Policy.Ignored, QSizePolicy.Policy.Preferred)
        info_row.addWidget(self.array_info_label, 1)
        layout.addLayout(info_row)

        self.json_input_frame = self._json_input_panel()
        layout.addWidget(self.json_input_frame)

        preview_card = QFrame()
        preview_card.setObjectName("card")
        preview_card.setMinimumWidth(0)
        preview_card.setSizePolicy(QSizePolicy.Policy.Ignored, QSizePolicy.Policy.Expanding)
        preview_layout = QVBoxLayout(preview_card)
        preview_layout.setContentsMargins(0, 0, 0, 0)
        preview_layout.setSpacing(0)

        preview_header = QFrame()
        preview_header.setObjectName("cardHeader")
        header_layout = QHBoxLayout(preview_header)
        header_layout.setContentsMargins(12, 9, 12, 9)
        header_layout.addWidget(QLabel("Preview"))
        self.preview_label = QLabel("Select an exportable array to preview rows.")
        self.preview_label.setObjectName("mutedOnPanel")
        self.preview_label.setMinimumWidth(0)
        self.preview_label.setSizePolicy(QSizePolicy.Policy.Ignored, QSizePolicy.Policy.Preferred)
        header_layout.addWidget(self.preview_label, 1, Qt.AlignmentFlag.AlignRight)
        preview_layout.addWidget(preview_header)

        self.cell_preview = QTextEdit()
        self.cell_preview.setObjectName("cellInspector")
        self.cell_preview.setReadOnly(True)
        self.cell_preview.setMaximumHeight(74)
        self.cell_preview.setPlaceholderText("Select a preview cell to inspect the full value.")
        self.cell_preview.setSizePolicy(QSizePolicy.Policy.Ignored, QSizePolicy.Policy.Fixed)
        preview_layout.addWidget(self.cell_preview)

        self.preview_table = QTableWidget()
        self.preview_table.setEditTriggers(QTableWidget.EditTrigger.NoEditTriggers)
        self.preview_table.setSelectionBehavior(QTableWidget.SelectionBehavior.SelectRows)
        self.preview_table.setAlternatingRowColors(False)
        self.preview_table.setWordWrap(False)
        self.preview_table.setCornerButtonEnabled(False)
        self.preview_table.setHorizontalScrollMode(QAbstractItemView.ScrollMode.ScrollPerPixel)
        self.preview_table.setVerticalScrollMode(QAbstractItemView.ScrollMode.ScrollPerPixel)
        self.preview_table.setSizeAdjustPolicy(QAbstractScrollArea.SizeAdjustPolicy.AdjustIgnored)
        self.preview_table.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)
        self.preview_table.setVerticalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)
        self.preview_table.setSizePolicy(QSizePolicy.Policy.Ignored, QSizePolicy.Policy.Expanding)
        self.preview_table.setMinimumWidth(0)
        self.preview_table.verticalHeader().setVisible(False)
        self.preview_table.horizontalHeader().setStretchLastSection(False)
        self.preview_table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Interactive)
        self.preview_table.horizontalHeader().setDefaultSectionSize(150)
        self.preview_table.horizontalHeader().setMinimumSectionSize(80)
        self.preview_table.itemSelectionChanged.connect(self._update_cell_preview)
        preview_layout.addWidget(self.preview_table, 1)

        layout.addWidget(preview_card, 1)
        return pane

    def _json_input_panel(self) -> QWidget:
        panel = QFrame()
        panel.setObjectName("card")
        layout = QVBoxLayout(panel)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(0)

        header = QFrame()
        header.setObjectName("cardHeader")
        header_layout = QHBoxLayout(header)
        header_layout.setContentsMargins(12, 9, 12, 9)
        header_layout.addWidget(QLabel("JSON Input"), 1)
        clear_button = QPushButton("Clear")
        clear_button.clicked.connect(self.clear_json_text)
        header_layout.addWidget(clear_button)
        layout.addWidget(header)

        body = QFrame()
        body.setObjectName("panelBody")
        body_layout = QVBoxLayout(body)
        body_layout.setContentsMargins(12, 0, 12, 12)
        self.json_text = QTextEdit()
        self.json_text.setPlaceholderText("Paste JSON here...")
        self.json_text.setMinimumHeight(150)
        self.json_text.setObjectName("jsonEditor")
        body_layout.addWidget(self.json_text)
        layout.addWidget(body)
        return panel

    def _right_pane(self) -> QWidget:
        pane = QFrame()
        pane.setObjectName("panel")
        pane.setFixedWidth(340)
        layout = QVBoxLayout(pane)
        layout.setContentsMargins(14, 14, 14, 14)
        layout.setSpacing(12)

        title = QLabel("STEP 3 - TRANSFORM & EXPORT")
        title.setObjectName("paneTitle")
        layout.addWidget(title)

        layout.addWidget(self._export_section())
        layout.addWidget(self._transform_section())
        layout.addWidget(self._warnings_section())
        layout.addStretch(1)
        layout.addWidget(self._project_section())
        return pane

    def _export_section(self) -> QWidget:
        group = QGroupBox("Export")
        form = QGridLayout(group)
        form.setContentsMargins(10, 18, 10, 10)
        form.setVerticalSpacing(8)

        form.addWidget(self._muted_label("Output format"), 0, 0, 1, 2)
        self.output_format_combo = QComboBox()
        self.output_format_combo.addItems(["Excel (.xlsx)", "CSV (.csv)", "JSON (.json)"])
        self.output_format_combo.setCurrentText("Excel (.xlsx)")
        self.output_format_combo.currentTextChanged.connect(self.on_format_selected)
        form.addWidget(self.output_format_combo, 1, 0, 1, 2)

        form.addWidget(self._muted_label("File name"), 2, 0, 1, 2)
        self.output_filename_edit = QLineEdit()
        self.output_filename_edit.textEdited.connect(self.on_filename_edited)
        form.addWidget(self.output_filename_edit, 3, 0)
        self.extension_label = QLabel(".xlsx")
        self.extension_label.setObjectName("mutedOnPanel")
        form.addWidget(self.extension_label, 3, 1)

        form.addWidget(self._muted_label("Save folder"), 4, 0, 1, 2)
        self.output_folder_edit = QLineEdit()
        form.addWidget(self.output_folder_edit, 5, 0)
        folder_button = QPushButton("Browse")
        folder_button.clicked.connect(self.browse_output_folder)
        form.addWidget(folder_button, 5, 1)

        self.process_button = QPushButton("Convert to File")
        self.process_button.setObjectName("primaryButton")
        self.process_button.setEnabled(False)
        self.process_button.clicked.connect(self.convert_to_file)
        form.addWidget(self.process_button, 6, 0)
        clear_button = QPushButton("Clear")
        clear_button.clicked.connect(self.clear_all)
        form.addWidget(clear_button, 6, 1)
        return group

    def _transform_section(self) -> QWidget:
        group = QGroupBox("Transform Options")
        layout = QVBoxLayout(group)
        layout.setContentsMargins(10, 18, 10, 10)
        layout.setSpacing(8)

        layout.addWidget(self._muted_label("Unfold level"))
        self.nested_candidate_combo = QComboBox()
        self.nested_candidate_combo.setEnabled(False)
        self.nested_candidate_combo.currentTextChanged.connect(self.refresh_selected_candidate)
        layout.addWidget(self.nested_candidate_combo)

        self.stringify_all_check = QCheckBox("Stringify everything")
        self.stringify_all_check.stateChanged.connect(self.refresh_selected_candidate)
        layout.addWidget(self.stringify_all_check)
        self.stringify_formulas_check = QCheckBox("Stringify formulas")
        self.stringify_formulas_check.stateChanged.connect(self.refresh_selected_candidate)
        layout.addWidget(self.stringify_formulas_check)
        self.parent_metadata_check = QCheckBox("Include parent metadata")
        self.parent_metadata_check.setEnabled(False)
        self.parent_metadata_check.stateChanged.connect(self.refresh_selected_candidate)
        layout.addWidget(self.parent_metadata_check)

        self.advanced_toggle_button = QPushButton("Show Advanced Options")
        self.advanced_toggle_button.clicked.connect(self.open_advanced_options)
        layout.addWidget(self.advanced_toggle_button)

        self.advanced_section_frame = self._advanced_section()
        self.advanced_section_frame.setVisible(False)
        layout.addWidget(self.advanced_section_frame)
        return group

    def _advanced_section(self) -> QWidget:
        section = QFrame()
        section.setObjectName("inlinePanel")
        layout = QGridLayout(section)
        layout.setContentsMargins(0, 10, 0, 0)
        layout.setVerticalSpacing(8)

        header = QLabel("Column Actions")
        layout.addWidget(header, 0, 0, 1, 2)
        layout.addWidget(self._muted_label("Column"), 1, 0)
        self.advanced_column_combo = QComboBox()
        self.advanced_column_combo.currentTextChanged.connect(self._load_column_action)
        layout.addWidget(self.advanced_column_combo, 1, 1)

        layout.addWidget(self._muted_label("Data Type"), 2, 0)
        self.advanced_type_combo = QComboBox()
        self.advanced_type_combo.addItems(["Keep", "Text", "Number", "Integer", "Boolean"])
        layout.addWidget(self.advanced_type_combo, 2, 1)

        layout.addWidget(self._muted_label("Find"), 3, 0)
        self.advanced_find_edit = QLineEdit()
        layout.addWidget(self.advanced_find_edit, 3, 1)

        layout.addWidget(self._muted_label("Replace"), 4, 0)
        self.advanced_replace_edit = QLineEdit()
        layout.addWidget(self.advanced_replace_edit, 4, 1)

        apply_button = QPushButton("Apply Column Action")
        apply_button.clicked.connect(self._save_column_action)
        layout.addWidget(apply_button, 5, 0)
        clear_button = QPushButton("Clear Column Action")
        clear_button.clicked.connect(self._clear_column_action)
        layout.addWidget(clear_button, 5, 1)

        self.advanced_status_label = QLabel("No column action selected")
        self.advanced_status_label.setObjectName("mutedOnPanel")
        self.advanced_status_label.setWordWrap(True)
        layout.addWidget(self.advanced_status_label, 6, 0, 1, 2)
        layout.addWidget(QLabel("Additional Actions"), 7, 0, 1, 2)
        placeholder = self._muted_label("Placeholder for future table-level actions.")
        placeholder.setWordWrap(True)
        layout.addWidget(placeholder, 8, 0, 1, 2)
        return section

    def _warnings_section(self) -> QWidget:
        group = QGroupBox("Detected Warnings")
        layout = QVBoxLayout(group)
        layout.setContentsMargins(10, 18, 10, 10)
        self.warning_label = self._muted_label("JSON not parsed yet")
        self.warning_label.setWordWrap(True)
        layout.addWidget(self.warning_label)
        return group

    def _project_section(self) -> QWidget:
        group = QGroupBox("Project")
        layout = QVBoxLayout(group)
        layout.setContentsMargins(10, 18, 10, 10)
        layout.addWidget(self._muted_label("ArrayMate on GitHub"))
        repo_button = QPushButton("Open Repository")
        repo_button.clicked.connect(self.open_repository)
        layout.addWidget(repo_button)
        return group

    def _muted_label(self, text: str) -> QLabel:
        label = QLabel(text)
        label.setObjectName("mutedOnPanel")
        return label

    def _apply_styles(self) -> None:
        self.setStyleSheet(
            """
            QLabel { color: #d4d4d4; }
            QMainWindow, #workspace, #center { background: #1e1e1e; color: #d4d4d4; }
            #header, #panel, QGroupBox, #card, #cardHeader, #panelBody { background: #252526; color: #d4d4d4; }
            #panel { border-left: 1px solid #3c3c3c; border-right: 1px solid #3c3c3c; }
            #appTitle { font-size: 18px; font-weight: 700; }
            #paneTitle { color: #9da5b4; font-size: 11px; font-weight: 700; letter-spacing: 1px; }
            #muted, #mutedOnPanel { color: #9da5b4; }
            #card { border: 1px solid #3c3c3c; border-radius: 6px; }
            #cardHeader { border-bottom: 1px solid #3c3c3c; }
            QGroupBox { border: 1px solid #3c3c3c; border-radius: 6px; margin-top: 8px; padding-top: 8px; }
            QGroupBox::title { subcontrol-origin: margin; left: 10px; padding: 0 4px; color: #d4d4d4; }
            QLineEdit, QComboBox, QTextEdit, QTableWidget, QTreeWidget {
                background: #1b1b1b; color: #d4d4d4; border: 1px solid #3c3c3c; border-radius: 4px;
                selection-background-color: #094771;
            }
            QTableWidget, QTreeWidget {
                gridline-color: #3c3c3c;
                alternate-background-color: #1b1b1b;
            }
            QTableWidget::viewport, QTreeWidget::viewport {
                background: #1b1b1b;
            }
            #cellInspector {
                background: #1b1b1b;
                color: #d4d4d4;
                border: 0;
                border-bottom: 1px solid #3c3c3c;
                font-family: Consolas;
                padding: 6px;
            }
            QTableWidget QTableCornerButton::section, QTreeWidget QTableCornerButton::section {
                background: #202020;
                border: 0;
                border-right: 1px solid #3c3c3c;
                border-bottom: 1px solid #3c3c3c;
            }
            QHeaderView { background: #202020; }
            QHeaderView::section { background: #202020; color: #9da5b4; border: 0; border-right: 1px solid #3c3c3c; border-bottom: 1px solid #3c3c3c; padding: 6px; }
            QScrollBar:horizontal, QScrollBar:vertical { background: #252526; border: 0; }
            QScrollBar::handle:horizontal, QScrollBar::handle:vertical { background: #4a4a4a; border-radius: 4px; min-width: 24px; min-height: 24px; }
            QScrollBar::handle:horizontal:hover, QScrollBar::handle:vertical:hover { background: #5a5a5a; }
            QScrollBar::add-line, QScrollBar::sub-line { width: 0; height: 0; }
            QPushButton { background: #2d2d30; color: #d4d4d4; border: 1px solid #3c3c3c; border-radius: 4px; padding: 6px 9px; }
            QPushButton:hover { background: #3a3d41; }
            QPushButton:disabled { color: #777; background: #252526; }
            #primaryButton { background: #007acc; border-color: #007acc; color: white; }
            #primaryButton:hover { background: #1688d1; }
            QCheckBox { color: #d4d4d4; spacing: 8px; }
            #jsonEditor { font-family: Consolas; }
            #statusBar { background: #007acc; color: white; padding-left: 10px; }
            QMessageBox { background: #252526; color: #d4d4d4; }
            QMessageBox QLabel { color: #d4d4d4; }
            QMessageBox QPushButton { min-width: 72px; }
            """
        )

    def browse_json_file(self) -> None:
        file_path, _ = QFileDialog.getOpenFileName(self, "Select JSON File", "", "JSON files (*.json);;All files (*.*)")
        if file_path:
            self.file_path_edit.setText(file_path)
            self.load_json_file()

    def load_json_file(self) -> None:
        file_path = self.file_path_edit.text().strip()
        if not file_path:
            self.browse_json_file()
            return

        try:
            with open(file_path, "r", encoding="utf-8") as file:
                json_text = file.read()
            self.suppress_text_auto_parse = True
            self.json_text.setPlainText(json_text)
            self.suppress_text_auto_parse = False
            self._load_json_from_text(show_errors=True, source_label="JSON file", clear_file_path=False)
        except Exception as e:
            QMessageBox.critical(self, "Error loading file", f"Error loading file: {e}")
            self.status_label.setText("Error loading file")

    def _schedule_json_auto_parse(self) -> None:
        if self.suppress_text_auto_parse:
            return
        if not self.json_text.toPlainText().strip():
            return
        self.json_parse_timer.start()

    def _auto_load_json_from_text(self) -> None:
        self._load_json_from_text(show_errors=False)

    def load_json_from_text(self) -> None:
        self._load_json_from_text(show_errors=True)

    def _load_json_from_text(
        self,
        show_errors: bool,
        source_label: str = "JSON data",
        clear_file_path: bool = True,
    ) -> None:
        json_text = self.json_text.toPlainText().strip()
        if not json_text:
            return

        try:
            load_result = self.service.load_text(json_text)
            if clear_file_path:
                self.file_path_edit.setText("")
            self._apply_load_result(load_result, source_label)
            if not load_result.array_candidates:
                self.warning_label.setText("No arrays found in the JSON data")
                self.status_label.setText("No arrays found in JSON data")
                if show_errors:
                    QMessageBox.warning(self, "No arrays found", "No arrays found in the JSON data")
        except json.JSONDecodeError as e:
            self.warning_label.setText(f"Invalid JSON format: {e}")
            self.status_label.setText("Waiting for valid JSON input")
            if show_errors:
                QMessageBox.critical(self, "Invalid JSON", f"Invalid JSON format: {e}")
        except Exception as e:
            self.warning_label.setText(f"Error parsing JSON: {e}")
            self.status_label.setText("Error parsing JSON")
            if show_errors:
                QMessageBox.critical(self, "Error parsing JSON", f"Error parsing JSON: {e}")

    def clear_json_text(self) -> None:
        self.suppress_text_auto_parse = True
        self.json_text.clear()
        self.suppress_text_auto_parse = False
        self.file_path_edit.clear()
        self.candidate_by_path = {}
        self.selected_array_key = ""
        self.effective_candidate_key = ""
        self.column_transforms = {}
        self.current_preview_columns = []
        self.service.clear()
        self.array_tree.clear()
        self._clear_preview()
        self.array_info_label.setText("No JSON loaded")
        self.warning_label.setText("JSON not parsed yet")
        self.status_label.setText("Ready to convert JSON arrays")
        self.process_button.setEnabled(False)
        self.parent_metadata_check.setChecked(False)
        self.parent_metadata_check.setEnabled(False)
        self._reset_column_action_form()
        self._clear_nested_candidate_action()

    def clear_all(self) -> None:
        self.file_path_edit.clear()
        self.output_folder_edit.clear()
        self.output_filename_edit.clear()
        self.output_format_combo.setCurrentText("Excel (.xlsx)")
        self.extension_label.setText(".xlsx")
        self.process_button.setText("Convert to File")
        self.process_button.setEnabled(False)
        self.auto_filename = True
        self.candidate_by_path = {}
        self.selected_array_key = ""
        self.effective_candidate_key = ""
        self.column_transforms = {}
        self.current_preview_columns = []
        self.service.clear()
        self.array_tree.clear()
        self._clear_preview()
        self.array_info_label.setText("No JSON loaded")
        self.warning_label.setText("JSON not parsed yet")
        self.status_label.setText("Ready to convert JSON arrays")
        self.stringify_all_check.setChecked(False)
        self.stringify_formulas_check.setChecked(False)
        self.parent_metadata_check.setChecked(False)
        self.parent_metadata_check.setEnabled(False)
        self._reset_column_action_form()
        self._clear_nested_candidate_action()
        self.clear_json_text()

    def _apply_load_result(self, load_result: LoadResult, source_label: str) -> None:
        self.candidate_by_path = {candidate.display_path: candidate for candidate in load_result.array_candidates}
        self.column_transforms = {}
        self._reset_column_action_form()
        self.array_tree.clear()
        self._clear_preview()

        for candidate in load_result.array_candidates:
            item = QTreeWidgetItem(
                [
                    candidate.display_path,
                    str(candidate.item_count),
                    str(candidate.column_count or ""),
                    self._candidate_status(candidate),
                ]
            )
            item.setData(0, Qt.ItemDataRole.UserRole, candidate.display_path)
            self.array_tree.addTopLevelItem(item)

        selected_candidate = self._default_candidate(load_result.array_candidates)
        if selected_candidate:
            items = self.array_tree.findItems(selected_candidate.display_path, Qt.MatchFlag.MatchExactly, 0)
            if items:
                self.array_tree.setCurrentItem(items[0])
            self._select_candidate(selected_candidate)
            self.status_label.setText(f"Found {len(load_result.array_candidates)} array candidate(s) in {source_label}")
        else:
            self.selected_array_key = ""
            self.array_info_label.setText(f"No arrays found in {source_label}")
            self.process_button.setEnabled(False)
            self.status_label.setText(f"No arrays found in {source_label}")

    def _default_candidate(self, candidates: list[ArrayCandidate]) -> Optional[ArrayCandidate]:
        return next((candidate for candidate in candidates if candidate.exportable), candidates[0] if candidates else None)

    def _candidate_status(self, candidate: ArrayCandidate) -> str:
        if candidate.exportable and candidate.warning:
            return f"Exportable, {candidate.warning.lower()}"
        if candidate.exportable:
            return "Exportable"
        return candidate.warning or "Not exportable"

    def on_array_selected(self) -> None:
        item = self.array_tree.currentItem()
        if item is None:
            return
        key = item.data(0, Qt.ItemDataRole.UserRole)
        candidate = self.candidate_by_path.get(key)
        if candidate:
            self._select_candidate(candidate)

    def refresh_selected_candidate(self) -> None:
        candidate = self.candidate_by_path.get(self.selected_array_key)
        if candidate:
            self._select_candidate(candidate, reset_unfold=False)

    def _select_candidate(self, candidate: ArrayCandidate, reset_unfold: bool = True) -> None:
        previous_key = self.selected_array_key
        if previous_key and previous_key != candidate.display_path:
            self.column_transforms = {}
            self._reset_column_action_form()
        self.selected_array_key = candidate.display_path
        self._update_nested_candidate_action(candidate, reset_unfold=reset_unfold or previous_key != candidate.display_path)
        unfold_key = self._unfold_key()
        effective_candidate = self.candidate_by_path.get(unfold_key, candidate) if unfold_key else candidate
        if unfold_key:
            self.parent_metadata_check.setChecked(False)
            self.parent_metadata_check.setEnabled(False)
        else:
            self._update_parent_metadata_option(candidate, is_new_selection=self.effective_candidate_key != candidate.display_path)
        self.effective_candidate_key = effective_candidate.display_path
        if not self.output_filename_edit.text():
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
                self.array_info_label.setText(self._candidate_detail_text(candidate, effective_candidate, preview))
                self.warning_label.setText(self._warning_text(effective_candidate, preview))
                self._render_preview(preview)
                self.process_button.setEnabled(True)
            except ArrayMateCoreError as e:
                self.warning_label.setText(str(e))
                self._clear_preview(str(e))
                self.process_button.setEnabled(False)
        else:
            self.array_info_label.setText(self._candidate_detail_text(candidate, effective_candidate))
            self.warning_label.setText(effective_candidate.warning or "This array is not exportable as a table.")
            self._clear_preview(effective_candidate.warning or "This array is not exportable as a table.")
            self.process_button.setEnabled(False)

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
        if self.stringify_formulas_check.isChecked():
            warnings.append("Formula protection enabled")
        return "\n".join(warnings)

    def _update_parent_metadata_option(self, candidate: ArrayCandidate, is_new_selection: bool) -> None:
        supports_parent_metadata = self._supports_parent_metadata(candidate)
        self.parent_metadata_check.setEnabled(supports_parent_metadata)
        if supports_parent_metadata and is_new_selection:
            self.parent_metadata_check.setChecked(True)
        elif not supports_parent_metadata:
            self.parent_metadata_check.setChecked(False)

    def _supports_parent_metadata(self, candidate: ArrayCandidate) -> bool:
        return any(segment is Ellipsis for segment in candidate.path)

    def _include_parent_metadata_for(self, candidate: ArrayCandidate) -> bool:
        return self._supports_parent_metadata(candidate) and self.parent_metadata_check.isChecked()

    def _table_transform_options(self) -> TableTransformOptions:
        return TableTransformOptions(
            stringify_all=self.stringify_all_check.isChecked(),
            stringify_formulas=self.stringify_formulas_check.isChecked(),
            column_transforms=tuple(self.column_transforms.values()),
        )

    def _unfold_key(self) -> Optional[str]:
        nested_key = self.nested_candidate_combo.currentText()
        if nested_key and nested_key != self.NO_UNFOLD_LABEL:
            return nested_key
        return None

    def _update_nested_candidate_action(self, candidate: ArrayCandidate, reset_unfold: bool) -> None:
        nested_candidates = self.service.get_nested_array_candidates(candidate.display_path, max_nested_levels=3)
        current_value = self.nested_candidate_combo.currentText()
        self.nested_candidate_combo.blockSignals(True)
        self.nested_candidate_combo.clear()
        if not nested_candidates:
            self.nested_candidate_combo.setEnabled(False)
            self.nested_candidate_combo.blockSignals(False)
            return
        values = [self.NO_UNFOLD_LABEL] + [nested_candidate.display_path for nested_candidate in nested_candidates]
        self.nested_candidate_combo.addItems(values)
        if not reset_unfold and current_value in values:
            self.nested_candidate_combo.setCurrentText(current_value)
        else:
            self.nested_candidate_combo.setCurrentText(self.NO_UNFOLD_LABEL)
        self.nested_candidate_combo.setEnabled(True)
        self.nested_candidate_combo.blockSignals(False)

    def _clear_nested_candidate_action(self) -> None:
        self.nested_candidate_combo.blockSignals(True)
        self.nested_candidate_combo.clear()
        self.nested_candidate_combo.setEnabled(False)
        self.nested_candidate_combo.blockSignals(False)

    def _render_preview(self, preview: TablePreview) -> None:
        columns = [column.name for column in preview.columns[: self.MAX_PREVIEW_COLUMNS]]
        self.current_preview_columns = [column.name for column in preview.columns]
        self._refresh_column_action_columns()
        self.preview_table.clear()
        self.preview_table.setColumnCount(len(columns))
        self.preview_table.setHorizontalHeaderLabels(columns)
        preview_rows = preview.preview_rows[:6]
        self.preview_table.setRowCount(len(preview_rows))
        for row_index, row in enumerate(preview_rows):
            for column_index, column_name in enumerate(columns):
                self.preview_table.setItem(row_index, column_index, QTableWidgetItem(self._format_preview_value(row.get(column_name))))
        self._size_preview_columns(len(columns))
        self._update_cell_preview()
        type_summary = ", ".join(sorted({column.inferred_type for column in preview.columns[: self.MAX_PREVIEW_COLUMNS]}))
        warning_text = f" | {'; '.join(preview.warnings)}" if preview.warnings else ""
        summary = f"{preview.rows} rows | {len(preview.columns)} columns"
        if type_summary:
            summary = f"{summary} | types: {type_summary}"
        self.preview_label.setText(f"{summary}{warning_text}")
        self.preview_label.setToolTip("\n".join(f"{column.name}: {column.inferred_type}" for column in preview.columns))

    def _size_preview_columns(self, column_count: int) -> None:
        for column_index in range(column_count):
            self.preview_table.setColumnWidth(column_index, 150)

    def _update_cell_preview(self) -> None:
        item = self.preview_table.currentItem()
        if item is None:
            self.cell_preview.clear()
            return
        column_item = self.preview_table.horizontalHeaderItem(item.column())
        column_name = column_item.text() if column_item is not None else f"Column {item.column() + 1}"
        self.cell_preview.setPlainText(f"{column_name}: {item.text()}")

    def _clear_preview(self, message: str = "Select an exportable array to preview rows.") -> None:
        self.preview_table.clear()
        self.preview_table.setRowCount(0)
        self.preview_table.setColumnCount(0)
        self.cell_preview.clear()
        self.current_preview_columns = []
        self._refresh_column_action_columns()
        self.preview_label.setText(message)
        self.preview_label.setToolTip("")

    def open_advanced_options(self) -> None:
        visible = not self.advanced_section_frame.isVisible()
        self.advanced_section_frame.setVisible(visible)
        self.advanced_toggle_button.setText("Hide Advanced Options" if visible else "Show Advanced Options")
        if visible:
            self._refresh_column_action_columns()

    def _refresh_column_action_columns(self) -> None:
        if self.advanced_column_combo is None:
            return
        current = self.advanced_column_combo.currentText()
        self.advanced_column_combo.blockSignals(True)
        self.advanced_column_combo.clear()
        self.advanced_column_combo.addItems(self.current_preview_columns)
        if current in self.current_preview_columns:
            self.advanced_column_combo.setCurrentText(current)
        self.advanced_column_combo.blockSignals(False)
        self._load_column_action()

    def _load_column_action(self) -> None:
        column = self.advanced_column_combo.currentText() if self.advanced_column_combo is not None else ""
        possible_types = self._possible_types_for_column(column)
        self.advanced_type_combo.blockSignals(True)
        self.advanced_type_combo.clear()
        self.advanced_type_combo.addItems(possible_types)
        self.advanced_type_combo.blockSignals(False)

        transform = self.column_transforms.get(column)
        if transform is None:
            self.advanced_type_combo.setCurrentText("Keep")
            self.advanced_find_edit.clear()
            self.advanced_replace_edit.clear()
            self.advanced_status_label.setText(self._column_type_hint(column, possible_types))
            return
        self.advanced_type_combo.setCurrentText(transform.data_type if transform.data_type in possible_types else "Keep")
        self.advanced_find_edit.setText(transform.find_text)
        self.advanced_replace_edit.setText(transform.replace_text)
        self.advanced_status_label.setText(f"{self._column_action_status_text(transform)} | {self._column_type_hint(column, possible_types)}")

    def _save_column_action(self) -> None:
        column = self.advanced_column_combo.currentText()
        if not column:
            QMessageBox.warning(self, "Missing column", "Please select a column")
            return
        transform = ColumnTransform(
            column=column,
            data_type=self.advanced_type_combo.currentText(),
            find_text=self.advanced_find_edit.text(),
            replace_text=self.advanced_replace_edit.text(),
        )
        next_transforms = dict(self.column_transforms)
        if transform.data_type == "Keep" and not transform.find_text:
            next_transforms.pop(column, None)
        else:
            next_transforms[column] = transform
        try:
            self._validate_column_transforms(next_transforms)
        except ArrayMateCoreError as e:
            self.advanced_status_label.setText(f"Cannot apply: {e}")
            self.status_label.setText(str(e))
            return
        self.column_transforms = next_transforms
        self.advanced_status_label.setText(
            f"No action set for {column}" if transform.data_type == "Keep" and not transform.find_text else self._column_action_status_text(transform)
        )
        self.refresh_selected_candidate()

    def _clear_column_action(self) -> None:
        column = self.advanced_column_combo.currentText()
        if column:
            self.column_transforms.pop(column, None)
        self.advanced_type_combo.setCurrentText("Keep")
        self.advanced_find_edit.clear()
        self.advanced_replace_edit.clear()
        self.advanced_status_label.setText(f"No action set for {column}" if column else "No column selected")
        self.refresh_selected_candidate()

    def _reset_column_action_form(self) -> None:
        if self.advanced_column_combo is not None:
            self.advanced_column_combo.clear()
        if self.advanced_type_combo is not None:
            self.advanced_type_combo.clear()
            self.advanced_type_combo.addItems(["Keep", "Text"])
        if hasattr(self, "advanced_find_edit"):
            self.advanced_find_edit.clear()
            self.advanced_replace_edit.clear()
            self.advanced_status_label.setText("No column action selected")

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
        candidate = self.candidate_by_path.get(self.selected_array_key)
        if candidate is None:
            return None
        return self.service.get_table_data(
            candidate.display_path,
            unfold_key=self._unfold_key(),
            include_parent_metadata=self._include_parent_metadata_for(candidate),
            transform_options=TableTransformOptions(
                stringify_all=self.stringify_all_check.isChecked(),
                stringify_formulas=self.stringify_formulas_check.isChecked(),
            ),
        )

    def _validate_column_transforms(self, column_transforms: dict[str, ColumnTransform]) -> None:
        candidate = self.candidate_by_path.get(self.selected_array_key)
        if candidate is None:
            return
        array_data = self.service.get_table_data(
            candidate.display_path,
            unfold_key=self._unfold_key(),
            include_parent_metadata=self._include_parent_metadata_for(candidate),
            transform_options=TableTransformOptions(
                stringify_all=self.stringify_all_check.isChecked(),
                stringify_formulas=self.stringify_formulas_check.isChecked(),
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
        selected_key = array_key or self.effective_candidate_key or self.selected_array_key
        if selected_key:
            self.output_filename_edit.setText(self._suggest_filename(selected_key))

    def on_format_selected(self) -> None:
        output_format = get_output_format(self.output_format_combo.currentText())
        self.extension_label.setText(output_format.extension)
        self.process_button.setText(f"Convert to {output_format.label}")

    def browse_output_folder(self) -> None:
        folder = QFileDialog.getExistingDirectory(self, "Select Output Folder")
        if folder:
            self.output_folder_edit.setText(folder)
            self._refresh_auto_filename()

    def on_filename_edited(self) -> None:
        self.auto_filename = False

    def convert_to_file(self) -> None:
        if not self.selected_array_key:
            QMessageBox.critical(self, "Missing array", "Please select an array to convert")
            return
        if not self.output_folder_edit.text():
            QMessageBox.critical(self, "Missing folder", "Please select an output folder")
            return
        if not self.output_filename_edit.text():
            QMessageBox.critical(self, "Missing filename", "Please enter a filename")
            return
        try:
            export_plan = self.service.create_export_plan(
                self.output_folder_edit.text(),
                self.output_filename_edit.text(),
                self.output_format_combo.currentText(),
            )
            if os.path.exists(export_plan.file_path):
                result = QMessageBox.question(
                    self,
                    "File Exists",
                    f"File '{export_plan.filename}' already exists in the selected folder.\nDo you want to overwrite it?",
                )
                if result != QMessageBox.StandardButton.Yes:
                    return
            selected_candidate = self.candidate_by_path.get(self.selected_array_key)
            unfold_key = self._unfold_key() if selected_candidate else None
            export_result = self.service.export_array(
                self.selected_array_key,
                export_plan,
                include_parent_metadata=bool(
                    selected_candidate and not unfold_key and self._include_parent_metadata_for(selected_candidate)
                ),
                unfold_key=unfold_key,
                transform_options=self._table_transform_options(),
            )
            QMessageBox.information(
                self,
                "Success",
                f"{export_result.output_format.label} file saved successfully!\n"
                f"File: {export_result.file_path}\n"
                f"Rows: {export_result.rows}\n"
                f"Columns: {export_result.columns}",
            )
            self.status_label.setText(f"{export_result.output_format.label} file saved: {export_result.filename}")
            self._open_exported_file(export_result.output_format.label, export_result.file_path)
        except ArrayMateCoreError as e:
            QMessageBox.critical(self, "Error", str(e))
            self.status_label.setText(str(e))
        except Exception as e:
            output_format = get_output_format(self.output_format_combo.currentText())
            QMessageBox.critical(self, "Error", f"Error creating {output_format.label} file: {e}")
            self.status_label.setText(f"Error creating {output_format.label} file")

    def _open_exported_file(self, output_label: str, file_path: str) -> None:
        if output_label in ("Excel", "CSV", "JSON"):
            self._open_file(file_path, output_label)
        else:
            self.open_file_location(file_path)

    def _open_file(self, file_path: str, label: str) -> None:
        try:
            system = platform.system()
            if system == "Windows":
                os.startfile(file_path)
            elif system == "Darwin":
                subprocess.run(["open", file_path], check=True)
            else:
                subprocess.run(["xdg-open", file_path], check=True)
            self.status_label.setText(f"{label} file opened: {os.path.basename(file_path)}")
        except Exception as e:
            QMessageBox.warning(
                self,
                "Warning",
                f"{label} file saved successfully, but could not open automatically.\n"
                f"File location: {file_path}\n"
                f"Error: {e}",
            )
            self.status_label.setText(f"{label} file saved, but could not open: {os.path.basename(file_path)}")

    def open_file_location(self, file_path: str) -> None:
        try:
            system = platform.system()
            if system == "Windows":
                subprocess.run(["explorer", "/select,", file_path], check=True)
            elif system == "Darwin":
                subprocess.run(["open", "-R", file_path], check=True)
            else:
                subprocess.run(["xdg-open", os.path.dirname(file_path)], check=True)
            self.status_label.setText(f"File location opened: {os.path.basename(file_path)}")
        except Exception:
            QMessageBox.information(self, "File Saved", f"File saved successfully!\nLocation: {file_path}")
            self.status_label.setText(f"File saved: {os.path.basename(file_path)}")

    def open_repository(self) -> None:
        webbrowser.open(self.REPOSITORY_URL)
        self.status_label.setText("Repository opened in browser")


def main() -> None:
    app = QApplication(sys.argv)
    window = ArrayMateWindow()
    window.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
