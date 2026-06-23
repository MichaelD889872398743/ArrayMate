"""Application workflow services for ArrayMate."""

from __future__ import annotations

import json
import os
from dataclasses import dataclass
from typing import Any, Optional

from arraymate.core import (
    ArrayCandidate,
    JsonData,
    JsonNode,
    OutputFormat,
    build_output_path,
    build_json_tree,
    discover_array_candidates,
    find_arrays,
    get_array_data,
    get_array_data_by_path,
    get_array_data_with_parent_metadata,
    get_unfolded_array_data,
    get_output_format,
    write_array_to_file,
)


@dataclass(frozen=True)
class LoadResult:
    """Result of loading JSON into the app workflow."""

    array_keys: list[str]
    selected_key: Optional[str]
    selected_array: Optional[list[Any]]
    json_tree: JsonNode
    array_candidates: list[ArrayCandidate]


@dataclass(frozen=True)
class ExportPlan:
    """Resolved export destination and format."""

    output_format: OutputFormat
    file_path: str
    filename: str


@dataclass(frozen=True)
class ExportResult:
    """Details of a completed export."""

    output_format: OutputFormat
    file_path: str
    filename: str
    rows: int
    columns: int


class ArrayMateService:
    """Stateful application workflow, independent of any UI toolkit."""

    def __init__(self) -> None:
        self.json_data: Optional[JsonData] = None
        self.array_keys: list[str] = []
        self.json_tree: Optional[JsonNode] = None
        self.array_candidates: list[ArrayCandidate] = []

    def clear(self) -> None:
        """Reset loaded JSON state."""
        self.json_data = None
        self.array_keys = []
        self.json_tree = None
        self.array_candidates = []

    def load_text(self, json_text: str) -> LoadResult:
        """Parse JSON text and load it into the workflow."""
        return self.load_data(json.loads(json_text))

    def load_file(self, file_path: str) -> LoadResult:
        """Read a JSON file and load it into the workflow."""
        with open(file_path, "r", encoding="utf-8") as file:
            return self.load_data(json.load(file))

    def load_data(self, data: JsonData) -> LoadResult:
        """Load parsed JSON data."""
        self.json_data = data
        self.array_keys = find_arrays(data)
        self.json_tree = build_json_tree(data)
        self.array_candidates = discover_array_candidates(data)
        selected_key = self.array_keys[0] if self.array_keys else None
        selected_array = self.get_array_data(selected_key) if selected_key else None
        return LoadResult(
            array_keys=self.array_keys,
            selected_key=selected_key,
            selected_array=selected_array,
            json_tree=self.json_tree,
            array_candidates=self.array_candidates,
        )

    def get_array_data(self, array_key: Optional[str], include_parent_metadata: bool = False) -> Optional[list[Any]]:
        """Return the selected array from the loaded JSON."""
        if not array_key or self.json_data is None:
            return None

        array_data = get_array_data(self.json_data, array_key)
        if array_data is not None:
            return array_data

        candidate = self.get_array_candidate(array_key)
        if candidate is None:
            return None

        if include_parent_metadata:
            return get_array_data_with_parent_metadata(self.json_data, candidate.path)

        return get_array_data_by_path(self.json_data, candidate.path)

    def get_table_data(
        self,
        array_key: Optional[str],
        unfold_key: Optional[str] = None,
        include_parent_metadata: bool = False,
    ) -> Optional[list[Any]]:
        """Return rows for the selected table, optionally unfolding a nested child table."""
        if not array_key or self.json_data is None:
            return None

        if unfold_key:
            parent = self.get_array_candidate(array_key)
            nested = self.get_array_candidate(unfold_key)
            if parent is None or nested is None:
                return None
            return get_unfolded_array_data(self.json_data, parent.path, nested.path)

        return self.get_array_data(array_key, include_parent_metadata=include_parent_metadata)

    def get_array_candidate(self, array_key: str) -> Optional[ArrayCandidate]:
        """Return candidate metadata by display path."""
        return next(
            (candidate for candidate in self.array_candidates if candidate.display_path == array_key),
            None,
        )

    def get_nested_array_candidates(self, array_key: str, max_nested_levels: int = 3) -> list[ArrayCandidate]:
        """Return exportable child arrays nested under a parent array candidate."""
        parent = self.get_array_candidate(array_key)
        if parent is None:
            return []

        parent_wildcard_count = parent.path.count(Ellipsis)
        nested_candidates = []
        for candidate in self.array_candidates:
            if candidate.display_path == parent.display_path:
                continue
            if not candidate.exportable:
                continue
            if not _is_descendant_candidate(parent.path, candidate.path):
                continue
            nested_level = candidate.path.count(Ellipsis) - parent_wildcard_count
            if 1 <= nested_level <= max_nested_levels:
                nested_candidates.append(candidate)

        return nested_candidates

    def create_export_plan(self, output_folder: str, filename: str, format_text: str) -> ExportPlan:
        """Resolve output format and file path before overwrite confirmation."""
        output_format = get_output_format(format_text)
        file_path = build_output_path(output_folder, filename, output_format.extension)
        return ExportPlan(
            output_format=output_format,
            file_path=file_path,
            filename=os.path.basename(file_path),
        )

    def export_array(
        self,
        array_key: str,
        export_plan: ExportPlan,
        include_parent_metadata: bool = False,
        unfold_key: Optional[str] = None,
    ) -> ExportResult:
        """Write the selected array to the planned output file."""
        array_data = self.get_table_data(
            array_key,
            unfold_key=unfold_key,
            include_parent_metadata=include_parent_metadata,
        )
        dataframe = write_array_to_file(array_data, export_plan.file_path, export_plan.output_format)
        return ExportResult(
            output_format=export_plan.output_format,
            file_path=export_plan.file_path,
            filename=export_plan.filename,
            rows=len(dataframe),
            columns=len(dataframe.columns),
        )


def _is_descendant_candidate(parent_path: tuple[Any, ...], candidate_path: tuple[Any, ...]) -> bool:
    expected_prefix = parent_path + (Ellipsis,)
    return candidate_path[: len(expected_prefix)] == expected_prefix
