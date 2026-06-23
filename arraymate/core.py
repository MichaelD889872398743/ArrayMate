"""
Core data handling for ArrayMate.

This module intentionally has no Tkinter dependencies so the JSON selection and
export behavior can be tested without starting the desktop UI.
"""

from __future__ import annotations

import json
import os
from dataclasses import dataclass
from pathlib import Path
from typing import Any, List, Optional, Sequence, Union

import pandas as pd


JsonData = Union[dict[str, Any], list[Any]]
WILDCARD = Ellipsis


@dataclass(frozen=True)
class JsonNode:
    """Tree node describing a JSON structure for the future explorer UI."""

    path: tuple[Any, ...]
    display_path: str
    label: str
    kind: str
    depth: int
    children: tuple["JsonNode", ...] = ()
    item_count: Optional[int] = None
    source_count: int = 1
    is_empty: bool = False
    is_object_array: bool = False
    is_primitive_array: bool = False
    has_nested_arrays: bool = False
    exportable: bool = False
    warning: Optional[str] = None


@dataclass(frozen=True)
class ArrayCandidate:
    """Array discovered in JSON, with export-oriented metadata."""

    path: tuple[Any, ...]
    display_path: str
    item_count: int
    source_count: int
    depth: int
    column_count: int = 0
    is_empty: bool = False
    is_object_array: bool = False
    is_primitive_array: bool = False
    has_nested_arrays: bool = False
    exportable: bool = False
    warning: Optional[str] = None


@dataclass(frozen=True)
class ColumnPreview:
    """Preview metadata for one table column."""

    name: str
    inferred_type: str
    contains_nested_values: bool = False


@dataclass(frozen=True)
class TablePreview:
    """Small table preview for a selected export candidate."""

    display_path: str
    rows: int
    columns: tuple[ColumnPreview, ...]
    preview_rows: tuple[dict[str, Any], ...]
    warnings: tuple[str, ...] = ()


@dataclass(frozen=True)
class ArraySummary:
    """Short description of a selected JSON array."""

    item_count: int
    column_count: int = 0
    contains_objects: bool = False


@dataclass(frozen=True)
class OutputFormat:
    """Normalized output format selected by the UI."""

    label: str
    extension: str


OUTPUT_FORMATS = {
    "Excel": OutputFormat(label="Excel", extension=".xlsx"),
    "CSV": OutputFormat(label="CSV", extension=".csv"),
    "JSON": OutputFormat(label="JSON", extension=".json"),
}


class ArrayMateCoreError(ValueError):
    """Raised when data cannot be converted safely."""


def find_arrays(data: JsonData, path: tuple[Any, ...] = ()) -> list[str]:
    """
    Recursively find all arrays in JSON data.

    Paths are displayed in the existing ArrayMate notation, such as
    ``orders[0].items``. The displayed path is kept for UI compatibility, while
    lookup uses the original object structure and does not parse the string.
    """
    arrays: list[str] = []

    if isinstance(data, dict):
        for key, value in data.items():
            current_path = path + (key,)
            if isinstance(value, list):
                arrays.append(format_path(current_path))
                for index, item in enumerate(value):
                    if isinstance(item, (dict, list)):
                        arrays.extend(find_arrays(item, current_path + (index,)))
            elif isinstance(value, dict):
                arrays.extend(find_arrays(value, current_path))

    elif isinstance(data, list):
        if not path:
            arrays.append("root")
        for index, item in enumerate(data):
            if isinstance(item, (dict, list)):
                arrays.extend(find_arrays(item, path + (index,)))

    return arrays


def get_array_data(data: JsonData, array_path: str) -> Optional[list[Any]]:
    """
    Return array data for a displayed path.

    The resolver first rediscovers arrays and matches by formatted path. This
    preserves support for keys containing spaces, hyphens, dots, and other
    characters that cannot be represented safely by a regex split.
    """
    if array_path == "root":
        return data if isinstance(data, list) else None

    for path in iter_array_paths(data):
        if format_path(path) == array_path:
            current: Any = data
            for segment in path:
                current = current[segment]
            return current if isinstance(current, list) else None

    return None


def get_array_data_by_path(data: JsonData, path: tuple[Any, ...]) -> Optional[list[Any]]:
    """Return array data for a structured path, including aggregate wildcard paths."""
    values = _resolve_path_values(data, path)
    if len(values) == 1 and isinstance(values[0], list):
        return values[0]

    if values and all(isinstance(value, list) for value in values):
        return [item for value in values for item in value]

    return None


def get_array_data_with_parent_metadata(data: JsonData, path: tuple[Any, ...]) -> Optional[list[Any]]:
    """Return wildcard array rows with scalar parent fields attached as metadata columns."""
    if WILDCARD not in path:
        return get_array_data_by_path(data, path)

    rows: list[Any] = []
    _collect_rows_with_parent_metadata(data, path, rows, (), parent_columns_first=False)
    return rows if rows else None


def get_unfolded_array_data(data: JsonData, parent_path: tuple[Any, ...], nested_path: tuple[Any, ...]) -> Optional[list[Any]]:
    """Return nested array rows expanded into their parent row context."""
    if not _is_nested_path(parent_path, nested_path):
        return None

    rows: list[Any] = []
    _collect_rows_with_parent_metadata(data, nested_path, rows, (), parent_columns_first=True)
    return rows if rows else None


def iter_array_paths(data: JsonData, path: tuple[Any, ...] = ()) -> list[tuple[Any, ...]]:
    """Return structured paths for every array in the JSON data."""
    paths: list[tuple[Any, ...]] = []

    if isinstance(data, dict):
        for key, value in data.items():
            current_path = path + (key,)
            if isinstance(value, list):
                paths.append(current_path)
                for index, item in enumerate(value):
                    if isinstance(item, (dict, list)):
                        paths.extend(iter_array_paths(item, current_path + (index,)))
            elif isinstance(value, dict):
                paths.extend(iter_array_paths(value, current_path))

    elif isinstance(data, list):
        if not path:
            paths.append(())
        for index, item in enumerate(data):
            if isinstance(item, (dict, list)):
                paths.extend(iter_array_paths(item, path + (index,)))

    return paths


def _resolve_path_values(value: Any, path: tuple[Any, ...]) -> list[Any]:
    if not path:
        return [value]

    segment = path[0]
    remaining_path = path[1:]
    if segment is WILDCARD:
        if not isinstance(value, list):
            return []
        resolved_values: list[Any] = []
        for item in value:
            resolved_values.extend(_resolve_path_values(item, remaining_path))
        return resolved_values

    if isinstance(segment, int):
        if isinstance(value, list) and 0 <= segment < len(value):
            return _resolve_path_values(value[segment], remaining_path)
        return []

    if isinstance(value, dict) and segment in value:
        return _resolve_path_values(value[segment], remaining_path)

    return []


def _collect_rows_with_parent_metadata(
    value: Any,
    path: tuple[Any, ...],
    rows: list[Any],
    parent_metadata: tuple[tuple[str, Any], ...],
    parent_columns_first: bool,
) -> None:
    if not path:
        if isinstance(value, list):
            for item in value:
                rows.append(_merge_parent_metadata(item, parent_metadata, parent_columns_first))
        return

    segment = path[0]
    remaining_path = path[1:]
    if segment is WILDCARD:
        if not isinstance(value, list):
            return
        for index, item in enumerate(value):
            metadata = parent_metadata + _parent_metadata(item, index)
            _collect_rows_with_parent_metadata(item, remaining_path, rows, metadata, parent_columns_first)
        return

    if isinstance(segment, int):
        if isinstance(value, list) and 0 <= segment < len(value):
            _collect_rows_with_parent_metadata(value[segment], remaining_path, rows, parent_metadata, parent_columns_first)
        return

    if isinstance(value, dict) and segment in value:
        _collect_rows_with_parent_metadata(value[segment], remaining_path, rows, parent_metadata, parent_columns_first)


def _parent_metadata(value: Any, index: int) -> tuple[tuple[str, Any], ...]:
    metadata: list[tuple[str, Any]] = []
    if isinstance(value, dict):
        for key, item in value.items():
            if not isinstance(item, (dict, list)):
                metadata.append((str(key), item))
    return tuple(metadata)


def _merge_parent_metadata(item: Any, parent_metadata: tuple[tuple[str, Any], ...], parent_columns_first: bool) -> Any:
    if not isinstance(item, dict):
        return item

    merged: dict[str, Any] = {}
    if parent_columns_first:
        for key, value in parent_metadata:
            output_key = key
            suffix = 2
            while output_key in item or output_key in merged:
                output_key = f"{key}_{suffix}"
                suffix += 1
            merged[output_key] = value
        merged.update(item)
        return merged

    merged = dict(item)
    for key, value in parent_metadata:
        output_key = key
        suffix = 2
        while output_key in merged:
            output_key = f"{key}_{suffix}"
            suffix += 1
        merged[output_key] = value
    return merged


def _is_nested_path(parent_path: tuple[Any, ...], nested_path: tuple[Any, ...]) -> bool:
    expected_prefix = parent_path + (WILDCARD,)
    return nested_path[: len(expected_prefix)] == expected_prefix


def format_path(path: Sequence[Any]) -> str:
    """Format a structured JSON path for display."""
    if not path:
        return "root"

    display = ""
    for segment in path:
        if isinstance(segment, int):
            display += f"[{segment}]"
        elif segment is WILDCARD:
            display += "[*]"
        elif str(segment).isidentifier():
            if display:
                display += f".{segment}"
            else:
                display = str(segment)
        elif display:
            display += f"[{json.dumps(str(segment))}]"
        else:
            display = f"[{json.dumps(str(segment))}]"

    return display


def build_json_tree(data: JsonData) -> JsonNode:
    """Build an aggregate JSON tree suitable for an explorer-style UI."""
    return _build_node(data, (), "root")


def discover_array_candidates(data: JsonData) -> list[ArrayCandidate]:
    """Return aggregate array candidates, avoiding repeated per-row duplicates."""
    candidates: list[ArrayCandidate] = []
    _collect_array_candidates(_build_node(data, (), "root"), candidates)
    return candidates


def build_table_preview(array_data: Optional[list[Any]], display_path: str, max_rows: int = 50) -> TablePreview:
    """Build preview metadata for an array of objects."""
    if array_data is None:
        raise ArrayMateCoreError("Selected array is invalid")
    if not array_data:
        return TablePreview(display_path=display_path, rows=0, columns=(), preview_rows=(), warnings=("Empty array",))
    if not all(isinstance(item, dict) for item in array_data):
        raise ArrayMateCoreError("Array must contain objects with key-value pairs")

    column_names = _column_names(array_data)
    columns = tuple(
        ColumnPreview(
            name=column_name,
            inferred_type=_infer_column_type([row.get(column_name) for row in array_data]),
            contains_nested_values=any(isinstance(row.get(column_name), (dict, list)) for row in array_data),
        )
        for column_name in column_names
    )
    warnings = []
    if any(column.contains_nested_values for column in columns):
        warnings.append("Some columns contain nested records or arrays")

    preview_rows = tuple(dict(row) for row in array_data[:max_rows])
    return TablePreview(
        display_path=display_path,
        rows=len(array_data),
        columns=columns,
        preview_rows=preview_rows,
        warnings=tuple(warnings),
    )


def _build_node(value: Any, path: tuple[Any, ...], label: str) -> JsonNode:
    kind = _value_kind(value)
    if isinstance(value, dict):
        children = tuple(_build_node(child_value, path + (key,), str(key)) for key, child_value in value.items())
        return JsonNode(
            path=path,
            display_path=format_path(path),
            label=label,
            kind=kind,
            depth=len(path),
            children=children,
        )

    if isinstance(value, list):
        return _build_array_node(value, path, label, source_count=1)

    return JsonNode(
        path=path,
        display_path=format_path(path),
        label=label,
        kind=kind,
        depth=len(path),
    )


def _build_array_node(values: list[Any], path: tuple[Any, ...], label: str, source_count: int) -> JsonNode:
    is_empty = not values
    is_object_array = bool(values) and all(isinstance(item, dict) for item in values)
    is_primitive_array = bool(values) and all(not isinstance(item, (dict, list)) for item in values)
    has_nested_arrays = any(_contains_array(item) for item in values)
    column_count = len(_column_names(values)) if is_object_array else 0
    exportable = is_object_array and not is_empty
    warning = _array_warning(is_empty, is_object_array, is_primitive_array, has_nested_arrays)

    children: tuple[JsonNode, ...] = ()
    if is_object_array:
        children = tuple(
            _build_aggregate_node([item[key] for item in values if key in item], path + (WILDCARD, key), str(key))
            for key in _column_names(values)
        )

    return JsonNode(
        path=path,
        display_path=format_path(path),
        label=label,
        kind="array",
        depth=len(path),
        children=children,
        item_count=len(values),
        source_count=source_count,
        is_empty=is_empty,
        is_object_array=is_object_array,
        is_primitive_array=is_primitive_array,
        has_nested_arrays=has_nested_arrays,
        exportable=exportable,
        warning=warning,
    )


def _build_aggregate_node(values: list[Any], path: tuple[Any, ...], label: str) -> JsonNode:
    if not values:
        return JsonNode(path=path, display_path=format_path(path), label=label, kind="unknown", depth=len(path))

    if all(isinstance(value, list) for value in values):
        merged_items = [item for value in values for item in value]
        node = _build_array_node(merged_items, path, label, source_count=len(values))
        return JsonNode(
            path=node.path,
            display_path=node.display_path,
            label=node.label,
            kind=node.kind,
            depth=node.depth,
            children=node.children,
            item_count=node.item_count,
            source_count=node.source_count,
            is_empty=node.is_empty,
            is_object_array=node.is_object_array,
            is_primitive_array=node.is_primitive_array,
            has_nested_arrays=node.has_nested_arrays,
            exportable=node.exportable,
            warning=node.warning,
        )

    if all(isinstance(value, dict) for value in values):
        keys = _object_keys(values)
        children = tuple(
            _build_aggregate_node([value[key] for value in values if key in value], path + (key,), str(key))
            for key in keys
        )
        return JsonNode(
            path=path,
            display_path=format_path(path),
            label=label,
            kind="object",
            depth=len(path),
            children=children,
        )

    kinds = {_value_kind(value) for value in values}
    kind = kinds.pop() if len(kinds) == 1 else "mixed"
    return JsonNode(
        path=path,
        display_path=format_path(path),
        label=label,
        kind=kind,
        depth=len(path),
    )


def _collect_array_candidates(node: JsonNode, candidates: list[ArrayCandidate]) -> None:
    if node.kind == "array":
        candidates.append(
            ArrayCandidate(
                path=node.path,
                display_path=node.display_path,
                item_count=node.item_count or 0,
                source_count=node.source_count,
                depth=node.depth,
                column_count=len(_table_column_nodes(node)),
                is_empty=node.is_empty,
                is_object_array=node.is_object_array,
                is_primitive_array=node.is_primitive_array,
                has_nested_arrays=node.has_nested_arrays,
                exportable=node.exportable,
                warning=node.warning,
            )
        )

    for child in node.children:
        _collect_array_candidates(child, candidates)


def _table_column_nodes(node: JsonNode) -> tuple[JsonNode, ...]:
    if not node.is_object_array:
        return ()
    return tuple(child for child in node.children if child.path and child.path[-1] is not WILDCARD)


def _value_kind(value: Any) -> str:
    if isinstance(value, dict):
        return "object"
    if isinstance(value, list):
        return "array"
    if value is None:
        return "null"
    if isinstance(value, bool):
        return "boolean"
    if isinstance(value, (int, float)):
        return "number"
    return "text"


def _contains_array(value: Any) -> bool:
    if isinstance(value, list):
        return True
    if isinstance(value, dict):
        return any(_contains_array(child_value) for child_value in value.values())
    return False


def _array_warning(is_empty: bool, is_object_array: bool, is_primitive_array: bool, has_nested_arrays: bool) -> Optional[str]:
    if is_empty:
        return "Empty array"
    if is_primitive_array:
        return "Primitive array; not directly exportable as a table"
    if not is_object_array:
        return "Mixed array; not directly exportable as a table"
    if has_nested_arrays:
        return "Contains nested arrays"
    return None


def _column_names(rows: list[Any]) -> list[str]:
    dict_rows = [row for row in rows if isinstance(row, dict)]
    names: list[str] = []
    for row in dict_rows:
        for key in row.keys():
            if key not in names:
                names.append(str(key))
    return names


def _object_keys(objects: list[dict[str, Any]]) -> list[str]:
    keys: list[str] = []
    for obj in objects:
        for key in obj.keys():
            if key not in keys:
                keys.append(str(key))
    return keys


def _infer_column_type(values: list[Any]) -> str:
    non_null_values = [value for value in values if value is not None]
    if not non_null_values:
        return "empty"

    inferred_types = {_value_kind(value) for value in non_null_values}
    if len(inferred_types) == 1:
        return inferred_types.pop()
    return "mixed"


def summarize_array(array_data: Optional[list[Any]]) -> Optional[ArraySummary]:
    """Create a display-friendly summary for an array."""
    if array_data is None:
        return None

    if not array_data:
        return ArraySummary(item_count=0)

    first_item = array_data[0]
    if isinstance(first_item, dict):
        return ArraySummary(
            item_count=len(array_data),
            column_count=len(first_item.keys()),
            contains_objects=True,
        )

    return ArraySummary(item_count=len(array_data))


def get_output_format(format_text: str) -> OutputFormat:
    """Normalize a UI format label such as ``Excel (.xlsx)``."""
    for key, output_format in OUTPUT_FORMATS.items():
        if key in format_text:
            return output_format

    return OUTPUT_FORMATS["Excel"]


def build_output_path(output_folder: str, filename: str, extension: str) -> str:
    """Build an output path and append the expected extension when missing."""
    clean_filename = filename.strip()
    if not clean_filename:
        raise ArrayMateCoreError("Please enter a filename")

    if not clean_filename.endswith(extension):
        clean_filename += extension

    return os.path.join(output_folder, clean_filename)


def records_to_dataframe(array_data: Optional[list[Any]]) -> pd.DataFrame:
    """Convert a JSON array of objects to a DataFrame."""
    if array_data is None:
        raise ArrayMateCoreError("Selected array is invalid")

    if not array_data:
        raise ArrayMateCoreError("Selected array is empty")

    if not isinstance(array_data[0], dict):
        raise ArrayMateCoreError("Array must contain objects with key-value pairs")

    return pd.DataFrame(array_data)


def write_array_to_file(array_data: Optional[list[Any]], file_path: str, output_format: OutputFormat) -> pd.DataFrame:
    """
    Write array data to disk and return the DataFrame that was written.

    Returning the DataFrame keeps the UI able to report rows and columns without
    duplicating conversion logic.
    """
    dataframe = records_to_dataframe(array_data)
    output_path = Path(file_path)

    if output_format.label == "Excel":
        dataframe.to_excel(output_path, index=False)
    elif output_format.label == "CSV":
        dataframe.to_csv(output_path, index=False)
    elif output_format.label == "JSON":
        dataframe.to_json(output_path, orient="records", indent=2)
    else:
        raise ArrayMateCoreError(f"Unsupported output format: {output_format.label}")

    return dataframe
