import json
import unittest
from pathlib import Path

from arraymate.core import (
    ArrayMateCoreError,
    OutputFormat,
    build_json_tree,
    build_table_preview,
    build_output_path,
    discover_array_candidates,
    find_arrays,
    get_array_data,
    records_to_dataframe,
    summarize_array,
    write_array_to_file,
)


class JsonPathTests(unittest.TestCase):
    def test_find_arrays_in_nested_json(self):
        data = {
            "orders": [
                {"id": 1, "items": [{"sku": "A"}]},
                {"id": 2, "items": [{"sku": "B"}]},
            ],
            "users": [{"name": "Ada"}],
        }

        self.assertEqual(
            find_arrays(data),
            ["orders", "orders[0].items", "orders[1].items", "users"],
        )

    def test_get_array_data_supports_keys_that_regex_paths_cannot_parse(self):
        data = {
            "user-list": [{"id": 1}],
            "obj": {
                "line.items": [{"sku": "A"}],
                "spaced key": [{"name": "Ada"}],
            },
        }

        self.assertEqual(find_arrays(data), ['["user-list"]', 'obj["line.items"]', 'obj["spaced key"]'])
        self.assertEqual(get_array_data(data, '["user-list"]'), [{"id": 1}])
        self.assertEqual(get_array_data(data, 'obj["line.items"]'), [{"sku": "A"}])
        self.assertEqual(get_array_data(data, 'obj["spaced key"]'), [{"name": "Ada"}])

    def test_display_paths_disambiguate_dotted_keys_from_nested_keys(self):
        data = {
            "obj": {
                "line": {"items": [{"id": "nested"}]},
                "line.items": [{"id": "literal"}],
            }
        }

        self.assertEqual(find_arrays(data), ["obj.line.items", 'obj["line.items"]'])
        self.assertEqual(get_array_data(data, "obj.line.items"), [{"id": "nested"}])
        self.assertEqual(get_array_data(data, 'obj["line.items"]'), [{"id": "literal"}])

    def test_get_array_data_supports_root_array(self):
        data = [{"id": 1}]

        self.assertEqual(find_arrays(data), ["root"])
        self.assertEqual(get_array_data(data, "root"), [{"id": 1}])

    def test_get_array_data_returns_none_for_missing_path(self):
        self.assertIsNone(get_array_data({"users": []}, "orders"))


class ConversionTests(unittest.TestCase):
    def test_summarize_array_describes_objects(self):
        summary = summarize_array([{"id": 1, "name": "Ada"}])

        self.assertEqual(summary.item_count, 1)
        self.assertEqual(summary.column_count, 2)
        self.assertTrue(summary.contains_objects)

    def test_records_to_dataframe_rejects_empty_array(self):
        with self.assertRaisesRegex(ArrayMateCoreError, "empty"):
            records_to_dataframe([])

    def test_records_to_dataframe_rejects_non_object_array(self):
        with self.assertRaisesRegex(ArrayMateCoreError, "objects"):
            records_to_dataframe(["one", "two"])

    def test_build_output_path_appends_extension(self):
        output_path = build_output_path("out", "users", ".csv")

        self.assertEqual(output_path, str(Path("out") / "users.csv"))

    def test_write_array_to_json_file(self):
        output_path = Path("test_core_users.json")
        try:
            dataframe = write_array_to_file(
                [{"id": 1, "name": "Ada"}],
                str(output_path),
                OutputFormat(label="JSON", extension=".json"),
            )

            self.assertEqual(len(dataframe), 1)
            self.assertEqual(json.loads(output_path.read_text(encoding="utf-8")), [{"id": 1, "name": "Ada"}])
        finally:
            output_path.unlink(missing_ok=True)


class DiscoveryModelTests(unittest.TestCase):
    def test_build_json_tree_aggregates_repeated_array_children(self):
        data = {
            "orders": [
                {"id": 1, "items": [{"sku": "A"}], "warnings": []},
                {"id": 2, "items": [{"sku": "B"}], "warnings": []},
            ],
            "tags": ["urgent", "api"],
        }

        tree = build_json_tree(data)
        orders = next(child for child in tree.children if child.label == "orders")
        child_paths = [child.display_path for child in orders.children]

        self.assertEqual(orders.display_path, "orders")
        self.assertEqual(orders.item_count, 2)
        self.assertTrue(orders.exportable)
        self.assertEqual(child_paths, ["orders[*].id", "orders[*].items", "orders[*].warnings"])

    def test_discover_array_candidates_classifies_exportability(self):
        data = {
            "users": [{"name": "Ada"}],
            "empty": [],
            "tags": ["urgent", "api"],
            "orders": [{"items": [{"sku": "A"}]}],
        }

        candidates = {candidate.display_path: candidate for candidate in discover_array_candidates(data)}

        self.assertTrue(candidates["users"].exportable)
        self.assertEqual(candidates["users"].column_count, 1)
        self.assertFalse(candidates["empty"].exportable)
        self.assertEqual(candidates["empty"].warning, "Empty array")
        self.assertFalse(candidates["tags"].exportable)
        self.assertEqual(candidates["tags"].warning, "Primitive array; not directly exportable as a table")
        self.assertTrue(candidates["orders"].exportable)
        self.assertEqual(candidates["orders"].warning, "Contains nested arrays")
        self.assertTrue(candidates["orders[*].items"].exportable)

    def test_nested_array_candidates_are_grouped_by_parent_shape(self):
        data = {
            "orders": [
                {"items": [{"sku": "A"}]},
                {"items": [{"sku": "B"}, {"sku": "C"}]},
            ]
        }

        candidates = {candidate.display_path: candidate for candidate in discover_array_candidates(data)}

        self.assertIn("orders[*].items", candidates)
        self.assertEqual(candidates["orders[*].items"].item_count, 3)
        self.assertEqual(candidates["orders[*].items"].source_count, 2)

    def test_build_table_preview_reports_columns_and_nested_values(self):
        preview = build_table_preview(
            [
                {"id": 1, "customer": {"name": "Ada"}, "total": 10.5},
                {"id": 2, "customer": {"name": "Grace"}, "total": 12},
            ],
            "orders",
        )

        self.assertEqual(preview.rows, 2)
        self.assertEqual([column.name for column in preview.columns], ["id", "customer", "total"])
        self.assertEqual([column.inferred_type for column in preview.columns], ["number", "object", "number"])
        self.assertTrue(preview.columns[1].contains_nested_values)
        self.assertEqual(preview.warnings, ("Some columns contain nested records or arrays",))


if __name__ == "__main__":
    unittest.main()
