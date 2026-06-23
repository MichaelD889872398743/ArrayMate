import json
import unittest
from pathlib import Path

from arraymate.service import ArrayMateService


class ArrayMateServiceTests(unittest.TestCase):
    def test_load_text_tracks_available_arrays_and_default_selection(self):
        service = ArrayMateService()

        result = service.load_text('{"users": [{"name": "Ada"}], "empty": []}')

        self.assertEqual(result.array_keys, ["users", "empty"])
        self.assertEqual(result.selected_key, "users")
        self.assertEqual(result.selected_array, [{"name": "Ada"}])
        self.assertEqual(result.json_tree.display_path, "root")
        self.assertEqual([candidate.display_path for candidate in result.array_candidates], ["users", "empty"])
        self.assertEqual(service.get_array_data("empty"), [])

    def test_create_export_plan_resolves_extension_and_filename(self):
        service = ArrayMateService()

        plan = service.create_export_plan("out", "users", "CSV (.csv)")

        self.assertEqual(plan.output_format.label, "CSV")
        self.assertEqual(plan.filename, "users.csv")
        self.assertEqual(plan.file_path, str(Path("out") / "users.csv"))

    def test_export_array_writes_selected_data(self):
        service = ArrayMateService()
        service.load_text('{"users": [{"id": 1, "name": "Ada"}]}')
        output_path = Path("test_arraymate_service_users.json")

        try:
            plan = service.create_export_plan(".", output_path.stem, "JSON (.json)")
            result = service.export_array("users", plan)

            self.assertEqual(result.rows, 1)
            self.assertEqual(result.columns, 2)
            self.assertEqual(json.loads(output_path.read_text(encoding="utf-8")), [{"id": 1, "name": "Ada"}])
        finally:
            output_path.unlink(missing_ok=True)

    def test_export_array_supports_grouped_nested_candidate(self):
        service = ArrayMateService()
        service.load_text(
            '{"orders": [{"items": [{"sku": "A"}]}, {"items": [{"sku": "B"}, {"sku": "C"}]}]}'
        )
        output_path = Path("test_arraymate_service_items.json")

        try:
            plan = service.create_export_plan(".", output_path.stem, "JSON (.json)")
            result = service.export_array("orders[*].items", plan)

            self.assertEqual(result.rows, 3)
            self.assertEqual(json.loads(output_path.read_text(encoding="utf-8")), [{"sku": "A"}, {"sku": "B"}, {"sku": "C"}])
        finally:
            output_path.unlink(missing_ok=True)

    def test_grouped_nested_candidate_can_include_parent_metadata(self):
        service = ArrayMateService()
        service.load_text(
            '{"orders": ['
            '{"order_id": "ORD001", "customer_id": 1, "items": [{"sku": "A"}, {"sku": "B"}]},'
            '{"order_id": "ORD002", "customer_id": 2, "items": [{"sku": "C"}]}'
            ']}'
        )

        rows = service.get_array_data("orders[*].items", include_parent_metadata=True)

        self.assertEqual(
            rows,
            [
                {"order_id": "ORD001", "customer_id": 1, "sku": "A"},
                {"order_id": "ORD001", "customer_id": 1, "sku": "B"},
                {"order_id": "ORD002", "customer_id": 2, "sku": "C"},
            ],
        )

    def test_parent_table_can_unfold_nested_array_rows(self):
        service = ArrayMateService()
        service.load_text(
            '{"orders": ['
            '{"order_id": "ORD001", "customer_id": 1, "items": [{"sku": "A"}, {"sku": "B"}]},'
            '{"order_id": "ORD002", "customer_id": 2, "items": [{"sku": "C"}]}'
            ']}'
        )

        rows = service.get_table_data("orders", unfold_key="orders[*].items")

        self.assertEqual(
            rows,
            [
                {"sku": "A", "order_id": "ORD001", "customer_id": 1},
                {"sku": "B", "order_id": "ORD001", "customer_id": 1},
                {"sku": "C", "order_id": "ORD002", "customer_id": 2},
            ],
        )
        self.assertEqual(list(rows[0].keys()), ["order_id", "customer_id", "sku"])

    def test_parent_table_can_unfold_multiple_nested_levels(self):
        service = ArrayMateService()
        service.load_text(
            '{"orders": ['
            '{"order_id": "ORD001", "items": [{"sku": "A", "serials": [{"serial": "S1"}]}]}'
            ']}'
        )

        rows = service.get_table_data("orders", unfold_key="orders[*].items[*].serials")

        self.assertEqual(rows, [{"order_id": "ORD001", "sku": "A", "serial": "S1"}])
        self.assertEqual(list(rows[0].keys()), ["order_id", "sku", "serial"])

    def test_get_nested_array_candidates_returns_child_tables(self):
        service = ArrayMateService()
        service.load_text(
            '{"orders": ['
            '{"items": [{"sku": "A", "serials": [{"id": 1}]}], "notes": []}'
            ']}'
        )

        candidates = service.get_nested_array_candidates("orders")

        self.assertEqual([candidate.display_path for candidate in candidates], ["orders[*].items", "orders[*].items[*].serials"])

    def test_get_nested_array_candidates_limits_nested_levels(self):
        service = ArrayMateService()
        service.load_text(
            '{"root_rows": ['
            '{"a": [{"b": [{"c": [{"d": [{"id": 1}]}]}]}]}'
            ']}'
        )

        candidates = service.get_nested_array_candidates("root_rows", max_nested_levels=2)

        self.assertEqual([candidate.display_path for candidate in candidates], ["root_rows[*].a", "root_rows[*].a[*].b"])


if __name__ == "__main__":
    unittest.main()
