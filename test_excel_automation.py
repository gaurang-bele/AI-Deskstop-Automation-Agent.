import sys
import unittest

from excel_automation import (
    ExcelAutomation,
    ExcelCommandParser,
    column_letter_to_index,
    index_to_column_letter,
    is_excel_command,
    normalize_function_name,
    parse_excel_actions_from_command,
    try_handle_excel_command,
    validate_excel_actions,
)


class ExcelAutomationParserTests(unittest.TestCase):
    def setUp(self):
        self.parser = ExcelCommandParser()

    def test_csv_path_and_delete_rows_parse_exactly(self):
        operations = self.parser.parse(
            "open C:\\Users\\DELL\\Downloads\\Orders_Table (2).csv and delete 6 rows"
        )

        self.assertEqual([operation.action for operation in operations], ["open_workbook", "delete_rows"])
        self.assertEqual(operations[0].params["reference"], r"C:\Users\DELL\Downloads\Orders_Table (2).csv")
        self.assertEqual(operations[1].params["start_row"], 1)
        self.assertEqual(operations[1].params["end_row"], 6)

    def test_open_and_aggregate_command_parses_with_context(self):
        operations = self.parser.parse("Open ABC excel. Find sum of column XYZ.")

        self.assertEqual([operation.action for operation in operations], ["open_workbook", "aggregate"])
        self.assertEqual(operations[0].params["reference"], "ABC")
        self.assertEqual(operations[1].params["function"], "sum")
        self.assertEqual(operations[1].params["source_type"], "column")
        self.assertEqual(operations[1].params["source"], "XYZ")
        self.assertEqual(operations[1].params["workbook"], "ABC")

    def test_sheet_and_cell_update_command_parses(self):
        operations = self.parser.parse("Use sheet Summary and set cell D2 to Approved")

        self.assertEqual([operation.action for operation in operations], ["select_sheet", "set_cell_value"])
        self.assertEqual(operations[1].params["target"], "D2")
        self.assertEqual(operations[1].params["value"], "Approved")
        self.assertEqual(operations[1].params["sheet"], "Summary")

    def test_numeric_column_aggregate_parses(self):
        operations = self.parser.parse(
            "open C:\\Users\\DELL\\Downloads\\Orders_Table (2).csv and write average of column 3"
        )

        self.assertEqual([operation.action for operation in operations], ["open_workbook", "aggregate"])
        self.assertEqual(operations[1].params["source"], "3")

    def test_edit_sort_and_filter_commands_parse(self):
        operations = self.parser.parse(
            "Set cell B2 to Approved and sort by column Status descending and filter column Status equals Approved"
        )

        self.assertEqual(
            [operation.action for operation in operations],
            ["set_cell_value", "sort_sheet", "filter_column"],
        )
        self.assertEqual(operations[0].params["target"], "B2")
        self.assertEqual(operations[0].params["value"], "Approved")
        self.assertEqual(operations[1].params["column"], "Status")
        self.assertEqual(operations[1].params["order"], "descending")
        self.assertEqual(operations[2].params["operator"], "equals")
        self.assertEqual(operations[2].params["value"], "Approved")


class ExcelAutomationHelperTests(unittest.TestCase):
    def test_parser_outputs_valid_shared_actions(self):
        actions = parse_excel_actions_from_command(
            "open C:\\Users\\DELL\\Downloads\\Orders_Table (2).csv and delete rows 2 to 10"
        )
        valid, note = validate_excel_actions(actions)

        self.assertTrue(valid, note)
        self.assertEqual(actions[0]["action"], "open_workbook")
        self.assertEqual(actions[1]["action"], "delete_rows")

    def test_column_index_round_trip(self):
        self.assertEqual(column_letter_to_index("A"), 1)
        self.assertEqual(column_letter_to_index("XYZ"), 16900)
        self.assertEqual(index_to_column_letter(1), "A")
        self.assertEqual(index_to_column_letter(16900), "XYZ")

    def test_excel_detection_and_fallback_behavior(self):
        self.assertTrue(is_excel_command("Find sum of column Sales in workbook Budget"))
        self.assertFalse(is_excel_command("Open chrome and search for cricket scores"))

        result = try_handle_excel_command("Open excel")
        self.assertFalse(result.handled)

    def test_filter_criteria_and_function_aliases(self):
        automation = ExcelAutomation()

        self.assertEqual(normalize_function_name("average"), "average")
        self.assertEqual(automation._build_filter_criteria("contains", "Done"), "*Done*")
        self.assertEqual(automation._build_filter_criteria(">=", 10), ">=10")


if __name__ == "__main__":
    suite = unittest.defaultTestLoader.loadTestsFromModule(sys.modules[__name__])
    runner = unittest.TextTestRunner(verbosity=2, stream=sys.stdout)
    result = runner.run(suite)

    print("\n" + "=" * 50)
    print("EXCEL AUTOMATION TEST SUMMARY")
    print("=" * 50)
    print(f"Ran     : {result.testsRun}")
    print(f"Passed  : {result.testsRun - len(result.failures) - len(result.errors)}")
    print(f"Failed  : {len(result.failures) + len(result.errors)}")
    print("=" * 50)

    raise SystemExit(0 if result.wasSuccessful() else 1)
