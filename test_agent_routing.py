import unittest

from agent import _is_web_research_command, _parse_research_request, _browser_for_command


class AgentRoutingTests(unittest.TestCase):
    def test_fetch_information_detected_without_google_keyword(self):
        self.assertTrue(_is_web_research_command("fetch information on titanium price and density"))

    def test_excel_command_not_treated_as_research(self):
        self.assertFalse(
            _is_web_research_command(
                r"open C:\Users\DELL\Downloads\Orders_Table (2).csv and find sum of column transaction_id"
            )
        )

    def test_parse_research_request_extracts_query_and_output(self):
        query, fields, output_file = _parse_research_request(
            "fetch information on bmw and save in summary.txt"
        )
        self.assertEqual(query.lower(), "bmw")
        self.assertEqual(output_file, "summary.txt")
        self.assertIn("overview", fields)
        self.assertIn("key facts", fields)

    def test_research_browser_honors_edge_request(self):
        self.assertEqual(_browser_for_command("open edge and fetch information about 0603YG105ZAT4A"), "edge")

    def test_parse_research_request_extracts_explicit_fields_after_its(self):
        query, fields, output_file = _parse_research_request(
            "open google and fetch information about 0603YG105ZAT2A, its Voltage, Temperature Coefficient, Operating Temperature"
        )
        self.assertEqual(query, "0603YG105ZAT2A")
        self.assertIn("overview", fields)
        self.assertIn("voltage", fields)
        self.assertIn("temperature coefficient", fields)
        self.assertIn("operating temperature", fields)
        self.assertEqual(output_file, "response.txt")


if __name__ == "__main__":
    unittest.main()
