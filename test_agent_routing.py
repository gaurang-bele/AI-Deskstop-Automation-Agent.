import unittest
from pathlib import Path
import tempfile
from unittest.mock import patch

from agent import (
    on_command,
    _build_email_intent_for_model,
    _extract_part_codes,
    _extract_findings_from_clipboard,
    _is_relevant_part_clipboard,
    _is_web_research_command,
    _parse_research_request,
    _save_part_code_result,
    _split_research_and_email_command,
)


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
        self.assertIn("price", fields)

    def test_split_combined_part_code_and_email_command(self):
        command = (
            'get me details of this product 0603YG105ZAT2A search after that '
            'open google and mail gaurangbele178@gmail.com subject "missing documents" '
            'body "you have missing documents" send it polish it'
        )
        research, email = _split_research_and_email_command(command)
        self.assertIsNotNone(research)
        self.assertIsNotNone(email)
        self.assertIn("0603YG105ZAT2A", research)
        self.assertTrue(email.lower().startswith("mail "))

    def test_noisy_clipboard_does_not_match_part_code(self):
        noisy = "samyak, + 1 other\n2 people\nYour video on\nmicrophone on\nLoading..."
        self.assertFalse(_is_relevant_part_clipboard(noisy, ["0603YG105ZAT2A"]))

    def test_description_does_not_trigger_on_microphone_word(self):
        noisy = "Your video on microphone on loading"
        findings = _extract_findings_from_clipboard(noisy, ["description", "manufacturer", "datasheet"])
        self.assertEqual(findings, {})

    def test_extracts_manufacturer_and_datasheet_from_part_text(self):
        text = (
            "0603YG105ZAT2A capacitor manufacturer KYOCERA AVX. "
            "Datasheet https://example.com/0603YG105ZAT2A.pdf"
        )
        findings = _extract_findings_from_clipboard(text, ["manufacturer", "description", "datasheet", "specifications"])
        self.assertIn("manufacturer", findings)
        self.assertIn("datasheet", findings)

    def test_extract_part_codes_deduplicates(self):
        command = "search 0603YG105ZAT2A and 0603YG105ZAT2A and ECJ1VF1C105Z"
        self.assertEqual(_extract_part_codes(command), ["0603YG105ZAT2A", "ECJ1VF1C105Z"])

    def test_save_part_code_result_keeps_per_code_findings(self):
        part_codes = ["0603YG105ZAT2A", "ECJ1VF1C105Z"]
        findings = {
            "0603YG105ZAT2A": {
                "manufacturer": ["KYOCERA"],
                "description": ["Capacitor for filtering."],
                "specifications": ["10uF", "16V"],
                "datasheet": ["https://example.com/0603YG105ZAT2A.pdf"],
            },
            "ECJ1VF1C105Z": {
                "manufacturer": ["PANASONIC"],
                "description": ["Automotive MLCC capacitor."],
                "specifications": ["X7R", "10uF", "16V"],
                "datasheet": ["https://example.com/ECJ1VF1C105Z.pdf"],
            },
        }

        with tempfile.TemporaryDirectory() as tmp_dir:
            output_file = str(Path(tmp_dir) / "part_code_results.txt")
            _save_part_code_result(part_codes, findings, output_file)
            written = Path(output_file).read_text(encoding="utf-8")

        self.assertIn("Part Code    : 0603YG105ZAT2A", written)
        self.assertIn("Description  : Capacitor for filtering.", written)
        self.assertIn("Part Code    : ECJ1VF1C105Z", written)
        self.assertIn("Description  : Automotive MLCC capacitor.", written)

    def test_email_intent_for_model_prefers_clean_subject_when_body_is_noisy(self):
        details = {
            "subject": "AI",
            "body": "open brave and mail user@example.com something about AI by yourself",
            "raw_text": "mail user@example.com about AI",
        }
        intent = _build_email_intent_for_model("mail user@example.com about AI", details)
        self.assertIn("about: AI", intent)

    def test_email_command_polishes_topic_and_adds_send_shortcut(self):
        captured: dict[str, object] = {}

        def fake_generate(intent_text: str, recipient_email: str | None = None) -> dict:
            captured["intent_text"] = intent_text
            captured["recipient_email"] = recipient_email
            return {"subject": "AI Innovation Update", "body": "Hello, here are top AI updates this week."}

        def fake_execute(actions: list[dict]) -> list[dict]:
            captured["actions"] = actions
            return [{"step": i + 1, "action": action.get("action"), "params": action, "success": True, "note": "ok"} for i, action in enumerate(actions)]

        command = (
            "open brave and mail samyakgajbhiye1965@gmail.com something about AI "
            "select the subject and main body by yourself the mail should be interesting"
        )

        with patch("agent.generate_professional_email", side_effect=fake_generate), \
             patch("agent._execute_actions_until_failure", side_effect=fake_execute), \
             patch("agent.write_thinking", return_value=None), \
             patch("agent.write_response", return_value=None), \
             patch("agent.get_usage_totals", return_value={"requests": 0, "prompt_tokens": 0, "completion_tokens": 0, "total_tokens": 0}):
            on_command(command)

        actions = captured.get("actions", [])
        self.assertEqual(captured.get("intent_text"), "Write an engaging professional email about: AI")
        self.assertEqual(captured.get("recipient_email"), "samyakgajbhiye1965@gmail.com")
        self.assertTrue(any(action.get("action") == "press_key" and action.get("keys") == ["ctrl", "enter"] for action in actions))

    def test_email_command_stops_when_model_draft_fails(self):
        command = (
            "open brave and mail samyakgajbhiye1965@gmail.com something about AI "
            "select the subject and main body by yourself the mail should be interesting"
        )

        captured: dict[str, object] = {}

        with patch("agent.generate_professional_email", side_effect=RuntimeError("401 Unauthorized")), \
             patch("agent._execute_actions_until_failure", side_effect=lambda actions: captured.setdefault("executed", True)), \
             patch("agent.write_thinking", return_value=None), \
             patch("agent.write_response", return_value=None), \
             patch("agent.write_error", side_effect=lambda cmd, err: captured.update({"error": err})), \
             patch("agent.get_usage_totals", return_value={"requests": 0, "prompt_tokens": 0, "completion_tokens": 0, "total_tokens": 0}):
            on_command(command)

        self.assertIn("Email drafting requested but model draft failed", str(captured.get("error", "")))
        self.assertNotIn("executed", captured)


if __name__ == "__main__":
    unittest.main()
