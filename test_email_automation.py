import unittest

from email_automation import parse_email_request, wants_model_polish


class EmailAutomationTests(unittest.TestCase):
    def test_auto_draft_request_triggers_model_polish(self):
        command = (
            "open brave and mail samyakgajbhiye1965@gmail.com something about AI "
            "select the subject and main body by yourself the mail should be interesting"
        )
        details = parse_email_request(command)

        self.assertTrue(wants_model_polish(command))
        self.assertEqual(details["subject"], "AI")
        self.assertTrue(details["send"])
        self.assertNotEqual(details["subject"].lower(), "and main")

    def test_do_not_send_overrides_implicit_send(self):
        command = "mail user@example.com about project status do not send"
        details = parse_email_request(command)
        self.assertFalse(details["send"])

    def test_subject_is_syntax_parses_correctly(self):
        command = "mail user@example.com subject is Quarterly Update body: Revenue grew 18 percent send it"
        details = parse_email_request(command)
        self.assertEqual(details["subject"], "Quarterly Update")
        self.assertEqual(details["body"], "Revenue grew 18 percent")
        self.assertTrue(details["send"])


if __name__ == "__main__":
    unittest.main()
