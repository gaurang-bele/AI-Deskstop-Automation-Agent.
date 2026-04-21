import unittest

from email_automation import build_email_actions, parse_email_request


class EmailAutomationParserTests(unittest.TestCase):
    def test_cc_bcc_are_parsed_into_correct_fields(self):
        command = (
            "open google and compose a mail to gaurangbele178@gmail.com "
            "also add gaurangbele123@gmail.com to cc and gaurangbele007@gmail.com to bcc "
            'subject "missing documents for application" '
            'body "you have missing documents like aadhar card, pan card etc"'
        )
        details = parse_email_request(command)
        self.assertEqual(details["to"], ["gaurangbele178@gmail.com"])
        self.assertEqual(details["cc"], ["gaurangbele123@gmail.com"])
        self.assertEqual(details["bcc"], ["gaurangbele007@gmail.com"])

    def test_attachment_path_is_parsed(self):
        command = (
            '"Send gaurang at gaurangbele178@gmail.com an email saying we are ready with the Q1 report '
            'and attach the file from C:\\Users\\DELL\\Downloads\\Orders_Table_cleaned.xlsx" polish it'
        )
        details = parse_email_request(command)
        self.assertEqual(details["to"], ["gaurangbele178@gmail.com"])
        self.assertEqual(details["attachments"], ["C:\\Users\\DELL\\Downloads\\Orders_Table_cleaned.xlsx"])
        self.assertNotIn("attach the file", details["body"].lower())

    def test_build_actions_include_attachment_before_send(self):
        details = {
            "to": ["gaurangbele178@gmail.com"],
            "cc": [],
            "bcc": [],
            "subject": "Q1 update",
            "body": "Please see attached.",
            "browser": "chrome",
            "attachments": ["C:\\Users\\DELL\\Downloads\\Orders_Table_cleaned.xlsx"],
            "send": True,
        }
        actions = build_email_actions(details)
        attach_index = next(i for i, item in enumerate(actions) if item.get("label") == "gmail_attachment_1")
        send_index = next(i for i, item in enumerate(actions) if item.get("keys") == ["ctrl", "enter"])
        attach_action_index = next(i for i, item in enumerate(actions) if item.get("action") == "attach_file")
        open_url_action = next(item for item in actions if item.get("action") == "open_url")
        open_url_index = next(i for i, item in enumerate(actions) if item.get("action") == "open_url")
        self.assertGreaterEqual(open_url_index, 0)
        self.assertIn("compose_url", actions[attach_action_index])
        self.assertEqual(open_url_action.get("app"), "chrome")
        self.assertEqual(actions[attach_action_index].get("browser"), "chrome")
        self.assertLess(attach_action_index, send_index)
        self.assertLess(attach_index, send_index)


if __name__ == "__main__":
    unittest.main()
