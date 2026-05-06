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

    def test_inbox_search_intent_and_query(self):
        command = 'search inbox query "from:billing@example.com is:unread" top 5'
        details = parse_email_request(command)
        self.assertEqual(details["intent"], "inbox_search")
        self.assertEqual(details["query"], "from:billing@example.com is:unread")
        self.assertEqual(details["max_results"], 5)

    def test_reply_intent_extracts_message_id(self):
        command = 'reply to message id 18c9abXYZ123 saying "Thanks for sharing" send now'
        details = parse_email_request(command)
        self.assertEqual(details["intent"], "message_reply")
        self.assertEqual(details["message_id"], "18c9abXYZ123")
        self.assertTrue(details["send"])

    def test_label_apply_intent_extracts_target(self):
        command = 'apply label "Finance" to message id 18c9abXYZ123'
        details = parse_email_request(command)
        self.assertEqual(details["intent"], "label_apply")
        self.assertEqual(details["label_name"], "Finance")
        self.assertEqual(details["message_id"], "18c9abXYZ123")

    def test_schedule_send_extracts_time_and_disables_immediate_send(self):
        command = (
            'schedule email to gaurangbele178@gmail.com subject "Reminder" '
            'body "Please check the report" in 10 minutes'
        )
        details = parse_email_request(command)
        self.assertEqual(details["intent"], "schedule_send")
        self.assertTrue(bool(details["schedule_at"]))
        self.assertFalse(details["send"])

    def test_template_and_signature_management_intents(self):
        template_details = parse_email_request(
            'save template follow up subject "Quick Follow-up" body "Checking in on this"'
        )
        signature_details = parse_email_request(
            'save signature work body "Regards, Gaurang"'
        )
        self.assertEqual(template_details["intent"], "template_save")
        self.assertEqual(template_details["template_name"], "follow up")
        self.assertEqual(signature_details["intent"], "signature_save")
        self.assertEqual(signature_details["signature_name"], "work")


if __name__ == "__main__":
    unittest.main()
