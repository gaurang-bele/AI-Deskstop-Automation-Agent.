import unittest

from email_automation import parse_email_request


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


if __name__ == "__main__":
    unittest.main()
