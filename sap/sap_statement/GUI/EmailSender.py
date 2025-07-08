import win32com.client 
import os
from typing import List, Dict


class EmailSender:
    def __init__(self):
        self.outlook = win32com.client.Dispatch('Outlook.Application')

    def setup_template(self, template_path : str) -> None:
        if not os.path.exists(template_path):
            raise FileNotFoundError(f'Template not found at: {template_path}')
        self._mail = self.outlook.CreateItemFromTemplate(template_path)

    def set_recipients(self, recipients : List[str]) -> None:
        self.email_exists()
        self._mail.To = recipients

    def set_subject(self, subject : str) -> None:
        self.email_exists()
        self._mail.Subject = subject

    def set_attachments(self, file_path : str) -> None:
        self.email_exists()
        self._mail.Attachments.Add(file_path)

    def update_body(self, replacements : Dict) -> None:
        self.email_exists()
        email_body = self._mail.HTMLBody
        for key, value in replacements.items():
            email_body = email_body.replace(key, value)
        self._mail.HTMLBody = email_body

    def send_email(self) -> None:
        self.email_exists()
        self._mail.Send()

    def display_email(self) -> None:
        self.email_exists()
        self._mail.Display()

    def email_exists(self) -> None:
        if not self._mail:
            raise ValueError('No email template specified')

