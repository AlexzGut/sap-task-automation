import win32com.client 
import os
from typing import List


class EmailSender:
    def __init__(self):
        self.outlook = win32com.client.Dispatch('Outlook.Application')

    def setup_template(self, template_path : str) -> None:
        if not os.path.exists(template_path):
            raise FileNotFoundError(f'Template not found at: {template_path}')
        self._mail = self.outlook.CreateItemFromTemplate(template_path)

    def set_recipients(self, recipients : List[str]) -> None:
        if not self._mail:
            raise ValueError('No email template specified')
        self._mail.To = ';'.join(recipients)

    def send_email(self) -> None:
        if not self._mail:
            raise ValueError('No email template specified')
        self._mail.Send()

    def display_email(self) -> None:
        if not self._mail:
            raise ValueError('No email template specified')
        self._mail.Send()

