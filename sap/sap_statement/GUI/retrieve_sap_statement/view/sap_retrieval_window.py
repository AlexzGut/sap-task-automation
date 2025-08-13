import os
from PyQt6.QtWidgets import QWidget, QVBoxLayout, QHBoxLayout, QLineEdit, QLabel, QPushButton, QRadioButton, QButtonGroup, QCheckBox, QMessageBox
from PyQt6.QtGui import QRegularExpressionValidator
from PyQt6.QtCore import QRegularExpression, Qt
from retrieve_sap_statement.controller.main_controller import execute
from retrieve_sap_statement.model.EmailSender import EmailSender
from retrieve_sap_statement.view.section.send_email import SendEmailSection
import calendar
from pypdf import PdfReader, PdfWriter
from datetime import datetime


basedir = os.path.dirname(__file__)

class SAP_Retrieval(QWidget):
    def __init__(self):
        super().__init__()
        # === Widgets ===
        # Text input
        self.le_account_number = QLineEdit()
        self.le_account_number.setFixedHeight(40)
        self.le_account_number.setPlaceholderText('Account Number')
        self.le_account_number.textChanged.connect(self.hide_label)
        # Because there is a inputMask() set on the line edit, the returnPressed() signal will only be emitted
        # when the input follow the inputMask()
        self.le_account_number.returnPressed.connect(self.retrieve_statement) 

        self.le_month = QLineEdit()
        self.le_month.setFixedHeight(40)
        self.le_month.setPlaceholderText('Month (MM)')
        self.le_month.textChanged.connect(self.hide_label)
        self.le_month.returnPressed.connect(self.retrieve_statement)

        # Text input validators
        rx_account_number = QRegularExpression('^[0-9]{10}$')
        self.le_account_number.setValidator(QRegularExpressionValidator(rx_account_number, self))
        rx_month = QRegularExpression('^(0?[1-9]|1[012])$')
        self.le_month.setValidator(QRegularExpressionValidator(rx_month, self))

        # Radio Buttons
        self.rb_sdm = QRadioButton('SDM')
        self.rb_sdm.setFixedHeight(40)
        self.rb_medisystem = QRadioButton('MediSystem')
        self.rb_medisystem.setChecked(True)
        # Radio Button Group (One selected at a time)
        self.company = QButtonGroup()
        self.company.addButton(self.rb_sdm)
        self.company.addButton(self.rb_medisystem)

        # Button
        button = QPushButton('Submit', clicked=self.retrieve_statement)
        button.setFixedHeight(40)

        # Label for messages
        self.lb_message = QLabel()
        self.lb_message.hide()
        self.lb_message.setFixedHeight(20)
        self.lb_message.setStyleSheet('''
            color: red;
        ''')

        # === Layouts ===
        # Create Vertical layout and add Widgets to the layout
        vertical_layout = QVBoxLayout()
        vertical_layout.addWidget(self.le_account_number)
        vertical_layout.addWidget(self.le_month)
        vertical_layout.addWidget(self.lb_message)

        # Create Horizontal layout for radio buttons
        horizontal_layout = QHBoxLayout()
        horizontal_layout.addWidget(self.rb_sdm)
        horizontal_layout.addWidget(self.rb_medisystem)

        # Add Horizontal layout to Vertical layout
        vertical_layout.addLayout(horizontal_layout)

        # Add Send Email section to vertical Layout
        self.email_section = SendEmailSection()
        vertical_layout.addWidget(self.email_section)

        # Add button
        vertical_layout.addWidget(button)

        main_layout = QVBoxLayout()
        main_layout.addStretch(1)
        main_layout.addLayout(vertical_layout)
        main_layout.addStretch(1)
        self.setLayout(main_layout)

    def retrieve_statement(self):
        if self.validate_account_number() and self.validate_month() and self.email_section.validate_email():
            month = self.parse_month()
            if self.rb_sdm.isChecked():
                company = 'sdm'
            if self.rb_medisystem.isChecked():
                company = 'medi'

            context = execute(self, self.le_account_number.text(), month, company)

            if self.email_section.check_box.checkState() == Qt.CheckState.Checked:
                if context.get('file_name'):
                    download_path = os.path.join('C:\\', 'Users', os.getlogin(), 'Downloads')
                    month_name = calendar.month_name[int(self.le_month.text())]                    
                    template_values = {'account_field' : f'<b>{self.le_account_number.text()}</b>',
                                        'month_field' : f'<b>{month_name}</b>'}
                    
                    statement_pdf = os.path.join(download_path, context.get('file_name'))
                    
                    statement_reader = PdfReader(statement_pdf)
                    statement_writer = PdfWriter(clone_from=statement_reader)

                    # Add a password to monthly statement (statement_pdf) 
                    statement_password = f'{company.upper()}{month_name}{datetime.today().year}'
                    statement_writer.encrypt(statement_password, algorithm="AES-256")

                    # Save the new PDF to a file
                    with open(statement_pdf, "wb") as f:
                        statement_writer.write(f)
                    
                    # Send encrypted statement
                    email_statement = EmailSender()
                    email_statement.setup_template(os.path.join(basedir, '..', 'resources', 'email_templates', company, 'monthly_statement_template.msg'))
                    email_statement.set_recipients(self.email_section.le_cx_email.text())
                    email_statement.set_subject(f'{month_name} Statement')
                    email_statement.set_attachments([statement_pdf])
                    email_statement.update_body(template_values)
                    email_statement.send_email()

                    # Send password email
                    template_values['password_field'] = f'<b>{statement_password}</b>'
                    email_password = EmailSender()
                    email_password.setup_template(os.path.join(basedir, '..', 'resources', 'email_templates', company, 'monthly_statement_password_template.msg'))
                    email_password.set_recipients(self.email_section.le_cx_email.text())
                    email_password.set_subject(f'{month_name} Statement - Password')
                    email_password.update_body(template_values)
                    email_password.send_email()
                    
                    QMessageBox.information(self,
                                    'Email Sent',
                                    f'Email to {self.email_section.le_cx_email.text()} was sent successfully')

            self.le_account_number.clear()
            self.le_month.clear()
            self.email_section.le_cx_email.clear()
            self.email_section.check_box.setCheckState(Qt.CheckState.Unchecked)

    # Field validations
    def validate_month(self):
        if not self.le_month.text():
            self.lb_message.show()
            self.lb_message.setText('Missing Month')
            self.le_month.setFocus()
            return False
    
        return True
    
    def validate_account_number(self):
        if not self.le_account_number.text():
            self.lb_message.show()
            self.lb_message.setText('Missing Account number')
            self.le_account_number.setFocus()
            return False
        
        if len(self.le_account_number.text()) < 10:
            self.lb_message.show()
            self.lb_message.setText('Account number must be 10 digits long')
            self.le_account_number.setFocus()
            return False
        
        return True

    def parse_month(self) -> str:
        """return a valid month (1-12 or 01-12)."""
        month = self.le_month.text()
        return month if len(month) == 2 else month.zfill(2)
    
    def hide_label(self):
        self.lb_message.hide()