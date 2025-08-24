import os
from PyQt6.QtWidgets import QWidget, QVBoxLayout, QListWidget, QListWidgetItem, QLabel, QLineEdit, QPushButton, QMessageBox, QFileDialog#, QProgressDialog
from PyQt6.QtGui import QIcon, QRegularExpressionValidator, QValidator
from PyQt6.QtCore import Qt, QRegularExpression
from time import sleep
from pypdf import PdfReader, PdfWriter, errors
from retrieve_sap_statement.view.section.send_email import SendEmailSection
from retrieve_sap_statement.model.EmailSender import EmailSender


basedir = os.path.dirname(__file__)

class Encryption(QWidget):
    def __init__(self):
        super().__init__()

        # === Widgets ===
        # Buttons
        self.btn_select_folder = QPushButton('\U0001F4C2 Select Files . . .')
        self.btn_select_folder.setFixedHeight(40)

        self.btn_submit= QPushButton('Submit')
        self.btn_submit.setFixedHeight(40)
        self.btn_submit.setEnabled(False)

        # Labels
        self.lb_files = QLabel()
        self.lb_files.setText('No files selected')

        # Label for messages
        self.lb_message = QLabel()
        self.lb_message.hide()
        self.lb_message.setFixedHeight(20)
        self.lb_message.setStyleSheet('''
            color: red;
        ''')

        # List
        self.li_files = QListWidget()
        self.li_files.setSelectionMode(QListWidget.SelectionMode.ExtendedSelection)

        # Text
        self.le_password = QLineEdit()
        self.le_password.setPlaceholderText('Encryption Password')
        self.le_password.setFixedHeight(40)

        rx_password = QRegularExpression('^(?=.*[A-Z])(?=.*[a-z])(?=.*[0-9]).{8,}$')
        self.le_password.setValidator(QRegularExpressionValidator(rx_password, self))

        # === Signals ===
        self.btn_select_folder.clicked.connect(self.handle_btn_select_folder)
        self.btn_submit.clicked.connect(self.handle_btn_submit)
        self.le_password.textChanged.connect(self.validate_rx_password)

        #=== Sections ===
        self.email_section = SendEmailEncryption()

        # === Create and add widgets to Vertical layout ===
        main_layout = QVBoxLayout()
        main_layout.addWidget(self.lb_files)
        main_layout.addWidget(self.li_files)
        main_layout.addWidget(self.btn_select_folder)
        main_layout.addWidget(self.le_password)
        main_layout.addWidget(self.lb_message)
        main_layout.addWidget(self.email_section)
        main_layout.addWidget(self.btn_submit)

        self.setLayout(main_layout)

    # === Slots ===
    def handle_btn_select_folder(self):
        self.li_files.clear()

        self.filenames, selected_filter = QFileDialog.getOpenFileNames(
            parent=self,
            caption='Select Files',
            directory=os.getcwd(),
            filter='PDF Files (*.pdf)',
            options=QFileDialog.Option.ShowDirsOnly
        )

        pdf_icon = QIcon(os.path.join(basedir, '..', 'resources', 'icons', 'acrobat-pdf.webp'))
        for file in self.filenames:
            self.li_files.addItem(QListWidgetItem(pdf_icon, os.path.basename(file)))

        self.lb_files.setText('Files to be processed:')
        self.btn_submit.setEnabled(True)

    def validate_rx_password(self):
        validator_state, _, _ = self.le_password.validator().validate(self.le_password.text(), 0)
        if validator_state != QValidator.State.Acceptable:
            self.lb_message.show()
            self.lb_message.setText('At least One Uppercase - One Lowercase - One number')
            return False
        
        self.lb_message.hide()
        return True

    def handle_btn_submit(self):
        # progress_bar = QProgressDialog('Processing Data', None, 0, 100, self)
        # progress_bar.setValue(0)
        # progress_increment = 100 // len(self.filenames)
        if not self.validate_encryption_password():
            return

        if self.email_section.check_box.checkState() == Qt.CheckState.Checked:
            if not self.email_section.validate_account_number() and not self.email_section.validate_email():
                return
        already_encrypted_files = []
        is_encrypted = False
        opened_files = []
        is_opened_file = False
        encryption_succesful = []

        for pdf_file in self.filenames:

            statement_reader = PdfReader(pdf_file)
            try:
                statement_writer = PdfWriter(clone_from=statement_reader)
            except errors.FileNotDecryptedError:
                is_encrypted = True
                already_encrypted_files.append(os.path.basename(pdf_file))
                continue

            # Add a password to the pdf file
            statement_password = self.le_password.text()
            statement_writer.encrypt(statement_password, algorithm="AES-256")

            # Save the new PDF to a file
            try:
                with open(pdf_file, "wb") as f:
                    statement_writer.write(f)
                    encryption_succesful.append(pdf_file)
            except PermissionError:
                is_opened_file = True
                opened_files.append(os.path.basename(pdf_file))
                continue
            # progress_bar.setValue(progress_bar.value() + progress_increment)

        if is_encrypted:
            QMessageBox.warning(
                self,
                'Files encryption',
                f'Already encrypted files:\n{'\n'.join(already_encrypted_files)}',
                defaultButton=QMessageBox.StandardButton.Ok)
        if is_opened_file:
            QMessageBox.warning(
                self,
                'Files encryption',
                f'Close\n{'\n'.join(opened_files)}\nand try again',
                defaultButton=QMessageBox.StandardButton.Ok)
        
        if encryption_succesful:
            QMessageBox.information(
                self,
                'Files encryption',
                f'Encryption Succesful:\n{'\n'.join([os.path.basename(file) for file in encryption_succesful])}',
                defaultButton=QMessageBox.StandardButton.Ok)
            
        if self.email_section.check_box.checkState() == Qt.CheckState.Checked:
            template_values = {'account_field' : f'<b>{self.email_section.le_account_number.text()}</b>'}
            email = EmailSender()
            email.setup_template(os.path.join(basedir, '..', 'resources', 'email_templates', 'medi', 'monthly_statements_template.msg'))
            email.set_recipients(self.email_section.le_cx_email.text())
            email.set_subject('Monthly Statements')
            email.set_attachments(encryption_succesful)
            email.update_body(template_values)

            template_values['password_field'] = f'<b>{self.le_password.text()}</b>'
            email_password = EmailSender()
            email_password.setup_template(os.path.join(basedir, '..', 'resources', 'email_templates', 'medi', 'monthly_statement_password_template.msg'))
            email_password.set_recipients(self.email_section.le_cx_email.text())
            email_password.set_subject('Monthly Statements')
            email_password.update_body(template_values)

            if is_encrypted or is_opened_file:
                email.display_email()
                email_password.display_email()
            else:
                email.send_email()
                email_password.send_email()
        
        self.li_files.clear()
        self.le_password.clear()
        self.email_section.le_account_number.clear()
        self.email_section.le_cx_email.clear()
        self.email_section.check_box.setCheckState(Qt.CheckState.Unchecked)

    def validate_encryption_password(self):
        if not self.le_password.text():
            self.lb_message.show()
            self.lb_message.setText('Missing Encryption Password')
            self.le_password.setFocus()
            return False
        
        if not self.validate_rx_password():
            self.lb_message.show()
            self.lb_message.setText('Encryption Password: Min 8 chars, 1 uppercase, 1 lowercase, 1 digit.')
            self.le_password.setFocus()
            return False
        
        return True


class SendEmailEncryption(SendEmailSection):
    def __init__(self):
        super().__init__()

        # === Widgets ===
        # Text box
        self.le_account_number = QLineEdit()
        self.le_account_number.setPlaceholderText('Account Number')
        self.le_account_number.setFixedHeight(40)
        self.le_account_number.setVisible(False)
        rx_account_number = QRegularExpression('^[0-9]{10}$')
        self.le_account_number.setValidator(QRegularExpressionValidator(rx_account_number, self))

        # Labels
        self.lb_account_error = QLabel()
        self.lb_account_error.hide()
        self.lb_account_error.setFixedHeight(20)
        self.lb_account_error.setStyleSheet('''
            color: red;
        ''')

        self.v_layout.addWidget(self.le_account_number)
        self.v_layout.addWidget(self.lb_account_error)

    def state_changed(self, state):
        if Qt.CheckState(state) == Qt.CheckState.Checked:
            self.le_cx_email.setVisible(True)
            self.le_account_number.setVisible(True)
        else:
            self.le_cx_email.setVisible(False)
            self.le_account_number.setVisible(False)
            self.le_cx_email.clear()
            self.le_account_number.clear()
        self.hide_label()
        self.lb_account_error.hide()

    def validate_account_number(self):
        if not self.le_account_number.text():
            self.lb_account_error.show()
            self.lb_account_error.setText('Missing Account number')
            self.le_account_number.setFocus()
            return False
        
        if len(self.le_account_number.text()) < 10:
            self.lb_account_error.show()
            self.lb_account_error.setText('Account number must be 10 digits long')
            self.le_account_number.setFocus()
            return False
        
        return True
