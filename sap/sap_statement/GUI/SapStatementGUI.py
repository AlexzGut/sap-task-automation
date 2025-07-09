import sys, os
from PyQt6.QtWidgets import QApplication, QWidget, QVBoxLayout, QHBoxLayout, QMainWindow, QLineEdit, QLabel, QPushButton, QRadioButton, QButtonGroup, QCheckBox, QMessageBox
from PyQt6.QtGui import QIcon, QRegularExpressionValidator
from PyQt6.QtCore import QRegularExpression, Qt
import SapStatement
import EmailSender
import calendar


basedir = os.path.dirname(__file__)

try:
    from ctypes import windll  # Only exists on Windows.
    myappid = 'loblaw.sap.statemetretrieval.01'
    windll.shell32.SetCurrentProcessExplicitAppUserModelID(myappid)
except ImportError:
    pass


class SapStatementRetrieval(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle('Retrieve SAP Statement')
        self.setWindowIcon(QIcon(os.path.join(basedir, 'icons', 'loblaw.ico')))
        self.setFixedSize(300, 300) # width, height

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
        self.lb_message.setFixedHeight(30)
        self.lb_message.setStyleSheet('''
            color: red;
        ''')

        # === Send Email Option ===
        # Check Box
        self.check_box = QCheckBox('Send Email')
        self.check_box .setFixedHeight(40)
        self.check_box.setCheckState(Qt.CheckState.Unchecked)
        # Check Box signals
        self.check_box.stateChanged.connect(self.state_changed)

        # Line Edit (Customer email)
        self.le_cx_email = QLineEdit()
        self.le_cx_email .setFixedHeight(40)
        self.le_cx_email.setPlaceholderText('email@address.com')
        rx_email = QRegularExpression('^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+.[a-zA-Z]{2,}$')
        self.le_cx_email.setValidator(QRegularExpressionValidator(rx_email, self))
        self.le_cx_email.hide()
        self.le_cx_email.textChanged.connect(self.hide_label)

        # === Layouts ===
        # Create Vertical layout and add Widgets to the layout
        vertical_layout = QVBoxLayout()
        vertical_layout.addWidget(self.le_account_number)
        vertical_layout.addWidget(self.le_month)
        vertical_layout.addWidget(self.lb_message)

        # Create Horizontal layout for radio buttons
        horizontal_layout = QHBoxLayout()
        # horizontalLayout.addWidget(self.rb_sdm)
        horizontal_layout.addWidget(self.rb_medisystem)

        email_layout = QVBoxLayout()
        email_layout.addWidget(self.check_box)
        email_layout.addWidget(self.le_cx_email)

        # Add Horizontal layout to Vertical layout
        vertical_layout.addLayout(horizontal_layout)
        vertical_layout.addLayout(email_layout)

        # Add button
        vertical_layout.addWidget(button)

        # Create a dummy QWidget to set the layout and add to the MainWindow 
        container = QWidget()
        container.setLayout(vertical_layout)
        self.setCentralWidget(container)

    def retrieve_statement(self):
        if self.validate_account_number() and self.validate_month() and self.validate_email():
            # Set retrieve_las
            month = self.parse_month()
            if self.rb_sdm.isChecked():
                company = 'sdm'
            if self.rb_medisystem.isChecked():
                company = 'medi'

            context = SapStatement.execute(self, self.le_account_number.text(), month, company)

            if self.check_box.checkState() == Qt.CheckState.Checked:
                if context.get('file_name'):
                    download_path = os.path.join('C:\\', 'Users', os.getlogin(), 'Downloads')
                    month_name = calendar.month_name[int(self.le_month.text())]                    
                    template_values = {'account_field' : f'<b>{self.le_account_number.text()}</b>',
		                               'month_field' : f'<b>{month_name}</b>'}
                    
                    email = EmailSender.EmailSender()
                    email.setup_template(os.path.join(basedir, 'email_templates', 'monthly_statement_template.msg'))
                    email.set_recipients(self.le_cx_email.text())
                    email.set_subject(f'{month_name} Statement')
                    email.set_attachments(os.path.join(download_path, context.get('file_name')))
                    email.update_body(template_values)
                    email.send_email()

                    QMessageBox.information(self,
                                    'Email Sent',
                                    f'Email to {self.le_cx_email.text()} was sent successfully')

            self.le_account_number.clear()
            self.le_month.clear()
            self.le_cx_email.clear()
            self.check_box.setCheckState(Qt.CheckState.Unchecked)

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
    
    def validate_email(self):
        if self.check_box.checkState() == Qt.CheckState.Checked:
            if not self.le_cx_email.text():
                self.lb_message.show()
                self.lb_message.setText('Missing Email Address')
                self.le_cx_email.setFocus()
                return False

            if not self.le_cx_email.hasAcceptableInput():
                self.lb_message.show()
                self.lb_message.setText('Invalid email address')
                self.le_cx_email.setFocus()
                return False
            
        return True

    def parse_month(self) -> str:
        """return a valid month (1-12 or 01-12)."""
        month = self.le_month.text()
        return month if len(month) == 2 else month.zfill(2)
    
    def hide_label(self):
        self.lb_message.hide()

    def state_changed(self, state):
        if Qt.CheckState(state) == Qt.CheckState.Checked:
            self.le_cx_email.show()
        else:
            self.le_cx_email.hide()
            self.le_cx_email.clear()
        self.hide_label()
      

if __name__ == '__main__':
    app = QApplication(sys.argv)

    mainWindow = SapStatementRetrieval()
    mainWindow.show()

    sys.exit(app.exec())