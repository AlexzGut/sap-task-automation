import sys, os
from PyQt6.QtWidgets import QApplication, QWidget, QVBoxLayout, QHBoxLayout, QMainWindow, QLineEdit, QLabel, QPushButton, QRadioButton, QButtonGroup
from PyQt6.QtGui import QIcon, QRegularExpressionValidator
from PyQt6.QtCore import QRegularExpression
import SapStatement


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
        self.setFixedSize(300, 200) # width, height

        # === Widgets ===
        # Text input
        self.le_account_number = QLineEdit()
        self.le_account_number.setPlaceholderText('Account Number')
        self.le_account_number.textChanged.connect(self.hide_label)
        # Because there is a inputMask() set on the line edit, the returnPressed() signal will only be emitted
        # when the input follow the inputMask()
        self.le_account_number.returnPressed.connect(self.retrieve_statement) 

        self.le_month = QLineEdit()
        self.le_month.setPlaceholderText('Month')
        self.le_month.textChanged.connect(self.hide_label)
        self.le_month.returnPressed.connect(self.retrieve_statement)

        # Text input validators
        rx_account_number = QRegularExpression('^[0-9]{10}$')
        self.le_account_number.setValidator(QRegularExpressionValidator(rx_account_number, self))
        rx_month = QRegularExpression('^(0?[1-9]|1[012])$')
        self.le_month.setValidator(QRegularExpressionValidator(rx_month, self))

        # Radio Buttons
        self.rb_sdm = QRadioButton('SDM')
        self.rb_medisystem = QRadioButton('MediSystem')
        self.rb_medisystem.setChecked(True)
        # Radio Button Group (One selected at a time)
        self.company = QButtonGroup()
        self.company.addButton(self.rb_sdm)
        self.company.addButton(self.rb_medisystem)

        # Button
        button = QPushButton('Submit', clicked=self.retrieve_statement)

        # Label for messages
        self.lb_message = QLabel()
        self.lb_message.hide()
        self.lb_message.setStyleSheet('''
            color: #ED4337
        ''')

        # Create Vertical layout and add Widgets to the layout
        verticalLayout = QVBoxLayout()
        verticalLayout.addWidget(self.le_account_number)
        verticalLayout.addWidget(self.le_month)
        verticalLayout.addWidget(self.lb_message)

        # Create Horizontal layout for radio buttons
        horizontalLayout = QHBoxLayout()
        horizontalLayout.addWidget(self.rb_sdm)
        horizontalLayout.addWidget(self.rb_medisystem)
        # Add Horizontal layout to Vertical layout
        verticalLayout.addLayout(horizontalLayout)

        # Add button
        verticalLayout.addWidget(button)

        # Create a dummy QWidget to set the layout and add to the MainWindow 
        container = QWidget()
        container.setLayout(verticalLayout)
        self.setCentralWidget(container)

    def retrieve_statement(self):
        if self.validate_required_field() and self.valid_month():
            # Set retrieve_las
            month = self.parse_month()
            if self.rb_sdm.isChecked():
                company = 'sdm'
            if self.rb_medisystem.isChecked():
                company = 'medi'

            SapStatement.execute(self, self.le_account_number.text(), month, company)
            
            self.le_account_number.clear()
            self.le_month.clear()

    # Field validations
    def validate_required_field(self):
        if self.le_account_number.text() == '':
            self.lb_message.show()
            self.lb_message.setText('Missing Account number')
            self.le_account_number.setFocus()
            return False

        if self.le_month.text() == '':
            self.lb_message.show()
            self.lb_message.setText('Missing Month')
            self.le_month.setFocus()
            return False
        return True

    def valid_month(self):
        if self.le_month.text() == '0':
            self.lb_message.show()
            self.lb_message.setText('Month must be 01 - 12')
            self.le_month.setFocus()
            return False
        return True

    def parse_month(self) -> str:
        """return a valid month (1-12 or 01-12)."""
        month = self.le_month.text()
        return month if len(month) == 2 else month.zfill(2)
    
    def hide_label(self):
        self.lb_message.hide()
      

if __name__ == '__main__':
    app = QApplication(sys.argv)
    app.setStyleSheet('''
        Qwidget {
            font-size: 25px;                  
        }
                      
        QLineEdit, QPushButton{
            height: 30px;
        }
    ''')
    mainWindow = SapStatementRetrieval()
    mainWindow.show()

    sys.exit(app.exec())