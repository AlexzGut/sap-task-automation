from PyQt6.QtWidgets import QWidget, QCheckBox, QLineEdit, QLabel, QVBoxLayout
from PyQt6.QtCore import QRegularExpression, Qt
from PyQt6.QtGui import QRegularExpressionValidator


class SendEmailSection(QWidget):
    def __init__(self):
        super().__init__()

        # === Widgets ===
        # Check Box
        self.check_box = QCheckBox('Send Email')
        self.check_box .setFixedHeight(40)
        self.check_box.setCheckState(Qt.CheckState.Unchecked)
        # Check Box signals
        self.check_box.stateChanged.connect(self.state_changed)

        # Line Edit (Customer email)
        self.le_cx_email = QLineEdit()
        self.le_cx_email .setFixedHeight(40)
        self.le_cx_email.setPlaceholderText('Email@Address.com')
        rx_email = QRegularExpression('^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+.[a-zA-Z]{2,}$')
        self.le_cx_email.setValidator(QRegularExpressionValidator(rx_email, self))
        self.le_cx_email.setVisible(False)
        self.le_cx_email.textChanged.connect(self.hide_label)

        # Label for messages
        self.lb_message = QLabel()
        self.lb_message.hide()
        self.lb_message.setFixedHeight(20)
        self.lb_message.setStyleSheet('''
            color: red;
        ''')

        # Add widgets to a layout
        self.v_layout = QVBoxLayout()
        self.v_layout.setContentsMargins(0,0,0,0)
        self.v_layout.addWidget(self.check_box)
        self.v_layout.addWidget(self.le_cx_email)
        self.v_layout.addWidget(self.lb_message)
        
        # Set the section layout
        self.setLayout(self.v_layout)

    def hide_label(self):
        self.lb_message.hide()

    def state_changed(self, state):
        if Qt.CheckState(state) == Qt.CheckState.Checked:
            self.le_cx_email.setVisible(True)
        else:
            self.le_cx_email.setVisible(False)
            self.le_cx_email.clear()
        self.hide_label()

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