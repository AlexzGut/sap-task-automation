from PyQt6.QtCore import QSize
from PyQt6.QtWidgets import QApplication, QWidget, QMainWindow, QVBoxLayout, QLabel, QLineEdit, QPushButton
import sys
import slots


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()

        self.setWindowTitle('Medical Expense Report')
        self.setFixedSize(QSize(300, 200))

        # === Widgets ===
        # labels
        lb_patient_id = QLabel('Patient ID')
        lb_customer_name = QLabel('Customer Name')

        # Inputs
        self.le_patient_id = QLineEdit()
        self.le_patient_id.setText('29949')
        self.le_customer_name = QLineEdit()
        self.le_customer_name.setText('John')

        # Buttons
        self.btn_submit = QPushButton('Submit')

        # === QLineEdit Signals ===

        # === QPushButton Signals ===
        self.btn_submit.clicked.connect(self.generate_report)

        # === Layouts ===
        v_layout = QVBoxLayout()
        v_layout.addWidget(lb_patient_id)
        v_layout.addWidget(self.le_patient_id)
        v_layout.addWidget(lb_customer_name)
        v_layout.addWidget(self.le_customer_name)
        v_layout.addWidget(self.btn_submit)

        # Container
        container = QWidget()
        container.setLayout(v_layout)

        # Set the central widget of the Main Window
        self.setCentralWidget(container)

    # === QPushButton Slots ===
    def generate_report(self):
        slots.download_mer(self.le_patient_id.text(), self.le_customer_name.text(), self)
        print('Completed!')


if __name__ == '__main__':
    app = QApplication(sys.argv)

    window = MainWindow()
    window.show()

    sys.exit(app.exec())