from PyQt6.QtWidgets import QWidget, QVBoxLayout, QLineEdit, QPushButton, QMessageBox
import os
import retrieve_sap_statement.controller.mer_controller as Mer
from retrieve_sap_statement.view.section.mer_parameters import AdvanceParametersSection


basedir = os.path.dirname(__file__)


class MedicalExpenseReport(QWidget):
    def __init__(self):
        super().__init__()

        self.setWindowTitle('Medical Expense Report')

        # === Widgets ===
        # Inputs
        self.le_patient_id = QLineEdit()
        self.le_patient_id.setPlaceholderText('Patient ID')
        self.le_patient_id.setFixedHeight(40)
 
        self.le_customer_name = QLineEdit()
        self.le_customer_name.setPlaceholderText('Customer Name (last name, first name)')
        self.le_customer_name.setFixedHeight(40)

        # Buttons
        self.btn_submit = QPushButton('Submit')
        self.btn_submit.setFixedHeight(40)

        # === QPushButton Signals ===
        self.btn_submit.clicked.connect(self.generate_report)

        # === Sections ===
        self.param_section = AdvanceParametersSection()

        # === Layouts ===
        v_main_layout = QVBoxLayout()

        # Add Widgets to the main layout
        v_main_layout.addStretch(1)
        v_main_layout.addWidget(self.le_patient_id)
        v_main_layout.addWidget(self.le_customer_name)
        v_main_layout.addWidget(self.param_section)
        v_main_layout.addWidget(self.btn_submit)
        v_main_layout.addStretch(1)

        # Container
        self.setLayout(v_main_layout)

    # === QPushButton Slots ===
    def generate_report(self):
        parameters = self.param_section.get_state()
        Mer.download_mer(self.le_patient_id.text(), self.le_customer_name.text(), parameters, self)
        QMessageBox.information(
            self, 
            'Report Complete', 
            f'Medical Expense Report generated for {self.le_customer_name.text()}'
        )