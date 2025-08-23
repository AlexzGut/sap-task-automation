from PyQt6.QtCore import QSize, QDate, Qt
from PyQt6.QtWidgets import QWidget, QVBoxLayout, QGridLayout, QLabel, QLineEdit, QCalendarWidget, QPushButton, QCheckBox, QDialog, QMessageBox
import os
import retrieve_sap_statement.controller.mer_controller as Mer


basedir = os.path.dirname(__file__)


class MedicalExpenseReport(QWidget):
    def __init__(self):
        super().__init__()

        self.setWindowTitle('Medical Expense Report')

        # === Widgets ===
        # labels
        lb_lower_date = QLabel('From')
        lb_upper_date = QLabel('To')

        # Inputs
        self.le_patient_id = QLineEdit()
        self.le_patient_id.setPlaceholderText('Patient ID')
 
        self.le_customer_name = QLineEdit()
        self.le_customer_name.setPlaceholderText('Customer Name (last name, first name)')

        # Check Boxes
        self.cb_filters = QCheckBox('Advanced Filters')

        # Filter Dates
        self.btn_lower_date = QPushButton(QDate.currentDate().toString('MMMM d, yyyy'))
        self.btn_upper_date = QPushButton(QDate.currentDate().toString('MMMM d, yyyy'))
        self.lower_date = QDate(2000,1,1)
        self.upper_date = QDate().currentDate()

        # Buttons
        self.btn_submit = QPushButton('Submit')

        # === QLineEdit Signals ===

        # === QCheckBox Signals ===
        self.cb_filters.stateChanged.connect(self.display_filters)

        # === QPushButton Signals ===
        self.btn_submit.clicked.connect(self.generate_report)

        # Show CalendarDialog
        self.btn_lower_date.clicked.connect(self.popup_lower_calendar)
        self.btn_upper_date.clicked.connect(self.popup_upper_calendar)

        # === Layouts ===
        v_main_layout = QVBoxLayout()

        filter_container = QWidget()
        filter_container.setFixedSize(QSize(260, 100))

        g_filters_layout = QGridLayout()
        g_filters_layout.setContentsMargins(0,0,0,0)
        g_filters_layout.addWidget(lb_lower_date, 0, 0)
        g_filters_layout.addWidget(lb_upper_date, 0, 1)
        g_filters_layout.addWidget(self.btn_lower_date, 1, 0)
        g_filters_layout.addWidget(self.btn_upper_date, 1, 1)

        # Hiding widgets in filters layout
        self.h_filters_widgets = [lb_lower_date, lb_upper_date, self.btn_lower_date, self.btn_upper_date]
        for widget in self.h_filters_widgets:
            widget.hide()

        filter_container.setLayout(g_filters_layout)

        # Add Widgets to the main layout
        v_main_layout.addStretch(1)
        v_main_layout.addWidget(self.le_patient_id)
        v_main_layout.addWidget(self.le_customer_name)
        v_main_layout.addWidget(self.cb_filters)
        v_main_layout.addWidget(filter_container)
        v_main_layout.addWidget(self.btn_submit)
        v_main_layout.addStretch(1)

        # Container
        self.setLayout(v_main_layout)

    # === QPushButton Slots ===
    def generate_report(self):
        Mer.download_mer(self.le_patient_id.text(), self.le_customer_name.text(), self.lower_date.toPyDate(), self.upper_date.toPyDate(), self)
        QMessageBox.information(
            self, 
            'Report Complete', 
            f'Medical Expense Report generated for {self.le_customer_name.text()}'
        )

    def popup_lower_calendar(self):
        calendar = CalendarDialog(self.lower_date, self)
        if calendar.exec():
            self.lower_date = calendar.select_date()
            self.btn_lower_date.setText(self.lower_date.toString('MMMM d, yyyy'))

    def popup_upper_calendar(self):
        calendar = CalendarDialog(self.upper_date, self)
        if calendar.exec():
            self.upper_date = calendar.select_date()
            self.btn_upper_date.setText(self.upper_date.toString('MMMM d, yyyy'))

    # === QCheckBox Slots ===
    def display_filters(self, s):
        if (s == Qt.CheckState.Checked.value):
            for widget in self.h_filters_widgets:
                self.lower_date = QDate().currentDate()
                self.upper_date = QDate().currentDate()
                widget.show()
        else:
            for widget in self.h_filters_widgets:
                self.lower_date = None
                self.upper_date = None
                widget.hide()


class CalendarDialog(QDialog):
    def __init__(self, selected_date : QDate=None, parent=None):
        super().__init__(parent)
        self.setWindowFlags(Qt.WindowType.Dialog | Qt.WindowType.FramelessWindowHint)

        self.calendar = QCalendarWidget(self)
        self.calendar.setVerticalHeaderFormat(QCalendarWidget.VerticalHeaderFormat.NoVerticalHeader)

        if selected_date:
            self.calendar.setSelectedDate(selected_date)

        self.resize(self.calendar.sizeHint())

        self.calendar.selectionChanged.connect(self.accept)

    def select_date(self):
        return self.calendar.selectedDate()