from PyQt6.QtWidgets import QWidget, QVBoxLayout, QGridLayout, QLabel, QCheckBox, QPushButton, QCalendarWidget, QDialog
from PyQt6.QtCore import QDate, Qt


class AdvanceParametersSection(QWidget):
    def __init__(self):
        super().__init__()

        # labels
        lb_lower_date = QLabel('From')
        lb_upper_date = QLabel('To')

        # Check Boxes
        self.cb_filters = QCheckBox('Advanced Filters')

        # Filter Dates
        self.btn_lower_date = QPushButton(QDate.currentDate().toString('MMMM d, yyyy'))
        self.btn_upper_date = QPushButton(QDate.currentDate().toString('MMMM d, yyyy'))
        self.lower_date = QDate(2000,1,1)
        self.upper_date = QDate().currentDate()

        # === QCheckBox Signals ===
        self.cb_filters.stateChanged.connect(self.display_filters)

        # === Button Signals ===
        self.btn_lower_date.clicked.connect(self.popup_lower_calendar)
        self.btn_upper_date.clicked.connect(self.popup_upper_calendar)

        # === Layouts ===
        main_layout = QVBoxLayout()
        main_layout.setContentsMargins(0,0,0,0)

        g_filters_layout = QGridLayout()
        g_filters_layout.addWidget(lb_lower_date, 0, 0)
        g_filters_layout.addWidget(lb_upper_date, 0, 1)
        g_filters_layout.addWidget(self.btn_lower_date, 1, 0)
        g_filters_layout.addWidget(self.btn_upper_date, 1, 1)

        main_layout.addWidget(self.cb_filters)
        main_layout.addLayout(g_filters_layout)

        # Hiding widgets in filters layout
        self.h_filters_widgets = [lb_lower_date, lb_upper_date, self.btn_lower_date, self.btn_upper_date]
        for widget in self.h_filters_widgets:
            widget.hide()

        self.setLayout(main_layout)

    # === Button Slots ===
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
                widget.show()
            self.lower_date = QDate().currentDate()
            self.upper_date = QDate().currentDate()
        else:
            for widget in self.h_filters_widgets:
                widget.hide()
            self.lower_date = QDate(2000,1,1)
            self.upper_date = QDate().currentDate()

    def get_state(self):
        return {
            'lower_date': self.lower_date.toPyDate(),
            'upper_date': self.upper_date.toPyDate(),
        }


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