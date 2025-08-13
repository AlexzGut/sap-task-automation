import os
from PyQt6.QtWidgets import QWidget, QVBoxLayout, QHBoxLayout, QMainWindow, QLineEdit, QLabel, QPushButton, QRadioButton, QButtonGroup, QCheckBox, QMessageBox, QTabWidget
from PyQt6.QtGui import QIcon
from retrieve_sap_statement.controller.main_controller import execute
from retrieve_sap_statement.view.sap_retrieval_window import SAP_Retrieval
from retrieve_sap_statement.view.encryption_window import Encryption


basedir = os.path.dirname(__file__)

try:
    from ctypes import windll  # Only exists on Windows.
    myappid = 'loblaw.sap.statemetretrieval.01'
    windll.shell32.SetCurrentProcessExplicitAppUserModelID(myappid)
except ImportError:
    pass


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle('Retrieve SAP Statement')
        self.setWindowIcon(QIcon(os.path.join(basedir, '..', 'resources', 'icons', 'loblaw.ico')))
        self.setFixedSize(414, 600) # width, height

        # Create a Tab layout
        tab_widget = QTabWidget()
        tab_widget.setTabPosition(QTabWidget.TabPosition.West)

        # Create QWidgets to add to the TabWidget 
        sap_retrieval_container = SAP_Retrieval()
        encryption = Encryption()

        tabs = {
            sap_retrieval_container: QIcon(os.path.join(basedir, '..', 'resources', 'icons', 'tabs', 'document-pdf.png')),
            encryption: QIcon(os.path.join(basedir, '..', 'resources', 'icons', 'tabs', 'lock.png'))
        }
        for tab, label in tabs.items():
            tab_widget.addTab(tab, label, None)
        tab_widget.setTabIcon

        # Set the central widget of the Main window. This widget is the QTabWidget
        self.setCentralWidget(tab_widget)