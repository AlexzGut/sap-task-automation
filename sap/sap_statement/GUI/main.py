from PyQt6.QtWidgets import QApplication
from retrieve_sap_statement.view.main_window import MainWindow
import sys


if __name__ == '__main__':
    app = QApplication(sys.argv)

    mainWindow = MainWindow()
    mainWindow.show()

    sys.exit(app.exec())