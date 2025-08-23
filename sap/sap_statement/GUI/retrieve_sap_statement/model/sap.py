import win32com.client
import pywintypes
from time import sleep
from PyQt6.QtWidgets import QMainWindow, QMessageBox


class SAPGUI:

    SAP_GRID_ID = "wnd[0]/usr/cntlGRID1/shellcont/shell"
    POP_UP_ID = "wnd[1]/tbar[0]/btn[0]"

    def get_connection(self, window : QMainWindow):
        """Returns the first SAP connection."""
        # logger = logging.getLogger(__name__)
        while (True):
            try:
                SapGuiAuto  = win32com.client.GetObject("SAPGUI")
                break
            except pywintypes.com_error:
                # logger.warning("SAP GUI is not running. Please start SAP GUI and try again.")
                answer = QMessageBox.warning(window,
                                    'SAP GUI not running',
                                    'SAP GUI is not running. Please start SAP GUI and try again.',
                                    QMessageBox.StandardButton.Ok | QMessageBox.StandardButton.Cancel)
                if answer == QMessageBox.StandardButton.Cancel:
                    return None
        application = SapGuiAuto.GetScriptingEngine
        return application.Connections(0)


    def accept_pop_up(self, session) -> None:
        """Accepts the pop-up window if it appears."""
        # logger = logging.getLogger(__name__)
        try:
            session.findById(self.__class__.POP_UP_ID).press()
            # logger.debug('Pop-up accepted')
        except pywintypes.com_error:
            # logger.debug('No pop-up to accept')
            pass


    def create_session(self, session) -> None:
        """Creates a new SAP session."""
        session.findById("wnd[0]").sendVKey(74) # Open a new SAP window

        # Click continue if Maximum number of SAP GUI sessions reached
        self.accept_pop_up(session)


    def get_last_session(self, window : QMainWindow, old_session, connection):
        """Returns the last SAP session."""
        # logger = logging.getLogger(__name__)
        
        self.create_session(old_session)
        # Access the last session (new window)
        n_children = connection.Sessions.Count
        # logger.debug(f'Number of SAP GUI sessions: {n_children}')
        while (n_children == 6): # Max number of sessions allowed are six.
            # logger.warning('Too many SAP GUI sessions!')
            answer = QMessageBox.warning(window,
                                    'Multiple SAP GUI Sessions',
                                    'Close the last SAP session',
                                    QMessageBox.StandardButton.Ok | QMessageBox.StandardButton.Cancel)
            if answer == QMessageBox.StandardButton.Cancel:
                return None
            self.create_session(old_session)
            n_children = connection.Sessions.Count
        
        # logger.debug('Waiting for SAP GUI session to be available')
        sleep(0.5)
        while(True):
            try:
                session = connection.Sessions(n_children)
                break
            except pywintypes.com_error:
                # logger.warning('SAP GUI session not available')
                sleep(0.5) #wait for half a second and try again
        # logger.debug(f'Connected to Last SAP GUI Session: {n_children}')
        return session
