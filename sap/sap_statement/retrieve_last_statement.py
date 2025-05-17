import win32com.client
import pywintypes
from time import sleep
import os
import logging
from logging.handlers import RotatingFileHandler
import sys


# == Global Constants ===
# Define a constant for the SAP grid path 
SAP_GRID_ID = "wnd[0]/usr/cntlGRID1/shellcont/shell"
POP_UP_ID = "wnd[1]/tbar[0]/btn[0]"


#== Main function ===
def main():
    # Initialize logging
    logger = setup_logging()

    logger.info('Opening new SAP GUI Session . . .')
    connection = get_sap_connection()
    session = connection.Sessions(0)
    logger.info('New SAP GUI Session opened')

    logger.info('Connecting with new SAP GUI Session. . . .')
    session = get_last_sap_session(session, connection)
    logger.info('Connected to SAP GUI Session')

    account_number = get_account_number()
    month = '0' + input('Enter the statement Month: ')

    access_tcode_fbl5n(session)
    while(True):
        try:
            communication_method = get_communication_method(session, account_number)
            break
        except pywintypes.com_error:
            logger.info('Customer NOT found - Incorrect account number')  
            account_number = get_account_number()          
    
    go_to_sap_access_screen(session)

    logger.info(f'Accessing {communication_method.strip()} directory . . .')
    access_tcode_al11(session)
    double_click(session, 'DIRNAME', '/interfaces')
    double_click(session, 'NAME', 'FII30040')
    double_click(session, 'NAME', 'archive')
    double_click(session, 'NAME', 'medi')
    double_click(session, 'NAME', 'others') if communication_method == 'E-Mail' else double_click(session, 'NAME', 'mail')

    logger.info('Filtering statements by account number . . .')
    filter_statements_by_account_number(session, account_number)
    statement_exists = statements_exists(session, month)

    if statement_exists:
        logger.info('Statement found')
        logger.info('Downloading statement . . .')
        session.findById("wnd[0]").sendVKey(14) # Access file attributes (Shift + F2)
        directory = session.findById(SAP_GRID_ID).GetCellValue(0, 'VALUE')
        file_name = session.findById(SAP_GRID_ID).GetCellValue(1, 'VALUE')

        go_to_sap_access_screen(session)

        download_statement_from_cg3y(session, directory, file_name)
        logger.info(f'Statement downloaded - {file_name}')
    else:
        go_to_sap_access_screen(session)
        logger.info('Statement not found')

    session.ActiveWindow.Close()
    input('Press Enter to exit . . .')


# == Logging setup ===
class UserFilter(logging.Filter):
    def filter(self, record):
        record.user = os.getlogin()
        return True
    
    
def setup_logging() -> logging.Logger:
    logger = logging.getLogger(__name__)
    logger.setLevel(logging.DEBUG)
    #Format: Date loggerLevel userID functionName: loggingMessage
    FH_FORMAT = '%(asctime)s %(levelname)-8s %(user)-10s %(funcName)s:  %(message)s'
    #Format: loggerLevel userID: loggingMessage
    CH_FORMAT = '%(levelname)-8s %(user)-10s:  %(message)s'

    if getattr(sys, 'frozen', False):
        # Running as PyInstaller bundle
        log_dir = os.path.dirname(sys.executable)
    else:
        log_dir = os.path.dirname(os.path.abspath(__file__))

    # File handler
    # Create a file handler and set level to DEBUG
    log_path = os.path.join(log_dir, 'retrieve_last_statement.log')
    file_handler = RotatingFileHandler(log_path, maxBytes=1_000_000, backupCount=5, encoding='utf-8')
    file_handler.setLevel(logging.DEBUG)
    # # Set formatter for file handler
    file_handler.setFormatter(logging.Formatter(FH_FORMAT))

    # Console handler
    # Create a console handler and set level to INFO
    console_handler = logging.StreamHandler()
    console_handler.setLevel(logging.INFO)
    # Set formatter for console handler
    console_handler.setFormatter(logging.Formatter(CH_FORMAT))

    # Add the console and file handler to the logger
    if not logger.hasHandlers():
        logger.addHandler(console_handler)
        logger.addHandler(file_handler)

    # Add the user filter to the logger
    user_filter = UserFilter()
    logger.addFilter(user_filter)

    return logger


# == Helper functions ===
def accept_pop_up(session):
    """Accepts the pop-up window if it appears."""
    logger = logging.getLogger(__name__)
    try:
        session.findById(POP_UP_ID).press()
        logger.debug('Pop-up accepted')
    except pywintypes.com_error:
        logger.debug('No pop-up to accept')
        pass


def double_click(session, column_name : str, row_value : str) -> None:
    """Double-clicks a row in the SAP grid where column matches row_value."""
    grid = session.FindById(SAP_GRID_ID)
    for i in range(grid.rowCount):
        if grid.GetCellValue(i, column_name) == row_value:
            grid.setCurrentCell(str(i), column_name)
            grid.selectedRows = str(i)
            grid.doubleClickCurrentCell()
            break


def go_to_sap_access_screen(session) -> None:
    """Navigates back to the SAP Easy Access screen."""
    for _ in range(3):
        session.findById("wnd[0]/tbar[0]/btn[12]").press()


def get_sap_connection():
    """Returns the first SAP connection."""
    logger = logging.getLogger(__name__)
    while (True):
        try:
            SapGuiAuto  = win32com.client.GetObject("SAPGUI")
            break
        except pywintypes.com_error:
            logger.warning("SAP GUI is not running. Please start SAP GUI and try again.")
            input('Press Enter to continue . . .') 
    application = SapGuiAuto.GetScriptingEngine
    return application.Connections(0)

def create_sap_session(session):
    """Creates a new SAP session."""
    session.findById("wnd[0]").sendVKey(74) # Open a new SAP window

    # Click continue if Maximum number of SAP GUI sessions reached
    accept_pop_up(session)

def get_last_sap_session(old_session, connection):
    """Returns the last SAP session."""
    logger = logging.getLogger(__name__)
    
    create_sap_session(old_session)
    # Access the last session (new window)
    n_children = connection.Sessions.Count
    logger.debug(f'Number of SAP GUI sessions: {n_children}')
    while (n_children == 6): # Max number of sessions allowed are six.
        logger.warning('Too many SAP GUI sessions!')
        input('Close a SAP window and press Enter to continue . . .')
        create_sap_session(old_session)
        n_children = connection.Sessions.Count
    
    logger.debug('Waiting for SAP GUI session to be available')
    sleep(0.5)
    while(True):
        try:
            session = connection.Sessions(n_children)
            break
        except pywintypes.com_error:
            logger.warning('SAP GUI session not available')
            print(n_children)
            sleep(0.5) #wait for half a second and try again
    logger.debug(f'Connected to Last SAP GUI Session: {n_children}')
    return session


def get_communication_method(session, account_number : str) -> str:
    """Gets the communication method for the given account number."""
    logger = logging.getLogger(__name__)

    while (True):
        logger.debug(f'Execute FBL5N search')
        session.findById("wnd[0]/usr/ctxtDD_KUNNR-LOW").text = account_number
        session.findById("wnd[0]/tbar[1]/btn[8]").press() #Execute search

        # Check for a pop up window and check ok if present.
        accept_pop_up(session)
        
        try:
            #Try to open XD03 tab
            session.findById("wnd[0]/tbar[1]/btn[34]").press()
            logger.debug('T-Code XD03 Accessed')
            break
        except pywintypes.com_error:
            raise pywintypes.com_error('Customer NOT found - Incorrect account number')

    logger.debug('Accessing communication method')
    return session.findById("wnd[0]/usr/subSUBTAB:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB01/ssubSUBSC:SAPLATAB:0201/subAREA1:SAPMF02D:7111/subADDRESS:SAPLSZA1:0300/subCOUNTRY_SCREEN:SAPLSZA1:0301/cmbADDR1_DATA-DEFLT_COMM").text


def is_account_valid(account_number):
    """Checks if the account number is a 10-digit string."""
    acct_number_length = 10
    return (len(account_number) == acct_number_length) and (account_number.isdigit())


def get_account_number():
    """Prompts the user for a valid 10-digit account number."""
    logger = logging.getLogger(__name__)
    account_number = input('Enter account number: ')

    while (not is_account_valid(account_number)):
        logger.warning('Invalid Account number - Must be 10 digits long')
        account_number = input('Enter account number: ')
    return account_number


def access_tcode_fbl5n(session):
    """Accesses the T-Code FBL5N."""
    logger = logging.getLogger(__name__)

    logger.debug('Accessing T-Code FBL5N')
    session.findById("wnd[0]/tbar[0]/okcd").text = "FBL5N"
    session.findById("wnd[0]").sendVKey(0)

    logger.debug('Setting parameters for FBL5N')
    # Set the initial parameters for the FBL5N transaction
    session.findById("wnd[0]/usr/ctxtDD_BUKRS-LOW").text = "L604"
    session.findById("wnd[0]/usr/ctxtDD_BUKRS-HIGH").text = "L607"
    session.findById("wnd[0]/usr/radX_AISEL").select()


def access_tcode_al11(session):
    """Accesses the T-Code AL11."""
    logger = logging.getLogger(__name__)

    logger.debug('Accessing T-Code AL11')
    session.findById("wnd[0]/tbar[0]/okcd").text = "AL11"
    session.findById("wnd[0]").sendVKey(0)
    

def download_statement_from_cg3y(session, source_file : str, target_file : str) -> None:
    """Downloads the statement from the T-Code CG3Y."""
    logger = logging.getLogger(__name__)

    logger.debug('Accessing T-Code CG3Y')
    # Open T-Code CG3Y
    session.findById("wnd[0]/tbar[0]/okcd").text = "CG3Y"
    session.findById("wnd[0]").sendVKey(0)

    logger.debug(f'Setting parameters for CG3Y Source: {source_file} Target: {target_file}')
    # Set the parameters for the download
    session.findById("wnd[1]/usr/txtRCGFILETR-FTAPPL").text = source_file + '/' + target_file
    session.findById("wnd[1]/usr/ctxtRCGFILETR-FTFRONT").text = target_file
    session.findById("wnd[1]/usr/chkRCGFILETR-IEFOW").selected = True

    logger.debug('Downloading file')
    # Download the file
    session.findById("wnd[1]/tbar[0]/btn[13]").press()
    session.findById("wnd[1]/tbar[0]/btn[12]").press()
    logger.debug('File downloaded')


def filter_statements_by_account_number(session, account_number : str) -> None:
    """Filters the statement by account number."""
    logger = logging.getLogger(__name__)

    logger.debug('Filtering statements by account number')
    grid = session.findById(SAP_GRID_ID)
    grid.contextMenu()
    grid.selectContextMenuItem("&FILTER")
    session.findById("wnd[1]/usr/ssub%_SUBSCREEN_FREESEL:SAPLSSEL:1105/ctxt%%DYN001-LOW").text = '*' + account_number + '*'
    accept_pop_up(session)


def statements_exists(session, month : str) -> bool:
    """Checks if statements exists for the given account number."""
    logger = logging.getLogger(__name__)

    column_name = 'NAME'
    grid = session.FindById(SAP_GRID_ID)
    for i in range(grid.rowCount):
        if grid.GetCellValue(i, column_name).startswith(month):
            grid.setCurrentCell(str(i), column_name)
            grid.selectedRows = str(i)
            logger.debug(f'Statement found')
            return True
    logger.debug(f'Statement not found')
    return False


if __name__ == '__main__':
    main()

