import win32com.client
import pywintypes
from time import sleep


# == Global Constants ===
# Define a constant for the SAP grid path 
SAP_GRID_ID = "wnd[0]/usr/cntlGRID1/shellcont/shell"


#== Main function ===
def main():
    print('Opening new SAP window . . .')
    connection = get_sap_connection()
    session = connection.Sessions(0)
    session.findById("wnd[0]").sendVKey(74) # Open a new SAP window
    print('New SAP window opened')
    print('Connecting to SAP Session . . .')
    session = get_last_sap_session(connection)
    print('Connected to SAP Session')

    account_number = get_account_number()
    month = '0' + input('Enter the statement Month: ')

    access_tcode_fbl5n(session)
    while(True):
        try:
            communication_method = get_communication_method(session, account_number)
            break
        except pywintypes.com_error:
            print('Customer NOT found - Incorrect account number')  
            account_number = get_account_number()          
    
    go_to_sap_access_screen(session)

    print(f'Accessing {communication_method.strip()} directory . . .')
    access_tcode_al11(session)
    double_click(session, 'DIRNAME', '/interfaces')
    double_click(session, 'NAME', 'FII30040')
    double_click(session, 'NAME', 'archive')
    double_click(session, 'NAME', 'medi')
    double_click(session, 'NAME', 'others') if communication_method == 'E-Mail' else double_click(session, 'NAME', 'mail')

    print('Filtering statements by account number . . .')
    filter_statements_by_account_number(session, account_number)
    statement_exists = statements_exists(session, month)

    if statement_exists:
        print('Statement found')
        print('Downloading statement . . .')
        session.findById("wnd[0]").sendVKey(14) # Access file attributes (Shift + F2)
        directory = session.findById(SAP_GRID_ID).GetCellValue(0, 'VALUE')
        file_name = session.findById(SAP_GRID_ID).GetCellValue(1, 'VALUE')

        go_to_sap_access_screen(session)

        download_statement_from_cg3y(session, directory, file_name)
        print('Statement downloaded')
    else:
        go_to_sap_access_screen(session)
        print('Statement not found')

    session.ActiveWindow.Close()
    input('Press Enter to exit . . .')


# == Helper functions ===
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
    SapGuiAuto  = win32com.client.GetObject("SAPGUI")
    application = SapGuiAuto.GetScriptingEngine
    return application.Connections(0)


def get_last_sap_session(connection):
    """Returns the last SAP session."""
    # Access the last session (new window)
    print('Connecting with new SAP GUI window.')
    n_children = connection.Sessions.Count
    while (n_children == 6): # Max number of sessions allowed are six.
        print('Too many sessions!')
        input('Close a SAP window and press Enter > ')
        n_children = connection.Sessions.Count
    
    sleep(0.5)
    while(True):
        try:
            session = connection.Sessions(n_children)
            break
        except pywintypes.com_error:
            sleep(0.5) #wait for half a second and try again
    return session


def get_communication_method(session, account_number : str) -> str:
    """Gets the communication method for the given account number."""

    while (True):
        session.findById("wnd[0]/usr/ctxtDD_KUNNR-LOW").text = account_number
        session.findById("wnd[0]/tbar[1]/btn[8]").press() #Execute search

        # Check for a pop up window and check ok if present.
        try:    
            session.findById("wnd[1]/tbar[0]/btn[0]").press()
        except Exception:
            pass

        try:
            #Try to open XD03 tab
            session.findById("wnd[0]/tbar[1]/btn[34]").press()
            break
        except pywintypes.com_error:
            raise pywintypes.com_error('Customer NOT found - Incorrect account number')

    return session.findById("wnd[0]/usr/subSUBTAB:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB01/ssubSUBSC:SAPLATAB:0201/subAREA1:SAPMF02D:7111/subADDRESS:SAPLSZA1:0300/subCOUNTRY_SCREEN:SAPLSZA1:0301/cmbADDR1_DATA-DEFLT_COMM").text


def is_account_valid(account_number):
    """Checks if the account number is a 10-digit string."""
    acct_number_length = 10
    return (len(account_number) == acct_number_length) and (account_number.isdigit())


def get_account_number():
    """Prompts the user for a valid 10-digit account number."""

    account_number = input('Enter account number: ')

    while (not is_account_valid(account_number)):
        print('Invalid Account number: must be 10 digits long')
        account_number = input('Enter account number: ')
    return account_number


def access_tcode_fbl5n(session):
    """Accesses the T-Code FBL5N."""
    session.findById("wnd[0]/tbar[0]/okcd").text = "FBL5N"
    session.findById("wnd[0]").sendVKey(0)

    # Set the initial parameters for the FBL5N transaction
    session.findById("wnd[0]/usr/ctxtDD_BUKRS-LOW").text = "L604"
    session.findById("wnd[0]/usr/ctxtDD_BUKRS-HIGH").text = "L607"
    session.findById("wnd[0]/usr/radX_AISEL").select()


def access_tcode_al11(session):
    """Accesses the T-Code AL11."""
    session.findById("wnd[0]/tbar[0]/okcd").text = "AL11"
    session.findById("wnd[0]").sendVKey(0)
    

def download_statement_from_cg3y(session, source_file : str, target_file : str) -> None:
    """Downloads the statement from the T-Code CG3Y."""
    # Open T-Code CG3Y
    session.findById("wnd[0]/tbar[0]/okcd").text = "CG3Y"
    session.findById("wnd[0]").sendVKey(0)

    # Set the parameters for the download
    session.findById("wnd[1]/usr/txtRCGFILETR-FTAPPL").text = source_file + '/' + target_file
    session.findById("wnd[1]/usr/ctxtRCGFILETR-FTFRONT").text = target_file
    session.findById("wnd[1]/usr/chkRCGFILETR-IEFOW").selected = True

    # Download the file
    session.findById("wnd[1]/tbar[0]/btn[13]").press()
    session.findById("wnd[1]/tbar[0]/btn[12]").press()

def filter_statements_by_account_number(session, account_number : str) -> None:
    """Filters the statement by account number."""
    grid = session.findById(SAP_GRID_ID)
    grid.contextMenu()
    grid.selectContextMenuItem("&FILTER")
    session.findById("wnd[1]/usr/ssub%_SUBSCREEN_FREESEL:SAPLSSEL:1105/ctxt%%DYN001-LOW").text = '*' + account_number + '*'
    session.findById("wnd[1]/tbar[0]/btn[0]").press()


def statements_exists(session, month : str) -> bool:
    """Checks if statements exists for the given account number."""
    column_name = 'NAME'
    grid = session.FindById(SAP_GRID_ID)
    for i in range(grid.rowCount):
        if grid.GetCellValue(i, column_name).startswith(month):
            grid.setCurrentCell(str(i), column_name)
            grid.selectedRows = str(i)
            return True
    return False


if __name__ == '__main__':
    main()

