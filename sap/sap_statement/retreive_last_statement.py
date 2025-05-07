import win32com.client

SapGuiAuto  = win32com.client.GetObject("SAPGUI")
application = SapGuiAuto.GetScriptingEngine
connection = application.Children(0)
session    = connection.Children(0)

# Functions
# Double click row in a grid
def double_click(column_name, row_value):
    grid = session.FindById("wnd[0]/usr/cntlGRID1/shellcont/shell")
    for i in range(grid.rowCount):
        if session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").GetCellValue(i, column_name) == row_value:
            session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").setCurrentCell(str(i), column_name)
            session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectedRows = str(i)
            session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").doubleClickCurrentCell()
            break

def sap_access_screen():
	session.findById("wnd[0]/tbar[0]/btn[12]").press()
	session.findById("wnd[0]/tbar[0]/btn[12]").press()
	session.findById("wnd[0]/tbar[0]/btn[12]").press()

account_number = input('Enter account number: ')
month = '0' + input('Enter the statement Month: ')

# Check communication method
session.findById("wnd[0]/tbar[0]/okcd").text = "FBL5N"
session.findById("wnd[0]").sendVKey(0)
session.findById("wnd[0]/usr/ctxtDD_KUNNR-LOW").text = account_number
session.findById("wnd[0]/usr/ctxtDD_BUKRS-LOW").text = "L604"
session.findById("wnd[0]/usr/ctxtDD_BUKRS-HIGH").text = "L607"
session.findById("wnd[0]/usr/ctxtDD_KUNNR-LOW").caretPosition = 10
session.findById("wnd[0]/usr/radX_AISEL").select()
session.findById("wnd[0]/tbar[1]/btn[8]").press()

try:    
    session.findById("wnd[1]/tbar[0]/btn[0]").press()
except:
     pass
session.findById("wnd[0]/tbar[1]/btn[34]").press()
communication_method = session.findById("wnd[0]/usr/subSUBTAB:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB01/ssubSUBSC:SAPLATAB:0201/subAREA1:SAPMF02D:7111/subADDRESS:SAPLSZA1:0300/subCOUNTRY_SCREEN:SAPLSZA1:0301/cmbADDR1_DATA-DEFLT_COMM").text
sap_access_screen()


# Functions
# Double click row in a grid
def double_click(column_name, row_value):
    grid = session.FindById("wnd[0]/usr/cntlGRID1/shellcont/shell")
    for i in range(grid.rowCount):
        if session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").GetCellValue(i, column_name) == row_value:
            session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").setCurrentCell(str(i), column_name)
            session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectedRows = str(i)
            session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").doubleClickCurrentCell()
            break

# Open T-Code AL11
session.findById("wnd[0]/tbar[0]/okcd").text = "AL11"
session.findById("wnd[0]").sendVKey(0)

# Get interfaces directory
double_click('DIRNAME', '/interfaces')
double_click('NAME', 'FII30040')
double_click('NAME', 'archive')
# double_click('NAME', 'medi')
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").currentCellRow = 6
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectedRows = "6"
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").doubleClickCurrentCell()


if communication_method == 'E-Mail':
    # email comuniction method 
    double_click('NAME', 'others')
else:
    # post comuniction method 
    double_click('NAME', 'mail')

# Filter statements by account number
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").currentCellRow = -1
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectColumn("NAME")
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").contextMenu()
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectContextMenuItem("&FILTER")
session.findById("wnd[1]/usr/ssub%_SUBSCREEN_FREESEL:SAPLSSEL:1105/ctxt%%DYN001-LOW").text = '*' + account_number + '*'
session.findById("wnd[1]/usr/ssub%_SUBSCREEN_FREESEL:SAPLSSEL:1105/ctxt%%DYN001-LOW").caretPosition = 12
session.findById("wnd[1]/tbar[0]/btn[0]").press()

# Select monthly statement
column_name = 'NAME'
grid = session.FindById("wnd[0]/usr/cntlGRID1/shellcont/shell")
file_exists = False
for i in range(grid.rowCount):
    if session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").GetCellValue(i, column_name).startswith(month):
        session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").setCurrentCell(str(i), column_name)
        session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectedRows = str(i)
        file_exists = True
        break

if file_exists:
    session.findById("wnd[0]").sendVKey(14)
    source_file = session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").GetCellValue(0, 'VALUE')
    target_file = session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").GetCellValue(1, 'VALUE')

    # Go Back to SAP easy Access
    sap_access_screen()

    # Open T-Code CG3Y
    session.findById("wnd[0]/tbar[0]/okcd").text = "CG3Y"
    session.findById("wnd[0]").sendVKey(0)
    session.findById("wnd[1]/usr/txtRCGFILETR-FTAPPL").text = source_file + '/' + target_file
    session.findById("wnd[1]/usr/ctxtRCGFILETR-FTFRONT").text = target_file
    session.findById("wnd[1]/usr/chkRCGFILETR-IEFOW").selected = True
    session.findById("wnd[1]/tbar[0]/btn[13]").press()
    session.findById("wnd[1]/tbar[0]/btn[12]").press()

    print('Statement downloaded')
else:
    sap_access_screen()
    print('Statement not found')





