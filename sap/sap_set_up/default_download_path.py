import win32com.client
import pywintypes

SapGuiAuto  = win32com.client.GetObject("SAPGUI")
application = SapGuiAuto.GetScriptingEngine
connection = application.Children(0)
session    = connection.Children(0)

# Open a new SAP window
session.findById("wnd[0]").sendVKey(74)
print('Connecting with new SAP GUI window . . .', end='')
while(True):
    try:
        application = SapGuiAuto.GetScriptingEngine
        connection = application.Children(0)
        session2 = connection.Children(1)
        break
    except pywintypes.com_error:
        pass
print('Done')

# Access T-Code SU3 
session2.findById("wnd[0]/tbar[0]/okcd").text = "SU3"
session2.findById("wnd[0]").sendVKey(0)

# Access Parameter Tab under SU3
session2.findById("wnd[0]/usr/tabsTABSTRIP1/tabpPARAM").select()
param_grid = session2.findById("wnd[0]/usr/tabsTABSTRIP1/tabpPARAM/ssubMAINAREA:SAPLSUID_MAINTENANCE:1104/cntlG_PARAMETER_CONTAINER/shellcont/shell")


for i in range(param_grid.RowCount):
    if param_grid.GetCellValue(i, 'PARID') == 'GR8':
        print(f'Default DOWNLOAD path found {param_grid.GetCellValue(i, 'PARVA')}')
        overrite_default_path = input('Would you like to change the DOWNLOAD path: Y/N ') in ['Y', 'y', 'Yes', 'yes']
        if overrite_default_path:
            new_path = input('Enter default DOWNLOAD path: ')
            param_grid.ModifyCell(i, 'PARVA', new_path)
        else:
            print(f'Path NOT modified')
        break
        
    if param_grid.GetCellValue(i, 'PARID') == '':
        new_path = input('Enter default DOWNLOAD path: ')
        param_grid.ModifyCell(i, 'PARID', 'GR8')
        param_grid.ModifyCell(i, 'PARVA', new_path)
        print(f'DOWNLOAD Path: {param_grid.GetCellValue(i, 'PARVA')}')
        print(f'DOWNLOAD Path modified succesfully')
        break

#Save changes
session2.findById("wnd[0]/tbar[0]/btn[11]").press()

input('Press any key to exit . . . ')