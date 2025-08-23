from retrieve_sap_statement.model.sap_zn30038 import ZN30038
from retrieve_sap_statement.model.sap import SAPGUI
from retrieve_sap_statement.controller.generate_mer import generate_medical_expense_report
from datetime import date
import os


def download_mer(patient_id : str, customer_name : str, lower_date : date=None, upper_date : date=None,  window=None):
    sap = SAPGUI()
    connection = sap.get_connection(window)
    if connection == None:
        return
    session = connection.Sessions(0)

    # Get last SAP session
    session = sap.get_last_session(window, session, connection)
    if session == None: 
        return 
    
    zn = ZN30038()
    zn.access_tcode(session)
    zn.search_by_patient_id(session, patient_id)
    user = os.getlogin()
    file_dir = os.path.join('C:\\', 'Users', user, 'Downloads')
    file_name = f'{patient_id}_{customer_name}.XLSX'
    zn.to_excel(session, file_dir, file_name)

    generate_medical_expense_report(os.path.join(file_dir, file_name), customer_name, lower_date, upper_date)