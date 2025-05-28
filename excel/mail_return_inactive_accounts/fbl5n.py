import win32com.client

class FBL5N:
    _CUSTOMER_NUMBER_ID = 'wnd[0]/usr/ctxtDD_KUNNR-LOW'
    _COMPANY_CODE_LOW_ID = 'wnd[0]/usr/ctxtDD_BUKRS-LOW'
    _COMPANY_CODE_HIGH_ID = 'wnd[0]/usr/ctxtDD_BUKRS-HIGH'

    def __init__(self, session):
        self.session = session

    def open_transaction(self):
        self.session.StartTransaction('FBL5N')

    def set_customer_number(self, number : str):
        self.session.FindById(self.__class__._CUSTOMER_NUMBER_ID).Text = number

    def set_company_codes(self, low : str = '', high : str = ''):
        self.session.FindById(self.__class__._COMPANY_CODE_LOW_ID).Text = low
        self.session.FindById(self.__class__._COMPANY_CODE_HIGH_ID).Text = high


class SAPGUI:
    def __init__(self):
        self.SapGuiAuto = win32com.client.GetObject('SAPGUI')
        self.application = self.SapGuiAuto.GetScriptingEngine
        self.connection = self.application.Connections(0)
        self.session = self.application.Sessions(0)

    def get_session(self):
        return self.session

    def get_last_session(self):
        pass

    def create_session(self):
        pass