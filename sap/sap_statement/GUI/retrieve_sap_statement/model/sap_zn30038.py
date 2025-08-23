class ZN30038:
    def access_tcode(self, session) -> None:
        """Accesses the T-Code FBL5N."""
        # logger = logging.getLogger(__name__)

        # logger.debug('Accessing T-Code FBL5N')
        session.findById("wnd[0]/tbar[0]/okcd").text = "ZN30038_TRANSRPT_01"
        session.findById("wnd[0]").sendVKey(0)

    def search_by_patient_id(self, session, patient_id : str) -> None:
        session.findById("wnd[0]/usr/txtS_PAT_ID-LOW").text = patient_id #"29949"
        session.findById("wnd[0]/tbar[1]/btn[8]").press()

    def to_excel(self, session, file_path : str, file_name : str) -> None:
        session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").contextMenu()
        session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectContextMenuItem("&XXL")
        session.findById("wnd[1]/tbar[0]/btn[0]").press()
        session.findById("wnd[1]/usr/ctxtDY_PATH").text = file_path
        session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = file_name
        session.findById("wnd[1]/tbar[0]/btn[0]").press() #Download Excel File
        session.ActiveWindow.Close()

