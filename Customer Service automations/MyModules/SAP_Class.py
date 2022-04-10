import time
from MyModules import utils
import re
import win32com.client


class VladSAP:
    # 1) Connect to SAP
    # 2) Open a new SAP window and connect to it.
    def __init__(self):
        self.SapGui = win32com.client.GetObject("SAPGUI").GetScriptingEngine
        self.sessionID = self.SapConnect()
        self.session = self.SapGui.FindById("ses[" + str(self.sessionID) + "]")

    def SapConnect(self):
        listOfOpenSess = self.check_connection()
        if not listOfOpenSess:
            utils.Mbox("Error", "Wasn't able connect to the SAP", 0)
            raise SystemExit(0)
        session = self.SapGui.FindById("ses[" + str(listOfOpenSess[-1]) + "]")
        # open new SAP window
        session.findById("wnd[0]").sendVKey(74)
        time.sleep(2)

        # Identify ID of new openned window.
        # Get list of new opened sessions
        listOfOpenSess2 = []
        SessionNr = 6
        for sessionID in range(SessionNr)[::-1]:
            try:
                self.SapGui.FindById("ses[" + str(sessionID) + "]")
                listOfOpenSess2.append(sessionID)
            except:
                continue
        # identify what session was open
        # by comparing 2 lists(before connection and after)
        for sessionID in listOfOpenSess2:
            if sessionID not in listOfOpenSess:
                print("Connected to SAP")
                return sessionID

    def check_connection(self):
        SessionNr = 6
        listOfOpenSess = []

        # Check what sessions are open and get list of openned sessions
        for i in range(SessionNr)[::-1]:
            try:
                self.SapGui.FindById("ses[" + str(i) + "]")
                listOfOpenSess.append(i)
            except:
                continue
        listOfOpenSess.sort()
        if len(listOfOpenSess) == 0:
            return False
        else:
            return listOfOpenSess

    # if Del is None then open tranaction and enter to delviery that first appear there.
    # open transaction
    def open_del_03(self, DelNr=None):
        self.session.findById("wnd[0]/tbar[0]/okcd").text = "/nvl03n"
        self.session.findById("wnd[0]").sendVKey(0)
        if DelNr: self.session.findById("wnd[0]/usr/ctxtLIKP-VBELN").text = DelNr
        self.session.findById("wnd[0]").sendVKey(0)
        # check if deliery exist. In case some error -
        # duplicate error massage from SAP to dialogbox
        info_msg = self.session.findById("wnd[0]/sbar").Text
        if str(info_msg).startswith("Delivery "):
            utils.Mbox("Error", self.session.findById("wnd[0]/sbar").text, 0)
            raise SystemExit(0)

    # open in change mode
    def open_del_02(self, DelNr=None):

        self.session.findById("wnd[0]/tbar[0]/okcd").Text = "/nvl02n"
        self.session.findById("wnd[0]").sendVKey(0)
        if DelNr: self.session.findById("wnd[0]/usr/ctxtLIKP-VBELN").Text = DelNr
        self.session.findById("wnd[0]").sendVKey(0)
        self.session.findById("wnd[0]").sendVKey(0)
        # Check if someone sitting in the order
        info_msg = self.session.findById("wnd[0]/sbar").Text
        # if delivery is blocked
        if str(info_msg).startswith("This delivery"):
            utils.Mbox("Error", self.session.findById("wnd[0]/sbar").text, 0)
            raise SystemExit(0)
        # if delivery doesn't exist
        if str(info_msg).startswith("Delivery "):
            utils.Mbox("Error", self.session.findById("wnd[0]/sbar").text, 0)
            raise SystemExit(0)

    def add_attachment(self, path, filename):

        self.session.findById("wnd[0]/titl/shellcont[1]/shell").pressButton("%GOS_TOOLBOX")
        self.session.findById("wnd[0]/shellcont/shell").pressContextButton("CREATE_ATTA")
        self.session.findById("wnd[0]/shellcont/shell").selectContextMenuItem("PCATTA_CREA")
        self.session.findById("wnd[1]/usr/ctxtDY_PATH").text = path
        self.session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = filename
        self.session.findById("wnd[1]/tbar[0]/btn[0]").press()

    # get Plant name and return country of origin,
    # where it's placed, from Excel file
    def getPlant(self):
        self.session.findById(r"wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\02").select()
        return self.session.findById(
            r"wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\02/ssubSUBSCREEN_BODY:SAPMV50A:1104/tblSAPMV50ATC_LIPS_PICK"
            "/ctxtLIPS-WERKS[2,0]").text

    def getShipTo(self):
        return self.session.findById("wnd[0]/usr/subSUBSCREEN_HEADER:SAPMV50A:1502/ctxtKUWEV-KUNNR").text

    # GET GOODS INFO
    # net value
    # gross value
    # trade name
    def getGInetValue(self, item):

        ValueInSAP = self.session.findById(
            r"wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\03/ssubSUBSCREEN_BODY:SAPMV50A:1106"
            "/tblSAPMV50ATC_LIPS_LOAD/txtLIPSD-G_LFIMG[2," + str(item) + "]").text

        Decription = self.session.findById(
            r"wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\03/ssubSUBSCREEN_BODY:SAPMV50A:1106"
            "/tblSAPMV50ATC_LIPS_LOAD/txtLIPS-ARKTX[12," + str(item) + "]").text

        Bag = re.findall(r"(\d+\s*)KG", Decription)
        # check if something was catched. In case of bulk there sure be nothing
        if len(Bag) == 0: Bag = ["KG"]
        Bag = Bag[0].strip()

        Unit = self.session.findById(r"wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\03/ssubSUBSCREEN_BODY:SAPMV50A"
                                     r":1106/tblSAPMV50ATC_LIPS_LOAD/ctxtLIPS-VRKME[3," + str(item) + "]").text
        print("Line", item, "Goods are:", ValueInSAP, Decription)

        if Unit == "KG":
            return "Net weight\n" + str(ValueInSAP) + "KG"
        # calculate Net value in case there is unit masuremnts are not KG
        # and add dot in value for better reading large values on documents
        # like from 20000 to 20.000
        elif Unit == "TO":
            ValueInSAP = ValueInSAP + ".000"
            return "Net weight\n" + str(ValueInSAP) + "KG"
        else:
            ValueInSAP = int(ValueInSAP) * int(Bag)
            ValueInSAP = str(ValueInSAP)
            if len(str(ValueInSAP)) == 4:
                ValueInSAP = str(ValueInSAP[:1]) + "." + str(ValueInSAP)[-3:]
            elif len(str(ValueInSAP)) == 5:

                ValueInSAP = str(ValueInSAP[:2]) + "." + str(ValueInSAP)[-3:]

        return "Net weight\n" + str(ValueInSAP) + "KG"

    def getGIgrosstValue(self, item):
        print("**Calculating Gross Value***")
        itemLine = 1
        grossValue2 = 0
        # if values is 0 then batch was assigned via batch split
        if self.session.findById(r"wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\03/ssubSUBSCREEN_BODY:SAPMV50A:1106"
                                 "/tblSAPMV50ATC_LIPS_LOAD/txtLIPS-BRGEW[4," + str(item) + "]").text == "0":
            # open batch split
            self.session.findById(r"wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\03/ssubSUBSCREEN_BODY:SAPMV50A:1106"
                                  "/tblSAPMV50ATC_LIPS_LOAD/btnRV50A-CHMULT[9," + str(item) + "]").press()
            # loop through every batch for item
            while self.session.findById(
                    r"wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\03/ssubSUBSCREEN_BODY:SAPMV50A:1106"
                    "/tblSAPMV50ATC_LIPS_LOAD/txtLIPS-BRGEW[4," + str(
                        itemLine) + "]").text != "___.___.___.___,___":

                grossValue = (self.session.findById(
                    r"wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\03/ssubSUBSCREEN_BODY:SAPMV50A:1106"
                    "/tblSAPMV50ATC_LIPS_LOAD/txtLIPS-BRGEW[4," + str(itemLine) + "]").text).replace(".", "")
                # replace don to comma in order or calculate total gross value
                grossValue = str(grossValue).replace(",", ".")
                # try to make value from string. When it's like 2200 KG and 2200,340 KG
                try:
                    grossValue2 = grossValue2 + int(grossValue)
                except:
                    grossValue2 = grossValue2 + float(grossValue)
                itemLine += 1
            # close batch split
            self.session.findById(r"wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\03/ssubSUBSCREEN_BODY:SAPMV50A:1106"
                                  "/tblSAPMV50ATC_LIPS_LOAD/btnRV50A-CHMULT[9,0]").press()
        # if batch was assigned manually
        else:
            grossValue = self.session.findById(
                r"wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\03/ssubSUBSCREEN_BODY:SAPMV50A:1106"
                "/tblSAPMV50ATC_LIPS_LOAD/txtLIPS-BRGEW[4," + str(item) + "]").text
            grossValue = str(grossValue).replace(".", "")
            grossValue = str(grossValue).replace(",", ".")
            try:
                grossValue2 = grossValue2 + int(grossValue)
            except:
                grossValue2 = grossValue2 + float(grossValue)

        # add zero's after comma
        grossValue2 = str(grossValue2).replace(".", ",")
        aftercoma = len(grossValue2) - grossValue2.find(",") - 1
        grossValue2 = grossValue2 + (3 - aftercoma) * "0"

        # assigne "." to the value of gross weight. so we would have
        # instead 20000123 --> 20.000,123
        stringlength = len(grossValue2)
        if stringlength == 4 or stringlength == 8:
            grossValue2 = grossValue2[:1] + "." + grossValue2[1:]
        elif stringlength == 5 or stringlength == 9:
            grossValue2 = grossValue2[:2] + "." + grossValue2[2:]
        return "Gross weight\n" + grossValue2 + "KG\n\n"

    def getTradeName(self, item):
        self.session.findById(
            r"wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\03/ssubSUBSCREEN_BODY:SAPMV50A:1106/tblSAPMV50ATC_LIPS_LOAD").getAbsoluteRow(
            str(item)).selected = -1
        self.session.findById("wnd[0]/mbar/menu[3]/menu[5]").select()
        TradeName = self.session.findById(
            "wnd[0]/usr/subCHARACTERISTICS:SAPLCEI0:1400/tblSAPLCEI0CHARACTER_VALUES/ctxtRCTMS-MWERT[1,1]").text
        print("Trade name : ", TradeName)
        self.session.findById("wnd[0]/tbar[0]/btn[3]").press()

        return TradeName

    # Get nr of Pallets and nr of packegies
    def getPackPall(self, item):
        self.session.findById(r"wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\03/ssubSUBSCREEN_BODY:SAPMV50A:1106"
                              "/tblSAPMV50ATC_LIPS_LOAD/txtLIPSD-G_LFIMG[2," + str(item) + "]").setFocus()
        self.session.findById(r"wnd[0]").sendVKey(2)
        self.session.findById(r"wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\12").select()  # select Addtional tab

        nrpackeges = self.session.findById(r"wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\12/ssubSUBSCREEN_BODY:SAPMV50A"
                                           r":3126/ssubCUSTOMER_SCREEN:SAPLZSD_LIPS_ITEM:9000/txtGV_ZZPACKAGING"
                                           r"").text
        nrpallets = self.session.findById(r"wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\12/ssubSUBSCREEN_BODY:SAPMV50A"
                                          r":3126/ssubCUSTOMER_SCREEN:SAPLZSD_LIPS_ITEM:9000/txtGV_ZZPALLET"
                                          r"").text
        self.session.findById("wnd[0]/tbar[0]/btn[3]").press()
        return "Nr of packages ", nrpackeges, " / Nr of pallets ", nrpallets

    def getGoodsInfo(self):
        list_of_items = []
        self.session.findById(r"wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\03").select()
        item = 0

        while self.session.findById(
                r"wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\03/ssubSUBSCREEN_BODY:SAPMV50A:1106"
                "/tblSAPMV50ATC_LIPS_LOAD/txtLIPSD-G_LFIMG[2," + str(
                    item) + "]").text != "_________________":
            list_of_items.append(self.getTradeName(item))
            list_of_items.append(self.getPackPall(item))
            list_of_items.append(self.getGInetValue(item))
            list_of_items.append(self.getGIgrosstValue(item))
            item += 1
        return list_of_items

    def getconCountry(self):
        self.session.findById("wnd[0]/usr/subSUBSCREEN_HEADER:SAPMV50A:1502/btnBT_WADR_T").press()
        conCOuntry = self.session.findById(
            "wnd[1]/usr/subGCS_ADDRESS:SAPLSZA1:0300/subCOUNTRY_SCREEN:SAPLSZA1:0301/txtT005T-LANDX").text
        self.session.findById("wnd[1]/tbar[0]/btn[0]").press()
        return conCOuntry

    # get Sales Order nr of Purchase Order nr from Documents Flow
    def getSOnr(self):
        self.session.findById(
            r"wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\03/ssubSUBSCREEN_BODY:SAPMV50A:1106/tblSAPMV50ATC_LIPS_LOAD"
            r"/ctxtLIPS-MATNR[1,0]").setFocus()
        self.session.findById("wnd[0]").sendVKey(2)
        self.session.findById(r"wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\10").select()
        # capture SO nr. Seems like there is different path for text of docs
        # SAPMV50A:3304 - SAPMV50A:3302
        try:
            sonr = self.session.findById(
                r"wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\10/ssubSUBSCREEN_BODY:SAPMV50A:3114/subSUBSCREEN_MAIN"
                r":SAPMV50A:3302/ctxtLIPS-VGBEL").text
        except:
            sonr = self.session.findById(
                r"wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\10/ssubSUBSCREEN_BODY:SAPMV50A:3114/subSUBSCREEN_MAIN"
                r":SAPMV50A:3304/ctxtLIPS-VGBEL").text
        self.session.findById("wnd[0]/tbar[0]/btn[3]").press()
        return sonr

    def select_item_hierar(self, item):
        # pass what item should be selected from Document Flow
        # Doc Flow should be opened at this point
        # return True when item found successfully
        DocFlow = self.session.findById("wnd[0]/usr/shell/shellcont[1]/shell[1]").GetAllNodeKeys()
        # iterate through each line in document flow to find Invoice
        for doc_line in range(len(DocFlow)):
            doc_line_str = str(self.session.findById("wnd[0]/usr/shell/shellcont[1]/shell[1]").GetNodeTextByKey(
                DocFlow(str(doc_line))))
            if doc_line_str.startswith(item):
                # if there are more then 9 lines - there is an additional space in "          "
                if doc_line < 9:
                    self.session.findById("wnd[0]/usr/shell/shellcont[1]/shell[1]").selectItem \
                        ("          " + str(doc_line + 1), "&Hierarchy")
                else:
                    self.session.findById("wnd[0]/usr/shell/shellcont[1]/shell[1]").selectItem \
                        ("         " + str(doc_line + 1), "&Hierarchy")
                # enter to document
                self.session.findById("wnd[0]/tbar[1]/btn[8]").press()
                return True

    def del_to_inv(self, filename="", change=True, output=True):
        # when in the delivery - open Document flow and select Invoice
        inv_types = {"Invoice 90": " Inv", "IC Stock": " IC Inv", "Intercompany sales": " IC Inv",
                     "Plants Abroad": " IC Inv"}
        # open document flow
        self.session.findById("wnd[0]/tbar[1]/btn[7]").press()

        for inv in inv_types.keys():
            if self.select_item_hierar(inv): break
        self.session.findById("wnd[0]/tbar[1]/btn[8]").press()
        # go to change mode
        if change:
            self.session.findById("wnd[0]").sendVKey(35)
            self.session.findById("wnd[0]").sendVKey(0)
        # open output window
        if output:
            print("Printing Invoice")
            self.session.findById("wnd[0]").sendVKey(20)
            self.print_output("ZINV")
            file_path = utils.global_variable().file_path() + filename
            utils.=Some_carrier_automation().handle_popup_foxit(text=file_path)

    def del_to_shpt(self, carrier, change_mode=True):
        self.session.findById(r"wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\02").select()

        if self.session.findById(
                r"wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\02/ssubSUBSCREEN_BODY:SAPMV50A:1104/txtVBSTT-KOSTA_BEZ").text \
                == "Fully picked": pickstatus = True
        self.session.findById("wnd[0]/tbar[1]/btn[7]").press()
        if self.select_item_hierar("Shipment"):
            # check if correct carrier assigned
            if self.session.findById(
                    "wnd[0]/usr/tabsHEADER_TABSTRIP1/tabpTABS_OV_PR/ssubG_HEADER_SUBSCREEN1:SAPMV56A:1021/ctxtVTTK-TDLNR").text != carrier:
                utils.Mbox("Verification check", "Wrong carrier is assigned.", 0)
                return False
            # Open change mode
            if change_mode:
                self.session.findById("wnd[0]/mbar/menu[0]/menu[1]").select()
                self.session.findById("wnd[0]").sendVKey(0)
                print("In the shipment now")
        else:
            self.add_carrier()
            self.del_to_shpt(carrier)

    def add_ship_cost(self, DelNr, carrier, CostPerDel):
        self.open_del_03(DelNr)
        self.del_to_shpt(carrier)

        # Check if Planning and PGI is completed
        PlanningStatus = False
        ShipmCOmpl = False

        if self.session.findById("wnd[0]/usr/tabsHEADER_TABSTRIP2/tabpTABS_OV_DE/ssubG_HEADER_SUBSCREEN2:SAPMV56A"
                                 ":1025/ctxtVTTK-DTABF").Text != "":
            ShipmCOmpl = True
            self.session.findById("wnd[0]/usr/tabsHEADER_TABSTRIP2/tabpTABS_OV_DE/ssubG_HEADER_SUBSCREEN2:SAPMV56A"
                                  ":1025/btn*RV56A-ICON_STABF").press()

        if self.session.findById("wnd[0]/usr/tabsHEADER_TABSTRIP2/tabpTABS_OV_DE/ssubG_HEADER_SUBSCREEN2:SAPMV56A"
                                 ":1025/ctxtVTTK-DTDIS").Text != "":
            PlanningStatus = True
            self.session.findById("wnd[0]/usr/tabsHEADER_TABSTRIP2/tabpTABS_OV_DE/ssubG_HEADER_SUBSCREEN2:SAPMV56A"
                                  ":1025/btn*RV56A-ICON_STDIS").press()
        self.session.findById("wnd[0]/usr/tabsHEADER_TABSTRIP2/tabpTABS_OV_LG").select()
        self.session.findById("wnd[0]/usr/tabsHEADER_TABSTRIP2/tabpTABS_OV_LG/ssubG_HEADER_SUBSCREEN2:SAPMV56A"
                              ":1032/subLEG_OVER:SAPLSTAG:1300/tblSAPLSTAGVTTS_ALL_CTRL/txtVTTS_ALL-BEZKOMPRZ[2,"
                              "0]").setFocus()
        self.session.findById("wnd[0]").sendVKey(2)
        try:
            self.session.findById("wnd[1]/usr/subDEPARTURE:SAPLSTAG:0510/txtVTTS-BELAD").text = CostPerDel
            self.session.findById("wnd[1]").sendVKey(0)
        except:
            #utils.Mbox("Verification check", "Delivery " + DelNr + " already PGI'ed.", 0)
            return

        self.session.findById("wnd[1]/tbar[0]/btn[3]").press()
        self.session.findById("wnd[0]/usr/tabsHEADER_TABSTRIP2/tabpTABS_OV_TE").select()
        self.session.findById("wnd[0]/usr/tabsHEADER_TABSTRIP2/tabpTABS_OV_TE/ssubG_HEADER_SUBSCREEN2:SAPMV56A"
                              ":1035/cmbVTTK-TNDRST").setFocus()
        self.session.findById("wnd[0]/usr/tabsHEADER_TABSTRIP2/tabpTABS_OV_TE/ssubG_HEADER_SUBSCREEN2:SAPMV56A"
                              ":1035/cmbVTTK-TNDRST").key = "NW"
        self.session.findById("wnd[0]/usr/tabsHEADER_TABSTRIP2/tabpTABS_OV_DE").select()

        if PlanningStatus:
            self.session.findById("wnd[0]/usr/tabsHEADER_TABSTRIP2/tabpTABS_OV_DE/ssubG_HEADER_SUBSCREEN2:SAPMV56A"
                                  ":1025/btn*RV56A-ICON_STDIS").press()
            self.session.findById("wnd[1]/tbar[0]/btn[0]").press()
            self.session.findById("wnd[0]/usr/tabsHEADER_TABSTRIP2/tabpTABS_OV_DE/ssubG_HEADER_SUBSCREEN2:SAPMV56A"
                                  ":1025/ctxtVTTK-DTDIS").text = \
                self.session.findById("wnd[0]/usr/tabsHEADER_TABSTRIP2/tabpTABS_OV_DE/ssubG_HEADER_SUBSCREEN2"
                                      ":SAPMV56A:1025/ctxtVTTK-DPREG").text

        if ShipmCOmpl:
            self.session.findById("wnd[0]/usr/tabsHEADER_TABSTRIP2/tabpTABS_OV_DE/ssubG_HEADER_SUBSCREEN2:SAPMV56A"
                                  ":1025/btn*RV56A-ICON_STABF").press()
            self.session.findById("wnd[0]").sendVKey(0)
            self.session.findById("wnd[0]/usr/tabsHEADER_TABSTRIP2/tabpTABS_OV_DE/ssubG_HEADER_SUBSCREEN2:SAPMV56A"
                                  ":1025/ctxtVTTK-DTABF").text = \
                self.session.findById("wnd[0]/usr/tabsHEADER_TABSTRIP2/tabpTABS_OV_DE/ssubG_HEADER_SUBSCREEN2"
                                      ":SAPMV56A:1025/ctxtVTTK-DPABF").text

        self.session.findById("wnd[0]/tbar[0]/btn[11]").press()
        print("Shipment cost added. Delivery ", DelNr, " Cost ", CostPerDel)
        return True

    # add carrier in Partner tab as SP - Forwarding Agent
    def add_carrier(self, carrier=""):
        self.open_del_02()
        self.session.findById("wnd[0]/tbar[1]/btn[8]").press()
        self.session.findById(r"wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\08").select()
        self.session.findById(
            r"wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\08/ssubSUBSCREEN_BODY:SAPMV50A:2114"
            "/subSUBSCREEN_PARTNER_OVERVIEW:SAPLV09C:1000/tblSAPLV09CGV_TC_PARTNER_OVERVIEW/cmbGVS_TC_DATA-REC-PARVW["
            "0,13]").Key = "SP "
        self.session.findById(
            r"wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\08/ssubSUBSCREEN_BODY:SAPMV50A:2114"
            "/subSUBSCREEN_PARTNER_OVERVIEW:SAPLV09C:1000/tblSAPLV09CGV_TC_PARTNER_OVERVIEW/ctxtGVS_TC_DATA-REC"
            "-PARTNER[1,13]").Text = carrier
        self.session.findById("wnd[0]").sendVKey(0)
        self.session.findById("wnd[0]").sendVKey(11)
        print("Carrier " + carrier + " was added. Waiting 20 sec to continiue")
        time.sleep(20)

    # generate raport zl06o
    def get_fake_del(self):
        print("Entering zl06o")
        self.session.findById("wnd[0]/tbar[0]/okcd").text = "/nzl06o"
        self.session.findById("wnd[0]").sendVKey(0)
        self.session.findById("wnd[0]/usr/ctxtIT_VBELN-LOW").text = ""
        self.session.findById("wnd[0]/usr/ctxtIT_VBELN-LOW").setFocus()
        self.session.findById("wnd[0]/usr/ctxtIT_VBELN-LOW").caretPosition = 8
        self.session.findById("wnd[0]/tbar[1]/btn[8]").press()

    def print_output(self, output_type):
        output = False
        while not output:
            for i in range(40):
                # if to long list, it should be scrolled. Otherwise bottom would not be visible
                if i > 16: self.session.findById("wnd[0]/usr/tblSAPDV70ATC_NAST3").verticalScrollbar.Position = 16
                if self.session.findById("wnd[0]/usr/tblSAPDV70ATC_NAST3/ctxtDNAST-KSCHL[1," + str(i) + "]").text == "":
                    self.session.findById(
                        "wnd[0]/usr/tblSAPDV70ATC_NAST3/ctxtDNAST-KSCHL[1," + str(i) + "]").text = output_type
                    self.session.findById("wnd[0]").sendVKey(0)
                    break
                elif self.session.findById(
                        "wnd[0]/usr/tblSAPDV70ATC_NAST3/ctxtDNAST-KSCHL[1," + str(i) + "]").text == output_type and \
                        self.session.findById(
                            "wnd[0]/usr/tblSAPDV70ATC_NAST3/cmbNAST-NACHA[3," + str(i) + "]").text == "1 Print output":
                    if self.session.findById(
                            "wnd[0]/usr/tblSAPDV70ATC_NAST3/txtNAST-DATVR[8," + str(i) + "]").text == "":
                        self.session.findById(
                            "wnd[0]/usr/tblSAPDV70ATC_NAST3/ctxtNAST-SPRAS[6," + str(i) + "]").text = "EN"
                        self.session.findById("wnd[0]").sendVKey(0)
                        self.session.findById("wnd[0]").sendVKey(0)
                        self.session.findById("wnd[0]/usr/tblSAPDV70ATC_NAST3").getAbsoluteRow(i).selected = -1
                        self.session.findById("wnd[0]").sendVKey(2)
                        self.session.findById("wnd[0]/usr/txtNAST-ANZAL").text = "1"
                        self.session.findById("wnd[0]/usr/chkNAST-DIMME").selected = True
                        self.session.findById("wnd[0]/usr/ctxtNAST-LDEST").text = "ALOCL"
                        self.session.findById("wnd[0]").sendVKey(3)
                        self.session.findById("wnd[0]/usr/tblSAPDV70ATC_NAST3").getAbsoluteRow(i).selected = -1
                        self.session.findById("wnd[0]").sendVKey(5)
                        self.session.findById("wnd[0]/usr/cmbNAST-VSZTP").key = "4"
                        self.session.findById("wnd[0]").sendVKey(3)
                        self.session.findById("wnd[0]").sendVKey(11)
                        # handle some unnecessary pop-ups. Looks like there is better way to do this.
                        try:
                            self.session.findById("wnd[0]").sendVKey(12)
                        except:
                            pass
                        try:
                            self.session.findById("wnd[1]/tbar[0]/btn[14]").press()
                        except:
                            pass
                        output = True
                        break
                    else:
                        self.session.findById("wnd[0]/usr/tblSAPDV70ATC_NAST3").getAbsoluteRow(i).selected = -1
                        self.session.findById("wnd[0]/tbar[1]/btn[6]").press()
                        break

    def download_report_ZL06O(self, variant, layout, filename, path=utils.global_variable().file_path()):
        print("Opening ZLO06O")
        self.session.findById("wnd[0]").maximize()
        self.session.findById("wnd[0]/tbar[0]/okcd").text = "/nzl06o"
        self.session.findById("wnd[0]").sendVKey(0)

        # select variant
        self.session.findById("wnd[0]").sendVKey(17)
        self.session.findById("wnd[1]/usr/txtV-LOW").text = variant
        self.session.findById("wnd[1]/usr/ctxtENVIR-LOW").text = ""
        self.session.findById("wnd[1]/usr/txtENAME-LOW").text = ""
        self.session.findById("wnd[1]/usr/txtAENAME-LOW").text = ""
        self.session.findById("wnd[1]/usr/txtMLANGU-LOW").text = ""
        self.session.findById("wnd[1]/tbar[0]/btn[8]").press()

        # self.session.findById("wnd[0]/usr/ctxtIT_LDDAT-LOW").text = "04.09.2020"
        # self.session.findById("wnd[0]/usr/ctxtIT_LDDAT-HIGH").text = "05.09.2020"
        self.session.findById("wnd[0]/usr/ctxtIT_LDDAT-HIGH").setFocus()
        self.session.findById("wnd[0]/usr/ctxtIT_LDDAT-HIGH").caretPosition = 2
        self.session.findById("wnd[0]/tbar[1]/btn[8]").press()
        print(self.session.findById("wnd[0]/sbar").text)
        if self.session.findById("wnd[0]/sbar").text == "No deliveries selected":
            utils.Mbox("Error", "There is no data in the report", 0)
            raise SystemExit(0)

        self.session.findById("wnd[0]/tbar[1]/btn[33]").press()
        self.session.findById("wnd[1]/tbar[0]/btn[71]").press()
        self.session.findById("wnd[2]/usr/chkSCAN_STRING-START").selected = 0
        self.session.findById("wnd[2]/usr/txtRSYSF-STRING").text = layout
        self.session.findById("wnd[2]/usr/chkSCAN_STRING-START").setFocus()
        self.session.findById("wnd[2]/tbar[0]/btn[0]").press()
        self.session.findById("wnd[3]/usr/lbl[1,2]").setFocus()
        self.session.findById("wnd[3]/usr/lbl[1,2]").caretPosition = 9
        self.session.findById("wnd[3]").sendVKey(2)
        self.session.findById("wnd[1]/usr/lbl[1,3]").caretPosition = 7
        self.session.findById("wnd[1]").sendVKey(2)

        if self.session.findById("wnd[0]/usr/lbl[2,3]").text == "List does not contain any data":
            utils.Mbox("Error", "There is no data in the report", 0)
            raise SystemExit(0)

        print("Downloading the report")
        self.session.findById("wnd[0]").sendVKey(43)
        self.session.findById("wnd[1]/usr/ctxtDY_PATH").text = path
        self.session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = filename
        self.session.findById("wnd[1]/tbar[0]/btn[11]").press()

    def close_window(self):
        # close window
        self.session.findById("wnd[0]").close()
