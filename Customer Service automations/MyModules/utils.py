import pywinauto
import time
from pywinauto.application import Application

import os


class global_variable:

    def printer(self):
        # default printer on user windows
        return ""

    def login(self):
        # user AD2 login
        return os.getlogin()

    def file_path(self):
        # OS should be in English !!!
        if os.path.isdir("C:\\Users\\" + self.login() + "\\Downloads\\"):
            return "C:\\Users\\" + self.login() + "\\Downloads\\"
        else:
            print("OS should be in English")


import codecs

    # pop-up similar to VBA msgbox


import ctypes


def Mbox(title, text, style):
    return ctypes.windll.user32.MessageBoxW(0, text, title, style)
    #  Styles:
    #  0 : OK
    #  1 : OK | Cancel
    #  2 : Abort | Retry | Ignore
    #  3 : Yes | No | Cancel
    #  4 : Yes | No
    #  5 : Retry | No
    #  6 : Cancel | Try Again | Continue


import win32com.client
from tkinter import *
from tkinter.scrolledtext import ScrolledText


class COO_automation:

    def InpoutBox(self, ShipToNr):
        # function to write down COnsingee Address and Name to the protected Excel file
        def setaddress(ShipToNr, name, address):
            UserLogin = os.getlogin()
            xlApp = win32com.client.Dispatch("Excel.Application")
            filename = 'C:\\Users\\' + UserLogin + '' \
                                                   'and ConsgineeAddress\\COOmain.xlsx '
            xlwb = xlApp.Workbooks.Open(filename, False, False, None, Password='')
            ws = xlwb.Worksheets('ListOfConsig')
            time.sleep(2)
            xlUp = -4162
            # define last record in existing table
            lastrow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row + 1
            ws.Cells(lastrow, 1).Value = ShipToNr
            ws.Cells(lastrow, 2).Value = name
            ws.Cells(lastrow, 3).Value = address
            xlwb.Close(True)

        # Tkinter user dialog
        mainwin = Tk()

        Label(mainwin, text="Consignee Name:").grid(row=0, column=0)
        ent = Entry(mainwin, width=90)
        ent.grid(row=0, column=1)
        Button(mainwin, text="OK", command=(lambda: setaddress(ShipToNr, ent.get(), st.get(1.0, END)))).grid(row=1,
                                                                                                             column=2,
                                                                                                             sticky="EW")

        Label(mainwin, text="Consignee Address:").grid(row=1, column=0)
        st = ScrolledText(mainwin, height=5)
        st.grid(row=1, column=1)
        Button(mainwin, text="Close", command=(lambda: mainwin.destroy())).grid(row=0, column=2, sticky="EW")

        mainwin.mainloop()

    # get delviery Nr from , created with VBA
    def getDelivery(self):
        MyTxtFile = codecs.open(r'', 'r', 'utf-16')
        mylist = MyTxtFile.readlines()
        return mylist[0]

    def get_plant_coo(self, plant):
        UserLogin = global_variable().login()
        xlApp = win32com.client.Dispatch("Excel.Application")
        filename = 'C:\\Users\\' + UserLogin + '' \
                                               'and ConsgineeAddress\\PlantCountry.xlsx '
        xlwb = xlApp.Workbooks.Open(filename)
        ws = xlwb.Worksheets('WEB_STD')
        time.sleep(2)
        xlUp = -4162
        lastrow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
        for i in range(1, lastrow):

            if ws.Cells(i, 1).Value == plant:
                print("Plant was found = ", plant, ". COO is: ", ws.Cells(i, 4).text)
                coo = ws.Cells(i, 4).text
                xlwb.Close(False)
                return coo
        print("Plant wasn't found")

    def get_credentials(self):
        xlwb = self.loadExcel()
        ws = xlwb.Worksheets('Accounts')
        xlUp = -4162
        lastrow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row + 1
        login, password = "", ""
        for i in range(1, lastrow):
            if ws.Cells(i, 1).text == global_variable().login():
                login = ws.Cells(i, 2).Value
                password = ws.Cells(i, 3).Value
        xlwb.Close(False)
        return login, password

    def GetSHipTOinfo(self, ShipToNr):
        xlwb = self.loadExcel()
        ws = xlwb.Worksheets('ListOfConsig')
        xlUp = -4162
        lastrow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row + 1
        ConsName, ConsAddress = "", ""
        for i in range(1, lastrow):
            if ws.Cells(i, 1).text == ShipToNr:
                ConsName = ws.Cells(i, 2).Value
                ConsAddress = ws.Cells(i, 3).Value
        xlwb.Close(False)
        return ConsName, ConsAddress

    def loadExcel(self):
        xlApp = win32com.client.Dispatch("Excel.Application")
        filename = 'C:\\Users\\' + global_variable().login() + ''
        xlwb = xlApp.Workbooks.Open(filename, False, True, None, Password='')
        return xlwb

    # get Invocie nr with path of the file form C:/Temp and enter it to the pop-up window
    def uploadInvoice(self):
        MyTxtFile = codecs.open(r'', 'r', 'utf-16')
        mylist = MyTxtFile.readlines()
        mytext = mylist[0]

        i = 0
        while i < 8:
            mywindows = pywinauto.findwindows.find_windows(title_re="File Upload")
            if len(mywindows) == 0:
                time.sleep(1)
                i += 1
            else:
                for handle in mywindows:
                    app = Application().connect(handle=handle)
                    navwin = app.window(handle=handle)
                    navwin.edit.type_keys(mytext + '.pdf', with_spaces=True)
                    navwin.Open.click()
                    i = 8
                    break


class Some_carrier_automation:

    def handle_popup_foxit(self, text):
        i = 0
        output = False
        while i < 15:
            if output: break
            mywindows = pywinauto.findwindows.find_windows(title_re=global_variable().printer())
            if len(mywindows) == 0:
                time.sleep(1)
                i += 1
            else:
                for handle in mywindows:
                    app = Application().connect(handle=handle)
                    navwin = app.window(handle=handle)
                    navwin.edit.type_keys(text, with_spaces=True)
                    navwin.Save.click()
                    output = True
                    break

    def handle_popup_Some_carrier(self, text):
        i = 0
        while i < 8:
            mywindows = pywinauto.findwindows.find_windows(title_re="Open")
            if len(mywindows) == 0:
                time.sleep(1)
                i += 1
            else:
                for handle in mywindows:
                    app = Application().connect(handle=handle)
                    navwin = app.window(handle=handle)
                    navwin.edit.type_keys(text, with_spaces=True)
                    navwin.Open.click()
                    break


import pandas as pd
import pickle


class write_read_data:

    # data i/o methods, filter rows

    def read_excel(self, path, sheet=None, print_flag=False):
        if print_flag:
            print("Loading excel", path)
        return pd.read_excel(path)

    def write_excel(self, path, data, print_flag=False):
        if print_flag:
            print("Write to excel", path)
        return data.to_excel(path, index=False)

    def read_pickle(self, path, print_flag=False):
        if print_flag:
            print("Loading pickle", path)
        with open(path, 'rb') as handle:
            return pickle.load(handle)

    def write_pickle(self, path, data, print_flag=False):
        if print_flag:
            print("Write to pickle", path)
        with open(path, 'wb') as handle:
            return pickle.dump(data, handle)
