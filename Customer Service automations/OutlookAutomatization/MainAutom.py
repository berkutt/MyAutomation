import re
from MyModules import utils
import os
from MyModules import SAP_Class


class OulAutom:

    def __init__(self):
        self.exception_track = False
        self.path = utils.global_variable().file_path()

    def get_deliveries(self, msg_subject):
        return re.findall("20[\d]{7}|83[\d]{6}|85[\d]{6}", msg_subject)

    def get_cost(self, msg_body):
        return re.search("\d+\.\d+", msg_body.replace(",", "")).group()

    def save_attachment(self, message, prefics):
        del_nr = self.get_deliveries(message.Subject)[0]
        new_name = prefics + del_nr + ".pdf"
        for att in message.Attachments:
            if att.FileName[-3:] in ['pdf', 'PDF']:
                # downdload file
                if att.FileName in os.listdir(self.path): os.remove(self.path + att.FileName)
                att.SaveAsFile(self.path + att.FileName)
                # rename file
                if new_name in os.listdir(self.path): os.remove(self.path + new_name)
                os.rename(self.path + att.FileName, self.path + new_name)
        return new_name

    def add_attachment(self, messeges, filename):
        # get all deliveries
        deliveries = self.get_deliveries(messeges.Subject)
        # for each deivery go to Invoice and attach the file
        SAP = SAP_Class.VladSAP()
        for deliv in deliveries:
            SAP.open_del_03(deliv)
            #try:
            SAP.del_to_inv(change=False, output=False)
            SAP.add_attachment(self.path, filename)
            #except:
            print("error with attaching Invoice. Maybe Invoice is missing. Check delivery ", deliv)
            self.exception_track = True
        SAP.close_window()
        os.remove(str(self.path) + '\\' + str(filename))
        return self.exception_track

    def add_cost(self, messeges):
        deliveries = self.get_deliveries(messeges.Subject)
        verif_list = list()
        freight_cost = int(float(self.get_cost(messeges.Body)) / len(deliveries))
        freight_cost = str(freight_cost) + ".00-GBP"
        if freight_cost:
            SAP = SAP_Class.VladSAP()
            for deliv in deliveries:
                print("Cost for ", deliv)
                verif_list.append(SAP.add_ship_cost(DelNr=deliv, carrier="", CostPerDel=freight_cost))
            SAP.close_window()
        if len(verif_list) == len(deliveries):
            return True
