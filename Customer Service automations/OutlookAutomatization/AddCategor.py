from win32com.client import Dispatch
import re
from collections import Counter
import time
from datetime import datetime
from MyModules import datasets
from OutlookAutomatization import MainAutom

Oul = MainAutom.OulAutom()
from mlmailclassify import mainML
# disable warnings
import sys
import warnings

if not sys.warnoptions:
    warnings.simplefilter("ignore")


## imports for ML to EXE
# import sklearn.utils._cython_blas
# import sklearn.neighbors.typedefs
# import sklearn.neighbors.quad_tree
# import sklearn.tree
# import sklearn.tree._utils


class MailCategorize:

    def __init__(self):

        outlook = Dispatch("Outlook.Application").GetNamespace("MAPI")

        folder = outlook.Folders.Item("")

        self.inbox = folder.Folders.Item("Inbox")
        self.outbox = folder.Folders.Item("Sent Items")
        self.oldmails = self.loadoldmails()
        self.sentitems = self.loadsentitems()
        dataset = datasets.Dataset_20200521()
        df = dataset.read_data('')
        self.team = df
        print(df.head(15))
        # self.team = self.updateTeam(df)
        self.exludedMails = ['']

    # update people who is working today. If nor add their backup
    def updateTeam(self, df):

        for index in range(len(df)):
            if df.StartOff[index] <= datetime.now() <= df.EndOff[index]:
                df.at[index, 'Category'] = df.Backup[index]
        return df

    def getrecentperson(self, mailbody):
        parsedmail = mailbody.split("shipments export")
        for mail in parsedmail:
            for index in range(len(self.team)):
                try:
                    if 'the exporter of the products covered by ' in mail: continue
                    if self.team['Text in mail'][index].lower() in mail:
                        return self.team.Category[index]
                except:
                    continue

    # load old messages into memory
    def loadoldmails(self):
        oldmails = []
        try:
            inbmsg = self.inbox.Folders.Item("CLOSED").Items
            inbmsg.Sort("[ReceivedTime]", True)
            for message in inbmsg:
                try:
                    oldmails.append(message)
                    if len(oldmails) > 9000: break
                except:
                    continue
            print('old data were loaded ')
            return oldmails
        except:
            print("RESTART OUTLOOK old")
            input(" ")
            raise SystemExit(0)

    def loadsentitems(self):
        sentitems = []
        try:
            senitems = self.outbox.Items
            senitems.Sort("[ReceivedTime]", True)
            for message in senitems:
                try:
                    sentitems.append(message)
                    if len(sentitems) > 500:
                        break
                except:
                    continue
            print('sent items were loaded')
            return sentitems
        except:
            print("RESTART OUTLOOK sent")
            SystemExit(0)

    def getcategor(self, categ_list):
        # get most frequent categor
        if categ_list:
            print("Previous meeseges were for: ", categ_list)
            catlist = filter(None, categ_list)
            c = Counter(catlist)
            return c.most_common(1)[0][0]

    def getrefnrs(self, reflist):
        # get all references from
        templist = []
        for nr in reflist:
            if not str(nr).startswith('4') and nr not in templist:
                templist.append(nr)
        return str(templist)

    def rhenusmrn(self, message):
        # MRN from Rhenus
        list_of_users = []
        if message.UnRead and message.SenderName == '' and \
                message.Categories == '' and not message.Subject.startswith("_mrn"):
            # fidn Rhenus ref number, starts from 94...
            rhenusref = re.findall(r'\d+', message.Subject)
            foundnr = [Oul.get_deliveries(oldmail.Subject) for oldmail in self.oldmails if
                       rhenusref[0] in oldmail.Subject]
            refnr = set(foundnr)
            deliveries = repr(refnr).replace("{", " ").replace("}", "").replace("'", "").replace(",", "")
            if len(refnr) > 0: message.Subject = '_mrn ' + str(message.Subject) + deliveries
            message.Save()
            time.sleep(2)

        # add attachemnt if there is relevant mail
        elif message.Subject.startswith("_mrn") and message.UnRead and \
                message.Categories in list_of_users:
            att_name = Oul.save_attachment(message, prefics="_mrn ")
            if not Oul.add_attachment(message, att_name):
                message.UnRead = False
                message.Save()

    # add categories
    def categorize(self, message):
        foundcat = list()
        delivery_nr = list()
        if message.UnRead and message.Categories == '' and message.SenderName not in self.exludedMails:
            # Oliwwia handles Shipping advises.
            if "kemira shipping advice" in str(message.Subject).lower():
                categ = ""
            else:
                categ = self.getrecentperson(message.Body.lower())
            # if category not in mailchain, then extract it based on ref number
            if not categ:
                # get ref number
                delivery_nr = Oul.get_deliveries(message.Subject)
                if not delivery_nr:
                    delivery_nr = Oul.get_deliveries(message.Body[:30])
                # find ref number in old mails
            if delivery_nr:
                [foundcat.append(oldmail.Categories) for oldmail in self.oldmails if delivery_nr[0] in oldmail.Subject]
                categ = self.getcategor(foundcat)
                # if still not category, check Outbound mails
                if not categ:
                    [foundcat.append(oldmail.Categories) for oldmail in self.sentitems if
                     delivery_nr[0] in oldmail.Subject]
                    categ = self.getcategor(foundcat)
                # if some category was found
            if categ:
                print("Subject: ", message.Subject)
                message.Categories = categ
                message.Save()
                time.sleep(2)
                print('***Category was assign to ', message.Categories, "\n")

    def ML_labels(self, message):
        if message.UnRead and message.Categories == "" and not str(message.Subject).startswith("_bl"):
            label = mainML.mail_predict(message)
            print(message.Subject)
            print("***ML think this is", label, '\n')
            if label == "Final_BL":
                att_name = Oul.save_attachment(message, prefics="_bl ")
                # in case Invoice wasn't issued
                if not Oul.add_attachment(message, att_name):
                    message.Subject = "_bl " + str(message.Subject)
                    message.Save()
            if label == "Draft_BL" and not message.IsMarkedAsTask:
                message.MarkAsTask(4)
                message.Save()

    def mail_handles(self, message):
        if message.UnRead and message.Categories == "" and \
                message.SenderName in ["", ""] and "£" in message.Body:
            if Oul.add_cost_BradUS(message):
                message.UnRead = False
                message.Save()

    def main(self):
        for message in self.inbox.Items:
            try:
                self.rhenusmrn(message)
                self.categorize(message)
                self.ML_labels(message)
                self.mail_handles(message)
            except:
                continue

if __name__ == '__main__':
    mailclasss = MailCategorize()
    mailclasss.loadoldmails()
    mailclasss.loadsentitems()
    for j in range(9):
        if j > 0:
            # update old mails
            mailclasss.loadoldmails()
            mailclasss.loadsentitems()
        for i in range(30):

            mailclasss.main()

            print("Coffee break")
            time.sleep(120)
