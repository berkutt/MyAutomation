import pandas as pd
import re
import DataSource


def cleanMailBody(rawfilw):
    '''
    Load data from excel file. Data without columns. Columents names will be added by script
    Since there is a mailchaing in each mail, so mail will be cut by DropWords
    Also message with chines latters and from Inrta will be dropped as well
    Everyhing will be saved to the excel then
    '''
    df = pd.read_excel(rawfilw)
    # ['Unnamed: 0', 'Unnamed: 1', 'Unnamed: 2', 'Unnamed: 3', 'Unnamed: 4','Unnamed: 5']
    # df = df.rename(
    #     columns={'Unnamed: 0': 'Subject', 'Unnamed: 1': 'Recieving time', 'Unnamed: 2': 'Sender Name',
    #              'Unnamed: 3': 'Body',
    #              'Unnamed: 4': 'To', 'Unnamed: 5': 'CC', 'Unnamed: 6': 'Categories','Unnamed: 6': 'ConversationID'})

    FinalLength = []
    FirstLength = []

    # this one was took from Sender List
    StopWord = ["Freight Operative Department", "From: ", "" ]
    ''' 
    ANALYZE THE TEXT
    Cut text and drop some unusfull messages
    delete empty messages and some unessasery onces
    '''
    lengthOfMessageDF = []
    index = 0
    for message in df['Body']:
        message = str(message)
        if len(message) <= 3 or message.find("INTTRA SI ACT, Rev 3.0") > 0 or \
                message.find("LATE CANCELLATION FEE: With effective from ") > 0 or \
                message.find("浴") > 0:
            df = df.drop(df.index[[index]])
            continue
        index += 1

    # Cut the message with some Key Words
    for message in df['Body']:
        message = str(message)

        ShortMass = re.sub(r'http\S+', '', message)  # delete wev links
        ShortMass = ShortMass.replace("\r\n", "")  # remove empty lines
        ShortMass = ShortMass.replace("\n\n", "")  # remove empty lines

        for dWord in StopWord:
            cutpoint = ShortMass.find(dWord)
            if cutpoint > 0: ShortMass = ShortMass[:cutpoint]

        try:
            FinalLength.append(len(ShortMass))
            FirstLength.append(len(message))
        except:
            print("most probably there is an empty value:", type(message))
        df['Body'] = df['Body'].replace(message, ShortMass)

    data = {'FirstLenght': FirstLength, 'SecondLength': FinalLength}
    lengthOfMasg = pd.DataFrame(data=data, index=range(0, len(df)))
    # print(lengthOfMasg.describe())
    # # print out top mail by length
    # print(lengthOfMasg.nlargest(10, 'SecondLength'))

    return df
    # df = df.to_excel(finalfile, index=False)


def cleanSubject(df):
    # df = pd.read_excel(finalfile)
    try:
        for mail in df.Subject:
            mail2 = str(mail).replace("RE: ", "")
            mail2 = str(mail2).replace("FW: ", "")
            df['Subject'] = df['Subject'].replace(mail, mail2)
    except:
        print("None type value")
    return df


def addCategories(df, team):
    # df = pd.read_excel(finalfile)
    MyCollegues2 = []
    MyCollegues = team.split(';')

    for Name in MyCollegues:
        MyCollegues2.append(Name.strip())

    MyMatch = bool
    index = 0
    NameList = []
    for message in df['Body']:
        message = str(message)
        # splitmessage = message.split()
        for Name2 in MyCollegues2:
            try:
                if message.lower().count(Name2.lower()) != 0:
                    MyMatch = True
                    NameList.append(Name2)
                    break
            except:
                break
        if not MyMatch:
            NameList.append("Unknown")
        MyMatch = False
        index += 1

    df['Categor'] = NameList
    return df
    # df = df.to_excel(finalfile, index=False)


def splitRecipient(df, finalfile):
    # df = pd.read_excel(finalfile)

    df2 = df.assign(To=df['To'].str.split(';')).explode('To')

    for recipient in df2.To:
        recipient2 = str(recipient).strip()
        df2['To'] = df2['To'].replace(recipient, recipient2)

    df2 = df2.to_excel(finalfile, sheet_name="Sent", index=False)


def addRefNr(Dataframe):
    DelNr = []
    for message in Dataframe['Subject']:
        message = str(message)
        DelNr.append(re.findall("([821][0-9]{7,8,9})",
                                message))  # find SAP ref. numbers, starting from 8,2,1 and 7-9 characters long
    ListOfDel = pd.DataFrame(data=DelNr, index=range(len(Dataframe['Subject'])))
    Dataframe['RefNr'] = ListOfDel.loc[:, 0].values

    return Dataframe


# DataSource.path Desktop\Projects\ML project\
def getrowdata(name, path=DataSource.path):
    return path + name


'''
Load raw file in Excel formate. 
cleanMailBody - cut mail body
addCategories - based on 'team' search team member in the mail body. If found - assigned name in new column
splitRecipient - Find all recipiens and add new line with duplicated values but with additional recipient
'''


def main():
    # load file with row data
    rawfile = getrowdata("Gather Mails.xlsm")

    # specify output file and it's path
    finalfile = DataSource.path + ''

    # load team members
    team = DataSource.iwexportteam

    # ****cleaning data
    DataFrame = cleanMailBody(rawfile)
    df = DataFrame.to_excel(rawfile, index=False)# in case of Inbox only


    # print("Body mail was claened")
    # DataFrame = addCategories(DataFrame, team)
    # print("Categories are added")
    # #DataFrame = cleanSubject(DataFrame)
    # splitRecipient(DataFrame, finalfile)
    # print("Recipients were splited. Scrip finished.")


if __name__ == '__main__':
    main()
