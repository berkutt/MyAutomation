import pandas as pd
import re

'''
Load data from excel file. Data without columns. Columents names will be added by script
Since there is a mailchaing in each mail, so mail will be cut by DropWords
Also massage with chines latters and from Inrta will be dropped as well
Everyhing will be saved to the excel then
'''

df = pd.read_excel(r'', 'Sent')
# ['Unnamed: 0', 'Unnamed: 1', 'Unnamed: 2', 'Unnamed: 3', 'Unnamed: 4','Unnamed: 5']
df = df.rename(
    columns={'Unnamed: 0': 'Subject', 'Unnamed: 1': 'Recieving time', 'Unnamed: 2': 'Sender Name', 'Unnamed: 3': 'Body',
             'Unnamed: 4': 'To', 'Unnamed: 5': 'CC', 'Unnamed: 6': 'Categories'})

FinalLength = []
FirstLength = []

#this one was took from Sender List
StopWord = ["From: ", ""]
''' 
ANALYZE THE TEXT
Cut text and drop some unusfull massages
delete empty massages and some unessasery onces
'''
lengthOfMessageDF = []
index = 0
for massage in df['Body']:
    massage = str(massage)
    if len(massage) <= 3:
        df = df.drop(df.index[[index]])
        continue
    if massage.find("LATE CANCELLATION FEE: With effective from ") > 0:
        df = df.drop(df.index[[index]])
        continue
    if massage.find("浴") > 0:
        df = df.drop(df.index[[index]])
        continue
    index += 1

# Cut the massage with some Key Words
for massage in df['Body']:
    massage = str(massage)

    ShortMass = re.sub(r'http\S+', '', massage)  # delete wev links
    ShortMass = ShortMass.replace("\r\n", "")  # remove empty lines
    ShortMass = ShortMass.replace("\n\n", "")  # remove empty lines

    for dWord in StopWord:
        cutpoint = ShortMass.find(dWord)
        if cutpoint > 0: ShortMass = ShortMass[:cutpoint]

    try:
        FinalLength.append(len(ShortMass))
        FirstLength.append(len(massage))
    except:
        print("most probably there is an empty value:", type(massage))
    df['Body'] = df['Body'].replace(massage, ShortMass)

print(df.describe())

# print(df.loc[200, :])
# raise SystemExit(0)

data = {'FirstLenght': FirstLength, 'SecondLength': FinalLength}
lengthOfMasg = pd.DataFrame(data=data, index=range(0, len(df)))
print(lengthOfMasg.describe())
#print out top mail by length
print(lengthOfMasg.nlargest(10, 'SecondLength'))

df = df.to_excel(r'',
                 index=False)
