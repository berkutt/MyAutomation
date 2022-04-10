import pandas as pd
from trash import DataSource


def addLabels(data):
    # Function to search key-words in massage. Return true if something found
    def checkword(ListOfWords, text):
        for word in ListOfWords:
            if word in text: return True


    data = pd.read_excel(DataSource.path + '')
    data = data[data.Folder != "Importante"]
    data = data.reset_index()
    data = data.drop(columns=["CC", "Folder", "ReceivedTime", "To", "Categories"])
    print(data.columns)

    Lable = []
    for index, row in data.iterrows():
        if row['Body'].find("I proceed with booking") > 0 and \
                (row["SenderName"] == "" or row["SenderName"] == "") and \
                row["AttachCount"] > 0:
            Lable.append("Booking Confirmation SanGG")
            continue
        elif row['Body'].find("PGI OK") > 0 and \
                row["SenderName"] == "":
            Lable.append("Issue Invoice and send it to ")
            continue
        elif row['Body'].startswith(" The content of this email is confidential ", 0, 100) and \
                row["SenderName"] == "" and \
                row["AttachCount"] > 0:
            Lable.append("Booking confirmation from ")
            continue
        elif str(row['Subject']).startswith("Order to book ", 0, 100) and \
                row["SenderName"] == "":
            Lable.append("Book delivery from ")
            continue
        elif str(row['Subject']).startswith("for ", 5, 20) and \
                row["SenderName"] == "":
            Lable.append("Book delivery from ")
            continue
        elif str(row['Subject']).startswith("Allocated PO", 0, 20) and \
                row["SenderName"] == "":
            Lable.append("Book delivery from ")
            continue
        elif checkword(("DateDocument", "ByDocument", "NumberMaterialMaterial"), row['Body']) and \
                ("" == row["SenderName"] or row["SenderName"] == ""):
            Lable.append("Book delivery from ")
            continue
        elif str(row['Body']).startswith(" ", 0, 30):
            Lable.append("Book delivery from ")
            continue
        elif row["SenderName"] == "":
            Lable.append("")
            continue
        elif checkword(("draft", "Draft", "HiDraft"), row['Body']) and \
                row["AttachCount"] > 0:
            Lable.append("Draft BL")
            continue
        else:
            Lable.append("Other")
            continue

    d = {x: Lable.count(x) for x in Lable}
    print(d)
    data["Labels"] = Lable
    # cut most rows with categry Other
    deletetarget = 0
    j = 0
    for i in range(len(data)):
        if deletetarget == 20000:
            break
        if data.iloc[j].Labels == "Other":
            deletetarget += 1
            data = data.drop(data.index[j])
        else:
            j += 1

    df2 = data.to_excel(DataSource.path + '', index=False)
