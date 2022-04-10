import datasets
from nltk.corpus import stopwords
from nltk.stem.porter import PorterStemmer
import re

#df2 - list of cleaned messeges
#df - initial df
# True = good row
def filter_row_1(row):
    message = str(row["Body"])
    if any(text in message for text in [
        "INTTRA SI ACT, Rev 3.0", "was processed by INTTRA",
        "LATE CANCELLATION FEE: With effective from ",
        "浴"]):
        return False
    return True


def clean_row_1(row):
    message = str(row["Body"])
    message = message.split('http')[0]  # delete web links like re.sub(r'http\S+', '', message) do
    message = '\n'.join(line for line in message.split('\n') if line)  # remove empty lines

    # cut message after these parts

    for sender in set(df.SenderName):
        message = message.split(sender.rstrip())[0]

    with open('../CutKeyWords.txt') as f:
        for message_cut_part in f:
            message = message.split(message_cut_part.rstrip())[0]

    row["Body"] = message

    row["Subject"] = str(row["Subject"]).replace("RE: ", "").replace("FW: ", "")

    # for column in ["CC", "ReceivedTime", "To", "Categories"]:  # sill don't get it
    #     row.pop(column)
    return row

class addLabels():

    def __init__(self):
        self.temp_indexes = []
        self.dataset = datasets.Dataset_20200521()
        self.df = self.get_df()
        self.df2 = self.get_df2()
        self.labels = {}

    def get_df(self):
        return self.dataset.read_data_pickle('initmails')

    def get_df2(self):
        return self.dataset.read_data_pickle('cleanedmails')

    from difflib import SequenceMatcher
    def similar(self, a, b):
        return self.SequenceMatcher(None, a, b).ratio()

    def add_cut_word(self):
        userinput = input("what cut key-word should be added ? : ")
        if userinput:
            with open('../CutKeyWords.txt', 'a') as f:
                f.write(userinput + '\n')

    def add_label(self):
        userinput = input("what label should be here ? : ")
        if userinput:
            labels = []
            for index in self.temp_indexes:
                labels.append(self.df2[index])
                self.labels[userinput] = labels
            self.dataset.write_data_pickle('labels', self.labels)

    def main(self):
        # df - list of body messeges (str)
        for i in range(0, len(self.df2)):
            if i > len(self.df2): break
            ratio_count = 0
            samplemsg = self.df2[i]  # get message
            for j in range(len(self.df2)):
                msg = self.df2[j]
                similarity = self.similar(msg, samplemsg)
                if similarity > 0.9:
                    ratio_count = ratio_count + 1
                    self.temp_indexes.append(j)

            print(ratio_count, i)

            if ratio_count > 100:

                print("messege is: ", self.df.Body[i][:300])
                print(self.df2[i][:300])

                user_decision = input("Would You like to add cut key-word or add label. cut/label: ")

                if user_decision == "label":
                    self.add_label()
                elif user_decision == "cut":
                    self.add_cut_word()

            # remove already checked messeges
            self.df2 = [ii for jj, ii in enumerate(self.df2) if
                        jj not in self.temp_indexes]  # remove values with similarity >0.9 from list of messeges
            self.df = self.df.drop(self.df.index[self.temp_indexes])
            # reset indexes in dataframe
            self.df = self.df.reset_index(drop=True)

            # save updated df and df2
            self.dataset.write_data_pickle('df_temp', self.df)
            self.dataset.write_data_pickle('df2_temp', self.df2)

            # reset list of indexes with high similarity
            self.temp_indexes = []
            # debug
            print("labels : ", self.labels)
            print(len(self.df), len(self.df2))

def filter_emptymails(row):
    return all(len(str(column).strip()) > 0 for column in row)


def pop_col_attachm(row):
    row.pop("AttachCount")
    return row

def textpreproces(df):
    corpus = []
    for i in range(0, len(df)):
        review = re.sub('[^a-zA-Z]', ' ', df['Body'][i])  # delete everything that not letters
        review = review.lower()
        review = review.split()
        ps = PorterStemmer()
        all_stopwords = stopwords.words('english')
        all_stopwords.remove('not')
        review = [ps.stem(word) for word in review if not word in set(all_stopwords)]
        review = ' '.join(review)
        corpus.append(review)
    return corpus

if __name__ == "__main__":
    dataset = datasets.Dataset_20200521()


    class Counter:
        def __init__(self, target):
            self.counter = 0
            self.target = target

        def increment(self):
            self.counter += 1

        def less(self):
            return self.counter < self.target

#todo DROP NA
    df = dataset.read_data('rawdata')

    #df = OutlookConn.getmails()

    rows_1 = []

    for row in dataset.iterrows(df):
        row = dict(row)
        # first filtering - drop some massagees that a not important or from folder = Importante
        if not filter_row_1(row):
            continue
        # cut massages by some Key words. Also clean subject from RE and FW
        row = clean_row_1(row)


        # row["Label"] = label_row(row)
        # row = pop_col_attachm(row)
        if not filter_emptymails(row):
            continue
        rows_1.append(row)
    df = dataset.from_list_dict(rows_1)
    #clean Body column
    df2 = textpreproces(df)

    dataset.write_data_pickle('initmails', df)
    dataset.write_data_pickle('cleanedmails', df2)

    addLabels().main()

    dataset.write_data('cleanlabel', df)
#todo define how much to drop: topLablecount < 2x 2nd Labcount
    other_counter = Counter(20000)
    rows_2 = []
    for row in rows_1:
        if row["Label"] == "Other":
            other_counter.increment()
            if other_counter.less():
                continue
        rows_2.append(row)
    df = dataset.from_list_dict(rows_2)
    dataset.write_data('cleanlabelcut', df)