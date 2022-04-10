from nltk.corpus import stopwords
from nltk.stem.porter import PorterStemmer
import re
import os
import pandas as pd
from MyModules import datasets
dataset = datasets.Dataset_20200521()

def get_clean_data(root_path, singl_mail=False, save_model=False):
    """
    :param root_path: path to nested folder wth mails
    :param singl_mail: when using model passing messege from Outllok as Object
    :return: dataframe
    """

    def textpreproces(Body):
        # remove eerything that are not latters
        # make everything lowercase and keep root of the word
        corpus = []
        review = re.sub('[^a-zA-Z]', ' ', Body)  # delete everything that not letters
        review = review.lower()
        review = review.split()
        # ps = PorterStemmer()
        # all_stopwords = stopwords.words('english')
        # review = [ps.stem(word) for word in review if not word in set(all_stopwords)]
        # review = ' '.join(review)
        corpus.append(review)
        return corpus[0]

    def cut_body(body, sender):
        # cut mailbody. Every massage here having whole mailchain in the body.
        bodystr = body.split(sender)[0]
        bodystr.lower()
        with open(r'') as f:
            for message_cut_part in f:
                if len(message_cut_part) > 5: bodystr = bodystr.split(message_cut_part.lower())[0]
        return bodystr

    # prepare data and create df in the end.
    PDFbool = []
    body = []
    zfrom = []
    Label = []
    Subject = []

    import win32com.client
    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    # for each folder (that are label) check all files inside. Files are .msg from Outlook
    for folder in os.listdir(root_path):
        messeges = os.listdir(root_path + folder)
        for item in messeges:
            # if i am checking one mail (using model and not preparing data for teaching)
            if singl_mail:
                msg = singl_mail
            else:
                msg = outlook.OpenSharedItem(root_path + folder + "\\" + item)
            # get True if PDF is attached
            pdfbool = False
            for ItemNr in range(1, msg.Attachments.Count + 1):
                if msg.Attachments.Item(ItemNr).FileName[-3:] in ['pdf', 'PDF']:
                    PDFbool.append('PDF')
                    pdfbool = True
                    break
                elif ItemNr == msg.Attachments.Count:
                    PDFbool.append('0')
                    pdfbool = True
            if not pdfbool:
                continue

            bodystr = cut_body(msg.Body, msg.SenderName)
            # save mail attributes
            body.append(textpreproces(bodystr))
            Subject.append(textpreproces(msg.Subject.replace("RE: ", "").replace("FW: ", "")))
            zfrom.append(msg.SenderName)

            if singl_mail:
                # data = {'Subject': Subject, 'SenderName': zfrom, 'Body': body,
                #         'AttachCount': PDFbool}
                data = { 'Body': body,
                        'AttachCount': PDFbool}
                df = pd.DataFrame(data=data)
                return df
            Label.append(folder)

    # data = {'Subject': Subject, 'SenderName': zfrom, 'Body': body,
    #         'AttachCount': PDFbool, 'Label': Label}
    data = {'Body': body,
            'AttachCount': PDFbool, 'Label': Label}

    df = pd.DataFrame(data=data)
    if save_model: dataset.write_data('cleanlabelcut', df)
    return df




