from mlmailclassify import DataPrep2, ML_model, test


root_path = ""


def mail_predict(singl_mail):
    df = DataPrep2.get_clean_data(root_path, singl_mail)
    msg_label = test.classify_mail(df)
    if msg_label is not None:
        return str(msg_label)


# just to teach the model
def train_model():
    DataPrep2.get_clean_data(root_path, save_model=True)
    test.teach_model()
