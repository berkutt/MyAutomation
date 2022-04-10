from trash import ML_models as ml


def get_data():
    import datasets
    dataset = datasets.Dataset_20200521()
    df = dataset.read_data('cleanlabelcut')


if __name__ == '__main__':
    from sklearn.tree import DecisionTreeClassifier

    classifier = DecisionTreeClassifier(criterion='entropy', random_state=0)

    file2 = ml.models('3105Cleaned.xlsm')
    X = file2.textpreproces()
    print('VerData loaded')
    file1 = ml.models('ShipExprLabledCut.xlsx')
    Y = file1.textpreproces()


    from sklearn.feature_extraction.text import CountVectorizer

    cv = CountVectorizer(max_features=20)
    traindata = cv.fit_transform(traindata)
    VerfData = cv.transform(VerfData)

    from sklearn.model_selection import train_test_split

    X_train, X_test = train_test_split(traindata, test_size=0.2, random_state=0)
    y_train, y_test = train_test_split(file1.dataset.iloc[:, -1].values, test_size=0.2, random_state=0)

    classifier.fit(X_train, y_train)
    y_pred = classifier.predict(VerfData)


#todo XGboos, CATboos.



# pickle.dump(classifier, open(DataSource.path + 'Decision Tree','wb'))
