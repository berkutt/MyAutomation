import pandas as pd
from trash import DataSource

'''
EXAMPLE HOW TO CALL: 
file = ml.models('ShipExprLabled.xlsx')
textfrfile = file.textpreproces()
file.NaiveBase(textfrfile,encoding=False) 
file.RandomForest(textfrfile,encoding=False)
'''


class models:

    def loadataset(self):
        print("Loading Excel file")
        return pd.read_excel(DataSource.path + self.filename)

    def __init__(self, filename):
        # self.pipeline = self.pipeline
        self.filename = filename
        self.dataset = self.loadataset()

    '''
    Functions:
    some helper functions
    main - define data to learn and what data to classify
    splitdata - split data to test and train set
    crossval -display accurency result for some random blocks
    
    Variables:
    classifier - model type. Pipeline will be colled from each model function. 
    dataset - excel file with labels as last column. 
    corpus - preprocessed text from models.textpreproces
    cvnr -  build a vocabulary that only consider the top max_features ordered by term frequency across the corpus.
    encoding - transform Labels to matrix. 
    test_size - size of test set and train set represented as decimal (0.25) 
    '''

    class pipeline:

        def main(self, classifier, dataset, corpus, cvnr, encoding, test_size):
            # create matrix where each column - word, presented as vector.
            if encoding:
                #y = dataset[["Labels"]]
                y = dataset.iloc[:, -1].values
                y = self.encodetext(y)
                #y = y[:, -1]
            else:
                y = dataset.iloc[:, -1].values

            X_train, X_test = self.splitdata(corpus, test_size)
            y_train, y_test = self.splitdata(y, test_size)

            from sklearn.feature_extraction.text import CountVectorizer
            cv = CountVectorizer(max_features=cvnr)  # for Cut vesrion 4500, for full - 9500
            X_train = cv.fit_transform(X_train)
            X_test = cv.transform(X_test)

            classifier.fit(X_train, y_train)
            y_pred = classifier.predict(X_test)
            # print(y_pred)
            # self.crossval(classifier, X_train, y_train)
            from sklearn.metrics import accuracy_score
            print("Accurancy score:")
            print(accuracy_score(y_test, y_pred))

        # change predicted columns with categries to matrix
        def encodetext(self, y):
            print("Encoding True")
            from sklearn.preprocessing import OneHotEncoder
            cat_encoder = OneHotEncoder()
            y = cat_encoder.fit_transform(y)
            return y.toarray()

        def splitdata(self, dataset, test_size):
            from sklearn.model_selection import train_test_split
            train, test = train_test_split(dataset, test_size=test_size, random_state=0)
            return train, test

        def crossval(self, classifier, X_train, y_train):
            print("Cross validation score")
            from sklearn.model_selection import cross_val_score
            accurecies = cross_val_score(estimator=classifier, X=X_train, y=y_train, cv=10)
            print(accurecies)

    '''
    *********
    MODELS
    *********
    '''

    def NaiveBase(self, corpus, cvnr=4500, test_size=0.25,
                  encoding=False):
        print("***Naive Base***")
        from sklearn.naive_bayes import GaussianNB
        classifier = GaussianNB()

        self.pipeline().main(classifier=classifier, dataset=self.dataset,
                             corpus=corpus, cvnr=cvnr, encoding=encoding, test_size=test_size)

    def RandomForest(self, corpus, cvnr=4500, test_size=0.25, n_estimators=10,
                     encoding=False):

        print("***RandomForest***")
        from sklearn.ensemble import RandomForestClassifier
        classifier = RandomForestClassifier(n_estimators=n_estimators, criterion='entropy', random_state=0)
        RandomForestClassifier(bootstrap=True, ccp_alpha=0.0, class_weight=None,
                               criterion='entropy', max_depth=None, max_features='auto',
                               max_leaf_nodes=None, max_samples=None,
                               min_impurity_decrease=0.0, min_impurity_split=None,
                               min_samples_leaf=1, min_samples_split=2,
                               min_weight_fraction_leaf=0.0, n_estimators=n_estimators,
                               n_jobs=None, oob_score=False, random_state=0, verbose=0,
                               warm_start=False)

        self.pipeline().main(classifier=classifier, dataset=self.dataset,
                             corpus=corpus, cvnr=cvnr, encoding=encoding, test_size=test_size)

    def LogisticRegression(self, corpus, cvnr=4500, test_size=0.25,
                           encoding=False):

        print("***Logistic Regression***")
        from sklearn.linear_model import LogisticRegression
        classifier = LogisticRegression(random_state=0)

        self.pipeline().main(classifier=classifier, dataset=self.dataset,
                             corpus=corpus, cvnr=cvnr, encoding=encoding, test_size=test_size)

    def K_NearestNeighbors(self, corpus, cvnr=4500, test_size=0.25,
                           encoding=False):

        print("***K Nearest Neighbors***")
        from sklearn.neighbors import KNeighborsClassifier
        classifier = KNeighborsClassifier(n_neighbors=5, metric='minkowski', p=2)

        self.pipeline().main(classifier=classifier, dataset=self.dataset,
                             corpus=corpus, cvnr=cvnr, encoding=encoding, test_size=test_size)

    def SupportVectorMachine(self, corpus, cvnr=4500, test_size=0.25,
                             encoding=False):

        print("***Support Vector Machine***")
        from sklearn.svm import SVC
        classifier = SVC(kernel='linear', random_state=0)

        self.pipeline().main(classifier=classifier, dataset=self.dataset,
                             corpus=corpus, cvnr=cvnr, encoding=encoding, test_size=test_size)

    def KernelSVM(self, corpus, cvnr=4500, test_size=0.25,
                  encoding=False):

        print("***Kernel SVM***")
        from sklearn.svm import SVC
        classifier = SVC(kernel='rbf', random_state=0)

        self.pipeline().main(classifier=classifier, dataset=self.dataset,
                             corpus=corpus, cvnr=cvnr, encoding=encoding, test_size=test_size)

    def DecisionTree(self, corpus, cvnr=4500, test_size=0.25,
                     encoding=False):

        print("***Decision Tree***")
        from sklearn.tree import DecisionTreeClassifier
        classifier = DecisionTreeClassifier(criterion='entropy', random_state=0)

        self.pipeline().main(classifier=classifier, dataset=self.dataset,
                             corpus=corpus, cvnr=cvnr, encoding=encoding, test_size=test_size)
