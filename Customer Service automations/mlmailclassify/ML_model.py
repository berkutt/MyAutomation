import pandas as pd

from sklearn.naive_bayes import GaussianNB
from sklearn.ensemble import RandomForestClassifier
from sklearn.neighbors import KNeighborsClassifier
from sklearn.tree import DecisionTreeClassifier

from MyModules import datasets
dataset = datasets.Dataset_20200521()

def teach_model():
    # test some popular models and store results in df
    def test_model():

        # nested list. Here will be stored results from running models
        results = []
        for features in range(30, 300, 10):
            from sklearn.feature_extraction.text import CountVectorizer
            cv = CountVectorizer(max_features=features)
            X = cv.fit_transform(df.Total).toarray()
            dataset.write_data_pickle('Vectorizer', cv)

            y = df.Label.values

            # Splitting the dataset into the Training set and Test set
            from sklearn.model_selection import train_test_split
            X_train, X_test, y_train, y_test = train_test_split(X, y, test_size=0.20, random_state=0)

            # loop through models and store results in the list
            for mode_classifier in models:
                classifier = mode_classifier()
                classifier.fit(X_train, y_train)
                y_pred = classifier.predict(X_test)

                from sklearn.metrics import accuracy_score
                results.append([mode_classifier, features, accuracy_score(y_test, y_pred)])

        return results

    # select model with best accurency
    def select_model(model, features):
        print('Model:', model, ' Nr of words:', features, 'Accuracy', df_models.iloc[0][2])
        # model.__name__ == "RandomForestClassifier"

        from sklearn.feature_extraction.text import CountVectorizer
        cv = CountVectorizer(max_features=features)
        X = cv.fit_transform(df.Total).toarray()

        # save CV, load it when classifying single mail
        dataset.write_data_pickle('Vectorizer', cv)
        y = df.Label.values

        # Splitting the dataset into the Training set and Test set
        from sklearn.model_selection import train_test_split
        X_train, X_test, y_train, y_test = train_test_split(X, y, test_size=0.20, random_state=0)

        classifier = model()
        classifier.fit(X_train, y_train)

        # display
        print("Importance of each mail attribute: ")
        for name, score in zip(df.columns, classifier.feature_importances_):
            print(name, score)
        # save model
        dataset.write_data_pickle('ML_model', classifier)

    models = [GaussianNB, RandomForestClassifier, KNeighborsClassifier, DecisionTreeClassifier]
    # get trained data
    df = dataset.read_data('cleanlabelcut')
    df['Total'] = df[df.columns[:-1]].apply(
        lambda x: ','.join(x.dropna().astype(str)),
        axis=1)

    # best_model - nested list
    best_models = test_model()
    # get DataFrame from nested list
    # looks like this: <class 'xgboost.sklearn.XGBClassifier'>  100  0.981818
    df_models = pd.DataFrame(best_models[0:], columns=['Model', 'CV', 'Accuracy'])
    df_models = df_models.sort_values('Accuracy', ascending=False)
    print(df_models[:5])
    # get best feature ( CV ) and model
    features = int(df_models.iloc[0][1])
    model = df_models.iloc[0][0]

    # return classifier
    select_model(model=model, features=features)


def classify_mail(my_mail, threshold=0.5):
    model = dataset.read_data_pickle('ML_model')
    cv = dataset.read_data_pickle('Vectorizer')
    my_mail['Total'] = my_mail[my_mail.columns[:]].apply(
        lambda x: ','.join(x.dropna().astype(str)),
        axis=1)

    VerfData = cv.transform(my_mail['Total']).toarray()
    # print(model.predict_proba(VerfData))

    if len(VerfData) != 0:
        predicted_proba = model.predict_proba(VerfData)
        # print probability for each mail type
        print([i for i in zip(predicted_proba[0], model.classes_)])
        for probs in predicted_proba:
            # Iterating over class probabilities
            for i in range(len(probs)):
                if probs[i] >= threshold:
                    return model.classes_[i]
    else:
        return "Trash"