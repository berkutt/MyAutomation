import numpy as np
import matplotlib.pyplot as plt
import pandas as pd
import DataSource

dataset = pd.read_excel(DataSource.path + 'ShipExprLabled.xlsx')

import re
#import nltk
#nltk.download('stopwords')
from nltk.corpus import stopwords
from nltk.stem.porter import PorterStemmer
corpus = []
for i in range(0, len(dataset)):
    review = re.sub('[^a-zA-Z]', ' ', dataset['Body'][i]) #delete everything that not letters
    review = review.lower()
    review = review.split()
    ps = PorterStemmer()
    all_stopwords = stopwords.words('english')
    all_stopwords.remove('not')
    review = [ps.stem(word) for word in review if not word in set(all_stopwords)]
    review = ' '.join(review)
    corpus.append(review)
dataset.Body = corpus

from sklearn.feature_extraction.text import CountVectorizer
cv = CountVectorizer(max_features = 4500) #for Cut vesrion 4500, for full - 9500
X = cv.fit_transform(corpus).toarray()
# X = dataset.iloc[:, -1].values #TODO adding more then one column. But what with cv then ?
#y = dataset.iloc[:, -1].values
y = dataset[["Labels"]]

#change predicted columns with categries to matrix
from sklearn.preprocessing import OneHotEncoder
cat_encoder = OneHotEncoder()
y = cat_encoder.fit_transform(y)
y = y.toarray()

#split dataset to train and test
from sklearn.model_selection import train_test_split
X_train, X_test, y_train, y_test = train_test_split(X, y, test_size = 0.25, random_state = 0)

#do magic
from sklearn.ensemble import RandomForestClassifier
classifier = RandomForestClassifier(n_estimators = 10, criterion = 'entropy', random_state = 0)
classifier.fit(X_train, y_train)

RandomForestClassifier(bootstrap=True, ccp_alpha=0.0, class_weight=None,
                       criterion='entropy', max_depth=None, max_features='auto',
                       max_leaf_nodes=None, max_samples=None,
                       min_impurity_decrease=0.0, min_impurity_split=None,
                       min_samples_leaf=1, min_samples_split=2,
                       min_weight_fraction_leaf=0.0, n_estimators=10,
                       n_jobs=None, oob_score=False, random_state=0, verbose=0,
                       warm_start=False)

y_pred = classifier.predict(X_test)
from sklearn.model_selection import cross_val_score
accurecies = cross_val_score(estimator=classifier,X = X_train, y = y_train, cv=10)
print(accurecies)
from sklearn.metrics import confusion_matrix, accuracy_score
#cm = confusion_matrix(y_test, y_pred)
#print(cm)
print(accuracy_score(y_test, y_pred))
'''
result of different test_size
0.8637837837837837 0.1
0.855803893294881 0.15000000000000002
0.8528934559221201 0.20000000000000004
0.8559065339679792 0.25000000000000006
0.853588171655247 0.30000000000000004
0.8529048207663782 0.3500000000000001
0.8539751216873986 0.40000000000000013
0.8574519230769231 0.45000000000000007
0.857204673301601 0.5000000000000001
0.8583792289535799 0.5500000000000002
Process finished with exit code 0

'''