import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
import pickle

from sklearn.model_selection import train_test_split, GridSearchCV, KFold
from sklearn.utils import shuffle
from sklearn.metrics import classification_report, accuracy_score, confusion_matrix

from sklearn.linear_model import LinearRegression
from sklearn.tree import DecisionTreeRegressor

data = shuffle(pd.read_csv('final1.csv', index_col=0))
# print(data.shape)
# print(data.head())
# print(data.columns.to_list())

data_new = data[['0','1','2','3','4','5','6','7','8','9','10','11','12','13','14','15','category']]
df = data_new[data_new.loc[:,"category"] == 0].drop(columns=['category'])

params = {'dx':0.0, 'dy':0.0, 'tx':0.0, 'ty':0.0}
df = df.assign(**params)
# print(df.head())
# print(df.shape)

for index in df.index:
    x = float(index[2:])
    if index[:2] == 'dx':
        df.at[index, 'dx'] = x
    if index[:2] == 'dy':
        df.at[index, 'dy'] = x
    if index[:2] == 'tx':
        df.at[index, 'tx'] = x
    if index[:2] == 'ty':
        df.at[index, 'ty'] = x


print(df.head())
print()

X = np.array(df.drop(columns=['dx', 'dy', 'tx', 'ty']))
y = np.array(df[['dx', 'dy', 'tx', 'ty']])

print(X.shape)
print(y.shape)

X_train, X_test, y_train, y_test = train_test_split(X, y, test_size = 0.3)

reg = DecisionTreeRegressor(max_depth=5).fit(X_train, y_train)

pickle.dump(reg, open('reg_pm.pkl', 'wb'))
y_pred = reg.predict(X_test)

print()
print(y_test[5], y_pred[5])
print()
score = 0.0
for i in range(len(y_test)):
    ss = 0.0
    for j in range(4):
        ss += abs(y_test[i][j] - y_pred[i][j])
    score += ss
print(score)

kf = KFold(n_splits = 3, shuffle = True)
reg = DecisionTreeRegressor(max_depth=3)
scores = []
for i in range(3):
    result = next(kf.split(X), None)
    X_train = X[result[0]]
    X_test = X[result[1]]
    y_train = y[result[0]]
    y_test = y[result[1]]
        
    model = reg.fit(X_train,y_train)
    y_pred = reg.predict(X_test)

    score = 0.0  
    for i in range(len(y_test)):
        score += np.mean(np.abs(y_test[i] - y_pred[i]))
  
    scores.append(score)

print('Scores from each Iteration: ', scores)
print('Average K-Fold Score :' , np.mean(scores))
print()
    