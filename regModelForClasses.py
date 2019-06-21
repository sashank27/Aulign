import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
import pickle
import argparse

from sklearn.model_selection import train_test_split, KFold
from sklearn.utils import shuffle

from sklearn.tree import DecisionTreeRegressor

def extract_data(filename, sel_class=0):
    data = shuffle(pd.read_csv(filename, index_col=0))
    data_new = data[['0','1','2','3','4','5','6','7','8','9','10','11','12','13','14','15','category','dx','dy','tx','ty']]
    df = data_new[data_new.loc[:,"category"] == int(sel_class)].drop(columns=['category'])

    # print(df.head())
    # print()

    X = np.array(df.drop(columns=['dx', 'dy', 'tx', 'ty']))
    y = np.array(df[['dx', 'dy', 'tx', 'ty']])

    # print(X.shape)
    # print(y.shape)

    return X, y

def model(X, y, save_model=False, sel_class=0, max_depth=5):
    X_train, X_test, y_train, y_test = train_test_split(X, y, test_size = 0.3)
    
    reg = DecisionTreeRegressor(max_depth=max_depth).fit(X_train, y_train)
    
    if save_model:
        pickle.dump(reg, open('models/reg{}.pkl'.format(sel_class), 'wb'))
    
    y_pred = reg.predict(X_test)

    score = 0.0
    # print(y_test[5], y_pred[5])

    for i in range(len(y_test)):
        score += np.mean(np.abs(y_test[i] - y_pred[i]))
    print('Mean Absolute Error in prediction:', score)
    print()

    kf = KFold(n_splits = 3, shuffle = True)
    reg = DecisionTreeRegressor(max_depth=max_depth)
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

if __name__=='__main__':
    parser = argparse.ArgumentParser()
    parser.add_argument(
        "-c", "--category", help="exact component to model the regression(0-PM, 1-SM, 2-Lens, 3-CCD)")
    args = parser.parse_args()

    c = args.category
    X, y = extract_data('data.csv', c)
    model(X,y, sel_class=c,save_model=True)

    


