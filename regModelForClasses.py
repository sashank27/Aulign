import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
import pickle
import argparse
import xgboost as xgb

from sklearn.model_selection import train_test_split, KFold
from sklearn.utils import shuffle
from sklearn.metrics import r2_score

from sklearn.tree import DecisionTreeRegressor
from sklearn.ensemble import RandomForestRegressor

np.random.seed(42)

"""
Imports the dataframe and creates a regression model for the component specified
in the arguments passed.
"""

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

def model(X, y, save_model=False, sel_class=0, max_depth=20):
    X_train, X_test, y_train, y_test = train_test_split(X, y, test_size = 0.3, random_state = 42)
    
    # reg = RandomForestRegressor(criterion='mae', max_depth=max_depth).fit(X_train, y_train)
    
    # if save_model:
    #     pickle.dump(reg, open('models/reg{}.pkl'.format(sel_class), 'wb'))
    
    dtrain = xgb.DMatrix(X_train, label=y_train)
    dtest = xgb.DMatrix(X_test, label=y_test)

    # # set xgboost params
    # param = {
    #     'max_depth': 3,  # the maximum depth of each tree
    #     'eta': 0.3,  # the training step for each iteration
    #     'silent': 1,  # logging mode - quiet
    #     'objective': 'multi:softprob',  # error evaluation for multiclass training
    #     'num_class': 3}  # the number of classes that exist in this datset
    
    # num_round = 20  # the number of training iterations
    
    reg = xgb.XGBRegressor(objective ='reg:linear', colsample_bytree = 0.3, learning_rate = 0.1,
                max_depth = 5, alpha = 10, n_estimators = 10).fit(X_train, y_train)

    y_pred = reg.predict(X_test)

    print(y_pred)

    epsilon = 1e-4

    scoredec = 0.0
    scoretilt = 0.0
    # print(y_test, y_pred)
    s1 = 0.0
    s2 = 0.0
    s3 = 0.0
    s4 = 0.0


    for i in range(len(y_test)):
        reldx = np.abs(y_test[i][0] - y_pred[i][0]) / (np.abs(y_test[i][0]) + epsilon)
        reldy = np.abs(y_test[i][1] - y_pred[i][1]) / (np.abs(y_test[i][1]) + epsilon)
        # print(reldy)
        s1 += reldx
        s2 += reldy
        scoredec += reldx + reldy

        reltx = np.abs(y_test[i][2] - y_pred[i][2]) / (np.abs(y_test[i][2]) + epsilon)
        relty = np.abs(y_test[i][3] - y_pred[i][3]) / (np.abs(y_test[i][3]) + epsilon)
        s3 += reltx
        s4 += relty
        
        scoretilt += reltx + relty
    
    print('Mean Relative Error in prediction of Decenter:', scoredec / len(y_test))
    print('Mean Relative Error in prediction of Tilt:', scoretilt / len(y_test))

    print(s1 / len(y_test))
    print(s2 / len(y_test))
    print(s3 / len(y_test))
    print(s4 / len(y_test))
    
    
    scoredec = r2_score(y_test[:][0:2], y_pred[:][0:2])
    scoretilt = r2_score(y_test[:][2:], y_pred[:][2:])
    print('R2 score for Decenter:', scoredec)
    print('R2 score for Tilt:', scoretilt)

    kf = KFold(n_splits = 3, shuffle = True)
    reg = DecisionTreeRegressor(max_depth=max_depth)
    
    scoresdec = []
    scorestilt = []

    for i in range(3):
        result = next(kf.split(X), None)
        X_train = X[result[0]]
        X_test = X[result[1]]
        y_train = y[result[0]]
        y_test = y[result[1]]
            
        model = reg.fit(X_train,y_train)
        y_pred = reg.predict(X_test)

        scoredec = 0.0  
        scoretilt = 0.0
        for i in range(len(y_test)):
            reldx = np.abs(y_test[i][0] - y_pred[i][0]) / np.abs(y_test[i][0] + epsilon)
            reldy = np.abs(y_test[i][1] - y_pred[i][1]) / np.abs(y_test[i][1] + epsilon)
            scoredec += reldx + reldy

            reltx = np.abs(y_test[i][2] - y_pred[i][2]) / np.abs(y_test[i][2] + epsilon)
            relty = np.abs(y_test[i][3] - y_pred[i][3]) / np.abs(y_test[i][3] + epsilon)
            scoretilt += reltx + relty

        scoresdec.append(scoredec)
        scorestilt.append(scoretilt)


    print('Decenter Scores from each Iteration: ', scoresdec)
    print('Average K-Fold Score for Decenter:' , np.mean(scoresdec))
    print()
    print('Tilt Scores from each Iteration: ', scorestilt)
    print('Average K-Fold Score for Tilt:' , np.mean(scorestilt))
    print()

if __name__=='__main__':
    parser = argparse.ArgumentParser()
    parser.add_argument(
        "-c", "--category", help="exact component to model the regression(0-PM, 1-SM, 2-Lens, 3-CCD)")
    args = parser.parse_args()

    c = args.category
    X, y = extract_data('data.csv', c)
    model(X,y, sel_class=c,save_model=True)

    


