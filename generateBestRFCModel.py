import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
import pickle
import seaborn as sns

from sklearn.model_selection import train_test_split, GridSearchCV, cross_val_score
from sklearn.utils import shuffle
from sklearn.metrics import classification_report, accuracy_score, confusion_matrix

from sklearn.ensemble import RandomForestClassifier

"""
Imports the dataframe and creates a classification model for selection
of the component causing the perturbation, and the possible proabability
"""

def import_data(filename):
    data = shuffle(pd.read_csv(filename, index_col=0))
    data_new = data[['0','1','2','3','4','5','6','7','8','9','10','11','12','13','14','15','category']]

    X = np.array(data_new.drop(['category'],axis=1))
    y = np.array(data_new['category'])
    # print(X.shape)
    # print(y.shape)

    X_train, X_test, y_train, y_test = train_test_split(X, y, test_size = 0.3)
    # print(X_train.shape)
    # print(X_test.shape)
    return X_train, X_test, y_train, y_test

# Function to plot Confusion Matrix
def plot_confusion_matrix(Y_pred,Y):
    cm = confusion_matrix(Y_pred, Y, labels=None, sample_weight=None)
    
    df_cm = pd.DataFrame(cm, range(4), range(4))
    sns.set(font_scale = 1.4) #for label size
    ax = sns.heatmap(df_cm, annot = True, annot_kws = {"size": 16})
    ax.set_title("Confusion Matrix")
    ax.set(xlabel = 'Predicted Class', ylabel = 'Actual Class')
    plt.show()

def perform_grid_search(X_train, y_train):

    n_estimators = [10, 50, 100, 300, 500, 800, 1200]
    max_depth = [3, 5, 8, 15, 25, 30]
    min_samples_split = [2, 5, 10, 15, 100]
    min_samples_leaf = [1, 2, 5, 10] 

    hyperF = dict(n_estimators = n_estimators, max_depth = max_depth,  
                min_samples_split = min_samples_split, 
                min_samples_leaf = min_samples_leaf)

    gridF = GridSearchCV(RandomForestClassifier(), hyperF, cv = 3, verbose = 1, n_jobs = -1).fit(X_train, y_train)
    best_params = gridF.best_params_
    print("Best parameters set found on development set:")
    print(best_params)
    print()

    rfc_model = RandomForestClassifier(n_estimators=best_params['n_estimators'],
                max_depth=best_params['max_depth'], min_samples_split = best_params['min_samples_split'],
                min_samples_leaf=best_params['min_samples_leaf'], random_state=0)

    pickle.dump(rfc_model, open('models/rfc_model.pkl', 'wb'))
    return rfc_model
    
def predict(X_train, y_train, X_test, y_test, use_previous_model=True):
    if use_previous_model:
        rfc_model = pickle.load(open("models/rfc_model.pkl", "rb"))
    else:
        rfc_model = perform_grid_search(X_train, y_train)

    rfc_model.fit(X_train, y_train)
    
    print('Cross Validation Score:-')
    print(cross_val_score(rfc_model, X_train, y_train, cv=5, n_jobs=-1)) 
    print()

    y_pred = rfc_model.predict(X_test) 
    # y_pred_proba = rfc_model.predict_proba(X_test)
  
    print('-------- CLASSIFICATION REPORT -----------')
    print(classification_report(y_test,y_pred))
    print('Accuracy:', accuracy_score(y_test, y_pred)*100)

    print("\n")
    print("Training set score for RFC: %f" % rfc_model.score(X_train, y_train))
    print("Testing set score for RFC: %f" % rfc_model.score(X_test, y_test ))

    plot_confusion_matrix(y_pred,y_test)

    print('Weight of feature importance:-')
    for i in range(len(rfc_model.feature_importances_)):
        print('Coefficient {} -> {}'.format(i, rfc_model.feature_importances_[i]))
    print()

if __name__=='__main__':
    X_train, X_test, y_train, y_test = import_data('data.csv')
    predict(X_train, y_train, X_test, y_test, use_previous_model=True)
