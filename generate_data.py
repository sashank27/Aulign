import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
import os

cwd = os.getcwd()
path = os.path.join(cwd, 'Data')

catlist = []
i = 0
category = []
data = pd.DataFrame()
for subdirs in os.listdir(path):
    catlist.append(subdirs)

for cat in catlist:
    
    for root, subdirs, files in os.walk(os.path.join(path, cat)):
        
        for filename in files:
            file_path = os.path.join(root, filename)
            if filename == "zernike.csv":
                extract = pd.read_csv(file_path, header = None, names = [root[root.rfind('/')+1:]])
                extract = extract.T
                data = data.append(extract)
                category.append(i)
    i=i+1

data['category'] = category
# print(data.shape)
# print(data.head())
# print(data['category'].unique())
data.to_csv(os.path.join(cwd, 'final1.csv'))