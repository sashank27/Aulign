import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
import os

"""
This file is used to iterate over the data files obtained from the ZOS-API,
and is used to create a collection of values as a pandas dataframe
"""

cwd = os.getcwd()
path = os.path.join(cwd, 'data')

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

params = {'dx':0.0, 'dy':0.0, 'tx':0.0, 'ty':0.0}
df = data.assign(**params)

for index in df.index:
    if index[:2] == 'dx':
        df.at[index, 'dx'] = float(index[2:])
    if index[:2] == 'dy':
        df.at[index, 'dy'] = float(index[2:])
    if index[:2] == 'tx':
        df.at[index, 'tx'] = float(index[2:])
    if index[:2] == 'ty':
        df.at[index, 'ty'] = float(index[2:])

# print(data.shape)
# print(data.head())
# print(data['category'].unique())

df.to_csv(os.path.join(cwd, 'data.csv'))
