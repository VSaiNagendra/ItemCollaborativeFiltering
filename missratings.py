from openpyxl import load_workbook
import csv
import time
import math
import numpy as np
import pandas as pd

book = load_workbook('ratingitemp.xlsx')
data = pd.read_csv('userrating.csv')
itemsim = data.pivot_table(index=['u_id'], columns=['m_id'], values='rating')
itemsim = itemsim.iloc[0:101, 0:501]
tempsim = itemsim.iloc[:, :]
sheet1 = book['original']
sheet2 = book['filtered']
for i in range(101):
    sheet1.append(itemsim.iloc[i].tolist())
for i in range(101):
    for j in range(0, 501):
        if not math.isnan(itemsim.iloc[i, j]):
            itemsim.iloc[i, j] = np.nan
            break
        print(i, j)

for i in range(101):
    sheet2.append(itemsim.iloc[i].tolist())
book.save('D:\\sai\\pca\\tfvector\\\itemcolloborative\\ratingitemp.xlsx')
