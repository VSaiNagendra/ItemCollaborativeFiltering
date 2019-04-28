from openpyxl import load_workbook
import csv
import time
import math
import pickle
import numpy as np
import pandas as pd
def cosineS(a, b):
    print(a)
    print(b)
    p = a * b
    c, d = p / b, p / a
    c[np.isnan(c)] = 0
    d[np.isnan(d)] = 0
    norm = np.linalg.norm(c) * np.linalg.norm(d)
    if not norm:
        return 0
    value = np.dot(c, d) / norm
    if(not math.isnan(value)):
        print('cosineS' + str(value))
    else:
        print('Nan')
    return value


def cosineMatrix(ratingMatrix):
    print(ratingMatrix)
    noitems = np.shape(ratingMatrix)[1]
    print(np.shape(ratingMatrix))
    sim = np.zeros((noitems, noitems))
    for i in range(noitems):
        for j in range(i, noitems):
            sim[i][j] = cosineS(ratingMatrix[:, i], ratingMatrix[:, j])
            print(i, j)
            sim[j][i] = sim[i][j]
            print(sim)
    sim.dump('tempsimilarity.pkl')
    return sim


book = load_workbook('ratingitemp.xlsx')
sheet = book['premiss']
data = pd.read_csv('userrating.csv')
itemsim = data.pivot_table(index=['u_id'], columns=['m_id'], values='rating')
itemsim = itemsim.iloc[27:31, 0:3]
sheetoriginal=book['original1']
for i in range(4):
    sheetoriginal.append(itemsim.iloc[i].tolist())
print(itemsim)
for i in range(4):
    for j in range(0,3):
        print(i,j)
        if not math.isnan(itemsim.iloc[i, j]):
            itemsim.iloc[i, j] = np.nan
            break
        print(i, j)

print(itemsim)

isim = itemsim.reset_index().values
isim = isim[:, 1:]
csim = cosineMatrix(isim)
simCand = pd.Series()
cum = 0
dcum = 0
cormatrix = pd.DataFrame(index=itemsim.iloc[0].index, columns=itemsim.iloc[0].index, data=csim[:, :])
print(cormatrix)
for k in range(0, len(itemsim.index)):
    newuser = itemsim.iloc[k]
    newratings = newuser.dropna()
    for i in newuser.index:
        print(i, end=' ')
    for i in newratings.index:
        print(i, end=' ')
    l = []
    for i in range(0, len(newuser.index)):
        cum = 0
        dcum = 0
        if newuser.index[i] not in newratings.index:
            for j in range(0, len(newratings.index)):
                if(not math.isnan(cormatrix[newuser.index[i]][newratings.index[j]])):
                    cum = cum + cormatrix[newuser.index[i]][newratings.index[j]] * newratings[newratings.index[j]]
                    dcum = dcum + cormatrix[newuser.index[i]][newratings.index[j]]
                print(cum, dcum)
            if(dcum == 0):
                l.append(-1)
            else:
                a = cum / dcum
                if(not math.isnan(a)):
                    l.append(a)
                else:
                    l.append(-2)
        else:
            l.append(newuser[newuser.index[i]])
    sheet.append(l)
book.save('D:\\sai\\pca\\tfvector\\\itemcolloborative\\item_collobarative_movies_data\\ratingitemp.xlsx')
