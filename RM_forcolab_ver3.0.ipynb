# pip install pulp
# !pip install --upgrade -q gspread
 # -*- coding: utf-8 -*-
"""
Created on Fri Jul 12 12:01:39 2019
@author: ryo watanabe 11th 
※入力ファイルは CSV UTF-8(コンマ区切り)(*.csv) で保存すること
"""
import pulp
import glob
import openpyxl
import numpy as np
import pandas as pd
import time
# import matplotlib.pyplot as plt
# import matplotlib.cm as cm
class pycolor:
    BLUE = '\033[34m'
    HIGHLIGHT = '\033[01m'
    END = '\033[0m'
start = time.time()

import_attend_path = "C:\\Users\\tai\\Desktop\\systemGUI\\attend.xlsx"
 
A_pd = pd.read_excel(import_attend_path,"d2")
A_pd = pd.DataFrame(A_pd)

A_pd.columns = list(A_pd.iloc[0,:])
A_pd = A_pd.drop(A_pd.index[0]).reset_index(drop=True)
A_np = A_pd.iloc[2:,2:].fillna("0").values.astype(int)


import_need_path = "C:\\Users\\tai\\Desktop\\systemGUI\\need.xlsx"

N_pd = pd.read_excel(import_need_path,"need")

N_pd.columns = list(N_pd.iloc[0,:])
N_pd = N_pd.drop(N_pd.index[0]).reset_index(drop=True)
N_np = N_pd.iloc[2:,3:].fillna("0").values.astype(int)
N_np1 = N_pd.iloc[2:,:].fillna("0").values


N_pd_d = N_pd.drop(N_pd.columns[0],axis=1)
a = pd.DataFrame()
a["Unnamed: 0"] = N_pd.iloc[:,0]
a["Unnamed: 1"] = N_pd.iloc[:,1]
N_pd_d = pd.concat([a, N_pd_d], axis=1).replace("","0")


#attendfile
attend = A_pd
column_name1 = []
attend_dash = attend.drop(attend.columns[0],axis = 1)
attend_dash = attend_dash.drop(1)
#needfile
need = N_pd_d
column_name = []
need_dash = need.drop(1)
need_dash = need_dash.drop(["Unnamed: 0","名前番号"],axis = 1)

menu_rate = need_dash.copy()#今後使う
menber = need.iloc[0,3:]#メンバーリスト
span = attend.iloc[2:,:2]#timeリスト

###
# attend_num = attend_dash.iloc[1:,1:].astype(int).sum(axis=1).values
# label=attend_dash.iloc[1:,0].values
# bar_num = []
# for i in range(len(label)):
#   bar_num.append(i+1)
# plt.figure(figsize=(10,8),dpi=50,facecolor="w")
# plt.bar(bar_num, attend_num,color="#0078D7", width=0.4, tick_label=label, align="center")
# plt.title("attend_num")
# plt.show()

K = 3 #被り人数許容上限
W = 5 #同時練習許容上限

namedic = {}#メンバー名辞書
for i in range(len(N_pd.iloc[0])-3):
    namedic[i+1] = N_pd.iloc[0][i+3]
traindic = {}#練習名辞書
for i in range(len(N_pd)-2):
    traindic[i+1] = N_pd.iloc[i+2][1]#
#j = やる練習 指定
training = []
for i in range(len(N_np1)):
    if N_np1[i][2] == "1":
        training.append(i+1)
#練習に必要なメンバー行列（N_np)
col = []
for i in range(len(N_np1[0])-3):
    col.append(i+3)
for i in range(len(N_np)):
    for j in range(len(N_np[0])):
        if N_np[i,j] == -1:
            N_np[i,j] = 0    
#i = member 指定
people = []
for i in range(len(N_np[0])):
    people.append(i+1)
#ｔ = コマ 指定
times = []
for i in range(len(A_np)):
    times.append(i+1)
#時間帯
timezone = A_pd.iloc[2:,1].values


#関数
m = pulp.LpProblem("bestpractice", pulp.LpMaximize)
x = pulp.LpVariable.dicts('X', (people, training, times), 0, 1, pulp.LpInteger)#tコマにiさんがj練習をするかどうか
y = pulp.LpVariable.dicts('Y', (training, times), 0, 1, pulp.LpInteger)#tコマにj練習をするかどうか
m += pulp.lpSum(x[i][j][t] for i in people for j in training for t in times ), "TotalPoint"
#制約
for t in times:#練習はまとめて4つまで　t期にする練習全部足したらw以下に
    m += pulp.lpSum(y[j][t] for j in training) <= W              
for j in training:#同じ練習は1回まで　jの練習に対して全期分足したら1以下に    
    m += pulp.lpSum(y[j][t] for t in times) <= 1            
for i in people:#やる練習にしか参加できない　参加しないのはあり    
    for j in training:
        for t in times:
            m += x[i][j][t] <= y[j][t]
for i in people:#必要な練習しかしない　必要な練習でもやらないのはあり
    for j in training:
        for t in times:
            m += x[i][j][t] <= N_np[j-1][i-1]
for i in people:#いる人しか参加しない　いる人で参加しないのはあり
    for j in training:
        for t in times:
            m += x[i][j][t] <= A_np[t-1][i-1]
for i in people:#tコマで1人ができる練習は1つまで
    for t in times:
        m += pulp.lpSum(x[i][j][t] for j in training) <= 1
for t in times:#各期の参加人数はいる人でやる練習に参加可の人の合計よりK人少ない人数以上必要
    m += pulp.lpSum(x[i][j][t] for i in people for j in training) >= pulp.lpSum(N_np[j-1][i-1]*A_np[t-1][i-1]*y[j][t] for i in people for j in training) - K
m += pulp.lpSum(y[j][t] for j in training for t in times) == len(training)#入れた練習は全て採用する

m.solve()

print(pulp.LpStatus[m.solve()])
print("練習数は"+str(len(training)))


print("=========================================")
if pulp.LpStatus[m.solve()] != "Infeasible":
    print (pycolor.HIGHLIGHT +"練習は入りきった！"+ pycolor.END)
    print("総練習人数は"+str(round(pulp.value(m.objective)))+"人") 
#結果
    print("=========================================")
    print(pycolor.HIGHLIGHT +"【メニュースケジュール】"+ pycolor.END)
    t1 = 0
    j1 = 0
    for t in times:
        tt = 0
        for j in training:
            if pulp.value(y[j][t]) == 1:
                if t1 != t:
                    t1 = t
                if j1 != j:
                    j1 = j
                    tt += 1
                    if tt == 1:
                        print(pycolor.BLUE +timezone[t-1] + pycolor.END)
                    print(traindic[j]) 
    print("=========================================")
    print(pycolor.HIGHLIGHT +"【被り】"+ pycolor.END)
    for t in times:
        tt = 0
        for j in training:
            for i in people:
                if pulp.value(x[i][j][t]) != pulp.value(N_np[j-1][i-1]*A_np[t-1][i-1]*y[j][t]):
                    tt += 1
                    if tt == 1:
                        print(pycolor.BLUE +timezone[t-1] + pycolor.END)
                    print(namedic[i])
    print("=========================================")
    print(pycolor.HIGHLIGHT +"【やることない】"+ pycolor.END)
    for t in times:
        tt = 0
        for i in people:
            if A_np[t-1][i-1] == 1:
                #print(pulp.value(pulp.lpSum(x[i][j][t] for j in training)))
                if pulp.value(pulp.lpSum(x[i][j][t] for j in training)) == 0:
                    tt += 1
                    if tt == 1:
                        print(pycolor.BLUE +timezone[t-1] + pycolor.END)
                    print(namedic[i])
    
else:
    print(pycolor.HIGHLIGHT +"練習は入り切らなかった．被り人数許容上限="+str(K)+",同時練習許容上限="+str(W)+ pycolor.END)
    print("K(被り人数許容上限)やW(同時練習許容上限)を大きくするか，練習するメニュー減らしてみてね")