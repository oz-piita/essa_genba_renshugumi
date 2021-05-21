#!/usr/bin/env python
# coding: utf-8

# In[1]:

'''
get_ipython().system('pip install pulp')
get_ipython().system('pip install --upgrade -q gspread')
 # -*- coding: utf-8 -*-
"""
Created on Fri Jul 12 12:01:39 2019
@author: ryo watanabe 11th 
※入力ファイルは CSV UTF-8(コンマ区切り)(*.csv) で保存すること
"""
# In[0]:準備
import pulp
import numpy as np
import pandas as pd
import time
class pycolor:
    BLUE = '\033[34m'
    HIGHLIGHT = '\033[01m'
    END = '\033[0m'
start = time.time()


# # **スプレッドシート入力**

# In[ ]:


from google.colab import auth
auth.authenticate_user()
import gspread
from oauth2client.client import GoogleCredentials
gc = gspread.authorize(GoogleCredentials.get_application_default())
#Attendスプレッドシートのキーをコピペ
worksheet_A = gc.open_by_key('1ADkGYGAdwSje21YB1Y32FfUz1PzVIqSWwZC78-E_Efs').worksheet('シート1') 
df_A = pd.DataFrame.from_records(worksheet_A.get_all_values())
df_A.columns = list(df_A.iloc[0,:])
df_A = df_A.drop(df_A.index[0]).reset_index(drop=True)
arr_A = df_A.iloc[2:,2:].replace("","0").values.astype(int)

#needスプレッドシートのキーをコピペ
worksheet_n = gc.open_by_key('1dscLNQZpXj1b-9tItO3v9C7KtZR1aKkIehkCTwNyuos').worksheet('シート1') 
df_N = pd.DataFrame.from_records(worksheet_n.get_all_values())
df_N.columns = list(df_N.iloc[0,:])
df_N = df_N.drop(df_N.index[0]).reset_index(drop=True)
arr_N = df_N.iloc[2:,3:].replace("","0").values.astype(int)
arr_N1 = df_N.iloc[2:,:].replace("","0").values


# # **メニュー組参考分析**

# In[ ]:


df_N_d = df_N.drop(df_N.columns[0], axis=1)
a = pd.DataFrame()
a["Unnamed: 0"] = df_N.iloc[:,0]
a["Unnamed: 1"] = df_N.iloc[:,1]
df_N_d = pd.concat([a, df_N_d], axis=1).replace("","0")


# In[ ]:


import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import matplotlib.cm as cm


#attendfile
attend = df_A
column_name1 = []
attend_dash = attend.drop(attend.columns[0],axis = 1)
attend_dash = attend_dash.drop(1)
#needfile
need = df_N_d
column_name = []
need_dash = need.drop(1)
need_dash = need_dash.drop(["Unnamed: 0","名前番号"],axis = 1)

menu_rate = need_dash.copy()#今後使う
menber = need.iloc[0,3:]#メンバーリスト
span = attend.iloc[2:,:2]#timeリスト


#  **時間帯別出席者数**

# In[ ]:


attend_num = attend_dash.iloc[1:,1:].astype(int).sum(axis=1).values
label=attend_dash.iloc[1:,0].values
bar_num = []
for i in range(len(label)):
  bar_num.append(i+1)
plt.figure(figsize=(10,8),dpi=50,facecolor="w")
plt.bar(bar_num, attend_num,color="#0078D7", width=0.4, tick_label=label, align="center")
plt.title("attend_num")
plt.show()


# **各時間の出席率と平均値(メニュー募集時の参考データ)**

# In[ ]:


"""
#menurate 加工
for time in range(len(attend_dash)-1):
  menu_rate.iloc[0,time+1] = attend_dash.iloc[time+1,0]
  for menu in range(len(need_dash)-1):
    menber_attend = 0
    menber_need = 0
    for menber in range(len(attend_dash.iloc[1,:])-1):
      menber_attend+=int(need_dash.iloc[menu+1,menber+1])*int(attend_dash.iloc[time+1,menber+1])
      menber_need += int(need_dash.iloc[menu+1,menber+1])
    menu_rate.iloc[menu+1,time+1] = (menber_attend/menber_need)*100

#本題
menu_rate = menu_rate.iloc[:,:len(attend_dash)]
columns = []
for i in range(len(menu_rate.iloc[0,:])):
    if i==0:
        columns.append("メニュー名")
    else:
        columns.append(menu_rate.iloc[0,i])
menu_rate.columns = columns
menu_rate=menu_rate.drop(0)
menu_rate["平均出席率"] = menu_rate.iloc[:,1:].mean(axis='columns')
menu_rate_alltime = menu_rate.sort_values(by='平均出席率', ascending=False)
menu_rate_alltime.round(1).reset_index(drop=True)
"""


# # **練習組最適化システム**

# 
# **パラメータ入力**

# In[ ]:


K = 3 #被り人数許容上限
W = 5 #同時練習許容上限


# 
# **準備**

# In[ ]:


name_dic = {}#メンバー名辞書
for i in range(len(df_N.iloc[0])-3):
    name_dic[i+1] = df_N.iloc[0][i+3]
train_dic = {}#練習名辞書
for i in range(len(df_N)-2):
    train_dic[i+1] = df_N.iloc[i+2][1]#
#j = やる練習 指定
train_num_list = []
for i in range(len(arr_N1)):
    if arr_N1[i][2] == "1":
        train_num_list.append(i+1)
#練習に必要なメンバー行列（arr_N)
col = []
for i in range(len(arr_N1[0])-3):
    col.append(i+3)
for i in range(len(arr_N)):
    for j in range(len(arr_N[0])):
        if arr_N[i,j] == -1:
            arr_N[i,j] = 0    
#i = member 指定
member_num_list = []
for i in range(len(arr_N[0])):
    member_num_list.append(i+1)
#ｔ = コマ 指定
koma_list = []
for i in range(len(arr_A)):
    koma_list.append(i+1)
#時間帯
timezone_list = df_A.iloc[2:,1].values
'''
# -*- coding: utf-8 -*-
"""
Created on Fri Jul 12 12:01:39 2019

@author: ryo
"""

from pulp import *
import pandas as pd
import time
import sys
sys.setrecursionlimit(2500)     # recursionの最大値が1000だと足りないみたいなので引き上げ
# 編集用。データフレームの表示フォーマットについて。
# pd.set_option('display.max_rows',10)
# pd.set_option('display.max_columns',None)

start = time.time()

Needfile = "C:\\Users\\tai\\Desktop\\renshugumi\\219n.csv"
Attendfile = "C:\\Users\\tai\\Desktop\\renshugumi\\219A.csv"
K = 2 #被り人数の上限数
W = 4 #同時にする練習数

df_N = pd.read_csv(Needfile, encoding = 'shift-jis', header=None).dropna(how="all").dropna(how="all",axis=1)

training = len(df_N.index)-3
member = len(df_N.columns)-3
name_dic = {}
train_dic = {}

# 全練習の数までの連番リストと、全メンバーのリストを獲得。
# それぞれ、りょーさんコードでtraining、peopleに対応。
train_num_list = list(range(1,training+1))
member_num_list = list(range(1,member+1))
# Need由来の２。メンバー名と番号の対応辞書。namedic。
for i in range(member):
    name_dic[i+1] = df_N.iloc[1][i+3]
# Need由来の３。トレイニング名と番号の辞書。traindic
for i in range(training):
    train_dic[i+1] = df_N.iloc[i+3][1]
# Need由来の４。参加者データをnd配列にしたもの。N_np。
arr_N =  df_N.fillna(0).drop(df_N.index[[0,1,2]]).drop(df_N.columns[[0,1,2]],axis=1).values
# Attendファイルの参照
df_A = pd.read_csv(Attendfile, encoding = 'shift-jis', header=None).dropna(how="all").dropna(how="all",axis=1)
# 練習時間のコマ番号をリストとして獲得。times。
koma_list = list(range(len(df_A.columns)-3))
# Attend由来。出席データを配列化したもの。A_np。
arr_A = df_A.fillna(0).drop(df_A.index[[0,1,2]]).drop(df_A.columns[[0,1]],axis=1).values
#コマ名リスト。timezone
timezone_list = df_A.iloc[3:,1].values
# ここまでデータ整理


# **最適化計算**

# In[ ]:


#関数
m = pulp.LpProblem("bestpractice", pulp.LpMaximize)
x = pulp.LpVariable.dicts('X', (member_num_list, train_num_list, koma_list), 0, 1, pulp.LpInteger)#tコマにiさんがj練習をするかどうか
y = pulp.LpVariable.dicts('Y', (train_num_list, koma_list), 0, 1, pulp.LpInteger)#tコマにj練習をするかどうか
m += pulp.lpSum(x[i][j][t] for i in member_num_list for j in train_num_list for t in koma_list ), "TotalPoint"
#制約
for t in koma_list:#練習はまとめて4つまで　t期にする練習全部足したらw以下に
    m += pulp.lpSum(y[j][t] for j in train_num_list) <= W              
for j in train_num_list:#同じ練習は1回まで　jの練習に対して全期分足したら1以下に    
    m += pulp.lpSum(y[j][t] for t in koma_list) <= 1            
for i in member_num_list:#やる練習にしか参加できない　参加しないのはあり    
    for j in train_num_list:
        for t in koma_list:
            m += x[i][j][t] <= y[j][t]
for i in member_num_list:#必要な練習しかしない　必要な練習でもやらないのはあり
    for j in train_num_list:
        for t in koma_list:
            m += x[i][j][t] <= arr_N[j-1][i-1]
for i in member_num_list:#いる人しか参加しない　いる人で参加しないのはあり
    for j in train_num_list:
        for t in koma_list:
            m += x[i][j][t] <= arr_A[t-1][i-1]
for i in member_num_list:#tコマで1人ができる練習は1つまで
    for t in koma_list:
        m += pulp.lpSum(x[i][j][t] for j in train_num_list) <= 1
for t in koma_list:#各期の参加人数はいる人でやる練習に参加可の人の合計よりK人少ない人数以上必要
    m += pulp.lpSum(x[i][j][t] for i in member_num_list for j in train_num_list) >= pulp.lpSum(arr_N[j-1][i-1]*arr_A[t-1][i-1]*y[j][t] for i in member_num_list for j in train_num_list) - K
m += pulp.lpSum(y[j][t] for j in train_num_list for t in koma_list) == len(train_num_list)#入れた練習は全て採用する

m.solve()

print(pulp.LpStatus[m.solve()])
print("練習数は"+str(len(train_num_list)))


# **結果**

# In[ ]:


print("=========================================")
if pulp.LpStatus[m.solve()] != "Infeasible":
    print ("練習は入りきった！")
    print("総練習人数は"+str(round(pulp.value(m.objective)))+"人") 
#結果
    print("=========================================")
    print("【メニュースケジュール】")
    t1 = 0
    j1 = 0
    for t in koma_list:
        tt = 0
        for j in train_num_list:
            if pulp.value(y[j][t]) == 1:
                if t1 != t:
                    t1 = t
                if j1 != j:
                    j1 = j
                    tt += 1
                    if tt == 1:
                        print(timezone_list[t-1] )
                    print(train_dic[j]) 
    print("=========================================")
    print("【被り】")
    for t in koma_list:
        tt = 0
        for j in train_num_list:
            for i in member_num_list:
                if pulp.value(x[i][j][t]) != pulp.value(arr_N[j-1][i-1]*arr_A[t-1][i-1]*y[j][t]):
                    tt += 1
                    if tt == 1:
                        print(timezone_list[t-1] )
                    print(name_dic[i])
    print("=========================================")
    print("【やることない】")
    for t in koma_list:
        tt = 0
        for i in member_num_list:
            if arr_A[t-1][i-1] == 1:
                #print(pulp.value(pulp.lpSum(x[i][j][t] for j in train_num_list)))
                if pulp.value(pulp.lpSum(x[i][j][t] for j in train_num_list)) == 0:
                    tt += 1
                    if tt == 1:
                        print(timezone_list[t-1] )
                    print(name_dic[i])
    
else:
    print("練習は入り切らなかった．被り人数許容上限="+str(K)+",同時練習許容上限="+str(W))
    print("K(被り人数許容上限)やW(同時練習許容上限)を大きくするか，練習するメニュー減らしてみてね")

