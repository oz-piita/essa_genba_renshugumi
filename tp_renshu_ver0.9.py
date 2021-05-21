# -*- coding: utf-8 -*-
"""
Created on Fri Jul 12 12:01:39 2019

@author: ryo
"""
#import numpy as np
import pulp
import pandas as pd
import time
import slackweb

slack = slackweb.Slack(url="webhookURL")
start = time.time()


Attendfile = "C:\\Users\\tai\\Desktop\\Renshugumi\\出欠調査(12_17-23).xlsx"
# Attendfile = "C:\\Users\\tai\\Desktop\\renshugumi\\219A.csv"
Needfile = "C:\\Users\\tai\\Desktop\\Renshugumi\\香盤表(卒公18).xlsx"
# Needfile = "C:\\Users\\tai\\Desktop\\renshugumi\\219n.csv"
Needsheet = "need"
Attendsheet = "d1"
K = 40                                     #被り人数の上限数
W = 5                                     #同時にする練習数
bikou = "やっほー！(備考欄)"   

if Needfile.find("xls")>1:
    df_N = pd.read_excel(Needfile,sheet_name=Needsheet,  header=None).dropna(how="all").dropna(how="all",axis=1).fillna(0)
else:
    df_N = pd.read_csv(Needfile, encoding = 'shift-jis', header=None).dropna(how="all").dropna(how="all",axis=1).fillna(0)

if Attendfile.find("xls")>1:
    df_A = pd.read_excel(Attendfile,sheet_name=Attendsheet, header=None).dropna(how="all").dropna(how="all",axis=1).fillna(0)
else:
    df_A = pd.read_csv(Attendfile, encoding = 'shift-jis', header=None).dropna(how="all").dropna(how="all",axis=1)


training = len(df_N.index)-3               # 全練習メニュー数の定数のつもり
member = len(df_N.columns)-3               # 全人数
koma_list = list(range(1,len(df_A.index)-2))    # 練習時間のコマ番号をリストとして獲得。times。
member_num_list = list(range(1,member+1))  # 全メンバーのリストを獲得。りょーさんコードでpeopleに対応。
name_dic = {}                              # Need由来の２。メンバー名と番号の対応辞書。namedic。
for i in range(member):
    name_dic[i+1] = df_N.iloc[1][i+3]
train_dic = {}                             # Need由来の３。トレーニング名と番号の辞書。traindic
for i in range(training):
    train_dic[i+1] = df_N.iloc[i+3][1]


train_num_list = []                        # Need由来の5。時間を割きたい練習のみの番号リスト。トリッキー。training。
for i in range(training):
    if int(df_N[2][i+3]) == 1:
        train_num_list.append(i+1)



timezone_list = df_A.iloc[3:3+len(koma_list),1].astype(str).values          #コマ名リスト。timezone。EXCEL入力時のみDatetimeになって秒が出るので苦肉の策
for n in range(len(timezone_list)):
    if len(timezone_list[n])==8 and timezone_list[n][-3:]==":00":
        timezone_list[n] = timezone_list[n][:5]

# Need由来の４。参加者データを0と1で表現してした配列。履修者名簿。N_np。
arr_N =  df_N.drop(df_N.index[[0,1,2]]).drop(df_N.columns[[0,1,2]],axis=1).astype("int").values
# Attend由来。出席データを０と１で表して配列化したもの。出席者名簿。A_np。
arr_A = df_A.fillna(0).drop(df_A.index[[0,1,2]]).drop(df_A.columns[[0,1]],axis=1).astype("int").values

# ここまでデータ整理

print(timezone_list)

# ここからPulpによる最適化
m = pulp.LpProblem("bestpractice", pulp.LpMaximize)
x = pulp.LpVariable.dicts('X', (member_num_list, train_num_list, koma_list), 0, 1, pulp.LpInteger)#tコマにiさんがj練習をするかどうか
y = pulp.LpVariable.dicts('Y', (train_num_list, koma_list), 0, 1, pulp.LpInteger)#tコマにj練習をするかどうか
m += pulp.lpSum(x[i][j][t] for i in member_num_list for j in train_num_list for t in koma_list ), "TotalPoint"
#制約
for t in koma_list:             #「１」練習はまとめて4つまで　t期にする練習全部足したらw以下に
    m += pulp.lpSum(y[j][t] for j in train_num_list) <= W              
for j in train_num_list:        #「２」同じ練習は1回まで　jの練習に対して全期分足したら1以下に    
    m += pulp.lpSum(y[j][t] for t in koma_list) <= 1            
for i in member_num_list:       #「３」やる練習にしか参加できない　参加しないのはあり    
    for j in train_num_list:
        for t in koma_list:
            m += x[i][j][t] <= y[j][t]
for i in member_num_list:       #「４」必要な練習しかしない　必要な練習でもやらないのはあり
    for j in train_num_list:
        for t in koma_list:
            m += x[i][j][t] <= arr_N[j-1][i-1]
for i in member_num_list:       #「５」いる人しか参加しない　いる人で参加しないのはあり
    for j in train_num_list:
        for t in koma_list:
            m += x[i][j][t] <= arr_A[t-1][i-1]
for i in member_num_list:       #「６」tコマで1人ができる練習は1つまで
    for t in koma_list:
        m += pulp.lpSum(x[i][j][t] for j in train_num_list) <= 1
for t in koma_list:             #「７」各期の参加人数はいる人でやる練習に参加可の人の合計よりK人少ない人数以上必要
    m += pulp.lpSum(x[i][j][t] for i in member_num_list for j in train_num_list) >= pulp.lpSum(arr_N[j-1][i-1]*arr_A[t-1][i-1]*y[j][t] for i in member_num_list for j in train_num_list) - K
m += pulp.lpSum(y[j][t] for j in train_num_list for t in koma_list) == len(train_num_list)#入れた練習は全て採用する

m.solve()
msg_list = []
msg_list.append(str(df_A.iloc[0][0]))
# msg_list.append(str(pulp.LpStatus[m.solve()]))
msg_list.append(str("練習数は"+str(len(train_num_list))))


# **結果**

# In[ ]:


msg_list.append(str("============================="))
if pulp.LpStatus[m.solve()] != "Infeasible":
    msg_list.append(str ("練習は入りきった！"))
    msg_list.append(str("総練習人数は"+str(round(pulp.value(m.objective)))+"人") )
#結果
    msg_list.append(str("============================="))
    msg_list.append(str("【メニュースケジュール】"))
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
                        msg_list.append("\n"+str(timezone_list[t-1]))
                    msg_list.append(str(train_dic[j]))
    msg_list.append(str("============================="))
    msg_list.append(str("【詳細】"))
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
                        msg_list.append("\n"+str(timezone_list[t-1]))
                    msg_list.append(str(train_dic[j]))
        msg_list.append(str("【被り】"))
        for j in train_num_list:
            for i in member_num_list:
                if pulp.value(x[i][j][t] ) != pulp.value(arr_N[j-1][i-1]*arr_A[t-1][i-1]*y[j][t]):
                    msg_list.append(str(name_dic[i]))
        msg_list.append(str("【やることない】"))
        for i in member_num_list:
            if arr_A[t-1][i-1] == 1:
                #msg_list.append(tr(pulp.value(pulp.lpSum(x[i][j][t] for j in train_num_list)))
                if pulp.value(pulp.lpSum(x[i][j][t] for j in train_num_list)) == 0:
                    msg_list.append(str(name_dic[i]))

    # msg_list.append(str("============================="))
    # msg_list.append(str("【被り】"))
    # for t in koma_list:
    #     tt = 0
    #     for j in train_num_list:
    #         for i in member_num_list:
    #             if pulp.value(x[i][j][t]) != pulp.value(arr_N[j-1][i-1]*arr_A[t-1][i-1]*y[j][t]):
    #                 tt += 1
    #                 if tt == 1:
    #                     msg_list.append(str("\n"+timezone_list[t-1] ))
    #                 msg_list.append(str(name_dic[i]))
    # msg_list.append(str("============================="))
    # msg_list.append(str("【やることない】"))
    # for t in koma_list:
    #     tt = 0
    #     for i in member_num_list:
    #         if arr_A[t-1][i-1] == 1:
    #             #msg_list.append(tr(pulp.value(pulp.lpSum(x[i][j][t] for j in train_num_list)))
    #             if pulp.value(pulp.lpSum(x[i][j][t] for j in train_num_list)) == 0:
    #                 tt += 1
    #                 if tt == 1:
    #                     msg_list.append(str("\n"+timezone_list[t-1]))
    #                 msg_list.append(str(name_dic[i]))
                    
else:
    msg_list.append(str("練習は入り切らなかった．被り人数許容上限="+str(K)+",同時練習許容上限="+str(W)))
    msg_list.append(str("K(被り人数許容上限)やW(同時練習許容上限)を大きくするか，練習するメニュー減らしてみてね"))
# 出力メッセージの結合
msg=str()
for i in msg_list:
    msg = msg + "\n" + i
print(bikou+msg)
# Slackへの送信
slack.notify(text=bikou+msg)
