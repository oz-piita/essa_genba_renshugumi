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

slack = slackweb.Slack(url="https://hooks.slack.com/services/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")
start = time.time()

date_id="d1"
file_path = "C:\\Users\\tai\\Desktop\\references\\essa_em_dbtest.xlsx"

Needsheet = "3_need"
Attendsheet = "d1"
overlap = 40                                     #被り人数の上限数
place = 5                                     #同時にする練習数
bikou = "やっほー！"   

df_N = pd.read_excel(file_path,sheet_name=Needsheet).dropna(how="all").fillna(int(0))
df_A = pd.read_excel(file_path,sheet_name=date_id).dropna(how="all").dropna(how="all",axis=1).fillna(int(0))

training = len(df_N.index)               # 全練習メニュー数の定数のつもり
member = len(df_A.columns)-3               # 全人数

df_N = df_N.iloc[:,:member+8]        # d2などが空欄だとdropnaで消し飛ばされてずれるので調整

koma_list = list(range(1,len(df_A.index)+1))    # 練習時間のコマ番号をリストとして獲得。times。

member_num_list = list(range(1,member+1))  # 全メンバーのリストを獲得。りょーさんコードでpeopleに対応。

name_dic = {}                              # メンバー名と番号の対応辞書。namedic。
for i in range(member):
    name_dic[i+1] = df_A.columns[i+3]

train_dic = {}                             # トレーニング名と番号の辞書。traindic
for i in range(training):
    train_dic[i+1] = df_N.iloc[i][0]

train_num_list = []                        # Need由来の5。時間を割きたい練習のみの番号リスト。トリッキー。training。
for i in range(training):
    if df_N.iat[i, df_N.columns.get_loc(date_id)] == 1:
        train_num_list.append(i+1)

timezone_list = df_A.iloc[0:len(koma_list),2].astype(str).values          #コマ名リスト。timezone。EXCELの時刻シリアルがDatetimeになって秒が出るので苦肉の策
for n in range(len(timezone_list)):
    if len(timezone_list[n])==8 and timezone_list[n][-3:]==":00":
        timezone_list[n] = timezone_list[n][:5]

# Need由来の４。参加者データを0と1で表現してした配列。履修者名簿。N_np。
arr_N =  df_N.drop(["メニュー名","d1","d2","d3","d4","d5","d6","d7"],axis=1).astype("int").values

# Attend由来。出席データを０と１で表して配列化したもの。出席者名簿。A_np。
arr_A = df_A.drop(df_A.columns[[0,1,2]],axis=1).astype("int").values

# ここまでデータ整理


# ここからPulpによる最適化
m = pulp.LpProblem("bestpractice", pulp.LpMaximize)
x = pulp.LpVariable.dicts('X', (member_num_list, train_num_list, koma_list), 0, 1, pulp.LpInteger)#tコマにiさんがj練習をするかどうか
y = pulp.LpVariable.dicts('Y', (train_num_list, koma_list), 0, 1, pulp.LpInteger)#tコマにj練習をするかどうか
m += pulp.lpSum(x[i][j][t] for i in member_num_list for j in train_num_list for t in koma_list ), "TotalPoint"
#制約
for t in koma_list:             #「１」練習はまとめて4つまで　t期にする練習全部足したらw以下に
    m += pulp.lpSum(y[j][t] for j in train_num_list) <= place              
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
    m += pulp.lpSum(x[i][j][t] for i in member_num_list for j in train_num_list) >= pulp.lpSum(arr_N[j-1][i-1]*arr_A[t-1][i-1]*y[j][t] for i in member_num_list for j in train_num_list) - overlap
m += pulp.lpSum(y[j][t] for j in train_num_list for t in koma_list) == len(train_num_list)#入れた練習は全て採用する

m.solve()

# **出力文**

msg_list = []
msg_list.append(str(df_A.iat[0,0])+"\n")
# msg_list.append(str(pulp.LpStatus[m.solve()]))
msg_list.append("練習数は"+str(len(train_num_list))+"\n")

msg_list.append("=============処理結果=============\n")
if pulp.LpStatus[m.solve()] != "Infeasible":
    msg_list.append("練習は入りきった！\n")
    msg_list.append("総練習人数は"+str(round(pulp.value(m.objective)))+"人"+"\n")
#結果
    msg_list.append("\n=============メニュー=============\n")
    msg_list.append(str(df_A.iat[0,0])+"\n参加者\n"+str(df_A.iat[0,1])+"\n")
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
                        msg_list.append("\n"+str(timezone_list[t-1])+"\n")
                    msg_list.append(str(train_dic[j])+"\n")
    msg_list.append("\n\n==============詳細==============")
    # msg_list.append("【詳細】")
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
                        msg_list.append("\n\n"+str(timezone_list[t-1])+"\n")
                    msg_list.append(str(train_dic[j]))
        msg_list.append("\n【被り】\n")
        for j in train_num_list:
            for i in member_num_list:
                if pulp.value(x[i][j][t] ) != pulp.value(arr_N[j-1][i-1]*arr_A[t-1][i-1]*y[j][t]):
                    msg_list.append(str(name_dic[i])+"　")
        msg_list.append("\n【やることない】\n")
        for i in member_num_list:
            if arr_A[t-1][i-1] == 1:
                #msg_list.append(tr(pulp.value(pulp.lpSum(x[i][j][t] for j in train_num_list)))
                if pulp.value(pulp.lpSum(x[i][j][t] for j in train_num_list)) == 0:
                    msg_list.append(str(name_dic[i])+"　")

else:
    msg_list.append("練習は入り切らなかった．被り人数許容上限="+str(overlap)+",同時練習許容上限="+str(place)+"\n")
    msg_list.append("K(被り人数許容上限)やW(同時練習許容上限)を大きくするか，練習するメニュー減らしてみてね")
# 出力メッセージの結合表示
msg=str()
for i in msg_list:
    msg = msg  + i
message = bikou+"\n" + msg
text_kekka.insert(tk.END,message)
Slackへの送信
if slack_check:
    try:
        slack.notify(text=message)
    except Forbidden:
        messagebox.showerror("エラー","web hookが無効です")
    return
print(message)
