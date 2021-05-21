import tkinter as tk
from tkinter import filedialog
from tkinter import messagebox
import slackweb
import pulp
import pandas as pd

slack = slackweb.Slack(url="Paste Your Webhook here")

# 実行ボタンの関数
def Clac_Menu(file_path, date_id, overlap, place, bikou, slack_check):
    Needsheet = "3_need"
    overlap = int(overlap)
    place = int(place)
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
    arr_N =  df_N.drop(["メニュー名","d1","d2","d3","d4","d5","d6","d7"],axis=1).astype(int).values

    # Attend由来。出席データを０と１で表して配列化したもの。出席者名簿。A_np。
    arr_A = df_A.drop(df_A.columns[[0,1,2]],axis=1).astype(int).values

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
    # Slackへの送信
    if slack_check:
        # try:
        slack.notify(text=message)
        # except Forbidden:
        #     messagebox.showerror("エラー","web hookが無効です")
        return

# ファイル選択用の関数
def OpenFileDlgA(tbox):
    Attend_path.delete(0,tk.END)
    ftype =[("","*")]       #タプルのリスト
    dir = "."
    filename= filedialog.askopenfilename(filetypes=ftype,initialdir=dir)
    tbox.insert(0,filename)




root= tk.Tk()
root.title("練習組みシステム")
root.geometry("800x300")

k_initial = 3
w_initial = 3

y1=10
y2=70
y3=130
y4=190
y5=260

# Excel選択ラベル
label_1 = tk.Label(root,text="Excelファイル")
label_1.place(x=30,y=y1)
Attend_path = tk.Entry(root,width=35)
Attend_path.place(x=30,y=y1+20)
fdlg_button = tk.Button(root,text="ファイル選択",command=lambda:OpenFileDlgA(Attend_path))
fdlg_button.place(x=250,y=y1+20)

label_1 = tk.Label(root,text="日付ID")
label_1.place(x=30,y=y2)
Attend_sheet = tk.Entry(root,width=15)
Attend_sheet.place(x=30,y=y2+20)
Attend_sheet.insert(tk.END,"d1")

# label_2 = tk.Label(root,text="Need Excelファイル")
# label_2.place(x=30,y=y2)
# Need_path = tk.Entry(root,width=20)
# Need_path.place(x=30,y=y2+20)
# fdlg_button = tk.Button(root,text="ファイル選択",command=lambda:OpenFileDlgN(Need_path))
# fdlg_button.place(x=160,y=y2+20)

# label_1 = tk.Label(root,text="Need シート名")
# label_1.place(x=250,y=y2)
# Need_sheet = tk.Entry(root,width=15)
# Need_sheet.place(x=250,y=y2+20)
# Need_sheet.insert(tk.END,"d1")

# かぶり人数許容上限K＝３、同時練習許容上限W=５
label_K = tk.Label(root,text="かぶり人数許容上限K")
label_K.place(x=30,y=y3)
entry_K = tk.Entry(width=15)
entry_K.place(x=30,y=y3+20)
entry_K.insert(tk.END,k_initial)
label_W = tk.Label(root,text="同時練習許容上限W")
label_W.place(x=200,y=y3)
entry_W = tk.Entry(width=15)
entry_W.place(x=200,y=y3+20)
entry_W.insert(tk.END,w_initial)

# Slackチェックボタン
bln = tk.BooleanVar()
bln.set(False)
chk1 = tk.Checkbutton(root, variable=bln,text="Slackへ送信")
chk1.place(x=30,y=y5)

# 実行ボタン
calc_button = tk.Button(root,text="実行",command=lambda:Clac_Menu(Attend_path.get(), Attend_sheet.get(), entry_K.get(), entry_W.get(), entry_bikou.get(), bln.get()))
calc_button.place(x=150,y=y5)

# # リセットボタン2つ
# reset_button =tk.Button(root,text="リセット",command=lambda:ClearPart())
# reset_button.place(x=250,y=y5)
# def ClearPart():
#     bln.set(False)
#     text_kekka.delete('1.0',tk.END)
#     entry_K.delete(0,tk.END)
#     entry_W.delete(0,tk.END)
#     entry_bikou.delete(0,tk.END)
#     entry_K.insert(tk.END,k_initial)
#     entry_W.insert(tk.END,w_initial)

reset_button =tk.Button(root,text="リセット",command=lambda:ClearAll())
reset_button.place(x=650,y=y5)
def ClearAll():
    bln.set(False)
    text_kekka.delete('1.0',tk.END)
    entry_bikou.delete(0,tk.END)
    Attend_path.delete(0,tk.END)
    # Need_path.delete(0,tk.END)
    Attend_sheet.delete(0,tk.END)
    Attend_sheet.insert(tk.END,"d1")
    # Need_sheet.delete(0,tk.END)
    # Need_sheet.insert(tk.END,"d1")
    entry_K.delete(0,tk.END)
    entry_W.delete(0,tk.END)
    entry_K.insert(tk.END,k_initial)
    entry_W.insert(tk.END,w_initial)

# 閉じるボタン
close_button =tk.Button(root,text="閉じる",command=lambda:DoExit())
close_button.place(x=730,y=y5)
# 閉じるcallback
def DoExit():
    root.destroy()

# 備考欄エントリー
label_bikou = tk.Label(root,text="備考")
label_bikou.place(x=30,y=y4)
entry_bikou = tk.Entry(root,width=50)
entry_bikou.place(x=30,y=y4+20)

# 出力結果
label_1 = tk.Label(root,text="出力結果")
label_1.place(x=430,y=y1)
text_kekka = tk.Text(root,width=50,height=15)
text_kekka.place(x=430,y=y1+20)

# # コピーボタン
# copy_button =tk.Button(root,text="コビー",command=lambda:CopyText())
# copy_button.place(x=600,y=y5)
# def CopyText():
#     tk.clipboard_append(text_kekka)
# # うまくいかないので保留

root.mainloop()