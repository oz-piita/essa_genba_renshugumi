import tkinter as tk
from tkinter import filedialog
from tkinter import messagebox
import plugins.my_input as ip
import plugins.my_calc as cl

# 実行ボタンの関数
def Clac_Menu(file_path, date_id, overlap, place, bikou):
    # 結果欄をクリア
    text_kekka.delete('1.0',tk.END)

    # エクセルから3回に分けてデータを呼び出し
    week_id, date, participant, lesson_minuete= ip.read_05_parameter_sheet(file_path, date_id)
    members_name_list, start_time_list, menu_name_list, need_menus, arr_N = ip.read_df_up_sheet(file_path, date_id)
    arr_A = ip.read_Form_Answer_sheet(file_path, week_id, date_id, members_name_list, start_time_list, lesson_minuete)
    # date                  日付と場所str。補足情報として出力に回す。
    # participant           参加者名簿str。補足情報として出力に回す。
    # week_id, date_id      入力の参照用。w05d06のように日付にIDを振っている。
    # lesson_minuete        1コマあたりの分数。参加者の出欠回答をコマ換算するために参照する。
    # ほかの変数はplugins/my_calcを確認のこと。

    # 計算用のパラメタをオブジェクト化
    rm = cl.Param(members_name_list, start_time_list, menu_name_list, need_menus, arr_N, arr_A, overlap, place)

    # 一回目の計算
    message, failed = rm.Calc(date, participant)

    # 2～5回の再計算。1回目が計算不可の場合に優先度の低いメニューを一つ削って再計算する
    cnt = 0                 # 繰り返し計算カウント
    limit = 5               # 繰り返し計算回数の上限
    rejected_menus = []     # 計算から爪弾きにしたメニュー名のインデックス
    if failed:
        while failed and cnt < limit:
            cnt += 1
            mx = max(rm.need_menus)
            for i, v in enumerate(rm.need_menus):
                if v == mx:             # need_menusの最大値（優先順位最大5から下）を0に置き換える
                    rm.need_menus[i] = 0
                    rejected_menus.append(i)
                    break
            message, failed = rm.Calc(date, participant)  # 再計算

    # 再計算のログを追加する
    if cnt == 0:
        message += "\n練習は入りきりました.　\nメニュー棄却（再計算）回数：0\n"
    elif 0 < cnt and not(failed):
        message += "\nメニュー棄却（再計算）回数："+str(cnt) + "\n"
        message += str(cnt) + "つの練習は入り切りませんでした．被り人数許容上限=" + \
            str(rm.overlap)+",同時練習許容上限="+str(rm.place)+"\n"
        message += "やり直すならばK(被り人数許容上限)やW(同時練習許容上限)を大きくするか，練習するメニューを減らしてください"
        for i in range(cnt):
            nn = rejected_menus[i]
            message += "\n棄却メニュー" + str(i+1)+"：" + rm.lessons[nn]
        message /= "\n"
    elif cnt == limit and failed:
        message += "\nメニュー棄却（再計算）回数："+str(cnt)
        message += "\n計算回数の上限に達した."
        for i in range(cnt):
            nn = rejected_menus[i]
            message += "\n棄却メニュー" + str(i+1)+"：" + rm.lessons[nn]
        message = "\n"
    else:
        message += "\nバグっています.システム担当者に連絡してください.\n"
    
    message += "\n" + bikou

    text_kekka.insert(tk.END, message)
    return

# ファイル選択用の関数
def OpenFileDlgA(tbox):
    Attend_path.delete(0,tk.END)
    ftype =[("","*")]       #タプルのリスト
    dir = "."
    filename= filedialog.askopenfilename(filetypes=ftype,initialdir=dir)
    tbox.insert(0,filename)

# 以下GUIの静的な部分

root= tk.Tk()
root.title("練習組みシステム")
root.geometry("800x300")

# ボックスのデフォルト値
k_initial = 3
w_initial = 3

# ボックスの表示位置
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
# bln = tk.BooleanVar()
# bln.set(False)
# chk1 = tk.Checkbutton(root, variable=bln,text="Slackへ送信")
# chk1.place(x=30,y=y5)

# 実行ボタン
calc_button = tk.Button(root,text="実行",command=lambda:Clac_Menu(Attend_path.get(), Attend_sheet.get(), entry_K.get(), entry_W.get(), entry_bikou.get()))# bln.get()))
calc_button.place(x=150,y=y5)

# リセットボタン
reset_button =tk.Button(root,text="リセット",command=lambda:ClearAll())
reset_button.place(x=650,y=y5)
def ClearAll():
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
    # bln.set(False)

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

root.mainloop()
