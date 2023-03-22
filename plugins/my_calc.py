'''bestpractice組合せ最適化'''

import pulp

class Param:
    def __init__(self,members_name_list,start_time_list,menu_name_list,need_menus,arr_N,arr_A,overlap,place):
        self.members_name_list = members_name_list          # 全現場関係者のstrリスト。
        self.start_time_list   = start_time_list            # その日のコマの開始時刻のdatetime.timeリスト。
        self.menu_name_list    = menu_name_list             # 全メニュー名のstrリスト。
        self.need_menus = need_menus                        # 要求メニューのintリスト。要求されるメニューのフィルタ。
        self.arr_N      = arr_N         # 香盤表の2次元numpy配列。indexはメニュー名（menu_name_list）、columnは人(members_name_list)。
        self.arr_A      = arr_A         # 出席簿の2次元numpy配列。indexはコマ（start_time_list）、columunは人(members_name_list)。
                                        # ※探索するのが start_time_list × menu_name_list配列のイメージ。
        self.overlap    = int(overlap)       # 制約式用のかぶり人数上限int。
        self.place      = int(place)         # 制約式用の同時練習数int。

    def Calc(self, datedata, participant):
        # 前処理 

        member_num_list = list(range(1,len(self.members_name_list)+1))     # メンバーのキー番号リスト。Pulpは1から参照することが多い。
        class_num_list = list(range(1,len(self.start_time_list)+1)) # コマのインデックスのリスト。1~
        menu_num_list = []                        # 時間を割きたい練習のフィルタリング。ナップザック問題の石
        for i in range(len(self.menu_name_list)):
            if self.need_menus[i] != 0:
                menu_num_list.append(i+1)

        # 最適化
        
        m = pulp.LpProblem("bestpractice", pulp.LpMaximize)
        x = pulp.LpVariable.dicts('X', (member_num_list, menu_num_list, class_num_list), 0, 1, pulp.LpInteger)  #tコマにiさんがj練習をするかどうか
        y = pulp.LpVariable.dicts('Y', (menu_num_list, class_num_list), 0, 1, pulp.LpInteger)                   #tコマにj練習をするかどうか
        m += pulp.lpSum(x[i][j][t] for i in member_num_list for j in menu_num_list for t in class_num_list ), "TotalPoint"
        #制約
        for t in class_num_list:            #「１」練習はまとめて4つまで　t期にする練習全部足したらw以下に
            m += pulp.lpSum(y[j][t] for j in menu_num_list) <= self.place              
        for j in menu_num_list:            #「２」同じ練習は1回まで　jの練習に対して全期分足したら1以下に    
            m += pulp.lpSum(y[j][t] for t in class_num_list) <= 1            
        for i in member_num_list:           #「３」やる練習にしか参加できない　参加しないのはあり    
            for j in menu_num_list:
                for t in class_num_list:
                    m += x[i][j][t] <= y[j][t]
        for i in member_num_list:           #「４」必要な練習しかしない　必要な練習でもやらないのはあり
            for j in menu_num_list:
                for t in class_num_list:
                    m += x[i][j][t] <= self.arr_N[j-1][i-1]
        for i in member_num_list:           #「５」いる人しか参加しない　いる人で参加しないのはあり
            for j in menu_num_list:
                for t in class_num_list:
                    m += x[i][j][t] <= self.arr_A[t-1][i-1]
        for i in member_num_list:           #「６」tコマで1人ができる練習は1つまで
            for t in class_num_list:
                m += pulp.lpSum(x[i][j][t] for j in menu_num_list) <= 1
        for t in class_num_list:            #「７」各期の参加人数はいる人でやる練習に参加可の人の合計よりK人少ない人数以上必要
            m += pulp.lpSum(x[i][j][t] for i in member_num_list for j in menu_num_list) >= pulp.lpSum(self.arr_N[j-1][i-1]*self.arr_A[t-1][i-1]*y[j][t] for i in member_num_list for j in menu_num_list) - self.overlap
        m += pulp.lpSum(y[j][t] for j in menu_num_list for t in class_num_list) == len(menu_num_list)#入れた練習は全て採用する
        # 演算
        m.solve()

        # **出力文**

        msg = str()
        msg += datedata + "\n"
        msg += ("入った練習数は"+str(len(menu_num_list))+"\n")
        msg += ("===========処理結果===========\n")
        infeasible = (pulp.LpStatus[m.solve()] == "Infeasible")
        if infeasible:    # Infeasibleは実行不可能の意。パラメタ調整のため抜ける
            msg += ("練習は入り切りませんでした．被り人数許容上限="+str(self.overlap)+",同時練習許容上限="+str(self.place)+"\n")
            msg += ("K(被り人数許容上限)やW(同時練習許容上限)を大きくするか，練習するメニュー減らしてみてください")
            return msg,infeasible
           
        
        msg += ("練習は入りきりました！\n")
        msg += ("総練習人数は"+str(round(pulp.value(m.objective)))+"人"+"\n\n")
        msg += ("===========メニュー===========\n")
        msg += (datedata+"\n"+ "参加者\n"+participant+"\n")
        t1 = 0
        j1 = 0
        for t in class_num_list:
            msg += "\n"
            tt = 0
            for j in menu_num_list:
                if pulp.value(y[j][t]) == 1:
                    if t1 != t:
                        t1 = t
                    if j1 != j:
                        j1 = j
                        tt += 1
                        if tt == 1:
                            msg += (self.start_time_list[t-1].strftime('%H:%M') +"\n")          # .time->strキャスト
                        msg += (self.menu_name_list[j-1]+"\n")
        
        msg +="\n" + ("============詳細============\n")
        t1 = 0
        j1 = 0
        for t in class_num_list:
            tt = 0
            for j in menu_num_list:
                if pulp.value(y[j][t]) == 1:
                    if t1 != t:
                        t1 = t
                    if j1 != j:
                        j1 = j
                        tt += 1
                        if tt == 1:
                            msg += (self.start_time_list[t-1].strftime('%H:%M')+"\n")
                        msg += (self.menu_name_list[j-1]+"\n")
            msg += ("【被り】\n")
            for j in menu_num_list:
                for i in member_num_list:
                    if pulp.value(x[i][j][t] ) != pulp.value(self.arr_N[j-1][i-1]*self.arr_A[t-1][i-1]*y[j][t]):
                        msg += (self.members_name_list[i-1]+"　")
            msg += ("\n")
            msg += ("【やることない】\n")
            for i in member_num_list:
                if self.arr_A[t-1][i-1] == 1:
                    if pulp.value(pulp.lpSum(x[i][j][t] for j in menu_num_list)) == 0:
                        msg += (self.members_name_list[i-1]+"　")
            msg += ("\n\n")

        return msg,infeasible
