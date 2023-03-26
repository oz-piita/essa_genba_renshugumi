'''Excelからのインプット'''
import pandas as pd
import datetime

date_ids = ["d1", "d2", "d3", "d4", "d5", "d6", "d7"]

# エクセルの読み込み1
def read_05_parameter_sheet(file_path, date_id):
    print("Pandas version: ", pd.__version__)

    df_05 = pd.read_excel(file_path, sheet_name = "05_parameter", header=None)

    week_id = df_05.iat[0, 1]
    date_index = date_ids.index(date_id)
    date = df_05.iat[2, 1 + date_index]
    participant = df_05.iat[4, 1 + date_index]
    lesson_minuete = df_05.iat[12, 1 + date_index]
    
    print(week_id, date, participant, lesson_minuete, "min/class")
    return week_id, date, participant, lesson_minuete

# エクセルの読み込み2
def read_df_up_sheet(file_path, date_id):
    df_origin = pd.read_excel(file_path,sheet_name="df_up",header=None)

    date_index = date_ids.index(date_id)
    # 生のデータからd1(～7)に関するデータ整理のための整形
    # 0全練習　1全名前　2＋index(2)コマ名　9＋index各(3)パラミタ　16+index(4)need
    df_up = df_origin[[0, 1, 2+date_index, 9+date_index, 16+date_index]]
    df_up = df_up.rename(
        columns={2+date_index: 2, 9+date_index: 3, 16+date_index: 4})
    df_up = df_up.drop(df_up.index[[0]])
    
    # シートに並べた各パラメータを拾う
    menu_name_list = df_up.iloc[:, 0].dropna(how='all').tolist()
    members_name_list = df_up.iloc[:, 1].dropna(how='all').tolist()
    start_time_list = df_up[2].dropna(how='all').tolist()

    need_menus = []                         # 時間を割きたい練習のみの番号リスト。
    for i in range(len(menu_name_list)):
        need_menus.append(int(df_up.fillna(0).iat[i, 4]))

    # 香盤表配列arr_N
    df_need = df_origin.drop(df_origin.columns[[range(23)]], axis=1)
    df_need = df_need.drop(df_need.index[[0]])
    arr_N = df_need.drop(df_need.index[range(len(menu_name_list), len(df_need.index))]).fillna(0).astype("int").values
    # lesson_minuete = int(df_up.iat[0, 3])   # 1コマ当たりの時間（分）.データ整理用
    # overlap = int(df_up.iat[1, 3])   # かぶり許容人数.
    # place = int(df_up.iat[2, 3])   # 同時練習数
    # datedata = df_up.iat[3, 3]   # 日付と場所情報.出力用
    # participant = df_up.iat[4, 3]   # 参加者列挙.出力用

             # 軸１            　　軸２　　     　　　軸３　      　　軸３’ 　表１vs3　表2vs1　　　o1        o2        p1      p2
    return members_name_list, start_time_list, menu_name_list, need_menus, arr_N


# エクセルの読み込み3
# フォーム回答シートを読み込み、出席簿配列arr_Aをまとめる
def read_Form_Answer_sheet(file_path, week_id, date_id, members_name_list, start_time_list, lesson_minuete):
    df_form_ans = pd.read_excel(file_path, sheet_name = week_id + "_", header=None).dropna(how="all")

    # 該当日のみを残す
    df_form_ans = cut_off_all_but_target_date(df_form_ans, date_id)
    
    # タイムスタンプが最新のものから30日より前のものは棄却
    last_ans = df_form_ans.at[len(df_form_ans.index)-1, "timestamp"]
    df_form_ans = df_form_ans.set_index("timestamp")[last_ans - datetime.timedelta(days=30):last_ans]
    
    # 重複した古いデータを削除し、名前でsortする
    df_form_ans = df_form_ans.drop_duplicates(subset="name", keep="last")
    df_mem = pd.DataFrame(members_name_list).dropna(how="all").rename(columns={0:"name"})
    df_form_ans = pd.merge(df_mem, df_form_ans, on="name", how="left")

    # 未回答者を欠席扱いにする
    df_form_ans["At/Ab"] = df_form_ans["At/Ab"].fillna("欠席 or 未定")

    # 参加早退の空欄を参加→00:00と早退→23:59:00と埋める。ただし欠席は参加時刻＝早退時刻＝00:00とする。
    df_form_ans =  fill_join_and_move_out_time(df_form_ans)


    # 練習時刻と出欠回答結果をすり合わせ、メンバーの出席簿array_Attendanceをまとめる
    arr_A = create_array_Attendance(df_form_ans, start_time_list, lesson_minuete)

    return arr_A

def cut_off_all_but_target_date(df_form_ans, date_id):
    date_index = date_ids.index(date_id)
    # date_idに従った日付だけを抽出する
    df_form_ans = df_form_ans[[0, 1, 3*date_index+2, 3*date_index+3, 3*date_index+4]]
    dic_col_form_ans = {0:"timestamp", 1:"name", 3*date_index+2:"At/Ab", 3*date_index+3:"join", 3*date_index+4:"move_out"}
    #                       0               1                       2                       3                         4
    df_form_ans = df_form_ans.rename(columns=dic_col_form_ans)
    # 「タイムスタンプ、名前」などの1行目を切り、番号を0はじまりに直す
    df_form_ans = df_form_ans.drop(df_form_ans.index[[0]])
    df_form_ans = df_form_ans.set_axis(list(range(0, len(df_form_ans.index))), axis=0)
    return df_form_ans

def fill_join_and_move_out_time(df_form_ans):
    zero_oclock = datetime.time(hour=0, minute=0)
    end_of_day = datetime.time(hour=23, minute=59)

    df_form_ans = df_form_ans.fillna(zero_oclock)
    for i in range(len(df_form_ans.index)):
        if df_form_ans.at[i, "At/Ab"] != "欠席 or 未定" and df_form_ans.at[i, "move_out"] == zero_oclock:
            # 出席　あるいは　事情アリかつ　、早退時刻が空欄なら終日参加に上書き
            df_form_ans.at[i, "move_out"] = end_of_day
    return df_form_ans

def create_array_Attendance(df_form_ans, start_time_list, lesson_minuete):
    total_lesson = len(start_time_list)
    df_attend = pd.DataFrame(columns=range(len(df_form_ans.index)), index=range(total_lesson))
    for i in range(total_lesson):
        for j in range(len(df_form_ans.index)):
            # timedeltaで計算するために.time型から.datetime型にキャストする必要がある。
            join = datetime.datetime.combine(datetime.date.today(), df_form_ans.at[j, "join"])
            move_out = datetime.datetime.combine(datetime.date.today(), df_form_ans.at[j, "move_out"])
            d_start = datetime.datetime.combine(datetime.date.today(), start_time_list[i])
            d_end = d_start + datetime.timedelta(minutes=lesson_minuete)
            if join <= d_start and d_end <= move_out:
                df_attend.iat[i, j] = 1
            else:
                df_attend.iat[i, j] = 0
    arr_A = df_attend.values   # 出席簿配列
    return arr_A
    