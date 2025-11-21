import streamlit as st
import pandas as pd
import datetime
import io
import re
from collections import Counter

# ==========================================
# ロジック部分
# ==========================================
def get_open_periods(date_obj):
    """
    日付ごとの開講コマ定義
    """
    m, d = date_obj.month, date_obj.day

    # 1. 1月7, 8, 9日は 3,4,5,6講
    if m == 1 and d in [7, 8, 9]:
        return [3, 4, 5, 6]

    # 2. 12/23, 24は 3-6講
    if m == 12 and d in [23, 24]:
        return [3, 4, 5, 6]

    # 3. 特定の日付の1,2講をバツにする
    if (m == 12 and d in [20, 21, 27]) or (m == 1 and d in [4, 10, 11]):
        return [3, 4, 5]
    
    if (m == 12 and d in [25, 26]) or (m == 1 and d == 6):
        return [3, 4, 5, 6]

    if m == 12 and d == 28:
        return [3, 4]

    # 4-6講のみ
    if (m == 12 and (2<=d<=5 or 9<=d<=12 or 16<=d<=19)) or \
       (m == 1 and (13<=d<=16 or 20<=d<=23 or 27<=d<=30)):
        return [4, 5, 6]

    # 2-5講のみ
    if (m == 12 and d in [6, 13]) or (m == 1 and d in [17, 24, 31]):
        return [2, 3, 4, 5]

    return []

def create_template_data(teacher_name, student_names_list):
    # 期間設定
    curr = datetime.date(2025, 12, 1)
    end = datetime.date(2026, 1, 31)
    
    # 日付リスト作成 (ヘッダー用)
    date_headers = []
    date_objs = []
    temp_curr = curr
    while temp_curr <= end:
        # 日付文字列 (例: 12/01(Mon))
        d_str = temp_curr.strftime("%m/%d(%a)")
        date_headers.append(d_str)
        date_objs.append(temp_curr)
        temp_curr += datetime.timedelta(days=1)

    # ---------------------------
    # 1. 先生シフト (縦:講, 横:日付)
    # ---------------------------
    # 行データを作成 (1講〜6講)
    rows_shift = []
    for p in range(1, 7):
        row_data = {"講": p}
        for d_str, d_obj in zip(date_headers, date_objs):
            open_periods = get_open_periods(d_obj)
            # 開講なら〇、閉講なら×
            row_data[d_str] = "〇" if p in open_periods else "×"
        rows_shift.append(row_data)
    
    # カラム順序を保証
    cols_order = ["講"] + date_headers
    df_template = pd.DataFrame(rows_shift)
    df_template = df_template[cols_order]

    # ---------------------------
    # 2. 生徒希望数
    # ---------------------------
    student_data = []
    for name in student_names_list:
        name = name.strip()
        if name:
            student_data.append({
                "生徒名": name, "国語": 0, "数学": 0, "英語": 0, "理科": 0, "社会": 0
            })
    if not student_data:
        student_data.append({"生徒名": "サンプル生", "国語": 0, "数学": 0, "英語": 0, "理科": 0, "社会": 0})
    df_req = pd.DataFrame(student_data)

    # ---------------------------
    # 3. Excel出力
    # ---------------------------
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        # 先生シフト
        df_template.to_excel(writer, sheet_name="先生シフト", index=False)
        
        # 生徒希望数
        df_req.to_excel(writer, sheet_name="生徒希望数", index=False)
        
        # 生徒ごとのシフトシート (先生と同じ形式)
        for s_info in student_data:
            sheet_name = f"シフト_{s_info['生徒名']}"[:31]
            df_template.to_excel(writer, sheet_name=sheet_name, index=False)

    return output.getvalue()

def process_schedule(uploaded_file, teacher_name):
    xl = pd.ExcelFile(uploaded_file)
    sheet_names = xl.sheet_names

    # 読み込み
    df_teacher = pd.read_excel(uploaded_file, sheet_name="先生シフト")
    df_req = pd.read_excel(uploaded_file, sheet_name="生徒希望数")

    # 日付カラムを特定する関数
    # (Excelで日付がシリアル値やdatetimeになったり文字列になったりするため)
    def is_date_column(col_name):
        # "講" 以外を日付とみなす
        return str(col_name) != "講"

    # 日付カラムのマッピング作成 (カラム名 -> datetimeオブジェクト)
    # 形式: "12/01(Mon)" -> datetime.date(2025, 12, 1)
    date_map = {}
    # 2025/12/1から開始と仮定してマッピング（簡易的だが確実）
    # もしExcelの日付ヘッダーが日付型で認識されている場合はそのまま使う
    
    # 列名リストから日付っぽいものを抽出
    date_cols = [c for c in df_teacher.columns if is_date_column(c)]
    
    # 列名から日付オブジェクトへの変換を試みる
    # ここではテンプレート通りの順番であると仮定して、開始日から割り当てるのが安全
    curr = datetime.date(2025, 12, 1)
    for col in date_cols:
        # もし列名自体がdatetime型ならそれを使う
        if isinstance(col, datetime.datetime):
            date_map[col] = col.date()
        else:
            # 文字列の場合は、ループ順に日付を割り当てる（テンプレートの仕様依存）
            date_map[col] = curr
            curr += datetime.timedelta(days=1)

    # A. 先生シフト解析 (横軸日付版)
    teacher_capacity = {}
    
    for _, row in df_teacher.iterrows():
        try:
            p_num = int(row['講'])
        except:
            continue # 講が数値でない行はスキップ
            
        for col in date_cols:
            d = date_map[col]
            val = str(row[col]).strip()
            
            # 開講日チェック
            if p_num not in get_open_periods(d):
                continue

            # 人数判定
            nums = re.findall(r'[0-9]+', val)
            if nums:
                teacher_capacity[(d, p_num)] = int(nums[0])
            elif any(x in val for x in ["〇", "○", "OK", "全"]):
                teacher_capacity[(d, p_num)] = 2 
            elif any(x in val for x in ["△", "▲", "半"]):
                teacher_capacity[(d, p_num)] = 1

    # B. 生徒データ & シフト解析
    students = {}
    student_availability = {} 

    for _, row in df_req.iterrows():
        name = row['生徒名']
        reqs = {k: row.get(k, 0) for k in ["国語", "数学", "英語", "理科", "社会"]}
        students[name] = {"reqs": reqs, "remaining": sum(reqs.values())}

        sheet_name = f"シフト_{name}"[:31]
        
        # シート名マッチング
        target_sheet = None
        if sheet_name in sheet_names:
            target_sheet = sheet_name
        
        if target_sheet:
            df_s = pd.read_excel(uploaded_file, sheet_name=target_sheet)
            # 生徒シフト読み込み
            s_date_cols = [c for c in df_s.columns if is_date_column(c)]
            
            # 生徒シートの日付マッピングも再構築
            s_date_map = {}
            curr_s = datetime.date(2025, 12, 1)
            for col in s_date_cols:
                if isinstance(col, datetime.datetime):
                    s_date_map[col] = col.date()
                else:
                    s_date_map[col] = curr_s
                    curr_s += datetime.timedelta(days=1)

            for _, s_row in df_s.iterrows():
                try:
                    p_num = int(s_row['講'])
                except:
                    continue
                
                for col in s_date_cols:
                    d = s_date_map[col]
                    val = str(s_row[col]).strip()
                    
                    if any(x in val for x in ["〇", "○", "OK", "1", "2", "3", "全"]):
                        student_availability[(name, d, p_num)] = True
                    else:
                        student_availability[(name, d, p_num)] = False

    # D. 作成
    # 結果格納用マップ: schedule_map[(date, period)] = [生徒名(科目), ...]
    schedule_map = {}

    curr = datetime.date(2025, 12, 1)
    end_date = datetime.date(2026, 1, 31)
    
    while curr <= end_date:
        periods = get_open_periods(curr)
        daily_counts = Counter()
        
        for p in periods:
            capacity = teacher_capacity.get((curr, p), 0)
            if capacity == 0: continue
            
            cands = []
            for s_name, data in students.items():
                if data["remaining"] <= 0: continue
                if daily_counts[s_name] >= 3: continue
                
                if not student_availability.get((s_name, curr, p), False):
                    continue

                cands.append(s_name)
            
            cands.sort(key=lambda x: students[x]["remaining"], reverse=True)
            
            assigned = []
            while len(assigned) < capacity and cands:
                s = cands.pop(0)
                items = sorted([(v, k) for k, v in students[s]["reqs"].items() if v > 0], reverse=True)
                if not items: continue
                subj = items[0][1]
                
                students[s]["reqs"][subj] -= 1
                students[s]["remaining"] -= 1
                daily_counts[s] += 1
                assigned.append(f"{s}({subj})")
            
            if assigned:
                schedule_map[(curr, p)] = assigned
                
        curr += datetime.timedelta(days=1)

    # E. 出力データ作成 (横軸日付形式)
    
    # 1. 時間割表 (Rows=講, Cols=日付)
    out_rows = []
    
    # 日付ヘッダー再作成
    out_date_headers = []
    out_dates = []
    temp_curr = datetime.date(2025, 12, 1)
    while temp_curr <= end_date:
        d_str = temp_curr.strftime("%m/%d(%a)")
        out_date_headers.append(d_str)
        out_dates.append(temp_curr)
        temp_curr += datetime.timedelta(days=1)
        
    for p in range(1, 7):
        row_data = {"講": p}
        for d_str, d_obj in zip(out_date_headers, out_dates):
            assigned_list = schedule_map.get((d_obj, p), [])
            if assigned_list:
                # セル内で改行して表示
                row_data[d_str] = "\n".join(assigned_list)
            else:
                # 開講してるけど誰もいないなら空欄、閉講なら斜線など
                if p in get_open_periods(d_obj):
                    row_data[d_str] = ""
                else:
                    row_data[d_str] = "×"
        out_rows.append(row_data)
        
    df_schedule = pd.DataFrame(out_rows)
    df_schedule = df_schedule[["講"] + out_date_headers]

    # 2. 未消化リスト
    unscheduled = []
    for s, data in students.items():
        for subj, cnt in data["reqs"].items():
            if cnt > 0: unscheduled.append({"生徒名": s, "科目": subj, "不足": cnt})
    df_unscheduled = pd.DataFrame(unscheduled)

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        # セル内改行を有効にするためのフォーマット設定は xlsxwriter の機能を使う必要があるが
        # Pandasのto_excelだけでは限界があるため、標準的な出力を行う
        df_schedule.to_excel(writer, sheet_name="時間割(横日付)", index=False)
        
        # 列幅調整などの見た目を整える（簡易的）
        workbook = writer.book
        worksheet = writer.sheets["時間割(横日付)"]
        wrap_format = workbook.add_format({'text_wrap': True, 'valign': 'top'})
        
        # データ範囲に折り返し設定を適用
        # (列数が多いのでざっくり全体に適用)
        worksheet.set_column(1, len(out_date_headers), 15, wrap_format)

        if not df_unscheduled.empty:
            df_unscheduled.to_excel(writer, sheet_name="未消化リスト", index=False)
    
    return output.getvalue()

# ==========================================
# Web画面 (Streamlit)
# ==========================================
st.title("個別指導塾 時間割作成ツール (横日付版)")
st.write("Excelの形式を「横軸＝日付」に変更しました。")

teacher_name = st.text_input("先生の名前を入力してください", "佐藤")

st.divider()

st.subheader("ステップ1: 入力用Excelを作る")

default_students = "山田くん\n田中さん\n高橋くん"
student_input = st.text_area("生徒の名前を入力してください（改行で区切る）", default_students, height=100)

if st.button("入力用ひな形をダウンロード"):
    student_list = [s.strip() for s in student_input.split('\n') if s.strip()]
    excel_data = create_template_data(teacher_name, student_list)
    
    st.download_button(
        label=f"📥 入力表をダウンロード",
        data=excel_data,
        file_name=f"入力表_{teacher_name}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
    st.info("Excelの「横方向」に日付が並んでいます。")

st.divider()

st.subheader("ステップ2: 作成実行")
uploaded_file = st.file_uploader("Excelをアップロード", type=["xlsx"])

if uploaded_file is not None:
    if st.button("時間割を作成する"):
        with st.spinner('計算中...'):
            try:
                excel_binary = process_schedule(uploaded_file, teacher_name)
                st.success("✅ 作成完了！")

                st.download_button(
                    label="📥 完成時間割をダウンロード",
                    data=excel_binary,
                    file_name=f"完成時間割_{teacher_name}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            except Exception as e:
                st.error(f"エラー: {e}")