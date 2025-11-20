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
    日付ごとの開講コマ定義 (最新版維持)
    """
    m, d = date_obj.month, date_obj.day

    # 1. 1月7, 8, 9日は 3,4,5,6講 (1,2講は休み)
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
    
    # 共通の空シフト表を作成
    rows_template = []
    temp_curr = curr
    while temp_curr <= end:
        open_p = get_open_periods(temp_curr)
        row = {"日付": temp_curr, "曜日": temp_curr.strftime("%a")}
        for p in range(1, 7):
            row[p] = "〇" if p in open_p else "×"
        rows_template.append(row)
        temp_curr += datetime.timedelta(days=1)
    
    df_template = pd.DataFrame(rows_template)[["日付", "曜日", 1, 2, 3, 4, 5, 6]]

    # 入力された生徒リストからデータを作成
    student_data = []
    for name in student_names_list:
        name = name.strip()
        if name:
            # デフォルト値を設定
            student_data.append({
                "生徒名": name, "国語": 0, "数学": 0, "英語": 0, "理科": 0, "社会": 0
            })
    
    if not student_data: # 空っぽの場合のダミー
        student_data.append({"生徒名": "サンプル生", "国語": 0, "数学": 0, "英語": 0, "理科": 0, "社会": 0})

    df_req = pd.DataFrame(student_data)

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        # 1. コーチシフト
        df_template.to_excel(writer, sheet_name="コーチシフト(〇をつける)", index=False)
        # 2. 生徒希望数
        df_req.to_excel(writer, sheet_name="生徒希望数", index=False)
        # 3. 生徒ごとのシフトシートを作成
        for s_info in student_data:
            sheet_name = f"シフト_{s_info['生徒名']}"
            # シート名が31文字を超えないようにカット（Excel制限）
            sheet_name = sheet_name[:31]
            df_template.to_excel(writer, sheet_name=sheet_name, index=False)

    return output.getvalue()

def process_schedule(uploaded_file, teacher_name):
    xl = pd.ExcelFile(uploaded_file)
    sheet_names = xl.sheet_names

    df_teacher_matrix = pd.read_excel(uploaded_file, sheet_name="コーチシフト(〇をつける)")
    df_req = pd.read_excel(uploaded_file, sheet_name="生徒希望数")

    # A. 先生シフト解析 (人数対応)
    teacher_capacity = {}
    p_cols = [c for c in df_teacher_matrix.columns if str(c) in ["1","2","3","4","5","6"]]
    
    for _, row in df_teacher_matrix.iterrows():
        d = row['日付'].date() if isinstance(row['日付'], pd.Timestamp) else row['日付']
        for p_col in p_cols:
            val = str(row[p_col]).strip()
            p_num = int(p_col)
            if p_num not in get_open_periods(d): continue
            
            # 人数判定
            nums = re.findall(r'[0-9]+', val)
            if nums:
                teacher_capacity[(d, p_num)] = int(nums[0])
            elif any(x in val for x in ["〇", "○", "OK", "全"]):
                teacher_capacity[(d, p_num)] = 2 
            elif any(x in val for x in ["△", "▲", "半"]):
                teacher_capacity[(d, p_num)] = 1

    # B. 生徒データ & 生徒シフト解析
    students = {}
    student_availability = {} 

    for _, row in df_req.iterrows():
        name = row['生徒名']
        reqs = {k: row.get(k, 0) for k in ["国語", "数学", "英語", "理科", "社会"]}
        students[name] = {"reqs": reqs, "remaining": sum(reqs.values())}

        sheet_name = f"シフト_{name}"[:31] # シート名の長さを合わせる
        
        # 完全一致または近いシート名を探す
        target_sheet = None
        if sheet_name in sheet_names:
            target_sheet = sheet_name
        
        if target_sheet:
            df_s_shift = pd.read_excel(uploaded_file, sheet_name=target_sheet)
            for _, s_row in df_s_shift.iterrows():
                d = s_row['日付'].date() if isinstance(s_row['日付'], pd.Timestamp) else s_row['日付']
                for p_col in p_cols:
                    val = str(s_row[p_col]).strip()
                    p_num = int(p_col)
                    if any(x in val for x in ["〇", "○", "OK", "1", "2", "3", "全"]):
                        student_availability[(name, d, p_num)] = True
                    else:
                        student_availability[(name, d, p_num)] = False
        else:
            # シートがない場合、とりあえずNGとする（または警告）
            pass

    # D. 作成
    schedule = []
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
                
                # 生徒のシフトチェック
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
                row_data = {
                    "日付": curr.strftime("%Y-%m-%d"), 
                    "曜日": curr.strftime("%a"), 
                    "講": p, 
                    "コーチ": teacher_name
                }
                for i, s_info in enumerate(assigned):
                    row_data[f"生徒{i+1}"] = s_info
                schedule.append(row_data)
                
        curr += datetime.timedelta(days=1)

    # E. 出力データ作成
    unscheduled = []
    for s, data in students.items():
        for subj, cnt in data["reqs"].items():
            if cnt > 0: unscheduled.append({"生徒名": s, "科目": subj, "不足": cnt})

    df_schedule = pd.DataFrame(schedule)
    if not df_schedule.empty:
        base_cols = ["日付", "曜日", "講", "コーチ"]
        student_cols = [c for c in df_schedule.columns if c.startswith("生徒")]
        student_cols.sort(key=lambda x: int(x.replace("生徒", "")))
        df_schedule = df_schedule[base_cols + student_cols]

    df_unscheduled = pd.DataFrame(unscheduled)

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df_schedule.to_excel(writer, sheet_name="時間割", index=False)
        if not df_unscheduled.empty:
            df_unscheduled.to_excel(writer, sheet_name="未消化リスト", index=False)
    
    return output.getvalue()

# ==========================================
# Web画面 (Streamlit)
# ==========================================
st.title("個別指導塾ゴールフリー 時間割作成ツール")


teacher_name = st.text_input("コーチの名前を入力してください", "")

st.divider()

st.subheader("ステップ1: 入力用Excelを作る")

# 生徒リスト入力欄の追加
default_students = "山田くん\n田中さん\n高橋くん"
student_input = st.text_area("生徒の名前を入力してください（改行で区切ると複数人になります）", default_students, height=150)

if st.button("入力用Excelをダウンロード"):
    # 入力されたテキストをリストに変換
    student_list = [s.strip() for s in student_input.split('\n') if s.strip()]
    
    excel_data = create_template_data(teacher_name, student_list)
    
    st.download_button(
        label=f"📥 {teacher_name}先生＆{len(student_list)}名分の入力表をダウンロード",
        data=excel_data,
        file_name=f"入力表_{teacher_name}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
    st.info(f"{len(student_list)}名分のシートを作成しました。\n各シートに行ける日時を入力してください。")

st.divider()

st.subheader("ステップ2: 編集したExcelをアップロードして作成")
uploaded_file = st.file_uploader("編集済みのExcelファイルをここにアップロード", type=["xlsx"])

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
                st.error(f"エラーが発生しました: {e}")
