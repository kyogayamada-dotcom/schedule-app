import streamlit as st
import pandas as pd
import datetime
import io
import re
import random
import pickle
import json
import os
from collections import Counter

# ==========================================
# 0. 設定・定数
# ==========================================
ADMIN_PASSWORD = "2020"
CONFIG_FILE = "admin_settings.json"

# ==========================================
# 1. 保存・読み込みロジック (JSON)
# ==========================================
def load_config():
    """サーバー上のファイルから設定を読み込む"""
    default_config = {
        "start_date": datetime.date(2025, 12, 1),
        "end_date": datetime.date(2026, 1, 31),
        "overrides": {}
    }
    
    if os.path.exists(CONFIG_FILE):
        try:
            with open(CONFIG_FILE, 'r', encoding='utf-8') as f:
                data = json.load(f)
                config = {}
                config["start_date"] = datetime.datetime.strptime(data["start_date"], "%Y-%m-%d").date()
                config["end_date"] = datetime.datetime.strptime(data["end_date"], "%Y-%m-%d").date()
                
                overrides = {}
                for k, v in data.get("overrides", {}).items():
                    d_key = datetime.datetime.strptime(k, "%Y-%m-%d").date()
                    overrides[d_key] = v
                config["overrides"] = overrides
                return config
        except Exception as e:
            st.error(f"設定読み込みエラー: {e}")
            return default_config
    else:
        return default_config

def save_config(current_config):
    """現在の設定をファイルに書き込む"""
    save_data = {
        "start_date": current_config["start_date"].strftime("%Y-%m-%d"),
        "end_date": current_config["end_date"].strftime("%Y-%m-%d"),
        "overrides": {k.strftime("%Y-%m-%d"): v for k, v in current_config["overrides"].items()}
    }
    try:
        with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
            json.dump(save_data, f, indent=4)
        return True
    except Exception as e:
        st.error(f"保存エラー: {e}")
        return False

# ==========================================
# 2. カレンダー・ロジック設定
# ==========================================
def get_base_open_periods(date_obj):
    m, d, w = date_obj.month, date_obj.day, date_obj.weekday()
    if m == 1 and d in [1, 2, 3]: return []
    if m == 12 and d == 31: return []
    if w in [5, 6]: return [2, 3, 4, 5, 6]
    return [4, 5, 6]

def get_open_periods(date_obj):
    overrides = st.session_state.calendar_config.get("overrides", {})
    if date_obj in overrides:
        return overrides[date_obj]
    return get_base_open_periods(date_obj)

def get_year_from_range(month, day, start_date, end_date):
    curr = start_date
    while curr <= end_date:
        if curr.month == month and curr.day == day:
            return curr.year
        curr += datetime.timedelta(days=1)
    return start_date.year

# ==========================================
# 3. データ処理・計算ロジック
# ==========================================
def check_sufficiency(student_weekly_data, req_df):
    warnings = []
    student_reqs = {}
    for _, row in req_df.iterrows():
        name = row['生徒名']
        total = sum(int(row.get(k, 0)) for k in ["国語", "数学", "英語", "理科", "社会"])
        student_reqs[name] = total
    student_avails = {name: 0 for name in student_reqs}
    start_date = st.session_state.calendar_config["start_date"]
    end_date = st.session_state.calendar_config["end_date"]
    for s_name, weekly_data in student_weekly_data.items():
        if s_name not in student_reqs: continue
        count = 0
        if not weekly_data: continue
        for week_label, df in weekly_data.items():
            for date_str in df.columns:
                match = re.search(r"(\d+)/(\d+)", date_str)
                if not match: continue
                m, d = int(match.group(1)), int(match.group(2))
                y = get_year_from_range(m, d, start_date, end_date)
                try: d_date = datetime.date(y, m, d)
                except: continue
                open_periods = get_open_periods(d_date)
                for p in range(1, 7):
                    if p not in open_periods: continue
                    try: val = str(df.loc[p, date_str])
                    except: continue
                    if any(x in val for x in ["〇", "○", "OK", "△", "▲", "1", "2", "3", "全"]):
                        count += 1
        student_avails[s_name] = count
    for name, req_num in student_reqs.items():
        avail_num = student_avails.get(name, 0)
        if avail_num < req_num:
            warnings.append(f"{name}：希望 {req_num}コマ > 空き {avail_num}コマ (不足確定: {req_num - avail_num})")
    return warnings

def calculate_schedule(teacher_weekly_data, req_df, student_weekly_data, teacher_name):
    teacher_capacity = {}
    start_date = st.session_state.calendar_config["start_date"]
    end_date = st.session_state.calendar_config["end_date"]
    for week_label, df in teacher_weekly_data.items():
        for date_str in df.columns:
            match = re.search(r"(\d+)/(\d+)", date_str)
            if not match: continue
            m, d = int(match.group(1)), int(match.group(2))
            y = get_year_from_range(m, d, start_date, end_date)
            try: d_date = datetime.date(y, m, d)
            except: continue
            open_periods = get_open_periods(d_date)
            for p in range(1, 7):
                try: val = str(df.loc[p, date_str])
                except: continue
                if p not in open_periods: continue
                if any(x in val for x in ["〇", "○", "OK", "全"]):
                    teacher_capacity[(d_date, p)] = 2
                elif any(x in val for x in ["△", "▲", "半", "1"]):
                    teacher_capacity[(d_date, p)] = 1
    all_slots = []
    for (d, p), cap in teacher_capacity.items():
        all_slots.append((d, p, cap))
    students = {}
    for _, row in req_df.iterrows():
        name = row['生徒名']
        reqs = {k: int(row.get(k, 0)) for k in ["国語", "数学", "英語", "理科", "社会"]}
        students[name] = {"reqs": reqs, "remaining": sum(reqs.values())}
    student_availability = {}
    for s_name, weekly_data in student_weekly_data.items():
        if not weekly_data: continue
        for week_label, df in weekly_data.items():
            for date_str in df.columns:
                match = re.search(r"(\d+)/(\d+)", date_str)
                if not match: continue
                m, d = int(match.group(1)), int(match.group(2))
                y = get_year_from_range(m, d, start_date, end_date)
                try: d_date = datetime.date(y, m, d)
                except: continue
                for p in range(1, 7):
                    try: val = str(df.loc[p, date_str])
                    except: continue
                    if any(x in val for x in ["〇", "○", "OK", "△", "▲", "1", "2", "3", "全"]):
                        student_availability[(s_name, d_date, p)] = True
                    else:
                        student_availability[(s_name, d_date, p)] = False
    schedule_map = { (d, p): [] for d, p, cap in all_slots }
    date_counts = Counter()
    daily_student_counts = Counter()
    random.seed(42)
    max_loops = 3000
    loop_count = 0
    while loop_count < max_loops:
        loop_count += 1
        assigned_in_this_loop = False
        def get_slot_priority(slot):
            d, p, cap = slot
            if len(schedule_map[(d, p)]) >= cap: return -99999
            score = 0
            if len(schedule_map.get((d, p-1), [])) > 0: score += 100
            if len(schedule_map.get((d, p+1), [])) > 0: score += 100
            score += date_counts[d] * 10
            score += random.random()
            return score
        all_slots.sort(key=get_slot_priority, reverse=True)
        for d, p, cap in all_slots:
            current_assigned = schedule_map[(d, p)]
            if len(current_assigned) >= cap: continue
            candidates = []
            for s_name, data in students.items():
                if data["remaining"] <= 0: continue
                if daily_student_counts[(s_name, d)] >= 3: continue
                if not student_availability.get((s_name, d, p), False): continue
                is_already_in = False
                for entry in current_assigned:
                    if entry.startswith(s_name + "("):
                        is_already_in = True; break
                if is_already_in: continue
                candidates.append(s_name)
            if not candidates: continue
            candidates.sort(key=lambda x: (students[x]["remaining"], random.random()), reverse=True)
            s = candidates[0]
            items = sorted([(v, k) for k, v in students[s]["reqs"].items() if v > 0], reverse=True)
            if not items: continue
            subj = items[0][1]
            students[s]["reqs"][subj] -= 1
            students[s]["remaining"] -= 1
            daily_student_counts[(s, d)] += 1
            date_counts[d] += 1
            schedule_map[(d, p)].append(f"{s}({subj})")
            assigned_in_this_loop = True
            break
        if not assigned_in_this_loop: break
    all_dates = sorted(list(set([x[0] for x in all_slots])))
    unscheduled = []
    for s, data in students.items():
        for subj, cnt in data["reqs"].items():
            if cnt > 0: unscheduled.append({"生徒名": s, "科目": subj, "不足": cnt})
    return schedule_map, all_dates, unscheduled

# ==========================================
# 4. UIヘルパー関数
# ==========================================
def get_week_ranges():
    start_date = st.session_state.calendar_config["start_date"]
    end_date = st.session_state.calendar_config["end_date"]
    weeks = []
    current_dates = []
    curr = start_date
    while curr <= end_date:
        current_dates.append(curr)
        if len(current_dates) == 7 or curr == end_date:
            label = f"{current_dates[0].strftime('%m/%d')} 〜 {current_dates[-1].strftime('%m/%d')}"
            weeks.append({"label": label, "dates": current_dates})
            current_dates = []
        curr += datetime.timedelta(days=1)
    return weeks

def create_weekly_df(dates):
    col_names = [d.strftime("%m/%d(%a)") for d in dates]
    data = {}
    for d_obj, col in zip(dates, col_names):
        open_periods = get_open_periods(d_obj)
        col_data = []
        for p in range(1, 7):
            val = "〇" if p in open_periods else "×"
            col_data.append(val)
        data[col] = col_data
    return pd.DataFrame(data, index=[1, 2, 3, 4, 5, 6])

def create_student_req_df(student_names):
    data = []
    for name in student_names:
        data.append({"生徒名": name, "国語": 0, "数学": 0, "英語": 0, "理科": 0, "社会": 0})
    return pd.DataFrame(data)

# ==========================================
# 5. メインアプリ (Streamlit)
# ==========================================
st.set_page_config(page_title="時間割作成 ", layout="wide")
st.title(" 個別指導塾 時間割作成")

if "calendar_config" not in st.session_state:
    st.session_state.calendar_config = load_config()

if "teacher_weekly_data" not in st.session_state: st.session_state.teacher_weekly_data = None
if "student_req_df" not in st.session_state: st.session_state.student_req_df = None
if "student_weekly_data" not in st.session_state: st.session_state.student_weekly_data = {}
if "student_list" not in st.session_state: st.session_state.student_list = []
if "teacher_name_default" not in st.session_state: st.session_state.teacher_name_default = "佐藤"

weeks_info = get_week_ranges()

# --- サイドバー ---
with st.sidebar:
    st.header("1. 基本設定")
    teacher_name = st.text_input("コーチの名前", value=st.session_state.teacher_name_default)
    
    st.subheader("生徒リスト")
    default_students = "\n".join(st.session_state.student_list) if st.session_state.student_list else "山田くん\n田中さん\n高橋くん"
    s_input = st.text_area("名前を入力 (改行区切り)", default_students, height=100)
    
    if st.button("入力を開始/リセット"):
        new_list = [s.strip() for s in s_input.split('\n') if s.strip()]
        st.session_state.student_list = new_list
        st.session_state.teacher_name_default = teacher_name
        
        t_data = {}
        for w in weeks_info: t_data[w["label"]] = create_weekly_df(w["dates"])
        st.session_state.teacher_weekly_data = t_data
        st.session_state.student_req_df = create_student_req_df(new_list)
        s_data_all = {}
        for s in new_list:
            s_weeks = {}
            for w in weeks_info: s_weeks[w["label"]] = create_weekly_df(w["dates"])
            s_data_all[s] = s_weeks
        st.session_state.student_weekly_data = s_data_all
        st.success("リセットしました。")

    # 管理者設定
    st.divider()
    st.subheader("🔧 管理者メニュー")
    pwd = st.text_input("パスワード", type="password", help="年度や期間、休講日を変更する場合に入力してください")
    
    if pwd == ADMIN_PASSWORD:
        st.success("認証成功")
        with st.expander("📅 期間・カレンダー設定", expanded=True):
            st.write("**講習期間の設定**")
            col_d1, col_d2 = st.columns(2)
            current_start = st.session_state.calendar_config["start_date"]
            current_end = st.session_state.calendar_config["end_date"]
            new_start = col_d1.date_input("開始日", current_start)
            new_end = col_d2.date_input("終了日", current_end)
            
            if new_start > new_end:
                st.error("終了日は開始日よりあとに設定してください")
            else:
                if new_start != current_start or new_end != current_end:
                    st.session_state.calendar_config["start_date"] = new_start
                    st.session_state.calendar_config["end_date"] = new_end
                    if save_config(st.session_state.calendar_config):
                        st.success("期間を保存しました。反映には「入力を開始」を押してください。")

            st.divider()
            st.write("**例外ルールの追加（特定日の変更）**")
            ex_date = st.date_input("日付を選択", new_start)
            current_periods = get_open_periods(ex_date)
            st.caption(f"現在の設定: {current_periods if current_periods else '全休'}")
            
            st.write("開講するコマを選択:")
            cols = st.columns(3)
            new_periods = []
            for p in range(1, 7):
                checked = p in current_periods
                if cols[(p-1)%3].checkbox(f"{p}講", value=checked, key=f"chk_{p}"):
                    new_periods.append(p)
            
            col_b1, col_b2 = st.columns(2)
            if col_b1.button("ルールを保存"):
                st.session_state.calendar_config["overrides"][ex_date] = new_periods
                if save_config(st.session_state.calendar_config):
                    st.success("設定を保存しました。")
            
            if col_b2.button("例外を削除"):
                if ex_date in st.session_state.calendar_config["overrides"]:
                    del st.session_state.calendar_config["overrides"][ex_date]
                    if save_config(st.session_state.calendar_config):
                        st.success("削除しました。")
    elif pwd != "":
        st.error("パスワードが違います")

    # 個人データ保存
    st.divider()
    st.subheader("💾 データの保存・復元")
    if st.session_state.teacher_weekly_data is not None:
        export_data = {
            "teacher_name": teacher_name,
            "student_list": st.session_state.student_list,
            "teacher_weekly_data": st.session_state.teacher_weekly_data,
            "student_req_df": st.session_state.student_req_df,
            "student_weekly_data": st.session_state.student_weekly_data,
            "calendar_config": st.session_state.calendar_config
        }
        try:
            pickle_byte = pickle.dumps(export_data)
            st.download_button(
                label="📥 データを保存 (.pkl)",
                data=pickle_byte,
                file_name=f"schedule_data_{datetime.date.today()}.pkl",
                mime="application/octet-stream"
            )
        except Exception as e:
            st.error(f"保存準備エラー: {e}")
    
    uploaded_file = st.file_uploader("📤 データを読み込む", type=["pkl"])
    if uploaded_file is not None:
        try:
            loaded_data = pickle.load(uploaded_file)
            st.session_state.student_list = loaded_data.get("student_list", [])
            st.session_state.teacher_weekly_data = loaded_data.get("teacher_weekly_data", None)
            st.session_state.student_req_df = loaded_data.get("student_req_df", None)
            st.session_state.student_weekly_data = loaded_data.get("student_weekly_data", {})
            if "teacher_name" in loaded_data:
                st.session_state.teacher_name_default = loaded_data["teacher_name"]
            if "calendar_config" in loaded_data:
                st.session_state.calendar_config = loaded_data["calendar_config"]
            st.success("復元完了！")
            st.rerun()
        except Exception as e:
            st.error(f"読み込み失敗: {e}")

# --- メインエリア ---
if st.session_state.teacher_weekly_data is None:
    st.info("👈 生徒名を入力して「入力を開始」を押してください。")
else:
    tab1, tab2, tab3, tab4 = st.tabs(["📅 コーチシフト", "🔢 生徒希望数", "🙋‍♂️ 生徒シフト", "🚀 作成＆結果"])

    with tab1:
        st.subheader(f"{teacher_name}コーチの予定")
        st.caption("「〇」=両配ok、「△」＝片配ok、「×」＝入れない")
        st.info("💡 入力後に必ず「保存」を押してください。")
        with st.form("teacher_form"):
            updated_weekly_data = {}
            for w in weeks_info:
                label = w["label"]
                st.write(f"**{label}**")
                original_df = st.session_state.teacher_weekly_data.get(label)
                if original_df is None: original_df = create_weekly_df(w["dates"])
                column_config = {}
                options = ["〇", "×", "△"]
                for col in original_df.columns:
                    column_config[col] = st.column_config.SelectboxColumn(col, options=options, width="small", required=True)
                edited_df = st.data_editor(original_df, column_config=column_config, width='stretch', key=f"teacher_edit_{label}", height=300)
                updated_weekly_data[label] = edited_df
                st.divider()
            if st.form_submit_button("💾 入力内容を保存する", type="primary"):
                st.session_state.teacher_weekly_data = updated_weekly_data
                st.success("保存しました！")

    with tab2:
        st.subheader("各教科の必要コマ数")
        st.info("💡 入力後に必ず「保存」を押してください。")
        with st.form("req_form"):
            edited_req_df = st.data_editor(st.session_state.student_req_df, hide_index=True, width='stretch')
            if st.form_submit_button("💾 希望数を保存する", type="primary"):
                st.session_state.student_req_df = edited_req_df
                st.success("保存しました！")

    with tab3:
        st.subheader("生徒の行ける日時")
        target_student = st.selectbox("生徒を選択", st.session_state.student_list)
        if target_student:
            st.caption(f"{target_student} の行ける時間")
            st.info("💡 入力後に必ず「保存」を押してください。")
            with st.form(f"student_form_{target_student}"):
                updated_s_weekly = {}
                for w in weeks_info:
                    label = w["label"]
                    st.write(f"**{label}**")
                    s_data_map = st.session_state.student_weekly_data.get(target_student, {})
                    s_df = s_data_map.get(label)
                    if s_df is None: s_df = create_weekly_df(w["dates"])
                    column_config_s = {}
                    options = ["〇", "×"]
                    for col in s_df.columns:
                        column_config_s[col] = st.column_config.SelectboxColumn(col, options=options, width="small", required=True)
                    edited_s_df = st.data_editor(s_df, column_config=column_config_s, width='stretch', key=f"student_edit_{target_student}_{label}", height=300)
                    
                    updated_s_weekly[label] = edited_s_df
                    st.divider()
                if st.form_submit_button(f"💾 {target_student} のシフトを保存する", type="primary"):
                    st.session_state.student_weekly_data[target_student] = updated_s_weekly
                    st.success("保存しました！")

    with tab4:
        st.subheader("時間割作成")
        if st.button("🚀 作成スタート", type="primary"):
            warnings = check_sufficiency(st.session_state.student_weekly_data, st.session_state.student_req_df)
            if warnings:
                st.warning("⚠️ 【注意】空きコマ不足の生徒がいます")
                for w in warnings: st.write(f"- {w}")
                st.divider()
            with st.spinner("計算中..."):
                try:
                    schedule_map, all_dates, unscheduled = calculate_schedule(
                        st.session_state.teacher_weekly_data,
                        st.session_state.student_req_df,
                        st.session_state.student_weekly_data,
                        teacher_name
                    )
                    st.success("✅ 完成しました！")
                    st.subheader("📅 完成時間割プレビュー")
                    
                    cal_dates = []
                    curr = st.session_state.calendar_config["start_date"]
                    end = st.session_state.calendar_config["end_date"]
                    while curr <= end:
                        cal_dates.append(curr)
                        curr += datetime.timedelta(days=1)

                    for i in range(0, len(cal_dates), 7):
                        week_dates = cal_dates[i : i+7]
                        week_data = {}
                        col_names = [d.strftime("%m/%d(%a)") for d in week_dates]
                        col_config = {}
                        for d_obj, col in zip(week_dates, col_names):
                            col_config[col] = st.column_config.TextColumn(col, width="medium")
                            col_content = []
                            for p in range(1, 7):
                                assigned = schedule_map.get((d_obj, p), [])
                                if assigned:
                                    col_content.append(", ".join(assigned))
                                else:
                                    open_periods = get_open_periods(d_obj)
                                    col_content.append("-" if p in open_periods else "×")
                            week_data[col] = col_content
                        df_week_view = pd.DataFrame(week_data, index=[f"{p}講" for p in range(1, 7)])
                        st.write(f"**{week_dates[0].strftime('%Y/%m/%d')} 週**")
                        # 色指定を削除し、通常のデータフレーム表示に戻しました
                        st.dataframe(df_week_view, column_config=col_config, width='stretch')
                        

                    if unscheduled:
                        st.error("⚠️ 入りきらなかった授業")
                        st.dataframe(pd.DataFrame(unscheduled), hide_index=True)
                    else:
                        st.info("🎉 全て完了！")

                    output = io.BytesIO()
                    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                        workbook = writer.book
                        worksheet = workbook.add_worksheet("時間割")
                        writer.sheets["時間割"] = worksheet
                        wrap_fmt = workbook.add_format({'text_wrap': True, 'valign': 'top', 'border': 1, 'align': 'center'})
                        header_fmt = workbook.add_format({'bold': True, 'bg_color': '#D9E1F2', 'border': 1, 'align': 'center'})
                        current_row = 0
                        for i in range(0, len(cal_dates), 7):
                            week_dates = cal_dates[i : i+7]
                            worksheet.write(current_row, 0, "講", header_fmt)
                            for col_idx, d_obj in enumerate(week_dates):
                                worksheet.write(current_row, col_idx + 1, d_obj.strftime("%m/%d(%a)"), header_fmt)
                            for p in range(1, 7):
                                row_idx = current_row + p
                                worksheet.write(row_idx, 0, p, wrap_fmt)
                                for col_idx, d_obj in enumerate(week_dates):
                                    assigned = schedule_map.get((d_obj, p), [])
                                    cell_text = "\n".join(assigned) if assigned else ("" if p in get_open_periods(d_obj) else "×")
                                    worksheet.write(row_idx, col_idx + 1, cell_text, wrap_fmt)
                            current_row += 8
                        worksheet.set_column(0, 0, 5); worksheet.set_column(1, 7, 18)
                        if unscheduled: pd.DataFrame(unscheduled).to_excel(writer, sheet_name="未消化リスト", index=False)
                    st.download_button(label="📥 Excel保存", data=output.getvalue(), file_name=f"完成時間割_{teacher_name}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
                except Exception as e:
                    st.error(f"エラー: {e}")