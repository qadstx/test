import base64
import io
import uuid
from datetime import datetime, timedelta, timezone

import pandas as pd
import requests
import streamlit as st
from streamlit_calendar import calendar

st.set_page_config(page_title="Task & Calendar", page_icon="📅", layout="wide")

# --- Timezone (Thailand UTC+7) ---
TH_TZ = timezone(timedelta(hours=7))


def today_th():
    return datetime.now(TH_TZ).date()


def parse_cal_date(date_str: str):
    if "T" in date_str:
        dt = pd.to_datetime(date_str)
        if dt.tzinfo is not None:
            return dt.tz_convert(TH_TZ).date()
        return dt.date()
    return pd.to_datetime(date_str).date()


# --- GitHub Config ---
TOKEN = st.secrets["GITHUB_TOKEN"]
REPO = st.secrets["GITHUB_REPO"]
BRANCH = st.secrets.get("GITHUB_BRANCH", "data")
FILE_PATH = st.secrets.get("EXCEL_PATH", "tasks_data.xlsx")

API_URL = f"https://api.github.com/repos/{REPO}/contents/{FILE_PATH}"
HEADERS = {
    "Authorization": f"Bearer {TOKEN}",
    "Accept": "application/vnd.github+json",
}

DEFAULT_TAGS = pd.DataFrame([
    {"Tag": "AML", "Color": "#8E44AD"},
    {"Tag": "FA Submit", "Color": "#E74C3C"},
    {"Tag": "General", "Color": "#2980B9"},
])
TASK_COLUMNS = ["id", "Task", "Tag", "PIC", "Due", "Status", "Note"]
EMPTY_TASKS = pd.DataFrame(columns=TASK_COLUMNS)

STATUS_LABELS = {
    "Pending": "🟠 Pending",
    "Done": "🟢 Done",
    "Late": "🔴 Late",
}
STATUS_OPTIONS = ["Pending", "Done", "Late"]


# --- GitHub Read / Write Functions ---
def normalize_tasks(df: pd.DataFrame) -> pd.DataFrame:
    for col in TASK_COLUMNS:
        if col not in df.columns:
            df[col] = ""
    df = df[TASK_COLUMNS].copy()
    if not df.empty:
        df["Due"] = pd.to_datetime(df["Due"], errors="coerce").dt.date
        df["id"] = df["id"].fillna("").astype(str).apply(lambda x: x if x else str(uuid.uuid4())[:8])
        df["Task"] = df["Task"].fillna("").astype(str)
        df["Tag"] = df["Tag"].fillna("General").astype(str)
        df["PIC"] = df["PIC"].fillna("").astype(str)
        df["Status"] = df["Status"].fillna("Pending").astype(str)
        df["Note"] = df["Note"].fillna("").astype(str)
    return df


def get_github_file():
    r = requests.get(API_URL, headers=HEADERS, params={"ref": BRANCH}, timeout=15)
    if r.status_code == 404:
        return EMPTY_TASKS.copy(), DEFAULT_TAGS.copy(), None
    if not r.ok:
        st.error(f"GitHub Read Error {r.status_code}: {r.text}")
        st.stop()
    data = r.json()
    raw_bytes = base64.b64decode(data["content"])
    excel = pd.ExcelFile(io.BytesIO(raw_bytes), engine="openpyxl")

    df_tasks = pd.read_excel(excel, sheet_name="Tasks") if "Tasks" in excel.sheet_names else EMPTY_TASKS.copy()
    df_tags = pd.read_excel(excel, sheet_name="Tags") if "Tags" in excel.sheet_names else DEFAULT_TAGS.copy()

    return normalize_tasks(df_tasks), df_tags, data["sha"]


def save_to_github(df_tasks: pd.DataFrame, df_tags: pd.DataFrame, commit_msg="Update tasks"):
    df_tasks = normalize_tasks(df_tasks)
    r_check = requests.get(API_URL, headers=HEADERS, params={"ref": BRANCH}, timeout=15)
    sha = r_check.json()["sha"] if r_check.status_code == 200 else None

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df_tasks.to_excel(writer, sheet_name="Tasks", index=False)
        df_tags.to_excel(writer, sheet_name="Tags", index=False)

    encoded = base64.b64encode(output.getvalue()).decode("utf-8")
    payload = {
        "message": commit_msg,
        "content": encoded,
        "branch": BRANCH,
    }
    if sha:
        payload["sha"] = sha

    r_put = requests.put(API_URL, headers=HEADERS, json=payload, timeout=15)
    if not r_put.ok:
        st.error(f"GitHub Save Error {r_put.status_code}: {r_put.text}")
        st.stop()

    st.session_state.df_tasks = df_tasks
    st.session_state.df_tags = df_tags
    st.session_state.data_ver = st.session_state.get("data_ver", 0) + 1


def text_color_for_bg(hex_color: str) -> str:
    try:
        hex_color = str(hex_color).lstrip("#")
        r, g, b = int(hex_color[0:2], 16), int(hex_color[2:4], 16), int(hex_color[4:6], 16)
        luminance = (0.299 * r + 0.587 * g + 0.114 * b)
        return "#000000" if luminance > 160 else "#FFFFFF"
    except Exception:
        return "#FFFFFF"


def get_effective_status(row) -> str:
    st_val = row["Status"]
    if st_val == "Done":
        return "Done"
    if st_val == "Late":
        return "Late"
    if pd.notna(row["Due"]) and row["Due"] < today_th():
        return "Late"
    return "Pending"


# --- Load Initial Data ---
if "df_tasks" not in st.session_state:
    with st.spinner("กำลังโหลดข้อมูลจาก GitHub..."):
        t_df, tag_df, _ = get_github_file()
        st.session_state.df_tasks = t_df
        st.session_state.df_tags = tag_df

st.session_state.df_tasks = normalize_tasks(st.session_state.df_tasks)
if "data_ver" not in st.session_state:
    st.session_state.data_ver = 0

df_tasks = st.session_state.df_tasks
df_tags = st.session_state.df_tags
tag_color_map = dict(zip(df_tags["Tag"], df_tags["Color"]))
tag_list = df_tags["Tag"].tolist()


# --- Header & Quick Add Bar ---
col_title, col_sync = st.columns([6, 1])
with col_title:
    st.title("📅 Task Manager & Calendar")
with col_sync:
    st.write("")
    if st.button("🔄 Sync GitHub", use_container_width=True):
        t_df, tag_df, _ = get_github_file()
        st.session_state.df_tasks = t_df
        st.session_state.df_tags = tag_df
        st.session_state.data_ver = st.session_state.get("data_ver", 0) + 1
        st.rerun()

with st.form("quick_add", clear_on_submit=True):
    c1, c2, c3, c4, c5 = st.columns([3.0, 1.6, 1.5, 1.5, 1.4])
    new_task = c1.text_input("ชื่องาน (Task)", placeholder="เช่น AML V6, MT1, FA Submit...")
    new_tag = c2.selectbox("Tag", tag_list if tag_list else ["General"])
    new_pic = c3.text_input("PIC (ผู้รับผิดชอบ)", placeholder="เช่น Ploy, Ae...")
    new_due = c4.date_input("Due Date", value=today_th())
    new_status = c5.selectbox("Status", STATUS_OPTIONS, format_func=lambda x: STATUS_LABELS[x])

    n1, n2 = st.columns([7.6, 1.4])
    new_note = n1.text_input("Note / รายละเอียดเพิ่มเติม (ไม่บังคับ)", placeholder="ใส่รายละเอียด, ลิงก์, หรือหมายเหตุ...")
    n2.write("")
    submitted = n2.form_submit_button("➕ เพิ่มงาน", type="primary", use_container_width=True)

    if submitted and new_task.strip():
        new_row = pd.DataFrame([{
            "id": str(uuid.uuid4())[:8],
            "Task": new_task.strip(),
            "Tag": new_tag,
            "PIC": new_pic.strip(),
            "Due": new_due,
            "Status": new_status,
            "Note": new_note.strip(),
        }])
        updated_tasks = pd.concat([df_tasks, new_row], ignore_index=True)
        save_to_github(updated_tasks, df_tags, f"Add task: {new_task.strip()}")
        st.toast("เพิ่มงานและเซฟลง GitHub แล้ว ✅")
        st.rerun()


# --- Main Navigation ---
page = st.radio(
    "เมนู",
    ["📋 Task List", "📅 Calendar", "⚙️ Tag & Color Settings"],
    horizontal=True,
    label_visibility="collapsed",
)
st.divider()


# === PAGE 1: TASK LIST ===
if page == "📋 Task List":
    if df_tasks.empty:
        st.info("ยังไม่มีงาน ลองเพิ่มงานจากช่องด้านบนได้เลยครับ")
    else:
        pic_options = sorted([p for p in df_tasks["PIC"].unique() if str(p).strip()])

        # Filter & Search Bar
        f1, f2, f3, f4, f5 = st.columns([2.2, 1.6, 1.6, 1.6, 1.5])
        search_q = f1.text_input("🔍 ค้นหา", placeholder="ค้นหางาน / PIC / Note...", label_visibility="collapsed")
        filter_tag = f2.multiselect("กรอง Tag", tag_list, placeholder="ทุก Tag", label_visibility="collapsed")
        filter_pic = f3.multiselect("กรอง PIC", pic_options, placeholder="ทุก PIC", label_visibility="collapsed")
        filter_status = f4.multiselect(
            "กรอง Status",
            STATUS_OPTIONS,
            format_func=lambda x: STATUS_LABELS[x],
            placeholder="ทุก Status",
            label_visibility="collapsed",
        )
        edit_mode = f5.toggle("แก้ตารางใหญ่", value=False)

        if edit_mode:
            edited = st.data_editor(
                df_tasks.sort_values("Due"),
                num_rows="dynamic",
                use_container_width=True,
                column_config={
                    "id": None,
                    "Task": st.column_config.TextColumn("Task", required=True),
                    "Tag": st.column_config.SelectboxColumn("Tag", options=tag_list, required=True),
                    "PIC": st.column_config.TextColumn("PIC"),
                    "Due": st.column_config.DateColumn("Due Date", format="YYYY-MM-DD", required=True),
                    "Status": st.column_config.SelectboxColumn("Status", options=STATUS_OPTIONS, required=True),
                    "Note": st.column_config.TextColumn("Note"),
                },
            )
            if st.button("💾 บันทึกการแก้ไขทั้งหมดลง GitHub", type="primary"):
                edited["id"] = edited["id"].fillna("").apply(lambda x: x if x else str(uuid.uuid4())[:8])
                save_to_github(edited.reset_index(drop=True), df_tags, "Bulk edit tasks")
                st.toast("บันทึกตารางลง GitHub แล้ว ✅")
                st.rerun()
        else:
            view_df = df_tasks.copy()
            view_df["EffStatus"] = view_df.apply(get_effective_status, axis=1)

            if search_q.strip():
                q = search_q.strip().lower()
                view_df = view_df[
                    view_df["Task"].str.lower().str.contains(q, na=False)
                    | view_df["PIC"].str.lower().str.contains(q, na=False)
                    | view_df["Note"].str.lower().str.contains(q, na=False)
                ]
            if filter_tag:
                view_df = view_df[view_df["Tag"].isin(filter_tag)]
            if filter_pic:
                view_df = view_df[view_df["PIC"].isin(filter_pic)]
            if filter_status:
                view_df = view_df[view_df["EffStatus"].isin(filter_status)]

            if view_df.empty:
                st.warning("ไม่พบงานตามเงื่อนไขที่กรองไว้")

            all_tags_ordered = tag_list + [t for t in view_df["Tag"].unique() if t not in tag_list]
            for tag_name in all_tags_ordered:
                group = view_df[view_df["Tag"] == tag_name].sort_values("Due")
                if group.empty:
                    continue

                bg = tag_color_map.get(tag_name, "#555555")
                fg = text_color_for_bg(bg)
                st.markdown(
                    f'<div style="background:{bg};color:{fg};padding:6px 14px;border-radius:6px;'
                    f'font-weight:700;margin-top:12px;margin-bottom:8px;display:inline-block;">'
                    f'{tag_name} ({len(group)})</div>',
                    unsafe_allow_html=True,
                )

                for _, row in group.iterrows():
                    tid = row["id"]
                    r1, r2, r_pic, r3, r4, r5, r6 = st.columns([0.4, 3.2, 1.3, 1.4, 1.6, 0.7, 0.6])

                    # 1. Quick Done Checkbox
                    is_done = row["Status"] == "Done"
                    checked = r1.checkbox("", value=is_done, key=f"chk_{tid}", label_visibility="collapsed")
                    if checked != is_done:
                        df_tasks.loc[df_tasks["id"] == tid, "Status"] = "Done" if checked else "Pending"
                        save_to_github(df_tasks, df_tags, f"Toggle status: {row['Task']}")
                        st.rerun()

                    # 2. Task Name + Note Preview
                    task_label = f"~~{row['Task']}~~" if is_done else f"**{row['Task']}**"
                    r2.markdown(task_label)
                    if str(row["Note"]).strip():
                        r2.caption(f"📝 {row['Note']}")

                    # 3. PIC
                    pic_text = str(row["PIC"]).strip()
                    r_pic.write(f"👤 {pic_text}" if pic_text else "-")

                    # 4. Due Date
                    due_str = str(row["Due"]) if pd.notna(row["Due"]) else "-"
                    r3.write(f"📅 {due_str}")

                    # 5. Quick Status Selector (Pending / Done / Late)
                    eff_st = row["EffStatus"]
                    current_idx = STATUS_OPTIONS.index(eff_st) if eff_st in STATUS_OPTIONS else 0
                    chosen_st = r4.selectbox(
                        "Status",
                        STATUS_OPTIONS,
                        index=current_idx,
                        format_func=lambda x: STATUS_LABELS[x],
                        key=f"st_{tid}_{st.session_state.data_ver}",
                        label_visibility="collapsed",
                    )
                    if chosen_st != eff_st:
                        df_tasks.loc[df_tasks["id"] == tid, "Status"] = chosen_st
                        save_to_github(df_tasks, df_tags, f"Change status: {row['Task']} -> {chosen_st}")
                        st.rerun()

                    # 6. Edit Popover
                    with r5.popover("✏️", use_container_width=True):
                        with st.form(key=f"edit_form_{tid}"):
                            st.markdown("**แก้ไขรายละเอียดงาน**")
                            e_task = st.text_input("ชื่องาน", value=row["Task"])
                            tag_opts = tag_list if row["Tag"] in tag_list else tag_list + [row["Tag"]]
                            e_tag = st.selectbox("Tag", tag_opts, index=tag_opts.index(row["Tag"]))
                            e_pic = st.text_input("PIC", value=row["PIC"])
                            e_due = st.date_input("Due Date", value=row["Due"] if pd.notna(row["Due"]) else today_th())
                            e_status = st.selectbox(
                                "Status",
                                STATUS_OPTIONS,
                                index=STATUS_OPTIONS.index(row["Status"]) if row["Status"] in STATUS_OPTIONS else 0,
                                format_func=lambda x: STATUS_LABELS[x],
                            )
                            e_note = st.text_area("Note", value=row["Note"], height=90)
                            if st.form_submit_button("💾 บันทึก", type="primary", use_container_width=True):
                                mask = df_tasks["id"] == tid
                                df_tasks.loc[mask, "Task"] = e_task.strip() or row["Task"]
                                df_tasks.loc[mask, "Tag"] = e_tag
                                df_tasks.loc[mask, "PIC"] = e_pic.strip()
                                df_tasks.loc[mask, "Due"] = e_due
                                df_tasks.loc[mask, "Status"] = e_status
                                df_tasks.loc[mask, "Note"] = e_note.strip()
                                save_to_github(df_tasks, df_tags, f"Edit task: {e_task.strip()}")
                                st.toast("บันทึกการแก้ไขแล้ว ✅")
                                st.rerun()

                    # 7. Delete Button
                    if r6.button("🗑️", key=f"del_{tid}", help="ลบงานนี้"):
                        df_tasks = df_tasks[df_tasks["id"] != tid].reset_index(drop=True)
                        save_to_github(df_tasks, df_tags, f"Delete task: {row['Task']}")
                        st.rerun()


# === PAGE 2: CALENDAR ===
elif page == "📅 Calendar":
    opt1, opt2 = st.columns([5, 2])
    color_mode = opt1.radio(
        "สีการ์ดในปฏิทิน",
        ["สีตาม Tag (Done เป็นสีเขียว)", "สีตาม Tag ทั้งหมด", "สีตาม Status (Done/Pending/Late)"],
        horizontal=True,
    )
    opt2.write("")
    show_pic_on_cal = opt2.checkbox("แสดงชื่อ PIC บนการ์ด", value=True)
    st.caption("💡 ลากการ์ดไปวางวันอื่นเพื่อเลื่อน Due Date ได้เลย หรือคลิกที่การ์ดเพื่อแก้ไขรายละเอียด")

    events = []
    for _, row in df_tasks.iterrows():
        if pd.isna(row["Due"]):
            continue
        eff_st = get_effective_status(row)
        prefix = {"Done": "✓ ", "Late": "⚠️ ", "Pending": ""}[eff_st]
        pic_str = f" [{row['PIC']}]" if show_pic_on_cal and str(row["PIC"]).strip() else ""
        note_icon = " 📝" if str(row["Note"]).strip() else ""

        if color_mode == "สีตาม Status (Done/Pending/Late)":
            bg = {"Done": "#27AE60", "Late": "#E74C3C", "Pending": "#F39C12"}[eff_st]
        elif color_mode == "สีตาม Tag (Done เป็นสีเขียว)" and eff_st == "Done":
            bg = "#27AE60"
        else:
            bg = tag_color_map.get(row["Tag"], "#3788D8")

        events.append({
            "id": str(row["id"]),
            "title": f"{prefix}{row['Task']}{pic_str}{note_icon}",
            "start": str(row["Due"]),
            "allDay": True,
            "backgroundColor": bg,
            "borderColor": bg,
            "textColor": text_color_for_bg(bg),
        })

    cal_options = {
        "initialView": "dayGridMonth",
        "initialDate": str(today_th()),
        "editable": True,
        "eventStartEditable": True,
        "eventDurationEditable": False,
        "headerToolbar": {
            "left": "prev,next today",
            "center": "title",
            "right": "dayGridMonth,listMonth",
        },
        "height": 680,
        "firstDay": 1,
    }
    cal_state = calendar(
        events=events,
        options=cal_options,
        callbacks=["eventClick", "eventChange"],
        key=f"main_calendar_{color_mode}_{show_pic_on_cal}_{st.session_state.data_ver}",
    )

    # ลากย้ายวันบนปฏิทิน -> อัปเดต Due Date ลง Task List & GitHub ทันที
    if cal_state and cal_state.get("callback") == "eventChange":
        changed_ev = cal_state.get("eventChange", {}).get("event", {})
        changed_id = str(changed_ev.get("id", ""))
        new_start = changed_ev.get("start")
        if changed_id and new_start:
            new_due = parse_cal_date(new_start)
            mask = df_tasks["id"] == changed_id
            if mask.any():
                old_due = df_tasks.loc[mask, "Due"].iloc[0]
                if old_due != new_due:
                    task_name = df_tasks.loc[mask, "Task"].iloc[0]
                    df_tasks.loc[mask, "Due"] = new_due
                    save_to_github(df_tasks, df_tags, f"Move task '{task_name}' to {new_due}")
                    st.toast(f"ย้าย '{task_name}' ไปวันที่ {new_due} แล้ว ✅")
                    st.rerun()

    # คลิกที่งานบนปฏิทินเพื่อดู Note / แก้ PIC / แก้ Status ได้ทันที
    if cal_state and cal_state.get("callback") == "eventClick":
        clicked_id = cal_state["eventClick"]["event"].get("id")
        matched = df_tasks[df_tasks["id"] == str(clicked_id)]
        if not matched.empty:
            row = matched.iloc[0]
            st.divider()
            st.subheader(f"✏️ จัดการงาน: {row['Task']}")
            with st.form(key=f"cal_edit_{clicked_id}"):
                cc1, cc2, cc3, cc4, cc5 = st.columns([2.6, 1.6, 1.5, 1.5, 1.6])
                ce_task = cc1.text_input("ชื่องาน", value=row["Task"])
                tag_opts = tag_list if row["Tag"] in tag_list else tag_list + [row["Tag"]]
                ce_tag = cc2.selectbox("Tag", tag_opts, index=tag_opts.index(row["Tag"]))
                ce_pic = cc3.text_input("PIC", value=row["PIC"])
                ce_due = cc4.date_input("Due Date", value=row["Due"] if pd.notna(row["Due"]) else today_th())
                eff_st = get_effective_status(row)
                ce_status = cc5.selectbox(
                    "Status",
                    STATUS_OPTIONS,
                    index=STATUS_OPTIONS.index(eff_st),
                    format_func=lambda x: STATUS_LABELS[x],
                )
                ce_note = st.text_area("Note", value=row["Note"], height=80)

                b1, b2, _ = st.columns([1.5, 1.2, 5])
                save_btn = b1.form_submit_button("💾 บันทึกการแก้ไข", type="primary", use_container_width=True)
                del_btn = b2.form_submit_button("🗑️ ลบงานนี้", use_container_width=True)

                if save_btn:
                    mask = df_tasks["id"] == str(clicked_id)
                    df_tasks.loc[mask, "Task"] = ce_task.strip() or row["Task"]
                    df_tasks.loc[mask, "Tag"] = ce_tag
                    df_tasks.loc[mask, "PIC"] = ce_pic.strip()
                    df_tasks.loc[mask, "Due"] = ce_due
                    df_tasks.loc[mask, "Status"] = ce_status
                    df_tasks.loc[mask, "Note"] = ce_note.strip()
                    save_to_github(df_tasks, df_tags, f"Edit from calendar: {ce_task.strip()}")
                    st.toast("อัปเดตงานจากปฏิทินแล้ว ✅")
                    st.rerun()

                if del_btn:
                    df_tasks = df_tasks[df_tasks["id"] != str(clicked_id)].reset_index(drop=True)
                    save_to_github(df_tasks, df_tags, f"Delete from calendar: {row['Task']}")
                    st.toast("ลบงานเรียบร้อยแล้ว 🗑️")
                    st.rerun()


# === PAGE 3: TAG & COLOR SETTINGS ===
elif page == "⚙️ Tag & Color Settings":
    st.subheader("เพิ่ม / แก้ไขสี Tag")
    with st.form("add_tag_form", clear_on_submit=True):
        s1, s2, s3 = st.columns([3, 1, 1.5])
        t_name = s1.text_input("ชื่อ Tag")
        t_color = s2.color_picker("เลือกสี", "#3498DB")
        add_t = s3.form_submit_button("💾 บันทึก Tag", type="primary", use_container_width=True)
        if add_t and t_name.strip():
            name_clean = t_name.strip()
            if name_clean in df_tags["Tag"].values:
                df_tags.loc[df_tags["Tag"] == name_clean, "Color"] = t_color
            else:
                df_tags = pd.concat([df_tags, pd.DataFrame([{"Tag": name_clean, "Color": t_color}])], ignore_index=True)
            save_to_github(df_tasks, df_tags, f"Update tag: {name_clean}")
            st.toast(f"เซฟ Tag '{name_clean}' ลง GitHub แล้ว ✅")
            st.rerun()

    st.divider()
    for idx, row in df_tags.iterrows():
        c1, c2, c3 = st.columns([3, 1.5, 1])
        bg = row["Color"]
        fg = text_color_for_bg(bg)
        c1.markdown(
            f'<span style="background:{bg};color:{fg};padding:5px 14px;border-radius:6px;font-weight:600;">'
            f'{row["Tag"]}</span>',
            unsafe_allow_html=True,
        )
        new_c = c2.color_picker("", bg, key=f"cp_{idx}_{row['Tag']}", label_visibility="collapsed")
        if new_c != bg:
            df_tags.loc[idx, "Color"] = new_c
            save_to_github(df_tasks, df_tags, f"Change color: {row['Tag']}")
            st.rerun()

        if c3.button("ลบ", key=f"del_tag_{idx}_{row['Tag']}"):
            df_tags = df_tags[df_tags["Tag"] != row["Tag"]].reset_index(drop=True)
            save_to_github(df_tasks, df_tags, f"Delete tag: {row['Tag']}")
            st.rerun()
