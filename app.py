import base64
import io
import uuid
from datetime import date

import pandas as pd
import requests
import streamlit as st
from streamlit_calendar import calendar

st.set_page_config(page_title="Task & Calendar", page_icon="📅", layout="wide")

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
EMPTY_TASKS = pd.DataFrame(columns=["id", "Task", "Tag", "Due", "Status"])


# --- GitHub Read / Write Functions ---
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

    if not df_tasks.empty:
        df_tasks["Due"] = pd.to_datetime(df_tasks["Due"], errors="coerce").dt.date
        df_tasks["id"] = df_tasks["id"].astype(str)
    return df_tasks, df_tags, data["sha"]


def save_to_github(df_tasks: pd.DataFrame, df_tags: pd.DataFrame, commit_msg="Update tasks"):
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


def text_color_for_bg(hex_color: str) -> str:
    try:
        hex_color = str(hex_color).lstrip("#")
        r, g, b = int(hex_color[0:2], 16), int(hex_color[2:4], 16), int(hex_color[4:6], 16)
        luminance = (0.299 * r + 0.587 * g + 0.114 * b)
        return "#000000" if luminance > 160 else "#FFFFFF"
    except Exception:
        return "#FFFFFF"


def get_display_status(row) -> str:
    if row["Status"] == "Done":
        return "Done"
    if pd.notna(row["Due"]) and row["Due"] < date.today():
        return "Late"
    return "Pending"


# --- Load Initial Data ---
if "df_tasks" not in st.session_state:
    with st.spinner("กำลังโหลดข้อมูลจาก GitHub..."):
        t_df, tag_df, _ = get_github_file()
        st.session_state.df_tasks = t_df
        st.session_state.df_tags = tag_df

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
        st.rerun()

with st.form("quick_add", clear_on_submit=True):
    c1, c2, c3, c4 = st.columns([4, 2, 2, 1.2])
    new_task = c1.text_input("ชื่องาน (Task)", placeholder="เช่น AML V6, MT1, FA Submit...")
    new_tag = c2.selectbox("Tag", tag_list if tag_list else ["General"])
    new_due = c3.date_input("Due Date", value=date.today())
    submitted = c4.form_submit_button("➕ เพิ่มงาน", type="primary", use_container_width=True)

    if submitted and new_task.strip():
        new_row = pd.DataFrame([{
            "id": str(uuid.uuid4())[:8],
            "Task": new_task.strip(),
            "Tag": new_tag,
            "Due": new_due,
            "Status": "Pending",
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
        f1, f2 = st.columns([2, 6])
        hide_done = f1.checkbox("ซ่อนงานที่ Done แล้ว", value=False)
        edit_mode = f2.toggle("โหมดแก้ไขตาราง (Bulk Edit)", value=False)

        if edit_mode:
            edited = st.data_editor(
                df_tasks.sort_values("Due"),
                num_rows="dynamic",
                use_container_width=True,
                column_config={
                    "id": None,
                    "Tag": st.column_config.SelectboxColumn("Tag", options=tag_list, required=True),
                    "Due": st.column_config.DateColumn("Due Date", format="YYYY-MM-DD", required=True),
                    "Status": st.column_config.SelectboxColumn("Status", options=["Pending", "Done"], required=True),
                },
            )
            if st.button("💾 บันทึกการแก้ไขทั้งหมดลง GitHub", type="primary"):
                edited["id"] = edited["id"].fillna("").apply(lambda x: x if x else str(uuid.uuid4())[:8])
                save_to_github(edited.reset_index(drop=True), df_tags, "Bulk edit tasks")
                st.toast("บันทึกตารางลง GitHub แล้ว ✅")
                st.rerun()
        else:
            view_df = df_tasks.copy()
            view_df["DisplayStatus"] = view_df.apply(get_display_status, axis=1)
            if hide_done:
                view_df = view_df[view_df["Status"] != "Done"]

            for tag_name in tag_list + [t for t in view_df["Tag"].unique() if t not in tag_list]:
                group = view_df[view_df["Tag"] == tag_name].sort_values("Due")
                if group.empty:
                    continue

                bg = tag_color_map.get(tag_name, "#555555")
                fg = text_color_for_bg(bg)
                st.markdown(
                    f'<div style="background:{bg};color:{fg};padding:6px 14px;border-radius:6px;'
                    f'font-weight:700;margin-top:12px;margin-bottom:6px;display:inline-block;">'
                    f'{tag_name} ({len(group)})</div>',
                    unsafe_allow_html=True,
                )

                for _, row in group.iterrows():
                    r1, r2, r3, r4, r5 = st.columns([0.5, 4.5, 2, 1.5, 0.8])
                    is_done = row["Status"] == "Done"
                    checked = r1.checkbox("", value=is_done, key=f"chk_{row['id']}", label_visibility="collapsed")
                    if checked != is_done:
                        df_tasks.loc[df_tasks["id"] == row["id"], "Status"] = "Done" if checked else "Pending"
                        save_to_github(df_tasks, df_tags, f"Toggle status: {row['Task']}")
                        st.rerun()

                    task_label = f"~~{row['Task']}~~" if is_done else f"**{row['Task']}**"
                    r2.markdown(task_label)
                    r3.write(f"📅 {row['Due']}")

                    st_badge = {"Done": "🟢 Done", "Late": "🔴 Late", "Pending": "🟠 Pending"}[row["DisplayStatus"]]
                    r4.write(st_badge)

                    if r5.button("🗑️", key=f"del_{row['id']}", help="ลบงานนี้"):
                        df_tasks = df_tasks[df_tasks["id"] != row["id"]].reset_index(drop=True)
                        save_to_github(df_tasks, df_tags, f"Delete task: {row['Task']}")
                        st.rerun()


# === PAGE 2: CALENDAR ===
elif page == "📅 Calendar":
    color_mode = st.radio(
        "สีการ์ดในปฏิทิน",
        ["สีตาม Tag (Done เป็นสีเขียว)", "สีตาม Tag ทั้งหมด"],
        horizontal=True,
    )

    events = []
    for _, row in df_tasks.iterrows():
        if pd.isna(row["Due"]):
            continue
        disp_st = get_display_status(row)
        prefix = {"Done": "✓ ", "Late": "⚠️ ", "Pending": ""}[disp_st]

        if color_mode == "สีตาม Tag (Done เป็นสีเขียว)" and disp_st == "Done":
            bg = "#27AE60"
        else:
            bg = tag_color_map.get(row["Tag"], "#3788D8")

        events.append({
            "title": f"{prefix}{row['Task']}",
            "start": str(row["Due"]),
            "allDay": True,
            "backgroundColor": bg,
            "borderColor": bg,
            "textColor": text_color_for_bg(bg),
        })

    cal_options = {
        "initialView": "dayGridMonth",
        "headerToolbar": {
            "left": "prev,next today",
            "center": "title",
            "right": "dayGridMonth,listMonth",
        },
        "height": 680,
        "firstDay": 1,
    }
    calendar(
        events=events,
        options=cal_options,
        key=f"main_calendar_{color_mode}_{len(events)}",
    )


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
