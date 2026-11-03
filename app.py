"""
Group Meeting Slot Booking App (Streamlit)

- Shows every 15-minute slot (Week 1 + Week 2) built from the office hours.
- A group picks its name and clicks a free slot -> the slot shows "Booked".
- One slot per group. A group can cancel and pick again.
- Admin (password) can cancel any booking, reset everything, download Excel.

Storage:
- If Google Sheets credentials are in .streamlit/secrets.toml -> uses Google Sheet
  (recommended when hosted on Streamlit Cloud, data is never lost).
- Otherwise -> uses a local file bookings.xlsx (good for testing on your PC).
"""

import io
import os
import threading
from datetime import datetime, timedelta

import pandas as pd
import streamlit as st

# ---------------------------------------------------------------------------
# 1. SETTINGS  (taken from your Excel: "Group Meeting Schedule")
# ---------------------------------------------------------------------------
TITLE = "Group Meeting Slot Booking"
NUM_GROUPS = 34
SLOT_MINUTES = 15
WEEKS = ["Week 1", "Week 2"]

# Day -> office hours (24-hour clock)
OFFICE_HOURS = {
    "Sunday":    ("11:45", "13:35"),
    "Monday":    ("10:10", "11:40"),
    "Tuesday":   ("11:15", "13:35"),
    "Wednesday": ("10:10", "11:40"),
    "Thursday":  ("11:45", "13:35"),
}

GROUPS = [f"Group {i}" for i in range(1, NUM_GROUPS + 1)]
COLUMNS = ["slot_id", "week", "day", "time", "group", "booked_at"]
LOCAL_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "bookings.xlsx")


def fmt12(t: datetime) -> str:
    """13:15 -> '1:15' (same style as your sheet)."""
    h = t.hour % 12 or 12
    return f"{h}:{t.minute:02d}"


def build_slots() -> pd.DataFrame:
    rows = []
    for week in WEEKS:
        for day, (start, end) in OFFICE_HOURS.items():
            t = datetime.strptime(start, "%H:%M")
            end_t = datetime.strptime(end, "%H:%M")
            while t + timedelta(minutes=SLOT_MINUTES) <= end_t:
                t2 = t + timedelta(minutes=SLOT_MINUTES)
                rows.append({
                    "slot_id": f"{week[-1]}-{day[:3]}-{t.strftime('%H%M')}",
                    "week": week,
                    "day": day,
                    "time": f"{fmt12(t)} - {fmt12(t2)}",
                })
                t = t2
    return pd.DataFrame(rows)


SLOTS = build_slots()


# ---------------------------------------------------------------------------
# 2. STORAGE  (Google Sheet or local Excel)
# ---------------------------------------------------------------------------
def use_gsheet() -> bool:
    try:
        return "gcp_service_account" in st.secrets and "sheet_url" in st.secrets
    except Exception:
        return False


@st.cache_resource
def get_worksheet():
    import gspread
    from google.oauth2.service_account import Credentials

    creds = Credentials.from_service_account_info(
        dict(st.secrets["gcp_service_account"]),
        scopes=["https://www.googleapis.com/auth/spreadsheets"],
    )
    sh = gspread.authorize(creds).open_by_url(st.secrets["sheet_url"])
    try:
        ws = sh.worksheet("bookings")
    except gspread.WorksheetNotFound:
        ws = sh.add_worksheet("bookings", rows=200, cols=len(COLUMNS))
    if ws.row_values(1) != COLUMNS:
        ws.clear()
        ws.append_row(COLUMNS)
    return ws


@st.cache_resource
def file_lock():
    return threading.Lock()


def load_bookings() -> pd.DataFrame:
    if use_gsheet():
        records = get_worksheet().get_all_records()
        df = pd.DataFrame(records, columns=COLUMNS)
    elif os.path.exists(LOCAL_FILE):
        df = pd.read_excel(LOCAL_FILE, dtype=str)
    else:
        df = pd.DataFrame(columns=COLUMNS)
    return df.fillna("").astype(str)


def save_local(df: pd.DataFrame):
    df[COLUMNS].to_excel(LOCAL_FILE, index=False)


def add_booking(slot: dict, group: str) -> tuple[bool, str]:
    """Book a slot. Re-checks right before saving so nothing is double booked."""
    row = {**{k: slot[k] for k in ["slot_id", "week", "day", "time"]},
           "group": group, "booked_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S")}

    if use_gsheet():
        df = load_bookings()
        if slot["slot_id"] in df["slot_id"].values:
            return False, "Sorry, this slot was just booked by another group."
        if group in df["group"].values:
            return False, f"{group} already has a slot. Cancel it first to change."
        ws = get_worksheet()
        ws.append_row([row[c] for c in COLUMNS])
        # Safety check: if two groups clicked at the same moment, first one wins.
        df = load_bookings()
        same = df[df["slot_id"] == slot["slot_id"]]
        if len(same) > 1 and same.iloc[0]["group"] != group:
            delete_rows(lambda r: r["slot_id"] == slot["slot_id"] and r["group"] == group)
            return False, "Sorry, another group booked this slot a moment before you."
        return True, "Booked!"

    with file_lock():
        df = load_bookings()
        if slot["slot_id"] in df["slot_id"].values:
            return False, "Sorry, this slot was just booked by another group."
        if group in df["group"].values:
            return False, f"{group} already has a slot. Cancel it first to change."
        df = pd.concat([df, pd.DataFrame([row])], ignore_index=True)
        save_local(df)
    return True, "Booked!"


def delete_rows(match):
    """Delete every booking where match(row) is True."""
    if use_gsheet():
        ws = get_worksheet()
        records = ws.get_all_records()
        # delete from bottom so row numbers stay correct (row 1 = header)
        for i in reversed(range(len(records))):
            if match({k: str(v) for k, v in records[i].items()}):
                ws.delete_rows(i + 2)
        return
    with file_lock():
        df = load_bookings()
        keep = df[~df.apply(lambda r: match(r.to_dict()), axis=1)] if len(df) else df
        save_local(keep)


def reset_all():
    if use_gsheet():
        ws = get_worksheet()
        ws.clear()
        ws.append_row(COLUMNS)
    else:
        with file_lock():
            save_local(pd.DataFrame(columns=COLUMNS))


def schedule_excel(bookings: pd.DataFrame) -> bytes:
    """Full schedule (all slots + who booked) as an Excel file."""
    full = SLOTS.merge(bookings[["slot_id", "group"]], on="slot_id", how="left")
    full["group"] = full["group"].fillna("")
    full["status"] = full["group"].apply(lambda g: "Booked" if g else "Available")
    out = full[["week", "day", "time", "status", "group"]]
    out.columns = ["Week", "Day", "Time", "Status", "Group"]
    by_group = pd.DataFrame({"Group": GROUPS}).merge(
        full[full["group"] != ""].rename(columns={"group": "Group", "week": "Week",
                                                  "day": "Day", "time": "Time"}),
        on="Group", how="left")[["Group", "Week", "Day", "Time"]].fillna("Not booked")
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as xw:
        out.to_excel(xw, sheet_name="All slots", index=False)
        by_group.to_excel(xw, sheet_name="By group", index=False)
        for ws in xw.book.worksheets:
            for col in ws.columns:
                ws.column_dimensions[col[0].column_letter].width = 16
    return buf.getvalue()


# ---------------------------------------------------------------------------
# 3. PAGE
# ---------------------------------------------------------------------------
import base64

UNIVERSITY = "East West University"
NAVY = "#0E2D6E"
LOGO_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "assets", "logo.png")


def logo_b64() -> str:
    try:
        with open(LOGO_PATH, "rb") as f:
            return base64.b64encode(f.read()).decode()
    except OSError:
        return ""


st.set_page_config(page_title=TITLE, page_icon=LOGO_PATH if os.path.exists(LOGO_PATH) else "📅",
                   layout="wide")

st.markdown(f"""
<style>
@import url('https://fonts.googleapis.com/css2?family=Poppins:wght@400;500;600;700&display=swap');
html, body, [class*="css"], .stMarkdown, button, input {{ font-family: 'Poppins', sans-serif; }}
.stApp {{ background: #F4F6FB; }}
.block-container {{ padding-top: 1.2rem; max-width: 1200px; }}
header[data-testid="stHeader"] {{ background: transparent; }}

/* ---------- header ---------- */
.hero {{ background:#fff; border-radius:18px; padding:18px 26px; display:flex; align-items:center;
  gap:24px; box-shadow:0 4px 18px rgba(14,45,110,.08); border-bottom:5px solid {NAVY}; }}
.hero img {{ height:86px; }}
.hero .divider {{ width:2px; align-self:stretch; background:#E3E8F3; }}
.hero h1 {{ color:{NAVY}; font-size:1.75rem; font-weight:700; margin:0; padding:0; line-height:1.2; }}
.hero p {{ color:#5B6785; margin:4px 0 0; font-size:.95rem; }}
.pill {{ display:inline-block; background:#EAF0FB; color:{NAVY}; border-radius:999px;
  padding:3px 12px; font-size:.78rem; font-weight:600; margin:8px 6px 0 0; }}

/* ---------- stat cards ---------- */
.stats {{ display:grid; grid-template-columns:repeat(3,1fr); gap:14px; margin:18px 0 6px; }}
.stat {{ background:#fff; border-radius:14px; padding:14px 18px; box-shadow:0 2px 10px rgba(14,45,110,.06); }}
.stat .n {{ font-size:1.7rem; font-weight:700; color:{NAVY}; line-height:1.1; }}
.stat .l {{ color:#6B7590; font-size:.82rem; }}
.bar {{ height:8px; background:#E3E8F3; border-radius:99px; overflow:hidden; margin-top:10px; }}
.bar > div {{ height:100%; background:linear-gradient(90deg,{NAVY},#3B6FD8); border-radius:99px; }}

/* ---------- section titles ---------- */
.step {{ display:flex; align-items:center; gap:10px; margin:22px 0 8px; }}
.step .num {{ background:{NAVY}; color:#fff; width:30px; height:30px; border-radius:50%;
  display:flex; align-items:center; justify-content:center; font-weight:700; font-size:.9rem; }}
.step .t {{ font-size:1.15rem; font-weight:600; color:#1D2740; }}

/* ---------- ticket (my booking) ---------- */
.ticket {{ background:linear-gradient(120deg,{NAVY} 0%,#2453B3 100%); color:#fff; border-radius:16px;
  padding:16px 22px; display:flex; align-items:center; gap:18px; box-shadow:0 6px 18px rgba(14,45,110,.25); }}
.ticket .ic {{ font-size:2rem; }}
.ticket .k {{ opacity:.8; font-size:.8rem; text-transform:uppercase; letter-spacing:.06em; }}
.ticket .v {{ font-size:1.25rem; font-weight:700; }}
.hint {{ background:#FFF8E6; border:1px solid #FFE1A1; color:#8A5A00; border-radius:12px;
  padding:10px 16px; font-size:.9rem; }}

/* ---------- legend ---------- */
.legend {{ display:flex; gap:18px; flex-wrap:wrap; font-size:.85rem; color:#4A5470; margin:4px 0 10px; }}
.legend span {{ display:inline-flex; align-items:center; gap:6px; }}
.dot {{ width:14px; height:14px; border-radius:4px; display:inline-block; }}

/* ---------- day columns + slots ---------- */
.dayhead {{ background:{NAVY}; color:#fff; border-radius:12px 12px 0 0; text-align:center;
  padding:8px 4px 6px; margin-bottom:10px; }}
.dayhead b {{ display:block; font-size:.98rem; }}
.dayhead small {{ opacity:.8; font-size:.75rem; }}
.slot {{ border-radius:10px; height:42px; display:flex; align-items:center; justify-content:center;
  font-size:.8rem; margin-bottom:12px; white-space:nowrap; overflow:hidden; }}
.slot-booked {{ background:#EEF0F4; color:#8A93A8; border:1px dashed #C9CFDB; }}
.slot-mine {{ background:#FFC93C; color:#4A3500; font-weight:700; border:2px solid #E0A800; }}
div[class*="st-key-slot_"] button {{ height:42px; border-radius:10px; background:#E9F8EF;
  border:1px solid #9ED9B4; color:#11643A; font-size:.8rem; font-weight:500; transition:all .15s; }}
div[class*="st-key-slot_"] button:hover {{ background:#11643A; color:#fff; border-color:#11643A;
  transform:translateY(-1px); box-shadow:0 4px 10px rgba(17,100,58,.25); }}
div[data-testid="stTabs"] button p {{ font-size:1rem; font-weight:600; }}
.foot {{ text-align:center; color:#8A93A8; font-size:.8rem; margin:30px 0 10px; }}

@media (max-width: 760px) {{
  .hero {{ flex-direction:column; text-align:center; gap:10px; }}
  .hero .divider {{ display:none; }}
  .stats {{ grid-template-columns:1fr; }}
}}
</style>
""", unsafe_allow_html=True)

# ---------- header ----------
logo = logo_b64()
logo_html = f"<img src='data:image/png;base64,{logo}'/>" if logo else ""
st.markdown(f"""
<div class="hero">
  {logo_html}
  <div class="divider"></div>
  <div>
    <h1>{TITLE}</h1>
    <p>Choose a 15-minute meeting slot for your group during office hours.</p>
    <span class="pill">👥 {NUM_GROUPS} groups</span>
    <span class="pill">⏱ {SLOT_MINUTES} min each</span>
    <span class="pill">📆 2 weeks · Sun–Thu</span>
  </div>
</div>
""", unsafe_allow_html=True)

bookings = load_bookings()
booked = dict(zip(bookings["slot_id"], bookings["group"]))
group_slot = dict(zip(bookings["group"], bookings["slot_id"]))
n_booked_groups = len(group_slot)
pct = int(100 * n_booked_groups / NUM_GROUPS)

st.markdown(f"""
<div class="stats">
  <div class="stat"><div class="n">{n_booked_groups} / {NUM_GROUPS}</div>
    <div class="l">Groups booked</div><div class="bar"><div style="width:{pct}%"></div></div></div>
  <div class="stat"><div class="n">{len(SLOTS) - len(booked)}</div><div class="l">Free slots left</div></div>
  <div class="stat"><div class="n">{NUM_GROUPS - n_booked_groups}</div><div class="l">Groups still to book</div></div>
</div>
""", unsafe_allow_html=True)


# ---------- confirm dialog ----------
@st.dialog("Confirm your booking")
def confirm_dialog(slot: dict, group: str):
    st.markdown(f"""
    <div class="ticket" style="margin-bottom:14px">
      <div class="ic">📌</div>
      <div><div class="k">{group}</div>
      <div class="v">{slot['week']} · {slot['day']}</div>
      <div>{slot['time']}</div></div>
    </div>""", unsafe_allow_html=True)
    st.write("Is this the slot you want? Other groups will see it as **booked**.")
    a, b = st.columns(2)
    if a.button("✅ Yes, book it", type="primary", use_container_width=True):
        ok, msg = add_booking(slot, group)
        st.session_state["flash"] = ("success" if ok else "error", msg)
        st.rerun()
    if b.button("Go back", use_container_width=True):
        st.rerun()


@st.dialog("Cancel your booking?")
def cancel_dialog(group: str):
    st.write(f"This will free the slot of **{group}** so another group can take it.")
    a, b = st.columns(2)
    if a.button("Yes, cancel", type="primary", use_container_width=True):
        delete_rows(lambda r: r["group"] == group)
        st.session_state["flash"] = ("info", "Booking cancelled. You can choose a new slot.")
        st.rerun()
    if b.button("Keep it", use_container_width=True):
        st.rerun()


flash = st.session_state.pop("flash", None)
if flash:
    kind, msg = flash
    if kind == "success":
        st.balloons()
        st.toast("🎉 " + msg)
    else:
        getattr(st, kind)(msg)

# ---------- step 1 ----------
st.markdown("<div class='step'><div class='num'>1</div><div class='t'>Select your group</div></div>",
            unsafe_allow_html=True)
group = st.selectbox("Your group", GROUPS, index=None, placeholder="🔎  Type or choose your group…",
                     label_visibility="collapsed")

if group and group in group_slot:
    s = SLOTS[SLOTS["slot_id"] == group_slot[group]].iloc[0]
    colA, colB = st.columns([5, 1.3], vertical_alignment="center")
    colA.markdown(f"""
    <div class="ticket"><div class="ic">🎟️</div>
      <div><div class="k">{group} · your meeting</div>
      <div class="v">{s['week']} · {s['day']} · {s['time']}</div></div></div>""",
                  unsafe_allow_html=True)
    if colB.button("Change / cancel", use_container_width=True):
        cancel_dialog(group)
elif group:
    st.markdown(f"<div class='hint'>👇 <b>{group}</b> has no slot yet. Click any "
                f"<b style='color:#11643A'>green</b> slot below to book it.</div>",
                unsafe_allow_html=True)

# ---------- step 2 ----------
st.markdown("<div class='step'><div class='num'>2</div><div class='t'>Pick a free slot</div></div>",
            unsafe_allow_html=True)
lg, rf = st.columns([5, 1], vertical_alignment="center")
lg.markdown("""<div class="legend">
  <span><i class="dot" style="background:#E9F8EF;border:1px solid #9ED9B4"></i>Free (click to book)</span>
  <span><i class="dot" style="background:#EEF0F4;border:1px dashed #C9CFDB"></i>Booked</span>
  <span><i class="dot" style="background:#FFC93C;border:1px solid #E0A800"></i>Your group</span>
</div>""", unsafe_allow_html=True)
if rf.button("🔄 Refresh", use_container_width=True):
    st.rerun()

for tab, week in zip(st.tabs([f"📅 {w}" for w in WEEKS]), WEEKS):
    with tab:
        cols = st.columns(len(OFFICE_HOURS))
        for col, day in zip(cols, OFFICE_HOURS):
            with col:
                day_slots = SLOTS[(SLOTS["week"] == week) & (SLOTS["day"] == day)]
                free = sum(sid not in booked for sid in day_slots["slot_id"])
                st.markdown(f"<div class='dayhead'><b>{day}</b><small>{free} of {len(day_slots)} free"
                            f"</small></div>", unsafe_allow_html=True)
                for _, s in day_slots.iterrows():
                    owner = booked.get(s["slot_id"])
                    if owner:
                        if owner == group:
                            st.markdown(f"<div class='slot slot-mine'>⭐ {s['time']}</div>",
                                        unsafe_allow_html=True)
                        else:
                            st.markdown(f"<div class='slot slot-booked'>🔒 {s['time']} · "
                                        f"{owner.replace('Group ', 'G')}</div>", unsafe_allow_html=True)
                    elif st.button(s["time"], key=f"slot_{s['slot_id']}", use_container_width=True):
                        if group is None:
                            st.toast("⚠️ Please select your group first (step 1).")
                        elif group in group_slot:
                            st.toast(f"⚠️ {group} already has a slot. Use 'Change / cancel' first.")
                        else:
                            confirm_dialog(s.to_dict(), group)

st.markdown(f"<div class='foot'>{UNIVERSITY} · Group meeting scheduler</div>", unsafe_allow_html=True)

# ---------- admin ----------
with st.sidebar:
    if logo:
        st.image(LOGO_PATH, width=140)
    st.header("🔒 Admin")
    try:
        admin_pw = st.secrets.get("admin_password", "admin123")
    except Exception:
        admin_pw = "admin123"
    pw = st.text_input("Password", type="password")
    if pw and pw == admin_pw:
        st.success("Admin mode")
        st.download_button("⬇️ Download schedule (Excel)", schedule_excel(bookings),
                           file_name="group_meeting_schedule.xlsx", use_container_width=True)
        if len(bookings):
            opts = {f"{r.group} → {r.week}, {r.day}, {r.time}": r.slot_id
                    for r in bookings.itertuples()}
            pick = st.selectbox("Cancel a booking", list(opts))
            if st.button("Cancel selected", use_container_width=True):
                sid = opts[pick]
                delete_rows(lambda r: r["slot_id"] == sid)
                st.rerun()
        st.divider()
        confirm = st.checkbox("I want to delete ALL bookings")
        if st.button("Reset all", disabled=not confirm, use_container_width=True):
            reset_all()
            st.rerun()
        st.caption("Storage: " + ("Google Sheet" if use_gsheet() else "local bookings.xlsx"))
    elif pw:
        st.error("Wrong password")
