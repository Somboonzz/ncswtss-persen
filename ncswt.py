import streamlit as st
import pandas as pd
import altair as alt
import datetime
import os
import pytz  # สำหรับโซนเวลา

st.set_page_config(page_title="HR Dashboard", layout="wide")

# -----------------------------
# โซนเวลาไทย
# -----------------------------
bangkok_tz = pytz.timezone("Asia/Bangkok")

def thai_date(dt):
    return dt.strftime(f"%d/%m/{dt.year + 543}")

thai_months = [
    "มกราคม", "กุมภาพันธ์", "มีนาคม", "เมษายน", "พฤษภาคม", "มิถุนายน",
    "กรกฎาคม", "สิงหาคม", "กันยายน", "ตุลาคม", "พฤศจิกายน", "ธันวาคม"
]

def format_thai_month(period):
    year = period.year + 543
    month = thai_months[period.month - 1]
    return f"{month} {year}"

# -----------------------------
# โหลดไฟล์ Excel
# -----------------------------
@st.cache_data(ttl=300)  # cache 5 นาที (300 วินาที)
def load_data(file_path="attendances.xlsx"):
    try:
        if file_path and os.path.exists(file_path):
            df = pd.read_excel(file_path, engine='openpyxl', dtype={'เข้างาน': str, 'ออกงาน': str})
            return df
        else:
            st.warning("❌ ไม่พบไฟล์ Excel: attendances.xlsx")
            return pd.DataFrame()
    except Exception as e:
        st.error(f"❌ อ่านไฟล์ Excel ไม่ได้: {e}")
        return pd.DataFrame()

df = load_data()

# -----------------------------
# ปุ่ม Refresh
# -----------------------------
if st.button("🔄 Refresh ข้อมูล (Manual)"):
    st.cache_data.clear()  # ล้าง cache
    st.rerun()

# -----------------------------
# นาฬิกา (แสดงเวลาตอนที่โหลดหน้าเว็บ)
# -----------------------------
bangkok_now = datetime.datetime.now(pytz.utc).astimezone(bangkok_tz)
st.markdown(
    f"<div style='text-align:right; font-size:50px; color:#FF5733; font-weight:bold;'>"
    f"🗓 {thai_date(bangkok_now)}  |  ⏰ {bangkok_now.strftime('%H:%M:%S')}</div>",
    unsafe_allow_html=True
)

# -----------------------------
# แสดง dashboard ถ้ามีข้อมูล
# -----------------------------
if not df.empty:
    for col in ["ชื่อ-สกุล", "แผนก", "ข้อยกเว้น"]:
        if col in df.columns:
            df[col] = df[col].astype(str).str.strip().str.replace(r"\s+", " ", regex=True)

    if "แผนก" in df.columns:
        df["แผนก"] = df["แผนก"].replace({"nan": "ไม่ระบุ", "": "ไม่ระบุ"})

    if "วันที่" in df.columns:
        df["วันที่"] = pd.to_datetime(df["วันที่"], errors='coerce')
        df["ปี"] = df["วันที่"].dt.year + 543
        df["เดือน"] = df["วันที่"].dt.to_period("M")
    
    if 'เข้างาน' in df.columns:
        df['เวลาเข้า'] = pd.to_datetime(df['เข้างาน'], format='%H:%M:%S', errors='coerce').dt.time
        df['เวลาเข้า'] = df['เวลาเข้า'].apply(lambda t: datetime.time(0, 0) if pd.isna(t) else t)
    if 'ออกงาน' in df.columns:
        df['เวลาออก'] = pd.to_datetime(df['ออกงาน'], format='%H:%M:%S', errors='coerce').dt.time
        df['เวลาออก'] = df['เวลาออก'].apply(lambda t: datetime.time(0, 0) if pd.isna(t) else t)

    df_filtered = df.copy()

    # --- Filter ปี
    years = ["-- แสดงทั้งหมด --"] + sorted(df["ปี"].dropna().unique(), reverse=True)
    selected_year = st.selectbox("📆 เลือกปี", years)
    if selected_year != "-- แสดงทั้งหมด --":
        df_filtered = df_filtered[df_filtered["ปี"] == int(selected_year)]

    # --- Filter เดือน
    if "เดือน" in df_filtered.columns and not df_filtered.empty:
        available_months = sorted(df_filtered["เดือน"].dropna().unique())
        month_options = ["-- แสดงทั้งหมด --"] + [format_thai_month(m) for m in available_months]
        selected_month = st.selectbox("📅 เลือกเดือน", month_options)
        if selected_month != "-- แสดงทั้งหมด --":
            mapping = {format_thai_month(m): str(m) for m in available_months}
            selected_period = mapping[selected_month]
            df_filtered = df_filtered[df_filtered["เดือน"].astype(str) == selected_period]

    # --- Filter แผนก
    departments = ["-- แสดงทั้งหมด --"] + sorted(df_filtered["แผนก"].dropna().unique())
    selected_dept = st.selectbox("🏢 เลือกแผนก", departments)
    if selected_dept != "-- แสดงทั้งหมด --":
        df_filtered = df_filtered[df_filtered["แผนก"] == selected_dept]

    # --- คำนวณประเภทการลา
    def leave_days(row):
        if "ครึ่งวัน" in str(row):
            return 0.5
        return 1

    df_filtered["ลาป่วย/ลากิจ"] = df_filtered["ข้อยกเว้น"].apply(
        lambda x: leave_days(x) if str(x) in ["ลาป่วย", "ลากิจ", "ลาป่วยครึ่งวัน", "ลากิจครึ่งวัน"] else 0
    )
    df_filtered["ขาด"] = df_filtered["ข้อยกเว้น"].apply(
        lambda x: leave_days(x) if str(x) in ["ขาด", "ขาดครึ่งวัน"] else 0
    )
    df_filtered["สาย"] = df_filtered["ข้อยกเว้น"].apply(lambda x: 1 if str(x) == "สาย" else 0)
    df_filtered["ลาพักผ่อน"] = df_filtered["ข้อยกเว้น"].apply(lambda x: 1 if str(x) == "ลาพักผ่อน" else 0)

    leave_types = ["ลาป่วย/ลากิจ", "ขาด", "สาย", "ลาพักผ่อน"]
    summary = df_filtered.groupby(["ชื่อ-สกุล", "แผนก"])[leave_types].sum().reset_index()

    st.title("📊 แดชบอร์ดการลา / ขาด / สาย")

    # --- ตัวกรองพนักงาน (คงค่าที่เลือกไว้) ---
    if 'selected_employee' not in st.session_state:
        st.session_state.selected_employee = '-- แสดงทั้งหมด --'

    all_names = ["-- แสดงทั้งหมด --"] + sorted(summary["ชื่อ-สกุล"].unique())
    
    selected_employee = st.selectbox(
        "🔍 ค้นหาชื่อพนักงาน",
        all_names,
        key='selected_employee',  # ใช้ key เพื่อจัดการ state ของ widget
    )

    colors = {
        "ลาป่วย/ลากิจ": "#FFC300",
        "ขาด": "#C70039",
        "สาย": "#FF5733",
        "พักผ่อน": "#33C4FF"
    }

    tabs = st.tabs(leave_types)
    for t, leave in zip(tabs, leave_types):
        with t:
            st.subheader(f"🏆 จัดอันดับ {leave}")

            # --- กรองข้อมูลตามชื่อที่เลือก ---
            if selected_employee != "-- แสดงทั้งหมด --":
                summary_filtered = summary[summary["ชื่อ-สกุล"] == selected_employee].reset_index(drop=True)
                person_data_full = df_filtered[df_filtered["ชื่อ-สกุล"] == selected_employee].reset_index(drop=True)
            else:
                summary_filtered = summary
                person_data_full = df_filtered

            # --- แสดงข้อมูลสรุป (เหมือนต้นฉบับ) ---
            st.markdown("### 📌 สรุปข้อมูล")
            st.dataframe(summary_filtered, use_container_width=True)

            # --- แสดงรายละเอียดวันลา (เมื่อเลือกพนักงาน) ---
            if selected_employee != "-- แสดงทั้งหมด --" and not person_data_full.empty:
                if leave == "ลาป่วย/ลากิจ":
                    relevant_exceptions = ["ลาป่วย", "ลากิจ", "ลาป่วยครึ่งวัน", "ลากิจครึ่งวัน"]
                elif leave == "ขาด":
                    relevant_exceptions = ["ขาด", "ขาดครึ่งวัน"]
                else:
                    relevant_exceptions = [leave]

                dates = person_data_full.loc[
                    person_data_full["ข้อยกเว้น"].isin(relevant_exceptions), ["วันที่", "เวลาเข้า", "เวลาออก", "ข้อยกเว้น"]
                ]

                if not dates.empty:
                    total_days = dates["ข้อยกเว้น"].apply(leave_days).sum()
                    with st.expander(f"ดูรายละเอียดวันที่ "):
                        for _, row in dates.iterrows():
                            entry_time = row['เวลาเข้า'].strftime('%H:%M')
                            exit_time = row['เวลาออก'].strftime('%H:%M')
                            label = f"• {row['วันที่'].strftime('%d/%m/%Y')} &nbsp;&nbsp; {entry_time} - {exit_time} &nbsp;&nbsp;&nbsp;&nbsp; {row['ข้อยกเว้น']}"
                            st.markdown(label, unsafe_allow_html=True)

            # --- ตารางอันดับ (เหมือนต้นฉบับ) ---
            ranking = summary_filtered[["ชื่อ-สกุล", "แผนก", leave]].sort_values(by=leave, ascending=False).reset_index(drop=True)
            ranking.insert(0, "อันดับ", range(1, len(ranking) + 1))
            
            ranking_display = ranking[ranking[leave] > 0] # กรองคนที่ไม่มียอดออก

            st.markdown("### 🏅 ตารางอันดับ")
            st.dataframe(ranking_display, use_container_width=True)

            # --- กราฟ (เหมือนต้นฉบับ) ---
            if not ranking_display.empty:
                chart_data = ranking_display if selected_employee != "-- แสดงทั้งหมด --" else ranking_display.head(20)
                chart_title = f"ข้อมูล {leave}" if selected_employee != "-- แสดงทั้งหมด --" else f"20 อันดับแรกของพนักงานที่ '{leave}'"

                chart = (
                    alt.Chart(chart_data)
                    .mark_bar(cornerRadius=5, color=colors.get(leave, "#C70039"))
                    .encode(
                        x=alt.X(leave + ":Q", title=f"จำนวน ({'วัน' if leave != 'สาย' else 'ครั้ง'})"),
                        y=alt.Y("ชื่อ-สกุล:N", sort="-x", title="ชื่อ-สกุล"),
                        tooltip=["อันดับ", "ชื่อ-สกุล", "แผนก", leave],
                    )
                    .properties(title=chart_title)
                )
                st.altair_chart(chart, use_container_width=True)
            else:
                st.info(f"ไม่พบข้อมูล '{leave}' ในช่วงเวลาที่เลือก")
else:
    st.info("กรุณาตรวจสอบว่ามีไฟล์ attendances.xlsx อยู่ในโฟลเดอร์เดียวกับโปรแกรม")

