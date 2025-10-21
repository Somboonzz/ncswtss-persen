import streamlit as st
import pandas as pd
import altair as alt
import datetime
import os
import pytz # สำหรับโซนเวลา

# ----------------------------------------------------------------------------------
# ตั้งค่าหน้า และ CSS
# ----------------------------------------------------------------------------------
st.set_page_config(page_title="HR Dashboard", layout="wide")
st.title("📊 แดชบอร์ดการลา / ขาด / สาย")
st.markdown("""
    <style>
        /* CSS สำหรับจัดตารางให้อยู่ตรงกลาง */
        div[data-testid="stDataframeHeader"] div {
            text-align: center !important;
            vertical-align: middle !important;
            justify-content: center !important;
        }
        
        div[data-testid="stDataframeCell"] {
            text-align: center !important;
            justify-content: center !important;
        }

        .stDataFrame {
            margin-left: 1rem;
            margin-right: 1rem;
        }
        /* Style สำหรับรายละเอียดวันลาที่ถูกปรับปรุง */
        .detail-row {
            display: flex;
            align-items: center;
            margin-bottom: 5px;
            font-size: 14px;
        }
        .detail-header {
            font-weight: bold;
            border-bottom: 1px solid #ddd;
            padding-bottom: 2px;
            margin-bottom: 5px;
        }
        .col-date {
            width: 120px;
        }
        .col-time {
            width: 150px;
            font-weight: bold;
        }
        .col-type {
            font-weight: bold;
        }
    </style>
    """, unsafe_allow_html=True)

# ----------------------------------------------------------------------------------
# โซนเวลาไทย และฟังก์ชันช่วย
# ----------------------------------------------------------------------------------
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

def parse_time(time_str):
    """พยายามแปลง string หรือ datetime.time ให้เป็น datetime.time object"""
    try:
        if pd.notna(time_str):
            if isinstance(time_str, datetime.time):
                return time_str
            # ลองแปลงเป็น H:M:S
            dt_obj = pd.to_datetime(str(time_str), format='%H:%M:%S', errors='ignore')
            if not pd.isna(dt_obj):
                return dt_obj.time()
    except:
        try:
            # ลองแปลงเป็น H:M
            dt_obj = pd.to_datetime(str(time_str), format='%H:%M', errors='ignore')
            if not pd.isna(dt_obj):
                return dt_obj.time()
        except:
            pass
    return None

# ----------------------------------------------------------------------------------
# โหลดไฟล์ Excel (Caching Level 1)
# ----------------------------------------------------------------------------------
@st.cache_data(ttl=300, show_spinner="⏳ กำลังโหลดไฟล์ข้อมูล...")
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

df_raw = load_data()

# ----------------------------------------------------------------------------------
# ประมวลผลและคำนวณสรุปทั้งหมด (Caching Level 2: ส่วนสำคัญที่ทำให้เร็วขึ้นมาก)
# ----------------------------------------------------------------------------------
@st.cache_data(ttl=3600, show_spinner="⚙️ กำลังประมวลผลข้อมูลและคำนวณสรุป...")
def preprocess_and_calculate_summary(df):
    if df.empty:
        return pd.DataFrame(), pd.DataFrame()

    df_base = df.copy()

    # --- การเตรียมข้อมูลเบื้องต้น ---
    for col in ["ชื่อ-สกุล", "แผนก", "ข้อยกเว้น"]:
        if col in df_base.columns:
            df_base[col] = df_base[col].astype(str).str.strip().str.replace(r"\s+", " ", regex=True)

    if "แผนก" in df_base.columns:
        df_base["แผนก"] = df_base["แผนก"].replace({"nan": "ไม่ระบุ", "": "ไม่ระบุ"})

    if "วันที่" in df_base.columns:
        df_base["วันที่"] = pd.to_datetime(df_base["วันที่"], errors='coerce')
        df_base.dropna(subset=['วันที่'], inplace=True)
        if df_base.empty: return pd.DataFrame(), pd.DataFrame() # เช็คซ้ำหลัง dropna
        
        df_base["ปี"] = df_base["วันที่"].dt.year + 543
        df_base["เดือน"] = df_base["วันที่"].dt.to_period("M")
    
    # แปลงเวลาเข้า/ออก
    if 'เข้างาน' in df_base.columns:
        df_base['เวลาเข้า'] = df_base['เข้างาน'].apply(parse_time)
    
    if 'ออกงาน' in df_base.columns:
        df_base['เวลาออก'] = df_base['ออกงาน'].apply(parse_time)

    # --- คำนวณประเภทการลา ---
    def leave_days(row):
        """คำนวณจำนวนวันลา (1 วันเต็ม หรือ 0.5 วัน)"""
        if "ครึ่งวัน" in str(row):
            return 0.5
        return 1

    df_base["ลาป่วย/ลากิจ"] = df_base["ข้อยกเว้น"].apply(
        lambda x: leave_days(x) if str(x) in ["ลาป่วย", "ลากิจ", "ลาป่วยครึ่งวัน", "ลากิจครึ่งวัน"] else 0
    )
    df_base["ขาด"] = df_base["ข้อยกเว้น"].apply(
        lambda x: leave_days(x) if str(x) in ["ขาด", "ขาดครึ่งวัน"] else 0
    )
    df_base["สาย"] = df_base["ข้อยกเว้น"].apply(lambda x: 1 if str(x) == "สาย" else 0)

    # 3 ประเภทหลัก
    leave_types = ["ลาป่วย/ลากิจ", "ขาด", "สาย"] 
    summary = df_base.groupby(["ชื่อ-สกุล", "แผนก"])[leave_types].sum().reset_index()

    return df_base, summary

df, summary = preprocess_and_calculate_summary(df_raw)

# ----------------------------------------------------------------------------------
# ปุ่ม Refresh และแสดงเวลา
# ----------------------------------------------------------------------------------
if st.button("🔄 Refresh ข้อมูล (Manual)"):
    # ล้างแคชทั้งหมด: ทำให้ load_data และ preprocess_and_calculate_summary ทำงานใหม่
    st.cache_data.clear() 
    st.rerun()

bangkok_now = datetime.datetime.now(pytz.utc).astimezone(bangkok_tz)
st.markdown(
    f"<div style='text-align:right; font-size:50px; color:#FF5733; font-weight:bold;'>"
    f"🗓 {thai_date(bangkok_now)} | ⏰ {bangkok_now.strftime('%H:%M:%S')}</div>",
    unsafe_allow_html=True
)

# ----------------------------------------------------------------------------------
# Dashboard หลัก (การกรองข้อมูล)
# ----------------------------------------------------------------------------------
if not df.empty:
    
    df_filtered = df.copy()
    summary_filtered = summary.copy()

    # --- Filter (จัดเรียงใน 3 คอลัมน์) ---
    col1, col2, col3 = st.columns(3)
    
    with col1:
        years = ["-- แสดงทั้งหมด --"] + sorted(df["ปี"].dropna().unique(), reverse=True)
        selected_year = st.selectbox("📆 เลือกปี", years)
        if selected_year != "-- แสดงทั้งหมด --":
            df_filtered = df_filtered[df_filtered["ปี"] == int(selected_year)]

    with col2:
        if "เดือน" in df_filtered.columns and not df_filtered.empty:
            available_months = sorted(df_filtered["เดือน"].dropna().unique())
            month_options = ["-- แสดงทั้งหมด --"] + [format_thai_month(m) for m in available_months]
            selected_month = st.selectbox("📅 เลือกเดือน", month_options)
            if selected_month != "-- แสดงทั้งหมด --":
                mapping = {format_thai_month(m): str(m) for m in available_months}
                selected_period = mapping[selected_month]
                df_filtered = df_filtered[df_filtered["เดือน"].astype(str) == selected_period]

    with col3:
        departments = ["-- แสดงทั้งหมด --"] + sorted(df_filtered["แผนก"].dropna().unique())
        selected_dept = st.selectbox("🏢 เลือกแผนก", departments)
        if selected_dept != "-- แสดงทั้งหมด --":
            df_filtered = df_filtered[df_filtered["แผนก"] == selected_dept]
    # -----------------------------------

    # --- คำนวณ summary ใหม่อีกครั้งหลังจากการกรอง ปี/เดือน/แผนก ---
    leave_types = ["ลาป่วย/ลากิจ", "ขาด", "สาย"]
    if not df_filtered.empty:
        summary_filtered = df_filtered.groupby(["ชื่อ-สกุล", "แผนก"])[leave_types].sum().reset_index()
    else:
        summary_filtered = pd.DataFrame(columns=["ชื่อ-สกุล", "แผนก"] + leave_types)
        
    # ----------------------------------------------------------------------------------
    # --- ตัวกรองพนักงาน ---
    if 'selected_employee' not in st.session_state:
        st.session_state.selected_employee = '-- แสดงทั้งหมด --'
    all_names = ["-- แสดงทั้งหมด --"] + sorted(summary_filtered["ชื่อ-สกุล"].unique())
    selected_employee = st.selectbox("🔍 ค้นหาชื่อพนักงาน", all_names, key='selected_employee')

    # --- สีกราฟ ---
    colors = {
        "ลาป่วย/ลากิจ": "#06BD1699", 
        "ขาด": "#C40FE9", 
        "สาย": "#0BCBEC", 
    }

    # ----------------------------------------------------------------------------------
    # ส่วนที่ 2: Tabs จัดอันดับ (***ส่วนนี้คือส่วนที่แก้ไข***)
    # ----------------------------------------------------------------------------------
    tabs = st.tabs(leave_types)
    for t, leave in zip(tabs, leave_types):
        with t:
            # กรองข้อมูลตามชื่อที่เลือก (ใช้ summary_filtered ที่อัปเดตแล้ว)
            if selected_employee != "-- แสดงทั้งหมด --":
                current_summary_display = summary_filtered[summary_filtered["ชื่อ-สกุล"] == selected_employee].reset_index(drop=True)
                person_data_full = df_filtered[df_filtered["ชื่อ-สกุล"] == selected_employee].reset_index(drop=True)
            else:
                current_summary_display = summary_filtered.reset_index(drop=True)
                person_data_full = df_filtered.reset_index(drop=True)

            # --- แสดงข้อมูลสรุป ---
            st.markdown("### 📌 สรุปข้อมูลรายบุคคล")
            st.dataframe(current_summary_display, use_container_width=True, hide_index=True)


            # --- แสดงรายละเอียดวันลา (เมื่อเลือกพนักงาน) ---
            if selected_employee != "-- แสดงทั้งหมด --" and not person_data_full.empty:
                if leave == "ลาป่วย/ลากิจ":
                    relevant_exceptions = ["ลาป่วย", "ลากิจ", "ลาป่วยครึ่งวัน", "ลากิจครึ่งวัน"]
                elif leave == "ขาด":
                    relevant_exceptions = ["ขาด", "ขาดครึ่งวัน"]
                else: # สาย
                    relevant_exceptions = [leave]

                dates = person_data_full.loc[
                    person_data_full["ข้อยกเว้น"].isin(relevant_exceptions), ["วันที่", "เวลาเข้า", "เวลาออก", "ข้อยกเว้น"]
                ].sort_values(by="วันที่", ascending=False)
                
                # คำนวณยอดรวมวันลา/ขาด/สาย
                total_days = dates.apply(lambda row: 0.5 if "ครึ่งวัน" in str(row['ข้อยกเว้น']) else 1, axis=1).sum()
                if leave == "สาย":
                    total_count = dates.shape[0]

                if not dates.empty:
                    with st.expander(f"ดูรายละเอียดวันที่ ({leave})"):
                        
                        for _, row in dates.iterrows():
                            # การจัดการวันที่และเวลา
                            date_str = row['วันที่'].strftime('%d/%m/%Y')
                            exception_text = row['ข้อยกเว้น']

                            entry_time_raw = row['เวลาเข้า']
                            exit_time_raw = row['เวลาออก']
                            
                            if exception_text in ["ลาป่วย", "ลากิจ", "ขาด"]:
                                time_period = '00:00 - 00:00'
                            else:
                                entry_time = entry_time_raw.strftime('%H:%M') if entry_time_raw is not None else '00:00'
                                exit_time = exit_time_raw.strftime('%H:%M') if exit_time_raw is not None else '00:00'
                                time_period = f"{entry_time} - {exit_time}"
                            
                            label = (
                                f"<div style='margin-bottom: 5px; font-size: 16px;'>"
                                f"• {date_str}&nbsp;&nbsp;&nbsp;&nbsp;{time_period}&nbsp;&nbsp;&nbsp;&nbsp;{exception_text}"
                                f"</div>"
                            )
                            st.markdown(label, unsafe_allow_html=True)
                            
                        st.markdown("<hr>", unsafe_allow_html=True)
                        
                        if leave == "สาย":
                            st.markdown(f"<div style='font-weight: bold; font-size: 16px; margin-top: 10px;'>ยอดรวม: {total_count} ครั้ง</div>", unsafe_allow_html=True)
                        else:
                            st.markdown(f"<div style='font-weight: bold; font-size: 16px; margin-top: 10px;'>ยอดรวม: {total_days:.0f} วัน</div>", unsafe_allow_html=True)

            # --- ตารางอันดับ ---
            ranking = current_summary_display[["ชื่อ-สกุล", "แผนก", leave]].sort_values(by=leave, ascending=False).reset_index(drop=True)
            ranking.insert(0, "อันดับ", range(1, len(ranking) + 1))
            
            if not ranking.empty:
                ranking = ranking.iloc[:-1]
                
            # ย้าย subheader มาไว้ตรงนี้ตามที่ต้องการ
            st.subheader(f"🏆 จัดอันดับ {leave}")
            st.dataframe(ranking, use_container_width=True, hide_index=True)

    # ----------------------------------------------------------------------------------
    # Pie Chart
    # ----------------------------------------------------------------------------------
    st.markdown("---")
    
    # --- 1. เตรียมข้อมูลสำหรับ Pie Chart โดยอิงตามตัวกรองทั้งหมด ---
    data_for_pie_chart = summary_filtered.copy()
    pie_chart_title = "สัดส่วนภาพรวม Pie Chart" # หัวข้อเริ่มต้น

    # ถ้ามีการเลือกพนักงาน ให้กรองข้อมูลและเปลี่ยนหัวข้อ
    if selected_employee != "-- แสดงทั้งหมด --":
        data_for_pie_chart = data_for_pie_chart[data_for_pie_chart["ชื่อ-สกุล"] == selected_employee]
        pie_chart_title = f"สัดส่วนของ: {selected_employee}"

    st.subheader(f"🥧 {pie_chart_title}")
    
    # --- 2. คำนวณยอดรวมจากข้อมูลที่กรองแล้ว ---
    total_summary = data_for_pie_chart[leave_types].sum().reset_index()
    total_summary.columns = ['ประเภท', 'ยอดรวม']
    total_summary = total_summary[total_summary['ยอดรวม'] > 0].reset_index(drop=True)

    # --- 3. สร้างกราฟ (เหมือนเดิม) ---
    if total_summary['ยอดรวม'].sum() > 0:
        total = total_summary['ยอดรวม'].sum()
        total_summary['Percentage'] = (total_summary['ยอดรวม'] / total * 100).round(1)
        total_summary['label'] = total_summary.apply(lambda x: f"{x['ประเภท']} ({x['Percentage']}%)", axis=1)

        base = alt.Chart(total_summary).encode(
            theta=alt.Theta("ยอดรวม", stack=True),
            color=alt.Color(
                "ประเภท",
                scale=alt.Scale(domain=list(colors.keys()), range=list(colors.values()))
            ),
            order=alt.Order("ยอดรวม", sort="descending"), 
            tooltip=[
                "ประเภท",
                alt.Tooltip("ยอดรวม", format=".1f", title="จำนวน (วัน/ครั้ง)"),
                alt.Tooltip("Percentage", format=".1f", title="เปอร์เซ็นต์ (%)")
            ]
        )

        pie = base.mark_arc(outerRadius=130, innerRadius=60)

        text_labels = base.mark_text(
            radius= 230,
            size=20, 
            fontWeight="bold",
        ).encode(
            text=alt.Text('label:N'),
            color=alt.Color('ประเภท:N', scale=alt.Scale(domain=list(colors.keys()), range=list(colors.values()))),
            tooltip=alt.value(None)
        )

        center_text_val = "100%" if total > 0 else "0%"
        center_text = alt.Chart(pd.DataFrame({'text': [f"รวม {center_text_val}"]})).mark_text(
            size=20, color='black', fontWeight='bold'
        ).encode(text='text:N')

        chart = pie + text_labels + center_text
        
        chart = chart.properties(
            width=500,
            height=500
        ).configure_legend(
            titleFontSize=14,
            labelFontSize=12
        )

        st.altair_chart(chart, use_container_width=True)

    else:
        st.info("ไม่พบข้อมูลเพื่อแสดงในแผนภูมิวงกลม ตามตัวกรองที่เลือก")

else:
    st.info("กรุณาตรวจสอบว่ามีไฟล์ attendances.xlsx อยู่ในโฟลเดอร์เดียวกับโปรแกรม")