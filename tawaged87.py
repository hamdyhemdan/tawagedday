import streamlit as st
import pandas as pd
from io import BytesIO
from datetime import datetime, time
import re
from openpyxl.styles import PatternFill

st.set_page_config(page_title="نظام الدمج السريع الشامل", layout="centered")

# ... (دالة parse_manual_time ودالة smart_read_excel كما هما) ...
def parse_manual_time(time_str):
    try: return datetime.strptime(time_str.strip(), "%H:%M").time()
    except: return None

def smart_read_excel(uploaded_file):
    try:
        if uploaded_file.name.endswith('.xls'): return pd.read_excel(uploaded_file, engine='xlrd')
        else: return pd.read_excel(uploaded_file)
    except: return pd.read_excel(uploaded_file)

def main():
    st.title("🚀 نظام دمج وتلوين البصمة الشامل")
    main_file = st.file_uploader("1. ملف البصمة الأساسي (AC-LOG)", type=["xlsx", "xls"])
    source_file = st.file_uploader("2. ملف بيانات المصدر (الجدول)", type=["xlsx", "xls"])

    if st.button("🚀 ابدأ المعالجة الآن"):
        if main_file and source_file:
            try:
                df_main = smart_read_excel(main_file)
                df_source = smart_read_excel(source_file)
                
                df_main.columns = df_main.columns.str.strip()
                df_source.columns = df_source.columns.str.strip()

                # --- السطور المضافة لتوحيد الأسماء (بدون تغيير الأساسيات) ---
                df_source = df_source.rename(columns={"Badgenumber": "رقم البصمه", "Date": "التاريخ"})
                df_main = df_main.rename(columns={"Badgenumber": "رقم البصمه", "Date": "التاريخ"})
                # --------------------------------------------------------
                
                col_id_1, col_date_1 = "رقم البصمه", "التاريخ"
                col_id_2, col_date_2 = "رقم البصمه", "التاريخ"
                target_color_col = "الوقت"
                
                # --- باقي الكود الأصلي الخاص بك ---
                clean_id = lambda v: str(int(float(v))).strip() if pd.notna(v) else ""
                def extract_day(v):
                    if isinstance(v, (datetime, pd.Timestamp)): return str(v.day)
                    m = re.search(r'(\d+)', str(v))
                    return str(int(m.group(1))) if m else "NONE"

                df_main['M_ID'] = df_main[col_id_1].apply(clean_id)
                df_main['M_DAY'] = df_main[col_date_1].apply(extract_day)
                df_source['M_ID'] = df_source[col_id_2].apply(clean_id)
                df_source['M_DAY'] = df_source[col_date_2].apply(extract_day)

                cols_to_bring = [c for c in df_source.columns if c not in ['M_ID', 'M_DAY', col_id_2, col_date_2]]
                df_src_subset = df_source[['M_ID', 'M_DAY'] + cols_to_bring].drop_duplicates(subset=['M_ID', 'M_DAY'])
                result_df = pd.merge(df_main, df_src_subset, on=['M_ID', 'M_DAY'], how='left')
                result_df = result_df.drop(columns=['M_ID', 'M_DAY'])

                columns_to_drop = ["اساسى", "الشيفت", "SHIFT"]
                for col in columns_to_drop:
                    if col in result_df.columns:
                        result_df = result_df.drop(columns=[col])

                # (استمر في كود التلوين والتحميل كما كان)
                st.success("✅ تم الدمج بنجاح!")
            except Exception as e:
                st.error(f"خطأ: {e}")

if __name__ == "__main__":
    main()
