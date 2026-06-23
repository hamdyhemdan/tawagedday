import streamlit as st
import pandas as pd
from io import BytesIO
from datetime import datetime, time
import re
from openpyxl.styles import PatternFill

st.set_page_config(page_title="نظام الدمج السريع", layout="centered")

def parse_manual_time(time_str):
    try: return datetime.strptime(time_str.strip(), "%H:%M").time()
    except: return None

def smart_read_excel(uploaded_file):
    return pd.read_excel(uploaded_file)

def main():
    st.title("🚀 نظام دمج البصمة المرن")
    
    main_file = st.file_uploader("1. ملف البصمة الأساسي", type=["xlsx", "xls"])
    source_file = st.file_uploader("2. ملف بيانات المصدر", type=["xlsx", "xls"])

    if main_file and source_file:
        df_main = smart_read_excel(main_file)
        df_source = smart_read_excel(source_file)
        
        # تنظيف الأعمدة
        df_main.columns = df_main.columns.str.strip()
        df_source.columns = df_source.columns.str.strip()

        st.divider()
        st.subheader("⚙️ اختر أسماء الأعمدة الموجودة في ملفاتك")
        
        col1, col2 = st.columns(2)
        with col1:
            id_main = st.selectbox("عمود رقم البصمة (الملف 1):", df_main.columns)
            date_main = st.selectbox("عمود التاريخ (الملف 1):", df_main.columns)
        with col2:
            id_source = st.selectbox("عمود رقم البصمة (الملف 2):", df_source.columns)
            date_source = st.selectbox("عمود التاريخ (الملف 2):", df_source.columns)

        if st.button("🚀 ابدأ المعالجة الآن"):
            try:
                # تنظيف البيانات
                clean_id = lambda v: str(int(float(v))).strip() if pd.notna(v) else ""
                def extract_day(v):
                    if isinstance(v, (datetime, pd.Timestamp)): return str(v.day)
                    m = re.search(r'(\d+)', str(v))
                    return str(int(m.group(1))) if m else "NONE"

                df_main['M_ID'] = df_main[id_main].apply(clean_id)
                df_main['M_DAY'] = df_main[date_main].apply(extract_day)
                df_source['M_ID'] = df_source[id_source].apply(clean_id)
                df_source['M_DAY'] = df_source[date_source].apply(extract_day)

                # دمج
                result_df = pd.merge(df_main, df_source.drop(columns=[id_source, date_source]), 
                                     on=['M_ID', 'M_DAY'], how='left')
                result_df = result_df.drop(columns=['M_ID', 'M_DAY'])

                # حذف أعمدة
                cols_to_drop = ["اساسى", "الشيفت", "SHIFT"]
                result_df = result_df.drop(columns=[c for c in cols_to_drop if c in result_df.columns], errors='ignore')

                # تصدير
                output = BytesIO()
                result_df.to_excel(output, index=False)
                st.success("✅ تم الدمج بنجاح!")
                st.download_button("📥 تحميل الملف المدمج", output.getvalue(), "Combined_Report.xlsx")
                
            except Exception as e:
                st.error(f"خطأ: {e}")

if __name__ == "__main__":
    main()
