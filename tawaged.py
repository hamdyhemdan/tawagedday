import streamlit as st
import pandas as pd
from io import BytesIO
from datetime import datetime, time
import re
from openpyxl.styles import PatternFill

st.set_page_config(page_title="نظام الدمج السريع الشامل", layout="centered")

# --- الإعدادات: يمكنك تعديل أسماء الأعمدة من هنا بسهولة ---
CONFIG = {
    "MAIN": {
        "ID": "رقم البصمه",
        "DATE": "التاريخ"
    },
    "SOURCE": {
        "ID": "Badgenumber",
        "DATE": "Date",
        "NAME": "Name",
        "DEPT": "Department",
        "SHIFT": "SHIFT"
    }
}
# --------------------------------------------------------

def parse_manual_time(time_str):
    try:
        return datetime.strptime(time_str.strip(), "%H:%M").time()
    except:
        return None

def smart_read_excel(uploaded_file):
    try:
        return pd.read_excel(uploaded_file)
    except Exception as e:
        raise ValueError(f"لم نتمكن من قراءة الملف {uploaded_file.name}. الخطأ: {e}")

def main():
    st.title("🚀 نظام دمج وتلوين البصمة الشامل")
    
    main_file = st.file_uploader("1. ملف البصمة الأساسي (AC-LOG)", type=["xlsx", "xls"])
    source_file = st.file_uploader("2. ملف بيانات المصدر (الجدول)", type=["xlsx", "xls"])

    st.divider()
    tc1, tc2 = st.columns(2)
    with tc1: start_t_raw = st.text_input("من وقت:", value="08:00")
    with tc2: end_t_raw = st.text_input("إلى وقت:", value="08:15")

    if st.button("🚀 ابدأ المعالجة الآن"):
        if main_file and source_file:
            try:
                with st.spinner('جاري المعالجة...'):
                    df_main = smart_read_excel(main_file)
                    df_source = smart_read_excel(source_file)
                    
                    df_main.columns = df_main.columns.str.strip()
                    df_source.columns = df_source.columns.str.strip()
                    
                    # دالة تنظيف
                    clean_id = lambda v: str(int(float(v))).strip() if pd.notna(v) else ""
                    def extract_day(v):
                        if isinstance(v, (datetime, pd.Timestamp)): return str(v.day)
                        m = re.search(r'(\d+)', str(v))
                        return str(int(m.group(1))) if m else "NONE"

                    # تجهيز البيانات
                    df_main['M_ID'] = df_main[CONFIG["MAIN"]["ID"]].apply(clean_id)
                    df_main['M_DAY'] = df_main[CONFIG["MAIN"]["DATE"]].apply(extract_day)
                    
                    df_source['M_ID'] = df_source[CONFIG["SOURCE"]["ID"]].apply(clean_id)
                    df_source['M_DAY'] = df_source[CONFIG["SOURCE"]["DATE"]].apply(extract_day)

                    # الدمج
                    result_df = pd.merge(df_main, df_source.drop(columns=[CONFIG["SOURCE"]["ID"], CONFIG["SOURCE"]["DATE"]]), 
                                         on=['M_ID', 'M_DAY'], how='left')
                    
                    result_df = result_df.drop(columns=['M_ID', 'M_DAY'])

                    # حذف الأعمدة غير المرغوب فيها (اساسى والشيفت)
                    # تم إضافة SHIFT المذكور في طلبك
                    columns_to_drop = ["اساسى", "الشيفت", CONFIG["SOURCE"]["SHIFT"]]
                    result_df = result_df.drop(columns=[c for c in columns_to_drop if c in result_df.columns], errors='ignore')

                    # التلوين (لعمود الوقت)
                    output = BytesIO()
                    with pd.ExcelWriter(output, engine='openpyxl') as writer:
                        result_df.to_excel(writer, index=False, sheet_name='Sheet1')
                        worksheet = writer.sheets['Sheet1']
                        yellow_fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
                        
                        # تلوين عمود "الوقت" إذا وجد
                        if "الوقت" in result_df.columns:
                            col_idx = result_df.columns.get_loc("الوقت") + 1
                            start_t = parse_manual_time(start_t_raw)
                            end_t = parse_manual_time(end_t_raw)
                            
                            for row_num in range(2, len(result_df) + 2):
                                cell = worksheet.cell(row=row_num, column=col_idx)
                                matches = re.findall(r'(\d{1,2}):(\d{2})', str(cell.value))
                                for h, m in matches:
                                    t = time(int(h), int(m))
                                    if start_t and end_t and start_t <= t <= end_t:
                                        cell.fill = yellow_fill

                    st.success("✅ تم الانتهاء بنجاح!")
                    st.download_button("📥 تحميل الملف", output.getvalue(), "Report.xlsx")
            
            except Exception as e:
                st.error(f"خطأ: {e}")
        else:
            st.warning("الرجاء رفع الملفين")

if __name__ == "__main__":
    main()
