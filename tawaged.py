import streamlit as st
import pandas as pd
from io import BytesIO
from datetime import datetime, time
import re
from openpyxl.styles import PatternFill

st.set_page_config(page_title="نظام الدمج السريع الشامل", layout="centered")

def parse_manual_time(time_str):
    try:
        return datetime.strptime(time_str.strip(), "%H:%M").time()
    except:
        return None

def smart_read_excel(uploaded_file):
    try:
        if uploaded_file.name.endswith('.xls'):
            return pd.read_excel(uploaded_file, engine='xlrd')
        else:
            return pd.read_excel(uploaded_file)
    except Exception as e:
        try:
            return pd.read_excel(uploaded_file)
        except:
            raise ValueError(f"لم نتمكن من قراءة الملف {uploaded_file.name}")

def main():
    st.title("🚀 نظام دمج وتلوين البصمة الشامل")
    
    main_file = st.file_uploader("1. ملف البصمة الأساسي (AC-LOG)", type=["xlsx", "xls"])
    source_file = st.file_uploader("2. ملف بيانات المصدر (الجدول)", type=["xlsx", "xls"])

    if st.button("🚀 ابدأ المعالجة الآن"):
        if main_file and source_file:
            try:
                df_main = smart_read_excel(main_file)
                df_source = smart_read_excel(source_file)
                
                # تنظيف أسماء الأعمدة
                df_main.columns = df_main.columns.str.strip()
                df_source.columns = df_source.columns.str.strip()
                
                # التحقق من وجود الأعمدة المطلوبة
                # إذا ظهر لك خطأ، انظر إلى الأسماء المطبوعة في الرسالة
                required_cols = ["Badgenumber", "Date"]
                for col in required_cols:
                    if col not in df_source.columns:
                        st.error(f"❌ خطأ: العمود '{col}' غير موجود في ملف المصدر.")
                        st.write("الأعمدة المتاحة حالياً في ملفك هي:", list(df_source.columns))
                        return

                # --- باقي الكود الأصلي ---
                col_id_1, col_date_1 = "رقم البصمه", "التاريخ"
                col_id_2, col_date_2 = "Badgenumber", "Date"
                target_color_col = "الوقت"
                
                clean_id = lambda v: str(int(float(v))).strip() if pd.notna(v) else ""
                def extract_day(v):
                    if isinstance(v, (datetime, pd.Timestamp)): return str(v.day)
                    m = re.search(r'(\d+)', str(v))
                    return str(int(m.group(1))) if m else "NONE"

                df_main['M_ID'] = df_main[col_id_1].apply(clean_id)
                df_main['M_DAY'] = df_main[col_date_1].apply(extract_day)
                df_source['M_ID'] = df_source[col_id_2].apply(clean_id)
                df_source['M_DAY'] = df_source[col_date_2].apply(extract_day)

                result_df = pd.merge(df_main, df_source.drop(columns=[col_id_2, col_date_2]), on=['M_ID', 'M_DAY'], how='left')
                result_df = result_df.drop(columns=['M_ID', 'M_DAY'])

                columns_to_drop = ["اساسى", "الشيفت", "SHIFT"]
                for col in columns_to_drop:
                    if col in result_df.columns:
                        result_df = result_df.drop(columns=[col])

                st.success("✅ تم الدمج بنجاح!")
                st.dataframe(result_df.head()) # لعرض جزء من النتيجة
                
            except Exception as e:
                st.error(f"حدث خطأ: {e}")
        else:
            st.warning("الرجاء رفع الملفين")

if __name__ == "__main__":
    main()
