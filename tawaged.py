import streamlit as st
import pandas as pd
from io import BytesIO
from datetime import datetime, time
import re
from openpyxl.styles import PatternFill

st.set_page_config(page_title="نظام الدمج السريع الشامل", layout="centered")

def parse_manual_time(time_str):
    try: return datetime.strptime(time_str.strip(), "%H:%M").time()
    except: return None

def main():
    st.title("🚀 نظام دمج وتلوين البصمة الشامل")
    main_file = st.file_uploader("1. ملف البصمة الأساسي", type=["xlsx", "xls"])
    source_file = st.file_uploader("2. ملف بيانات المصدر", type=["xlsx", "xls"])

    if st.button("🚀 ابدأ المعالجة"):
        if main_file and source_file:
            try:
                df_main = pd.read_excel(main_file)
                df_source = pd.read_excel(source_file)
                
                df_main.columns = [str(c).strip() for c in df_main.columns]
                df_source.columns = [str(c).strip() for c in df_source.columns]

                # --- الأساسيات: التوفيق بين المسميات ---
                # نحدد الأسماء المحتملة للأعمدة
                id_keys = ["رقم البصمه", "Badgenumber"]
                date_keys = ["التاريخ", "Date"]
                
                # إيجاد الأعمدة الموجودة فعلياً في ملفك
                def get_col(df, keys):
                    for k in keys:
                        if k in df.columns: return k
                    return None

                col_id_1 = get_col(df_main, id_keys)
                col_date_1 = get_col(df_main, date_keys)
                col_id_2 = get_col(df_source, id_keys)
                col_date_2 = get_col(df_source, date_keys)

                # تنظيف البيانات
                clean_id = lambda v: str(int(float(v))).strip() if pd.notna(v) else ""
                df_main['M_ID'] = df_main[col_id_1].apply(clean_id)
                df_main['M_DAY'] = df_main[col_date_1].astype(str).str.extract(r'(\d+)')[0]
                
                df_source['M_ID'] = df_source[col_id_2].apply(clean_id)
                df_source['M_DAY'] = df_source[col_date_2].astype(str).str.extract(r'(\d+)')[0]

                # دمج كل الأعمدة (بما فيها Name, Department, SHIFT)
                result_df = pd.merge(df_main, df_source.drop(columns=[col_id_2, col_date_2]), on=['M_ID', 'M_DAY'], how='left')
                result_df = result_df.drop(columns=['M_ID', 'M_DAY'])

                # حذف الأعمدة غير المرغوبة (مع SHIFT)
                drop_cols = ["اساسى", "الشيفت", "SHIFT"]
                result_df = result_df.drop(columns=[c for c in drop_cols if c in result_df.columns], errors='ignore')

                output = BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    result_df.to_excel(writer, index=False)
                
                st.success("✅ تم الدمج بنجاح!")
                st.download_button("📥 تحميل الملف", output.getvalue(), "Combined_Result.xlsx")
            except Exception as e:
                st.error(f"خطأ أثناء المعالجة: {e}")
        else:
            st.warning("الرجاء رفع الملفين")

if __name__ == "__main__":
    main()
