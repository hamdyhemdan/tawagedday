import streamlit as st
import pandas as pd
from io import BytesIO
from datetime import datetime, time
import re
from openpyxl.styles import PatternFill

# --- الإعدادات ---
CONFIG = {
    "MAIN": {"ID": "رقم البصمه", "DATE": "التاريخ"},
    "SOURCE": {"ID": "Badgenumber", "DATE": "Date", "NAME": "Name", "DEPT": "Department", "SHIFT": "SHIFT"}
}

def smart_read_excel(uploaded_file):
    return pd.read_excel(uploaded_file)

def main():
    st.title("🚀 نظام دمج وتلوين البصمة")
    
    main_file = st.file_uploader("1. ملف البصمة الأساسي", type=["xlsx", "xls"])
    source_file = st.file_uploader("2. ملف بيانات المصدر", type=["xlsx", "xls"])

    if st.button("🚀 ابدأ المعالجة"):
        if main_file and source_file:
            try:
                df_main = smart_read_excel(main_file)
                df_source = smart_read_excel(source_file)
                
                # تنظيف المسافات من أسماء الأعمدة
                df_main.columns = df_main.columns.str.strip()
                df_source.columns = df_source.columns.str.strip()

                # التحقق من وجود الأعمدة
                required_cols = [CONFIG["SOURCE"]["ID"], CONFIG["SOURCE"]["DATE"]]
                missing = [c for c in required_cols if c not in df_source.columns]
                
                if missing:
                    st.error(f"❌ الخطأ: الأعمدة التالية غير موجودة في ملف المصدر: {missing}")
                    st.write("الأعمدة المكتشفة في الملف هي:", list(df_source.columns))
                    st.info("تأكد أن الاسم مطابق تماماً (مثلاً: تأكد من عدم وجود مسافة مخفية في نهاية الاسم داخل ملف الإكسيل).")
                    return

                # (تكملة الكود كما كان...)
                st.success("تم العثور على الأعمدة بنجاح! جاري المعالجة...")
                # ... باقي منطق الدمج والتلوين ...

            except Exception as e:
                st.error(f"حدث خطأ: {e}")

if __name__ == "__main__":
    main()
