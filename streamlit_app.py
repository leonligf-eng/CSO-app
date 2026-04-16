import streamlit as st

# 1. 宣告你的兩個頁面檔案在哪裡，並設定左側選單的名稱與 Icon
page_planner = st.Page("ATE_Smart_Capacity.py", title="ATE Smart Capacity and Allocation", icon="📟", default=True)
page_OEE_Analyzer = st.Page("OEE_Analyzer.py", title="ATE OEE Analyzer", icon="📈")
# 2. 把頁面綁定到導覽列
pg = st.navigation([page_planner, page_OEE_Analyzer])

# 3. 設定全域的網頁樣式 (必須放在最前面)
st.set_page_config(page_title="Zion OSAT Tools", layout="wide")

# 4. 執行導覽
pg.run()