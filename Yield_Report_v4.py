import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
import datetime
import re

# --- 頁面配置 ---
st.set_page_config(page_title="AOI & FT Product Analysis", layout="wide")

st.markdown("""
    <style>
    .kpi-card { background-color: #f8f9fa; border-radius: 10px; padding: 20px; border-left: 5px solid #1E3A8A; box-shadow: 2px 2px 5px rgba(0,0,0,0.05); }
    .kpi-value { font-size: 30px; font-weight: bold; color: #1E3A8A; }
    .kpi-label { font-size: 14px; color: #666; font-weight: bold; }
    .demo-banner { background-color: #e2e3e5; color: #383d41; padding: 10px; border-radius: 5px; text-align: center; font-weight: bold; margin-bottom: 15px; border: 1px solid #d6d8db; }
    </style>
""", unsafe_allow_html=True)

st.title("🚀 Product-Centric Yield & Production Dashboard")

# ==========================================
# 假資料產生器 (加入 ProductNo)
# ==========================================
@st.cache_data
def generate_mock_data():
    np.random.seed(42)
    now = datetime.datetime.now()
    
    # 定義 3 個主要產品型號
    product_list = ["PN-9980-A1", "PN-8820-B2", "PN-7710-C0"]
    lots = [f"LOT-{10000 + i}" for i in range(80)]
    
    stations = ["AOI", "FT1", "FT2"]
    report_rows = []
    
    for i, lot in enumerate(lots):
        p_no = np.random.choice(product_list, p=[0.5, 0.3, 0.2]) # 設定不同產品權重
        prog = f"{p_no}_V1.2"
        current_qty = np.random.randint(10000, 25000)
        lot_start_time = now - datetime.timedelta(days=np.random.randint(0, 5), hours=np.random.randint(0, 24))
        
        for stn in stations:
            tester = f"{stn}_ATE0{np.random.randint(1, 4)}"
            in_time = lot_start_time
            out_time = in_time + datetime.timedelta(hours=np.random.uniform(2, 6))
            
            # 產品良率差異化模擬
            if p_no == "PN-7710-C0": yd = np.random.uniform(0.88, 0.94) # 此產品良率較低
            else: yd = np.random.uniform(0.96, 0.99)
                
            pass_qty = int(current_qty * yd)
            
            report_rows.append({
                "Lot": lot, "ProductNo": p_no, "Station": stn, "Tester": tester, "ProgramName": prog,
                "CheckInTime": in_time, "CheckOutTime": out_time,
                "TestQty": current_qty, "PassQty": pass_qty
            })
            current_qty = pass_qty
            lot_start_time = out_time + datetime.timedelta(hours=np.random.uniform(0.5, 2))
            
    return pd.DataFrame(report_rows)

# ==========================================
# 1. 核心邏輯 (UPD 切割)
# ==========================================
@st.cache_data
def calculate_apportioned_upd(df, cutoff_time):
    df = df.dropna(subset=['CheckInTime', 'CheckOutTime', 'TestQty']).copy()
    cutoff_delta = datetime.timedelta(hours=cutoff_time.hour, minutes=cutoff_time.minute)
    apportioned_results = []
    for _, row in df.iterrows():
        in_t, out_t, qty = row['CheckInTime'], row['CheckOutTime'], row['TestQty']
        eff_in, eff_out = in_t - cutoff_delta, out_t - cutoff_delta
        s_date, e_date = eff_in.date(), eff_out.date()
        dur = (out_t - in_t).total_seconds()
        if s_date == e_date or dur <= 0:
            row['ProductionDate'], row['WeightedQty'] = s_date, qty
            apportioned_results.append(row)
        else:
            boundary = datetime.datetime.combine(e_date, cutoff_time)
            if boundary < in_t: boundary += datetime.timedelta(days=1)
            r1, r2 = row.copy(), row.copy()
            r1['ProductionDate'], r1['WeightedQty'] = s_date, qty * ((boundary - in_t).total_seconds() / dur)
            r2['ProductionDate'], r2['WeightedQty'] = e_date, qty * ((out_t - boundary).total_seconds() / dur)
            apportioned_results.extend([r1, r2])
    return pd.DataFrame(apportioned_results)

# ==========================================
# 2. 資料載入
# ==========================================
uploaded_file = st.sidebar.file_uploader("上傳報表 (.xlsx)", type=["xlsx"])
cutoff_time = st.sidebar.time_input("每日結算基準時間 (Cutoff)", datetime.time(0, 0))

if uploaded_file:
    xls = pd.ExcelFile(uploaded_file)
    df_report = pd.read_excel(xls, "Report")
else:
    df_report = generate_mock_data()
    st.markdown("<div class='demo-banner'>ℹ️ 預覽模式：已自動加入 ProductNo (產品型號) 維度分析</div>", unsafe_allow_html=True)

# 前處理
df_report['CheckInTime'] = pd.to_datetime(df_report['CheckInTime'])
df_report['CheckOutTime'] = pd.to_datetime(df_report['CheckOutTime'])
df_report['Yield'] = (df_report['PassQty'] / df_report['TestQty'] * 100).fillna(0)
df_apportioned = calculate_apportioned_upd(df_report, cutoff_time)

# --- Dashboard ---
tabs = st.tabs(["📦 產品組合分析", "📊 每日 UPD (by Product)", "🏗️ Build & Yield"])

# Tab 1: Product Mix Analysis (新增加的)
with tabs[0]:
    st.subheader("Product Mix & Volume Distribution")
    col1, col2 = st.columns(2)
    
    product_summary = df_report.groupby('ProductNo').agg({'TestQty': 'sum', 'Yield': 'mean'}).reset_index()
    
    with col1:
        fig_pie = px.pie(product_summary, values='TestQty', names='ProductNo', title="各產品投入總量比例 (Input Volume %)", hole=0.4)
        st.plotly_chart(fig_pie, use_container_width=True)
    
    with col2:
        fig_product_yield = px.bar(product_summary, x='ProductNo', y='Yield', title="各產品平均良率對比", color='Yield', color_continuous_scale='RdYlGn')
        st.plotly_chart(fig_product_yield, use_container_width=True)

# Tab 2: UPD by Product
with tabs[1]:
    st.subheader("每日 UPD 分配趨勢 (依產品型號區分)")
    # 使用者可以選擇要看哪個維度
    group_by_opt = st.radio("UPD 分群方式:", ["ProductNo", "Tester"], horizontal=True)
    
    upd_data = df_apportioned.groupby(['ProductionDate', group_by_opt])['WeightedQty'].sum().reset_index()
    fig_upd = px.bar(upd_data, x='ProductionDate', y='WeightedQty', color=group_by_opt, 
                     title=f"每日產出趨勢 (Group by {group_by_opt})", text_auto='.2s')
    st.plotly_chart(fig_upd, use_container_width=True)

# Tab 3: Build & Yield
with tabs[2]:
    st.subheader("機台良率穩定度 (依產品篩選)")
    target_p = st.multiselect("篩選產品型號", options=df_report['ProductNo'].unique(), default=df_report['ProductNo'].unique())
    filtered_df = df_report[df_report['ProductNo'].isin(target_p)]
    
    fig_health = px.box(filtered_df, x='Tester', y='Yield', color='ProductNo', title="各機台執行不同產品的良率表現")
    st.plotly_chart(fig_health, use_container_width=True)