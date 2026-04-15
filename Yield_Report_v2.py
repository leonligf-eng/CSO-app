import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
import datetime
import re

# --- 頁面配置 ---
st.set_page_config(page_title="AOI Advanced Yield Report", layout="wide")

st.markdown("""
    <style>
    .kpi-card { background-color: #f8f9fa; border-radius: 10px; padding: 20px; border-left: 5px solid #1E3A8A; box-shadow: 2px 2px 5px rgba(0,0,0,0.05); }
    .kpi-value { font-size: 30px; font-weight: bold; color: #1E3A8A; }
    .kpi-label { font-size: 14px; color: #666; font-weight: bold; }
    .demo-banner { background-color: #fff3cd; color: #856404; padding: 10px; border-radius: 5px; text-align: center; font-weight: bold; margin-bottom: 15px; border: 1px solid #ffeeba; }
    </style>
""", unsafe_allow_html=True)

st.title("🚀 AOI Yield & Production Dashboard")

# ==========================================
# 假資料產生器 (Mock Data Generator)
# ==========================================
@st.cache_data
def generate_mock_data():
    np.random.seed(42) # 固定亂數種子，讓每次討論的畫面一致
    now = datetime.datetime.now()
    
    # 產生 150 筆假的 Lot 數據
    lots = [f"LOT-{10000 + i}" for i in range(150)]
    testers = np.random.choice(["ATE01", "ATE02", "ATE03", "ATE04", "ATE05"], 150)
    progs = np.random.choice(["ZC13_P1.0", "ZC13_EVT1.0(A0)", "ZC13_EVT1.1_patch", "ZC13_DVT", "ZC13_MP"], 150)
    
    # 隨機產生過去 5 天內的進出站時間 (模擬跨日)
    in_times = [now - datetime.timedelta(days=np.random.randint(0, 5), hours=np.random.randint(0, 24)) for _ in range(150)]
    out_times = [t + datetime.timedelta(hours=np.random.uniform(2, 16)) for t in in_times] 
    
    test_qty = np.random.randint(5000, 25000, 150)
    # 製造一些低良率的 outlier
    yields = np.where(np.random.rand(150) > 0.1, np.random.uniform(0.92, 0.99, 150), np.random.uniform(0.75, 0.88, 150))
    pass_qty = (test_qty * yields).astype(int)
    
    df_report = pd.DataFrame({
        "Lot": lots, "Tester": testers, "ProgramName": progs,
        "CheckInTime": in_times, "CheckOutTime": out_times,
        "TestQty": test_qty, "PassQty": pass_qty
    })
    
    # 產生假的 Hold Lot 數據 (約 15 筆)
    df_hold = pd.DataFrame({
        "Lot": np.random.choice(lots, 15, replace=False),
        "HoldReason": np.random.choice(["Contact Yield Low", "Hardware Alarm", "Setup Error", "Bin 1 Limit", "Site Fail"], 15),
        "CheckInTime": [now - datetime.timedelta(hours=np.random.uniform(2, 48)) for _ in range(15)]
    })
    
    return df_report, df_hold

# ==========================================
# 1. 側邊欄設定 (控制參數)
# ==========================================
st.sidebar.header("📁 數據上傳與設定")
uploaded_file = st.sidebar.file_uploader("上傳 AOI 報表 (.xlsx)", type=["xlsx"])

st.sidebar.divider()
st.sidebar.subheader("⏰ 每日結算參數")
cutoff_time = st.sidebar.time_input("每日結算基準時間 (Cutoff)", datetime.time(0, 0), help="決定每日產出計算的切分點。")

st.sidebar.subheader("🎯 目標設定")
yield_target = st.sidebar.slider("Yield Target (%)", 85.0, 100.0, 95.0)

# ==========================================
# 2. 核心邏輯引擎
# ==========================================
@st.cache_data
def calculate_apportioned_upd(df, cutoff_time):
    df = df.dropna(subset=['CheckInTime', 'CheckOutTime', 'TestQty']).copy()
    cutoff_delta = datetime.timedelta(hours=cutoff_time.hour, minutes=cutoff_time.minute)
    apportioned_results = []
    
    for _, row in df.iterrows():
        in_t, out_t, qty = row['CheckInTime'], row['CheckOutTime'], row['TestQty']
        effective_in, effective_out = in_t - cutoff_delta, out_t - cutoff_delta
        start_date, end_date = effective_in.date(), effective_out.date()
        total_duration = (out_t - in_t).total_seconds()
        
        if start_date == end_date or total_duration <= 0:
            row['ProductionDate'] = start_date
            row['WeightedQty'] = qty
            apportioned_results.append(row)
        else:
            boundary = datetime.datetime.combine(end_date, cutoff_time)
            if boundary < in_t: boundary += datetime.timedelta(days=1)
            
            r1, r2 = row.copy(), row.copy()
            r1['ProductionDate'], r1['WeightedQty'] = start_date, qty * ((boundary - in_t).total_seconds() / total_duration)
            r2['ProductionDate'], r2['WeightedQty'] = end_date, qty * ((out_t - boundary).total_seconds() / total_duration)
            apportioned_results.extend([r1, r2])
            
    return pd.DataFrame(apportioned_results)

def get_build_phase(prog):
    prog = str(prog).upper()
    phase_map = [("P1.0", "P1.0"), ("P1.1", "P1.1"), ("EVT1.0(A0)", "EVT1.0(A0)"), ("EVT1.0(B0)", "EVT1.0(B0)"),
                 ("EVT1.1", "EVT1.1"), ("EVT", "EVT"), ("DVT", "DVT"), ("PVT", "PVT"), ("MP", "MP")]
    for key, label in phase_map:
        if key in prog: return label
    return "Other"

# ==========================================
# 3. 資料載入 (判斷 真實 vs 假資料)
# ==========================================
is_demo = False
if uploaded_file:
    try:
        xls = pd.ExcelFile(uploaded_file)
        df_report = pd.read_excel(xls, "Report")
        df_hold = pd.read_excel(xls, "Hold lot")
    except Exception as e:
        st.error(f"檔案讀取失敗: {e}")
        st.stop()
else:
    is_demo = True
    df_report, df_hold = generate_mock_data()
    st.markdown("<div class='demo-banner'>⚠️ 目前尚未上傳檔案，系統使用「展示用假資料 (Mock Data)」進行框架預覽體驗。</div>", unsafe_allow_html=True)

# ==========================================
# 4. 網頁呈現與分析
# ==========================================
# 前處理
df_report['CheckInTime'] = pd.to_datetime(df_report['CheckInTime'])
df_report['CheckOutTime'] = pd.to_datetime(df_report['CheckOutTime'])
df_report['BuildPhase'] = df_report['ProgramName'].apply(get_build_phase)
if 'Yield' not in df_report.columns:
    df_report['Yield'] = (df_report['PassQty'] / df_report['TestQty'] * 100).fillna(0)

df_apportioned = calculate_apportioned_upd(df_report, cutoff_time)

# KPI 計算
total_input = int(df_report['TestQty'].sum())
avg_yield = (df_report['PassQty'].sum() / total_input * 100) if total_input > 0 else 0

# --- Dashboard Header ---
kpi_cols = st.columns(4)
with kpi_cols[0]: st.markdown(f"<div class='kpi-card'><div class='kpi-label'>總測試產出 (Total Qty)</div><div class='kpi-value'>{total_input:,}</div></div>", unsafe_allow_html=True)
with kpi_cols[1]: st.markdown(f"<div class='kpi-card'><div class='kpi-label'>平均良率 (Avg Yield)</div><div class='kpi-value'>{avg_yield:.2f}%</div></div>", unsafe_allow_html=True)
with kpi_cols[2]: st.markdown(f"<div class='kpi-card'><div class='kpi-label'>當前 ATE 機台數</div><div class='kpi-value'>{df_report['Tester'].nunique()}</div></div>", unsafe_allow_html=True)
with kpi_cols[3]: st.markdown(f"<div class='kpi-card'><div class='kpi-label'>Build Phase 數量</div><div class='kpi-value'>{df_report['BuildPhase'].nunique()}</div></div>", unsafe_allow_html=True)

st.divider()

# --- 多分頁分析 ---
tabs = st.tabs(["📊 產出與 UPD", "🏗️ Build Phase 進度", "🔧 機台健康度", "🛑 Hold Lot 追蹤"])

# Tab 1: 每日 UPD 趨勢
with tabs[0]:
    st.subheader("每日 UPD 分配趨勢 (依機台)")
    upd_data = df_apportioned.groupby(['ProductionDate', 'Tester'])['WeightedQty'].sum().reset_index()
    fig_upd = px.bar(upd_data, x='ProductionDate', y='WeightedQty', color='Tester',
                     title=f"各機台每日產出 (Cutoff: {cutoff_time.strftime('%H:%M')})", text_auto='.2s')
    fig_upd.update_layout(barmode='stack', xaxis_title="生產日期", yaxis_title="分配後產出 Units")
    st.plotly_chart(fig_upd, use_container_width=True)

# Tab 2: Build Phase 產出分析
with tabs[1]:
    st.subheader("各 Build Phase 階段產出佔比")
    phase_data = df_report.groupby('BuildPhase').agg({'TestQty': 'sum', 'PassQty': 'sum'}).reset_index()
    phase_data['Yield'] = (phase_data['PassQty'] / phase_data['TestQty'] * 100)
    
    col_p1, col_p2 = st.columns(2)
    with col_p1:
        fig_phase_pie = px.pie(phase_data, values='TestQty', names='BuildPhase', title="各 Build 測試總量佔比", hole=0.4)
        st.plotly_chart(fig_phase_pie, use_container_width=True)
    with col_p2:
        fig_phase_yield = px.bar(phase_data, x='BuildPhase', y='Yield', title="各 Build 平均良率", color='Yield', 
                                 color_continuous_scale='RdYlGn', range_color=[80, 100])
        fig_phase_yield.add_hline(y=yield_target, line_dash="dash", line_color="red", annotation_text="Target")
        st.plotly_chart(fig_phase_yield, use_container_width=True)

# Tab 3: 機台效能與健康度
with tabs[2]:
    st.subheader("ATE Performance Matrix (機台穩定度)")
    fig_health = px.box(df_report, x='Tester', y='Yield', color='Tester', title="各機台良率分佈 (Box Plot)", points="all", hover_data=['Lot', 'ProgramName'])
    fig_health.add_hline(y=avg_yield, line_dash="dot", line_color="green", annotation_text="全廠平均")
    st.plotly_chart(fig_health, use_container_width=True)

# Tab 4: Hold Lot 追蹤
with tabs[3]:
    st.subheader("Hold Lot 分析與追蹤")
    if not df_hold.empty:
        col_h1, col_h2 = st.columns([1, 1])
        with col_h1:
            hold_pareto = df_hold['HoldReason'].value_counts().reset_index()
            hold_pareto.columns = ['Reason', 'Count']
            fig_hold = px.bar(hold_pareto, x='Reason', y='Count', title="卡控原因排名 (Pareto)", text='Count')
            st.plotly_chart(fig_hold, use_container_width=True)
        with col_h2:
            st.write("📋 **當前卡控清單：**")
            # 模擬計算卡控時間
            if 'CheckInTime' in df_hold.columns:
                df_hold['CheckInTime'] = pd.to_datetime(df_hold['CheckInTime'])
                df_hold['Hold_Hours'] = ((datetime.datetime.now() - df_hold['CheckInTime']).dt.total_seconds() / 3600).round(1)
                st.dataframe(df_hold[['Lot', 'HoldReason', 'Hold_Hours']].sort_values('Hold_Hours', ascending=False), use_container_width=True, hide_index=True)
    else:
        st.success("🎉 當前無任何 Hold Lot！")