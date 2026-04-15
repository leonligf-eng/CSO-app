import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
import datetime
import re

# --- 頁面配置 ---
st.set_page_config(page_title="AOI & FT Yield Report", layout="wide")

st.markdown("""
    <style>
    .kpi-card { background-color: #f8f9fa; border-radius: 10px; padding: 20px; border-left: 5px solid #1E3A8A; box-shadow: 2px 2px 5px rgba(0,0,0,0.05); }
    .kpi-value { font-size: 30px; font-weight: bold; color: #1E3A8A; }
    .kpi-label { font-size: 14px; color: #666; font-weight: bold; }
    .demo-banner { background-color: #e2e3e5; color: #383d41; padding: 10px; border-radius: 5px; text-align: center; font-weight: bold; margin-bottom: 15px; border: 1px solid #d6d8db; }
    </style>
""", unsafe_allow_html=True)

st.title("🚀 Multi-Station Yield & Production Dashboard")

# ==========================================
# 假資料產生器 (升級為跨站點真實流片邏輯)
# ==========================================
@st.cache_data
def generate_mock_data():
    np.random.seed(42)
    now = datetime.datetime.now()
    
    lots = [f"LOT-{10000 + i}" for i in range(60)] # 60 批貨
    progs = np.random.choice(["ZC13_P1.0", "ZC13_EVT1.0(A0)", "ZC13_DVT", "ZC13_MP"], 60)
    stations = ["AOI", "FT1", "FT2"] # 站點流程
    
    report_rows = []
    hold_rows = []
    
    for i, lot in enumerate(lots):
        prog = progs[i]
        # 初始投片量
        current_qty = np.random.randint(10000, 25000)
        lot_start_time = now - datetime.timedelta(days=np.random.randint(0, 5), hours=np.random.randint(0, 24))
        
        for stn in stations:
            tester = f"{stn}_ATE0{np.random.randint(1, 4)}"
            in_time = lot_start_time
            out_time = in_time + datetime.timedelta(hours=np.random.uniform(2, 8))
            
            # 各站點良率模擬 (AOI通常較低，FT2通常較高)
            if stn == "AOI": yd = np.random.uniform(0.92, 0.98)
            elif stn == "FT1": yd = np.random.uniform(0.95, 0.99)
            else: yd = np.random.uniform(0.98, 0.999)
            
            # 製造一點點 Outlier
            if np.random.rand() > 0.95: yd -= 0.1 
                
            pass_qty = int(current_qty * yd)
            
            report_rows.append({
                "Lot": lot, "Station": stn, "Tester": tester, "ProgramName": prog,
                "CheckInTime": in_time, "CheckOutTime": out_time,
                "TestQty": current_qty, "PassQty": pass_qty
            })
            
            # 隨機產生 Hold
            if np.random.rand() > 0.92:
                hold_rows.append({
                    "Lot": lot, "Station": stn, "HoldReason": np.random.choice(["Contact Low", "Hardware Alarm", "Setup Error"]),
                    "CheckInTime": out_time - datetime.timedelta(hours=np.random.uniform(1, 24))
                })
            
            # 這站的 PassQty 變成下一站的 TestQty (真實流片邏輯)
            current_qty = pass_qty
            lot_start_time = out_time + datetime.timedelta(hours=np.random.uniform(0.5, 4))
            
    return pd.DataFrame(report_rows), pd.DataFrame(hold_rows)

# ==========================================
# 1. 側邊欄設定
# ==========================================
st.sidebar.header("📁 數據上傳與設定")
uploaded_file = st.sidebar.file_uploader("上傳報表 (.xlsx)", type=["xlsx"])

st.sidebar.divider()
st.sidebar.subheader("⏰ 分析參數設定")
cutoff_time = st.sidebar.time_input("每日結算基準時間 (Cutoff)", datetime.time(0, 0))
# 讓使用者決定站點順序 (這對累加良率計算很重要)
station_order_input = st.sidebar.text_input("站點流程順序 (逗號分隔)", "AOI, FT1, FT2", help="請依據實際生產流程排列，用逗號隔開")
station_order = [s.strip() for s in station_order_input.split(',')]

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
    for key in ["P1.0", "P1.1", "EVT1.0(A0)", "EVT1.0(B0)", "EVT1.1", "EVT", "DVT", "PVT", "MP"]:
        if key in prog: return key
    return "Other"

# ==========================================
# 3. 資料載入與預處理
# ==========================================
if uploaded_file:
    try:
        xls = pd.ExcelFile(uploaded_file)
        df_report = pd.read_excel(xls, "Report")
        df_hold = pd.read_excel(xls, "Hold lot") if "Hold lot" in xls.sheet_names else pd.DataFrame()
    except Exception as e:
        st.error(f"檔案讀取失敗: {e}")
        st.stop()
else:
    df_report, df_hold = generate_mock_data()
    st.markdown("<div class='demo-banner'>ℹ️ 預覽模式：目前使用模擬的跨站點流片數據 (AOI ➜ FT1 ➜ FT2)</div>", unsafe_allow_html=True)

# 防呆檢查：如果真實資料沒有 Station，自動補上預設值
if 'Station' not in df_report.columns:
    df_report['Station'] = "AOI"

# 前處理
df_report['CheckInTime'] = pd.to_datetime(df_report['CheckInTime'])
df_report['CheckOutTime'] = pd.to_datetime(df_report['CheckOutTime'])
df_report['BuildPhase'] = df_report['ProgramName'].apply(get_build_phase)
df_report['Yield'] = (df_report['PassQty'] / df_report['TestQty'] * 100).fillna(0)

df_apportioned = calculate_apportioned_upd(df_report, cutoff_time)

# ==========================================
# 累加良率 (Cum. Yield) 計算邏輯
# ==========================================
# 1. 算出各站的總良率
stn_summary = df_report.groupby('Station').agg(TotalTest=('TestQty','sum'), TotalPass=('PassQty','sum')).reset_index()
stn_summary['Station_Yield'] = stn_summary['TotalPass'] / stn_summary['TotalTest']

# 2. 依照使用者定義的順序排序
stn_summary['Station'] = pd.Categorical(stn_summary['Station'], categories=station_order, ordered=True)
stn_summary = stn_summary.sort_values('Station').dropna()

# 3. 計算累加良率 (Cumprod)
stn_summary['Cum_Yield'] = stn_summary['Station_Yield'].cumprod() * 100
stn_summary['Station_Yield'] = stn_summary['Station_Yield'] * 100

final_cum_yield = stn_summary['Cum_Yield'].iloc[-1] if not stn_summary.empty else 0

# --- Dashboard Header ---
kpi_cols = st.columns(4)
with kpi_cols[0]: st.markdown(f"<div class='kpi-card'><div class='kpi-label'>總投片量 (Initial Input)</div><div class='kpi-value'>{stn_summary['TotalTest'].iloc[0]:,}</div></div>", unsafe_allow_html=True)
with kpi_cols[1]: st.markdown(f"<div class='kpi-card'><div class='kpi-label'>總產出量 (Final Output)</div><div class='kpi-value'>{stn_summary['TotalPass'].iloc[-1]:,}</div></div>", unsafe_allow_html=True)
with kpi_cols[2]: st.markdown(f"<div class='kpi-card'><div class='kpi-label'>跨站直通率 (Cum. Yield)</div><div class='kpi-value' style='color:#dc3545;'>{final_cum_yield:.2f}%</div></div>", unsafe_allow_html=True)
with kpi_cols[3]: st.markdown(f"<div class='kpi-card'><div class='kpi-label'>涵蓋站點數</div><div class='kpi-value'>{len(stn_summary)} Stations</div></div>", unsafe_allow_html=True)

st.divider()

# --- 多分頁分析 ---
tabs = st.tabs(["🎯 累加良率與站點漏斗", "📊 產出與 UPD", "🏗️ Build Phase 分析", "🔧 機台與異常追蹤"])

# Tab 1: 累加良率 (新增的核心需求)
with tabs[0]:
    st.subheader("站點直通率漏斗 (Rolled Throughput Yield)")
    
    fig_funnel = go.Figure()
    # 畫出單站良率長條圖
    fig_funnel.add_trace(go.Bar(x=stn_summary['Station'], y=stn_summary['Station_Yield'], name='單站良率 (Single Yield %)', marker_color='#6b8cd4', text=stn_summary['Station_Yield'].apply(lambda x: f"{x:.1f}%"), textposition='auto'))
    # 畫出累加良率折線圖
    fig_funnel.add_trace(go.Scatter(x=stn_summary['Station'], y=stn_summary['Cum_Yield'], mode='lines+markers+text', name='累加良率 (Cum. Yield %)', line=dict(color='#dc3545', width=4), marker=dict(size=10), text=stn_summary['Cum_Yield'].apply(lambda x: f"{x:.1f}%"), textposition='top right'))
    
    fig_funnel.update_layout(title="單站良率 vs. 累加良率衰退趨勢", yaxis=dict(title="Yield (%)", range=[min(70, stn_summary['Cum_Yield'].min()-5), 105]), legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1))
    st.plotly_chart(fig_funnel, use_container_width=True)

# Tab 2: 每日 UPD 趨勢 (可選站點)
with tabs[1]:
    sel_station = st.selectbox("選擇要檢視 UPD 的站點", ["All Stations"] + list(df_apportioned['Station'].unique()))
    filtered_upd = df_apportioned if sel_station == "All Stations" else df_apportioned[df_apportioned['Station'] == sel_station]
    
    st.subheader(f"每日 UPD 分配趨勢 ({sel_station})")
    upd_data = filtered_upd.groupby(['ProductionDate', 'Tester'])['WeightedQty'].sum().reset_index()
    fig_upd = px.bar(upd_data, x='ProductionDate', y='WeightedQty', color='Tester', title=f"各機台每日產出 (Cutoff: {cutoff_time.strftime('%H:%M')})")
    st.plotly_chart(fig_upd, use_container_width=True)

# Tab 3: Build Phase 產出分析
with tabs[2]:
    st.subheader("各 Build Phase 階段產出佔比")
    phase_data = df_report.groupby('BuildPhase').agg({'TestQty': 'sum', 'PassQty': 'sum'}).reset_index()
    phase_data['Yield'] = (phase_data['PassQty'] / phase_data['TestQty'] * 100)
    
    col_p1, col_p2 = st.columns(2)
    with col_p1:
        fig_phase_pie = px.pie(phase_data, values='TestQty', names='BuildPhase', title="各 Build 測試總量佔比", hole=0.4)
        st.plotly_chart(fig_phase_pie, use_container_width=True)
    with col_p2:
        fig_phase_yield = px.bar(phase_data, x='BuildPhase', y='Yield', title="各 Build 平均良率", color='Yield', color_continuous_scale='RdYlGn', range_color=[80, 100])
        st.plotly_chart(fig_phase_yield, use_container_width=True)

# Tab 4: 機台健康度與 Hold Lot
with tabs[3]:
    col_t1, col_t2 = st.columns([1.5, 1])
    with col_t1:
        st.subheader("ATE 機台良率分佈")
        fig_health = px.box(df_report, x='Tester', y='Yield', color='Station', title="各機台良率穩定度 (依站點區分)")
        st.plotly_chart(fig_health, use_container_width=True)
    
    with col_t2:
        st.subheader("Hold Lot 狀態")
        if not df_hold.empty:
            hold_pareto = df_hold['HoldReason'].value_counts().reset_index()
            hold_pareto.columns = ['Reason', 'Count']
            fig_hold = px.bar(hold_pareto, x='Reason', y='Count', title="卡控原因排名", text='Count')
            fig_hold.update_layout(height=300)
            st.plotly_chart(fig_hold, use_container_width=True)
        else:
            st.success("🎉 當前無任何 Hold Lot！")