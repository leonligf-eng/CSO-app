import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
from datetime import datetime, timedelta, time

st.set_page_config(page_title="AOI Yield & Output Report", layout="wide")

# --- Custom CSS ---
st.markdown("""
    <style>
    .kpi-card {
        background-color: #ffffff;
        border: 1px solid #e0e0e0;
        padding: 20px;
        border-radius: 10px;
        box-shadow: 2px 2px 5px rgba(0,0,0,0.05);
        text-align: center;
    }
    .kpi-title { color: #6c757d; font-size: 14px; font-weight: bold; text-transform: uppercase; }
    .kpi-value { color: #1E3A8A; font-size: 32px; font-weight: bold; margin-top: 5px; }
    </style>
""", unsafe_allow_html=True)

st.title("📈 AOI Yield & Output Report")

# ==============================================================================
# --- 1. Sidebar Setup & Data Loading ---
# ==============================================================================
st.sidebar.header("📁 Data Input")
uploaded_file = st.sidebar.file_uploader("Upload Yield Report (Excel)", type=["xlsx", "xls"])

st.sidebar.header("⚙️ Output Settings")
cutoff_hour = st.sidebar.slider("Daily Cutoff Time (Hour)", min_value=0, max_value=23, value=0, help="0 = 24:00 (Midnight). If set to 8, a day is 08:00 AM to 08:00 AM next day.")

@st.cache_data
def generate_mock_data():
    """產生測試用的假資料，確保沒有上傳檔案時也能預覽畫面"""
    np.random.seed(42)
    now = datetime.now()
    data = []
    testers = [f"93K_{i:02d}" for i in range(1, 10)]
    programs = ["P1.0", "P1.1", "EVT1.0(A0)", "EVT1.0(B0)", "EVT1.1", "DVT", "PVT"]
    
    for i in range(100):
        tester = np.random.choice(testers)
        prog = np.random.choice(programs)
        qty = np.random.randint(500, 5000)
        yield_rate = np.random.uniform(0.85, 0.99)
        pass_qty = int(qty * yield_rate)
        fail_qty = qty - pass_qty
        
        # 模擬跨日時間
        start_time = now - timedelta(days=np.random.randint(0, 5), hours=np.random.randint(0, 24))
        end_time = start_time + timedelta(hours=np.random.uniform(2, 12))
        
        data.append([f"LOT_{i:04d}", tester, start_time, end_time, qty, pass_qty, fail_qty, yield_rate, prog])
        
    return pd.DataFrame(data, columns=['LotID', 'Tester', 'CheckInTime', 'CheckOutTime', 'TestQty', 'PassQty', 'FailQty', 'Yield', 'ProgramName'])

def load_data(file):
    if file is not None:
        try:
            df = pd.read_excel(file, sheet_name="Report")
            # Ensure datetime format
            df['CheckInTime'] = pd.to_datetime(df['CheckInTime'])
            df['CheckOutTime'] = pd.to_datetime(df['CheckOutTime'])
            return df
        except Exception as e:
            st.error(f"Error loading file: {e}")
            return None
    return generate_mock_data()

raw_df = load_data(uploaded_file)

if raw_df is None or raw_df.empty:
    st.warning("No data available.")
    st.stop()

# ==============================================================================
# --- 2. Data Preprocessing & Proportional UPD Logic ---
# ==============================================================================
# (Req 1) 跨日時間比例切割演算法
def split_cross_day_lots(df, cutoff_time):
    # 🌟 防呆：強制移除 Excel 中可能重複的欄位名稱
    df = df.loc[:, ~df.columns.duplicated()]
    
    # 過濾掉沒有時間或數量的異常資料
    df = df.dropna(subset=['CheckInTime', 'CheckOutTime', 'TestQty']).copy()
    
    cutoff_hour = cutoff_time.hour
    cutoff_minute = cutoff_time.minute
    
    new_rows = []
    
    for _, row in df.iterrows():
        in_time = row['CheckInTime']
        out_time = row['CheckOutTime']
        qty = row['TestQty']
        
        shift_in_time = in_time - pd.Timedelta(hours=cutoff_hour, minutes=cutoff_minute)
        shift_out_time = out_time - pd.Timedelta(hours=cutoff_hour, minutes=cutoff_minute)
        
        start_date = shift_in_time.date()
        end_date = shift_out_time.date()
        
        total_seconds = (out_time - in_time).total_seconds()
        
        # 🌟 改用 to_dict()，避開 Pandas Series 複製時的標籤衝突，且處理速度更快
        row_dict = row.to_dict()
        
        if start_date == end_date or total_seconds <= 0:
            # 沒有跨日，直接歸屬
            row_dict['ProductionDate'] = start_date
            row_dict['ApportionedQty'] = qty
            new_rows.append(row_dict)
        else:
            # 發生跨日！精準計算第一天與第二天的秒數比例
            boundary_time = pd.Timestamp(datetime.datetime.combine(end_date, cutoff_time))
            if boundary_time < in_time: 
                boundary_time += pd.Timedelta(days=1)
                
            sec_day1 = (boundary_time - in_time).total_seconds()
            sec_day2 = (out_time - boundary_time).total_seconds()
            
            ratio_day1 = sec_day1 / total_seconds
            ratio_day2 = sec_day2 / total_seconds
            
            # 建立第一天的紀錄
            row1 = row_dict.copy()
            row1['ProductionDate'] = start_date
            row1['ApportionedQty'] = qty * ratio_day1
            new_rows.append(row1)
            
            # 建立第二天的紀錄
            row2 = row_dict.copy()
            row2['ProductionDate'] = end_date
            row2['ApportionedQty'] = qty * ratio_day2
            new_rows.append(row2)
            
    return pd.DataFrame(new_rows)

# 執行跨日拆分
df_split = split_cross_day_lots(raw_df, cutoff_hour)

# 定義 Build 順序以利圖表排序
build_order = ["P1.0", "P1.1", "EVT1.0(A0)", "EVT1.0(B0)", "EVT1.1", "DVT", "PVT", "MP"]
# 嘗試確保 ProgramName 轉換為 Categorical 以固定順序
if 'ProgramName' in df_split.columns:
    existing_builds = [b for b in build_order if b in df_split['ProgramName'].unique()]
    df_split['ProgramName'] = pd.Categorical(df_split['ProgramName'], categories=existing_builds, ordered=True)

# ==============================================================================
# --- 3. Dashboard UI ---
# ==============================================================================
st.markdown("### 📊 Overall Performance")

# --- KPI Cards ---
total_input = int(df_split['AllocatedQty'].sum())
total_pass = int(raw_df['PassQty'].sum())  # Pass Qty usually tracked by complete lot
overall_yield = (total_pass / raw_df['TestQty'].sum()) * 100 if raw_df['TestQty'].sum() > 0 else 0
active_testers = raw_df['Tester'].nunique()

c1, c2, c3, c4 = st.columns(4)
with c1:
    st.markdown(f"<div class='kpi-card'><div class='kpi-title'>Total UPD (Units)</div><div class='kpi-value'>{total_input:,}</div></div>", unsafe_allow_html=True)
with c2:
    st.markdown(f"<div class='kpi-card'><div class='kpi-title'>Total Pass</div><div class='kpi-value'>{total_pass:,}</div></div>", unsafe_allow_html=True)
with c3:
    st.markdown(f"<div class='kpi-card'><div class='kpi-title'>Overall Yield</div><div class='kpi-value'>{overall_yield:.2f}%</div></div>", unsafe_allow_html=True)
with c4:
    st.markdown(f"<div class='kpi-card'><div class='kpi-title'>Active Testers</div><div class='kpi-value'>{active_testers}</div></div>", unsafe_allow_html=True)

st.divider()

# --- Charts Area ---
col1, col2 = st.columns(2)

with col1:
    st.markdown("#### 📅 Daily Output (UPD) by Tester")
    st.caption(f"Based on Cutoff Time: {cutoff_hour:02d}:00")
    daily_upd = df_split.groupby(['ProdDate', 'Tester'])['AllocatedQty'].sum().reset_index()
    fig1 = px.bar(daily_upd, x='ProdDate', y='AllocatedQty', color='Tester', 
                  title="Daily Output Trend", barmode='stack')
    fig1.update_layout(xaxis_title="Production Date", yaxis_title="Units Tested")
    st.plotly_chart(fig1, use_container_width=True)

with col2:
    st.markdown("#### 🧬 Yield Progression by Build Phase")
    st.caption("Tracking NPI Release Flow")
    # 對原始資料計算 Yield
    build_yield = raw_df.groupby('ProgramName').agg({'TestQty':'sum', 'PassQty':'sum'}).reset_index()
    build_yield['Yield %'] = (build_yield['PassQty'] / build_yield['TestQty']) * 100
    
    # 確保按照 Build Order 排序
    build_yield['ProgramName'] = pd.Categorical(build_yield['ProgramName'], categories=build_order, ordered=True)
    build_yield = build_yield.sort_values('ProgramName')
    
    fig2 = px.line(build_yield, x='ProgramName', y='Yield %', markers=True, 
                   title="Yield Trend Across Builds")
    fig2.update_traces(line_color='#1E3A8A', marker=dict(size=10))
    fig2.update_yaxes(range=[min(80, build_yield['Yield %'].min()-5), 100])
    st.plotly_chart(fig2, use_container_width=True)

col3, col4 = st.columns(2)

with col3:
    st.markdown("#### 🤖 Tester Workload Distribution")
    tester_load = df_split.groupby(['Tester', 'ProgramName'])['AllocatedQty'].sum().reset_index()
    fig3 = px.bar(tester_load, y='Tester', x='AllocatedQty', color='ProgramName', 
                  orientation='h', title="Units Tested by ATE & Build")
    st.plotly_chart(fig3, use_container_width=True)

with col4:
    st.markdown("#### ⏱️ Average Test Time by Build")
    # 計算各 Lot 的總時長 (小時)
    raw_df['Duration_Hr'] = (raw_df['CheckOutTime'] - raw_df['CheckInTime']).dt.total_seconds() / 3600
    avg_time = raw_df.groupby('ProgramName')['Duration_Hr'].mean().reset_index()
    
    fig4 = px.bar(avg_time, x='ProgramName', y='Duration_Hr', 
                  title="Avg Duration per Lot (Hours)", color_discrete_sequence=['#28a745'])
    st.plotly_chart(fig4, use_container_width=True)

# --- Raw Data Expander ---
with st.expander("🔍 View Raw Proportional Splitting Data"):
    st.dataframe(df_split[['LotID', 'Tester', 'ProdDate', 'ProgramName', 'AllocatedQty', 'CheckInTime', 'CheckOutTime']].sort_values(by=['CheckInTime']))