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
def split_cross_day_lots(df, cutoff_hour):
    """將跨日的 Lot 依據時間比例拆分成多筆紀錄"""
    split_records = []
    
    for _, row in df.iterrows():
        start = row['CheckInTime']
        end = row['CheckOutTime']
        if pd.isnull(start) or pd.isnull(end) or start >= end:
            continue
            
        total_seconds = (end - start).total_seconds()
        qty = row['TestQty']
        
        # 進行時間平移，讓邏輯計算簡化為"午夜換日"
        # 例如 cutoff 是 8am，就把所有時間往前推 8 小時，這樣 08:00 就會變成 00:00 (換日點)
        shifted_start = start - timedelta(hours=cutoff_hour)
        shifted_end = end - timedelta(hours=cutoff_hour)
        
        current_shifted = shifted_start
        while current_shifted < shifted_end:
            # 找出今天的結束時間 (當天 23:59:59 的下一秒，也就是隔天 00:00:00)
            next_day_shifted = (current_shifted + timedelta(days=1)).replace(hour=0, minute=0, second=0, microsecond=0)
            
            # 這一段時間的結束點
            segment_end_shifted = min(shifted_end, next_day_shifted)
            
            # 計算這一段時間佔總時間的比例
            segment_seconds = (segment_end_shifted - current_shifted).total_seconds()
            ratio = segment_seconds / total_seconds if total_seconds > 0 else 0
            segment_qty = qty * ratio
            
            # 還原為真實日期標籤 (Production Date)
            prod_date = current_shifted.date()
            
            new_row = row.copy()
            new_row['ProdDate'] = prod_date
            new_row['AllocatedQty'] = segment_qty
            split_records.append(new_row)
            
            current_shifted = next_day_shifted
            
    return pd.DataFrame(split_records)

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