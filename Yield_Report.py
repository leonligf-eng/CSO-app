import streamlit as st
import pandas as pd
import numpy as np
from datetime import datetime, timedelta, time

st.set_page_config(page_title="ATE Tester Capacity Analysis", layout="wide")

# --- Custom CSS (完美還原 KPI 卡片與綠色高亮文字) ---
st.markdown("""
    <style>
    /* KPI Card Style */
    .kpi-card {
        background-color: #ffffff;
        border: 1px solid #e0e0e0;
        padding: 20px;
        border-radius: 8px;
        box-shadow: 2px 2px 5px rgba(0,0,0,0.05);
        text-align: center;
        margin-bottom: 25px;
    }
    .kpi-title { color: #6c757d; font-size: 12px; font-weight: bold; text-transform: uppercase; letter-spacing: 1px;}
    .kpi-value { color: #1E3A8A; font-size: 32px; font-weight: bold; margin-top: 10px; }
    
    /* Summary Text Style */
    .summary-text { font-size: 16px; line-height: 2.2; }
    .highlight-val { color: #28a745; font-weight: bold; font-size: 18px;}
    </style>
""", unsafe_allow_html=True)

# ==============================================================================
# --- 0. Mock Data Generator ---
# ==============================================================================
@st.cache_data
def generate_mock_data():
    np.random.seed(42)
    now = datetime.now()
    data = []
    testers = [f"HP93K-EXA{i:02d}" for i in range(2, 17)]
    ops = ["CP1", "CP2", "FT1", "FT2"]
    programs_map = {
        "CP1": ["P1.0", "P1.1"],
        "CP2": ["EVT1.0(A0)", "EVT1.0(B0)"],
        "FT1": ["EVT1.1", "DVT"],
        "FT2": ["PVT", "MP"]
    }
    
    for i in range(120):
        tester = np.random.choice(testers)
        op = np.random.choice(ops)
        prog = np.random.choice(programs_map[op])
        qty = np.random.randint(1500, 5000)
        yield_rate = np.random.uniform(0.85, 0.99)
        pass_qty = int(qty * yield_rate)
        
        start_time = now - timedelta(days=np.random.randint(0, 7), hours=np.random.randint(0, 24))
        end_time = start_time + timedelta(hours=np.random.uniform(2, 16)) 
        
        data.append([f"LOT_{i:04d}", op, prog, tester, start_time, end_time, qty, pass_qty])
        
    return pd.DataFrame(data, columns=['LotID', 'OpNo', 'ProgramName', 'Tester', 'CheckInTime', 'CheckOutTime', 'TestQty', 'PassQty'])

def load_data(file):
    if file is not None:
        try:
            # 讀取真實檔案的 Report 分頁
            df = pd.read_excel(file, sheet_name="Report")
            
            # 🌟 防呆：清除所有標題欄位前後可能不小心多打的空白鍵
            df.columns = df.columns.str.strip()
            
            # 確保時間格式正確 (吃你的 CheckInTime, CheckOutTime)
            df['CheckInTime'] = pd.to_datetime(df['CheckInTime'], errors='coerce')
            df['CheckOutTime'] = pd.to_datetime(df['CheckOutTime'], errors='coerce')
            
            # 如果資料有缺漏值，進行基本的填補防呆
            df['TestQty'] = pd.to_numeric(df['TestQty'], errors='coerce').fillna(0)
            df['PassQty'] = pd.to_numeric(df['PassQty'], errors='coerce').fillna(0)
            
            return df
        except Exception as e:
            st.error(f"檔案讀取失敗，請確認是否包含 Report 分頁。錯誤細節: {str(e)}")
            return None
            
    # 若沒有上傳，則執行我們先前寫好的 generate_mock_data()
    return generate_mock_data()

# ==============================================================================
# --- 1. Sidebar & Filters ---
# ==============================================================================
st.sidebar.header("📁 1. Data Input")
uploaded_file = st.sidebar.file_uploader("Upload Yield Report (Excel)", type=["xlsx", "xls"])
raw_df = load_data(uploaded_file)

if raw_df is None or raw_df.empty:
    st.warning("No data available.")
    st.stop()

st.sidebar.divider()
st.sidebar.header("🔍 2. Data Filters")

op_options = sorted(raw_df['OpNo'].dropna().unique().tolist())
selected_ops = st.sidebar.multiselect("選擇站點 (OpNo)", options=op_options, default=op_options)

filtered_by_op = raw_df[raw_df['OpNo'].isin(selected_ops)] if selected_ops else raw_df
prog_options = sorted(filtered_by_op['ProgramName'].dropna().unique().tolist())
selected_progs = st.sidebar.multiselect("選擇程式 (ProgramName)", options=prog_options, default=prog_options)

st.sidebar.divider()
st.sidebar.header("⚙️ 3. Calculation Settings")
cutoff_hour = st.sidebar.slider("Daily Cutoff Time (Hour)", min_value=0, max_value=23, value=0, help="0 = 24:00 (Midnight).")

if not selected_ops or not selected_progs:
    st.info("👈 請在左側側邊欄選擇 **OpNo** 與 **ProgramName**。")
    st.stop()

# ==============================================================================
# --- 2. Data Processing Engine ---
# ==============================================================================
filtered_df = raw_df[(raw_df['OpNo'].isin(selected_ops)) & (raw_df['ProgramName'].isin(selected_progs))].copy()

@st.cache_data
def split_cross_day_lots(df, cutoff_time):
    df = df.loc[:, ~df.columns.duplicated()]
    df = df.dropna(subset=['CheckInTime', 'CheckOutTime', 'TestQty']).copy()
    
    cutoff_hour = cutoff_time.hour
    cutoff_minute = cutoff_time.minute
    new_rows = []
    
    for _, row in df.iterrows():
        in_time, out_time, qty = row['CheckInTime'], row['CheckOutTime'], row['TestQty']
        
        shift_in_time = in_time - pd.Timedelta(hours=cutoff_hour, minutes=cutoff_minute)
        shift_out_time = out_time - pd.Timedelta(hours=cutoff_hour, minutes=cutoff_minute)
        
        start_date, end_date = shift_in_time.date(), shift_out_time.date()
        total_seconds = (out_time - in_time).total_seconds()
        total_duration_hr = total_seconds / 3600
        
        row_dict = row.to_dict()
        
        if start_date == end_date or total_seconds <= 0:
            row_dict['ProductionDate'] = start_date
            row_dict['ApportionedQty'] = qty
            row_dict['ApportionedDuration_Hr'] = total_duration_hr
            new_rows.append(row_dict)
        else:
            boundary_time = pd.Timestamp(datetime.combine(end_date, cutoff_time))
            if boundary_time < in_time: boundary_time += pd.Timedelta(days=1)
                
            ratio_day1 = (boundary_time - in_time).total_seconds() / total_seconds
            ratio_day2 = (out_time - boundary_time).total_seconds() / total_seconds
            
            row1, row2 = row_dict.copy(), row_dict.copy()
            row1['ProductionDate'], row1['ApportionedQty'] = start_date, qty * ratio_day1
            row1['ApportionedDuration_Hr'] = total_duration_hr * ratio_day1
            new_rows.append(row1)
            
            row2['ProductionDate'], row2['ApportionedQty'] = end_date, qty * ratio_day2
            row2['ApportionedDuration_Hr'] = total_duration_hr * ratio_day2
            new_rows.append(row2)
            
    return pd.DataFrame(new_rows)

cutoff_time_obj = time(cutoff_hour, 0)
df_split = split_cross_day_lots(filtered_df, cutoff_time_obj)

if df_split.empty:
    st.warning("所選條件下無生產資料，請重新調整 Filter。")
    st.stop()

# ==============================================================================
# --- 3. Dashboard UI ---
# ==============================================================================

# ---------------------------------------------------------
# Part A: Overall Performance (Top KPI Cards)
# ---------------------------------------------------------
st.markdown("### 📊 Overall Performance")

# 確保 PassQty 與 TestQty 抓取的是不重複的批次，避免跨日切割造成的重複計算
unique_lots = filtered_df.drop_duplicates(subset=['LotID'])
total_upd = int(df_split['ApportionedQty'].sum())
total_pass = int(unique_lots['PassQty'].sum())
total_test = int(unique_lots['TestQty'].sum())
overall_yield = (total_pass / total_test) * 100 if total_test > 0 else 0
active_testers = df_split['Tester'].nunique()

c1, c2, c3, c4 = st.columns(4)
with c1: st.markdown(f"<div class='kpi-card'><div class='kpi-title'>Total UPD (Units)</div><div class='kpi-value'>{total_upd:,}</div></div>", unsafe_allow_html=True)
with c2: st.markdown(f"<div class='kpi-card'><div class='kpi-title'>Total Pass</div><div class='kpi-value'>{total_pass:,}</div></div>", unsafe_allow_html=True)
with c3: st.markdown(f"<div class='kpi-card'><div class='kpi-title'>Overall Yield</div><div class='kpi-value'>{overall_yield:.2f}%</div></div>", unsafe_allow_html=True)
with c4: st.markdown(f"<div class='kpi-card'><div class='kpi-title'>Active Testers</div><div class='kpi-value'>{active_testers}</div></div>", unsafe_allow_html=True)

st.divider()

# ---------------------------------------------------------
# Part B: ATE Tester Capacity Analysis
# ---------------------------------------------------------
st.markdown("## ATE Tester 產能分析與驗證")

# 數據計算
daily_stats = df_split.groupby(['Tester', 'ProductionDate']).agg(
    Daily_UPD=('ApportionedQty', 'sum'),
    Daily_Duration=('ApportionedDuration_Hr', 'sum')
).reset_index()

daily_stats['Daily_OEE'] = (daily_stats['Daily_Duration'] / 24.0)

tester_summary = daily_stats.groupby('Tester').agg(
    Avg_UPD=('Daily_UPD', 'mean'),
    Avg_OEE=('Daily_OEE', 'mean')
).reset_index()

tester_summary['Avg_UPW'] = tester_summary['Avg_UPD'] * 7
lot_counts = filtered_df.groupby('Tester')['LotID'].nunique().reset_index().rename(columns={'LotID': 'Lots'})
tester_summary = tester_summary.merge(lot_counts, on='Tester', how='left')
tester_summary = tester_summary.sort_values(by='Avg_UPD', ascending=False)

# 摘要計算
avg_actual_upd = tester_summary['Avg_UPD'].mean()
avg_actual_upw = avg_actual_upd * 7
avg_actual_oee = tester_summary['Avg_OEE'].mean() * 100
max_upd, min_upd = tester_summary['Avg_UPD'].max(), tester_summary['Avg_UPD'].min()

st.markdown("#### 1. 關鍵分析結果摘要")
st.markdown(f"""
<div class='summary-text'>
<ul>
    <li>平均實際 UPD： <span class='highlight-val'>{avg_actual_upd:,.0f}</span></li>
    <li>平均實際 UPW： <span class='highlight-val'>{avg_actual_upw:,.0f}</span></li>
    <li>平均 OEE： <span class='highlight-val'>{avg_actual_oee:.1f}%</span></li>
    <li>UPD 範圍： <span class='highlight-val'>{min_upd:,.0f} ~ {max_upd:,.0f}</span></li>
</ul>
</div>
""", unsafe_allow_html=True)

st.markdown("<br>", unsafe_allow_html=True)

# ---------------------------------------------------------
# Part C: Tester Summary Table
# ---------------------------------------------------------
st.markdown("#### 2. 各機台表現分析 (Tester Summary)")

display_df = tester_summary.copy()
display_df = display_df.rename(columns={
    'Tester': 'Tester',
    'Avg_UPD': '平均 UPD',
    'Avg_UPW': '平均 UPW',
    'Avg_OEE': '平均 OEE',
    'Lots': '批次數量 (Lots)'
})

display_df['平均 UPD'] = display_df['平均 UPD'].apply(lambda x: f"{x:,.0f}")
display_df['平均 UPW'] = display_df['平均 UPW'].apply(lambda x: f"{x:,.0f}")
display_df['平均 OEE'] = display_df['平均 OEE'].apply(lambda x: f"{x*100:.1f}%")

st.dataframe(display_df, use_container_width=True, hide_index=True)