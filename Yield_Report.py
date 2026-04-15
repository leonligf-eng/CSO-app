import streamlit as st
import pandas as pd
import numpy as np
from datetime import datetime, timedelta, time

st.set_page_config(page_title="ATE Tester Capacity Analysis", layout="wide")

# --- Custom CSS ---
st.markdown("""
    <style>
    .kpi-card { background-color: #ffffff; border: 1px solid #e0e0e0; padding: 20px; border-radius: 8px; box-shadow: 2px 2px 5px rgba(0,0,0,0.05); text-align: center; margin-bottom: 25px; }
    .kpi-title { color: #6c757d; font-size: 12px; font-weight: bold; text-transform: uppercase; letter-spacing: 1px;}
    .kpi-value { color: #1E3A8A; font-size: 32px; font-weight: bold; margin-top: 10px; }
    .summary-card { background-color: #f8f9fa; border-left: 5px solid #28a745; padding: 15px 20px; border-radius: 6px; box-shadow: 1px 1px 4px rgba(0,0,0,0.04); text-align: left; margin-bottom: 20px; }
    .summary-title { color: #6c757d; font-size: 11px; font-weight: bold; text-transform: uppercase; letter-spacing: 0.5px;}
    .summary-value { color: #28a745; font-size: 26px; font-weight: bold; margin-top: 5px; }
    .summary-value-small { color: #28a745; font-size: 20px; font-weight: bold; margin-top: 10px; } 
    
    /* Notification/Insight Box Style */
    .insight-box { background-color: #e8f4f8; border-left: 5px solid #17a2b8; padding: 15px 20px; border-radius: 4px; margin-bottom: 20px; font-size: 15px; line-height: 1.6; color: #333;}
    .insight-highlight { font-weight: bold; color: #0c5460; }
    
    .stMultiSelect [data-baseweb="tag"] { max-width: 100% !important; }
    .stMultiSelect [data-baseweb="tag"] span { white-space: normal !important; max-width: none !important; overflow: visible !important; text-overflow: clip !important; }
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
    ops = ["FT1"]
    programs_map = {"FT1": ["PROD_GS631_Z..."]}
    
    for i in range(150):
        tester = np.random.choice(testers)
        op = np.random.choice(ops)
        prog = np.random.choice(programs_map[op])
        qty = np.random.randint(100, 5000)
        yield_rate = np.random.uniform(0.85, 0.99)
        pass_qty = int(qty * yield_rate)
        
        start_time = now - timedelta(days=np.random.randint(0, 10), hours=np.random.randint(0, 24))
        end_time = start_time + timedelta(hours=np.random.uniform(1, 10)) 
        
        data.append([f"LOT_{i:04d}", op, prog, tester, start_time, end_time, qty, pass_qty])
        
    return pd.DataFrame(data, columns=['LotNo', 'OpNo', 'ProgramName', 'Tester', 'CheckInTime', 'CheckOutTime', 'TestQty', 'PassQty'])

def load_data(file):
    if file is not None:
        try:
            df = pd.read_excel(file, sheet_name="Report")
            df.columns = df.columns.str.strip()
            df['CheckInTime'] = pd.to_datetime(df['CheckInTime'], errors='coerce')
            df['CheckOutTime'] = pd.to_datetime(df['CheckOutTime'], errors='coerce')
            df['TestQty'] = pd.to_numeric(df['TestQty'], errors='coerce').fillna(0)
            df['PassQty'] = pd.to_numeric(df['PassQty'], errors='coerce').fillna(0)
            return df
        except Exception as e:
            st.error(f"Failed to load file. Error: {str(e)}")
            return None
    return generate_mock_data()

# ==============================================================================
# --- 1. Sidebar Settings ---
# ==============================================================================
st.sidebar.header("📁 1. Data Input")
uploaded_file = st.sidebar.file_uploader("Upload Yield Report (Excel)", type=["xlsx", "xls"])
raw_df = load_data(uploaded_file)

if raw_df is None or raw_df.empty:
    st.warning("No data available.")
    st.stop()

st.sidebar.divider()
st.sidebar.header("⚙️ 2. Data Cleaning Rules")
min_lot_size = st.sidebar.number_input("Exclude Lots smaller than (Qty)", value=500, step=100)
cutoff_hour = st.sidebar.slider("Daily Cutoff Time (Hour)", min_value=0, max_value=23, value=0)

st.sidebar.divider()
st.sidebar.header("📐 3. Planning & Targets")
st.sidebar.caption("Define baselines for Capacity Planning.")
theo_max_upd = st.sidebar.number_input("Theoretical Max UPD (100% OEE)", value=4240, step=10)
# 🌟 NEW: 加入使用者的規劃基準
planned_upd = st.sidebar.number_input("Planned Target UPD", value=2800, step=100, help="產能規劃時使用的安全基準值")

# ==============================================================================
# --- 2. Main Area Filters ---
# ==============================================================================
st.markdown("### 🔍 Data Filters")
filter_col1, filter_col2 = st.columns([1, 2])

with filter_col1:
    op_options = sorted(raw_df['OpNo'].dropna().unique().tolist())
    selected_ops = st.multiselect("Select Operation (OpNo)", options=op_options)

filtered_by_op = raw_df[raw_df['OpNo'].isin(selected_ops)] if selected_ops else raw_df
prog_options = sorted(filtered_by_op['ProgramName'].dropna().unique().tolist())

with filter_col2:
    selected_progs = st.multiselect("Select Program (ProgramName)", options=prog_options)

if not selected_ops or not selected_progs:
    st.info("👆 Please select **OpNo** and **ProgramName** above to generate the report.")
    st.stop()

st.divider()

# ==============================================================================
# --- 3. Data Processing Engine ---
# ==============================================================================
filtered_df = raw_df[
    (raw_df['OpNo'].isin(selected_ops)) & 
    (raw_df['ProgramName'].isin(selected_progs)) &
    (raw_df['TestQty'] >= min_lot_size)
].copy()

@st.cache_data
def split_cross_day_lots(df, cutoff_time):
    df = df.loc[:, ~df.columns.duplicated()]
    df = df.dropna(subset=['CheckInTime', 'CheckOutTime', 'TestQty', 'PassQty']).copy()
    
    cutoff_hour = cutoff_time.hour
    cutoff_minute = cutoff_time.minute
    new_rows = []
    
    for _, row in df.iterrows():
        in_time, out_time = row['CheckInTime'], row['CheckOutTime']
        test_qty, pass_qty = row['TestQty'], row['PassQty']
        
        shift_in_time = in_time - pd.Timedelta(hours=cutoff_hour, minutes=cutoff_minute)
        shift_out_time = out_time - pd.Timedelta(hours=cutoff_hour, minutes=cutoff_minute)
        
        start_date, end_date = shift_in_time.date(), shift_out_time.date()
        total_seconds = (out_time - in_time).total_seconds()
        
        row_dict = row.to_dict()
        
        if start_date == end_date or total_seconds <= 0:
            row_dict['ProductionDate'] = start_date
            row_dict['ApportionedTestQty'] = test_qty
            row_dict['ApportionedPassQty'] = pass_qty
            new_rows.append(row_dict)
        else:
            boundary_time = pd.Timestamp(datetime.combine(end_date, cutoff_time))
            if boundary_time < in_time: boundary_time += pd.Timedelta(days=1)
                
            ratio_day1 = (boundary_time - in_time).total_seconds() / total_seconds
            ratio_day2 = (out_time - boundary_time).total_seconds() / total_seconds
            
            row1, row2 = row_dict.copy(), row_dict.copy()
            row1['ProductionDate'] = start_date
            row1['ApportionedTestQty'] = test_qty * ratio_day1
            row1['ApportionedPassQty'] = pass_qty * ratio_day1
            new_rows.append(row1)
            
            row2['ProductionDate'] = end_date
            row2['ApportionedTestQty'] = test_qty * ratio_day2
            row2['ApportionedPassQty'] = pass_qty * ratio_day2
            new_rows.append(row2)
            
    return pd.DataFrame(new_rows)

cutoff_time_obj = time(cutoff_hour, 0)
df_split = split_cross_day_lots(filtered_df, cutoff_time_obj)

if df_split.empty:
    st.warning("No production data available for the selected filters.")
    st.stop()

df_split['ProductionDate'] = pd.to_datetime(df_split['ProductionDate'])

# ==============================================================================
# --- 4. Dashboard UI ---
# ==============================================================================

# ---------------------------------------------------------
# Part A: Overall Performance
# ---------------------------------------------------------
st.markdown("### 📊 Overall Performance")

unique_lots = filtered_df.drop_duplicates(subset=['LotNo'])
total_test = int(unique_lots['TestQty'].sum())
total_pass = int(unique_lots['PassQty'].sum())
overall_yield = (total_pass / total_test) * 100 if total_test > 0 else 0
active_testers = df_split['Tester'].nunique()
valid_lots_count = len(unique_lots)

c1, c2, c3, c4 = st.columns(4)
with c1: st.markdown(f"<div class='kpi-card'><div class='kpi-title'>Total TestQty (Gross)</div><div class='kpi-value'>{total_test:,}</div></div>", unsafe_allow_html=True)
with c2: st.markdown(f"<div class='kpi-card'><div class='kpi-title'>Total PassQty (Net)</div><div class='kpi-value'>{total_pass:,}</div></div>", unsafe_allow_html=True)
with c3: st.markdown(f"<div class='kpi-card'><div class='kpi-title'>Overall Yield</div><div class='kpi-value'>{overall_yield:.2f}%</div></div>", unsafe_allow_html=True)
with c4: st.markdown(f"<div class='kpi-card'><div class='kpi-title'>Active Testers</div><div class='kpi-value'>{active_testers}</div></div>", unsafe_allow_html=True)

st.divider()

# ---------------------------------------------------------
# Part B: Key Analysis Summary
# ---------------------------------------------------------
st.markdown("## ATE Tester Capacity Analysis & Validation")

daily_stats = df_split.groupby(['Tester', 'ProductionDate']).agg(
    Daily_Gross_UPD=('ApportionedTestQty', 'sum'),
    Daily_Net_UPD=('ApportionedPassQty', 'sum')
).reset_index()

tester_summary = daily_stats.groupby('Tester').agg(
    Avg_Gross_UPD=('Daily_Gross_UPD', 'mean'),
    Avg_Net_UPD=('Daily_Net_UPD', 'mean')
).reset_index()

tester_summary['Avg_OEE'] = tester_summary['Avg_Gross_UPD'] / theo_max_upd

lot_counts = filtered_df.groupby('Tester')['LotNo'].nunique().reset_index().rename(columns={'LotNo': 'Lots'})
tester_summary = tester_summary.merge(lot_counts, on='Tester', how='left')
tester_summary = tester_summary.sort_values(by='Tester', ascending=True)

avg_gross_upd = tester_summary['Avg_Gross_UPD'].mean()
avg_net_upd = tester_summary['Avg_Net_UPD'].mean()
avg_oee = tester_summary['Avg_OEE'].mean() * 100
max_upd, min_upd = tester_summary['Avg_Gross_UPD'].max(), tester_summary['Avg_Gross_UPD'].min()

st.markdown("#### 1. Period Performance Summary")
st.caption(f"共計 {valid_lots_count} 個有效批次 (排除 < {min_lot_size} ea)。")

sc1, sc2, sc3 = st.columns(3)
with sc1: st.markdown(f"<div class='summary-card'><div class='summary-title'>Avg Actual UPD (TestQty)</div><div class='summary-value'>{avg_gross_upd:,.0f}</div></div>", unsafe_allow_html=True)
with sc2: st.markdown(f"<div class='summary-card'><div class='summary-title'>Avg Effective UPD (PassQty)</div><div class='summary-value'>{avg_net_upd:,.0f}</div></div>", unsafe_allow_html=True)
with sc3: st.markdown(f"<div class='summary-card'><div class='summary-title'>Overall OEE</div><div class='summary-value'>{avg_oee:.1f}%</div></div>", unsafe_allow_html=True)

# ---------------------------------------------------------
# Part C: Capacity Planning Insights (🌟 NEW)
# ---------------------------------------------------------
st.markdown("#### 2. Capacity Planning Insights (結論與規劃建議)")

# 計算統計與規劃指標
buffer_pct = ((avg_gross_upd - planned_upd) / avg_gross_upd) * 100 if avg_gross_upd > 0 else 0
implied_oee = (planned_upd / theo_max_upd) * 100 if theo_max_upd > 0 else 0

# 計算規劃值在歷史每日資料中的百分位數 (PR值)
all_daily_upds = daily_stats['Daily_Gross_UPD']
percentile_rank = (all_daily_upds < planned_upd).mean() * 100 if len(all_daily_upds) > 0 else 0

# 動態產出結論文字
if avg_gross_upd >= planned_upd:
    insight_text = f"""
    驗證結果顯示，目前的 ATE Tester 實際產能平均約為 <span class='insight-highlight'>{avg_gross_upd:,.0f} ea/day</span>，高於您預估的規劃基準 <span class='insight-highlight'>{planned_upd:,.0f}</span>。<br>
    如果您是以 {planned_upd:,.0f} 作為產能規劃 (Capacity Planning) 的基準，這是一個安全的設定，保留了約 <span class='insight-highlight'>{buffer_pct:.1f}%</span> 的緩衝空間以應對機台異常或換線損失。
    """
else:
    insight_text = f"""
    ⚠️ <b>注意：</b> 目前的實際產能平均約為 <span class='insight-highlight'>{avg_gross_upd:,.0f} ea/day</span>，<b>低於</b>您預估的規劃基準 <span class='insight-highlight'>{planned_upd:,.0f}</span>。<br>
    建議調降規劃基準，或檢視產線是否有異常 Downtime 導致產能未達標。
    """

st.markdown(f"<div class='insight-box'>{insight_text}</div>", unsafe_allow_html=True)

# 產出規劃比較表
insight_data = {
    "指標 (Metric)": ["報表平均 UPD", "您預估的 UPD", "理論最大 UPD", f"隱含 OEE (於 {planned_upd})"],
    "數值 (Value)": [f"{avg_gross_upd:,.0f}", f"{planned_upd:,.0f}", f"{theo_max_upd:,.0f}", f"{implied_oee:.1f}%"],
    "備註 (Note)": [
        "實際產出表現",
        f"約為實際數據的 PR{percentile_rank:.0f} 分位數 (保守估計)",
        "100% OEE, 無重測/無異常",
        "排程預設包含換料、異常、重測之折損"
    ]
}
insight_df = pd.DataFrame(insight_data)
st.dataframe(insight_df, use_container_width=True, hide_index=True)


# ---------------------------------------------------------
# Part D: Tester Summary Table
# ---------------------------------------------------------
st.markdown("#### 3. Tester Performance Details (各機台表現明細)")

display_df = tester_summary[['Tester', 'Lots', 'Avg_Gross_UPD', 'Avg_Net_UPD', 'Avg_OEE']].copy()
display_df = display_df.rename(columns={
    'Tester': 'Tester',
    'Lots': 'Lot Count',
    'Avg_Gross_UPD': 'Avg UPD (TestQty)',
    'Avg_Net_UPD': 'Avg UPD (PassQty)',
    'Avg_OEE': 'Avg OEE'
})

display_df['Avg UPD (TestQty)'] = display_df['Avg UPD (TestQty)'].apply(lambda x: f"{x:,.0f}")
display_df['Avg UPD (PassQty)'] = display_df['Avg UPD (PassQty)'].apply(lambda x: f"{x:,.0f}")
display_df['Avg OEE'] = display_df['Avg OEE'].apply(lambda x: f"{x*100:.1f}%")

st.dataframe(display_df, use_container_width=True, hide_index=True)
