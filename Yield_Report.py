import streamlit as st
import pandas as pd
import numpy as np
from datetime import datetime, timedelta, time

st.set_page_config(page_title="ATE Tester Capacity Analysis", layout="wide")

# --- Custom CSS (Added new styles for Sub-KPI cards) ---
st.markdown("""
    <style>
    /* Top KPI Card Style */
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
    
    /* Sub-KPI Card Style for Key Analysis */
    .summary-card {
        background-color: #f8f9fa;
        border-left: 5px solid #28a745;
        padding: 15px 20px;
        border-radius: 6px;
        box-shadow: 1px 1px 4px rgba(0,0,0,0.04);
        text-align: left;
        margin-bottom: 20px;
    }
    .summary-title { color: #6c757d; font-size: 11px; font-weight: bold; text-transform: uppercase; letter-spacing: 0.5px;}
    .summary-value { color: #28a745; font-size: 26px; font-weight: bold; margin-top: 5px; }
    .summary-value-small { color: #28a745; font-size: 20px; font-weight: bold; margin-top: 10px; } /* For Range text */
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
    
    for i in range(200):
        tester = np.random.choice(testers)
        op = np.random.choice(ops)
        prog = np.random.choice(programs_map[op])
        qty = np.random.randint(1500, 5000)
        yield_rate = np.random.uniform(0.85, 0.99)
        pass_qty = int(qty * yield_rate)
        
        start_time = now - timedelta(days=np.random.randint(0, 21), hours=np.random.randint(0, 24))
        end_time = start_time + timedelta(hours=np.random.uniform(2, 16)) 
        
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
            st.error(f"Failed to load file. Please ensure the 'Report' sheet exists. Error: {str(e)}")
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
st.sidebar.header("⚙️ 2. Calculation Settings")
cutoff_hour = st.sidebar.slider("Daily Cutoff Time (Hour)", min_value=0, max_value=23, value=0, help="0 = 24:00 (Midnight).")

# ==============================================================================
# --- 2. Main Area: Help Section & Filters ---
# ==============================================================================
with st.expander("ℹ️ Help: Formula & Parameter Definitions"):
    st.markdown("""
    ### 📖 Calculation Logic
    To ensure precision in daily capacity and equipment efficiency, this system employs a **"Cross-Day Apportionment"** algorithm to prevent data distortion caused by lots crossing the daily cutoff time.

    #### Phase 1: Cross-Day Apportionment
    For any lot that crosses the designated Daily Cutoff Time, the system splits its output and duration proportionally:
    * **Time Ratio (R):** `Seconds tested on Current Day / Total Test Seconds`
    * **Apportioned Qty:** `TestQty * R`
    * **Apportioned Duration (Hours):** `Total Duration (Hours) * R`

    #### Phase 2: Tester Summary Metrics
    After the foundational apportionment, metrics are aggregated per ATE:
    * **Daily UPD:** `SUM(Apportioned Qty)` for a specific day.
    * **Avg UPD:** `SUM(Daily UPD) / Days with production activity`.
    * **Real UPW:** Aggregates production into actual calendar weeks, then calculates the `AVERAGE(Weekly Output)`.
    * **Daily OEE:** `SUM(Apportioned Duration) / 24` *(Using 24 hours as full load base)*.
    * **Avg OEE:** `(SUM(Daily OEE) / Days with production activity) * 100%`.
    * **Lot Count:** Number of unique lots processed by the tester during the selected period.
    """)

st.markdown("### 🔍 Data Filters")

filter_col1, filter_col2 = st.columns(2)

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
    st.warning("No production data available for the selected filters.")
    st.stop()

# Essential formatting for Real UPW logic
df_split['ProductionDate'] = pd.to_datetime(df_split['ProductionDate'])

# ==============================================================================
# --- 4. Dashboard UI ---
# ==============================================================================

# ---------------------------------------------------------
# Part A: Overall Performance (Top KPI Cards)
# ---------------------------------------------------------
st.markdown("### 📊 Overall Performance")

unique_lots = filtered_df.drop_duplicates(subset=['LotNo'])
total_upd = int(df_split['ApportionedQty'].sum())
total_pass = int(unique_lots['PassQty'].sum())
total_test = int(unique_lots['TestQty'].sum())
overall_yield = (total_pass / total_test) * 100 if total_test > 0 else 0
active_testers = df_split['Tester'].nunique()

c1, c2, c3, c4 = st.columns(4)
with c1: st.markdown(f"<div class='kpi-card'><div class='kpi-title'>Total Output (Units)</div><div class='kpi-value'>{total_upd:,}</div></div>", unsafe_allow_html=True)
with c2: st.markdown(f"<div class='kpi-card'><div class='kpi-title'>Total Pass</div><div class='kpi-value'>{total_pass:,}</div></div>", unsafe_allow_html=True)
with c3: st.markdown(f"<div class='kpi-card'><div class='kpi-title'>Overall Yield</div><div class='kpi-value'>{overall_yield:.2f}%</div></div>", unsafe_allow_html=True)
with c4: st.markdown(f"<div class='kpi-card'><div class='kpi-title'>Active Testers</div><div class='kpi-value'>{active_testers}</div></div>", unsafe_allow_html=True)

st.divider()

# ---------------------------------------------------------
# Part B: ATE Tester Capacity Analysis
# ---------------------------------------------------------
st.markdown("## ATE Tester Capacity Analysis & Validation")

# Calculate Daily metrics
daily_stats = df_split.groupby(['Tester', 'ProductionDate']).agg(
    Daily_UPD=('ApportionedQty', 'sum'),
    Daily_Duration=('ApportionedDuration_Hr', 'sum')
).reset_index()

daily_stats['Daily_OEE'] = (daily_stats['Daily_Duration'] / 24.0)

tester_summary = daily_stats.groupby('Tester').agg(
    Avg_UPD=('Daily_UPD', 'mean'),
    Avg_OEE=('Daily_OEE', 'mean')
).reset_index()

# Real UPW Calculation
weekly_stats = df_split.groupby(['Tester', pd.Grouper(key='ProductionDate', freq='W')]).agg(
    Weekly_UPW=('ApportionedQty', 'sum')
).reset_index()

avg_weekly = weekly_stats.groupby('Tester').agg(
    Real_UPW=('Weekly_UPW', 'mean')
).reset_index()

tester_summary = tester_summary.merge(avg_weekly, on='Tester', how='left')

# Get Lot Counts
lot_counts = filtered_df.groupby('Tester')['LotNo'].nunique().reset_index().rename(columns={'LotNo': 'Lots'})
tester_summary = tester_summary.merge(lot_counts, on='Tester', how='left')

# Sort by Tester ascending
tester_summary = tester_summary.sort_values(by='Tester', ascending=True)

# Global Summary Aggregation
avg_actual_upd = tester_summary['Avg_UPD'].mean()
avg_actual_upw = tester_summary['Real_UPW'].mean()
avg_actual_oee = tester_summary['Avg_OEE'].mean() * 100
max_upd, min_upd = tester_summary['Avg_UPD'].max(), tester_summary['Avg_UPD'].min()

st.markdown("#### 1. Key Analysis Summary")

# 🌟 NEW: Summary Sub-KPI Cards Design
sc1, sc2, sc3, sc4 = st.columns(4)

with sc1:
    st.markdown(f"<div class='summary-card'><div class='summary-title'>Avg Actual UPD</div><div class='summary-value'>{avg_actual_upd:,.0f}</div></div>", unsafe_allow_html=True)
with sc2:
    st.markdown(f"<div class='summary-card'><div class='summary-title'>Real Avg UPW</div><div class='summary-value'>{avg_actual_upw:,.0f}</div></div>", unsafe_allow_html=True)
with sc3:
    st.markdown(f"<div class='summary-card'><div class='summary-title'>Avg OEE</div><div class='summary-value'>{avg_actual_oee:.1f}%</div></div>", unsafe_allow_html=True)
with sc4:
    # Use smaller font for range so it fits nicely
    st.markdown(f"<div class='summary-card'><div class='summary-title'>UPD Range</div><div class='summary-value-small'>{min_upd:,.0f} ~ {max_upd:,.0f}</div></div>", unsafe_allow_html=True)


# ---------------------------------------------------------
# Part C: Tester Summary Table
# ---------------------------------------------------------
st.markdown("#### 2. Tester Performance Analysis (Summary)")

display_df = tester_summary.copy()
display_df = display_df.rename(columns={
    'Tester': 'Tester',
    'Avg_UPD': 'Avg UPD',
    'Real_UPW': 'Real UPW',
    'Avg_OEE': 'Avg OEE',
    'Lots': 'Lot Count'
})

display_df['Avg UPD'] = display_df['Avg UPD'].apply(lambda x: f"{x:,.0f}")
display_df['Real UPW'] = display_df['Real UPW'].apply(lambda x: f"{x:,.0f}")
display_df['Avg OEE'] = display_df['Avg OEE'].apply(lambda x: f"{x*100:.1f}%")

st.dataframe(display_df, use_container_width=True, hide_index=True)