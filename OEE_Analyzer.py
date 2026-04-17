import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
from datetime import datetime, timedelta, time

st.set_page_config(page_title="ATE Capacity & OEE Analyzer", layout="wide")

# --- Custom CSS ---
st.markdown("""
    <style>
    .kpi-card { background-color: #ffffff; border: 1px solid #e0e0e0; padding: 20px; border-radius: 8px; box-shadow: 2px 2px 5px rgba(0,0,0,0.05); text-align: center; margin-bottom: 25px; }
    .kpi-title { color: #6c757d; font-size: 12px; font-weight: bold; text-transform: uppercase; letter-spacing: 1px;}
    .kpi-value { color: #1E3A8A; font-size: 32px; font-weight: bold; margin-top: 10px; }
    .summary-card { background-color: #f8f9fa; border-left: 5px solid #28a745; padding: 15px 20px; border-radius: 6px; box-shadow: 1px 1px 4px rgba(0,0,0,0.04); text-align: left; margin-bottom: 20px; }
    .summary-title { color: #6c757d; font-size: 11px; font-weight: bold; text-transform: uppercase; letter-spacing: 0.5px;}
    .summary-value { color: #28a745; font-size: 26px; font-weight: bold; margin-top: 5px; }
    .insight-box { background-color: #e8f4f8; border-left: 5px solid #17a2b8; padding: 15px 20px; border-radius: 4px; margin-bottom: 20px; font-size: 15px; line-height: 1.6; color: #333;}
    .insight-highlight { font-weight: bold; color: #0c5460; }
    
    .stMultiSelect [data-baseweb="tag"] { 
        max-width: 100% !important; 
        background-color: #e6f2ff !important;
        color: #0056b3 !important;
        border: 1px solid #b8daff !important;
    }
    .stMultiSelect [data-baseweb="tag"] span { 
        white-space: normal !important; 
        max-width: none !important; 
        overflow: visible !important; 
        text-overflow: clip !important; 
    }
    .stMultiSelect [data-baseweb="tag"] svg {
        fill: #0056b3 !important;
    }
    
    .stTabs [data-baseweb="tab-list"] {
        gap: 8px;
        border-bottom: 2px solid #e0e0e0;
    }
    .stTabs [data-baseweb="tab"] {
        font-size: 18px !important;
        font-weight: 600 !important;
        padding: 12px 24px !important;
        background-color: #f8f9fa;
        border: 1px solid #dee2e6;
        border-bottom: none;
        border-radius: 8px 8px 0 0;
        color: #6c757d;
        transition: all 0.2s ease-in-out;
    }
    .stTabs [data-baseweb="tab"]:hover {
        background-color: #e2e6ea;
        color: #495057;
    }
    .stTabs [aria-selected="true"] {
        background-color: #ffffff !important;
        color: #0056b3 !important;
        border-top: 4px solid #0056b3 !important;
        border-left: 1px solid #dee2e6 !important;
        border-right: 1px solid #dee2e6 !important;
        box-shadow: 0 -3px 6px rgba(0,0,0,0.04);
    }
    
    /* 🌟 強制 Streamlit Dataframe 內容與表頭置中 */
    [data-testid="stDataFrame"] th {
        text-align: center !important;
    }
    [data-testid="stDataFrame"] td {
        text-align: center !important;
    }
    
    /* 🌟 優化 Quick Action 按鈕的外觀 */
    div[data-testid="column"] button {
        padding: 0.2rem 0.5rem !important;
        font-size: 0.85rem !important;
        min-height: 32px !important;
        border-radius: 4px !important;
    }
    </style>
""", unsafe_allow_html=True)

st.markdown("# 📈 ATE OEE Analyzer")

# ==============================================================================
# --- 0. Mock Data Generator ---
# ==============================================================================
@st.cache_data
def generate_mock_data():
    np.random.seed(42)
    now = datetime.now()
    data = []
    testers = [f"HP93K-EXA{i:02d}" for i in range(2, 10)]
    ops = ["FT1", "FTA", "MT1"]
    
    # 🌟 微調了 Mock Data，加入 Phase 關鍵字以便展示自動 Mapping 功能
    programs_map = {
        "FT1": ["PROD_GS631_FT1_Proto1.0", "PROD_GS631_FT1_EVT1.0(A0)", "PROD_GS631_FT1_DVT"], 
        "FTA": ["PROD_GS631_FTA_Proto1.1", "PROD_GS631_FTA_EVT1.1", "PROD_GS631_FTA_PVT"], 
        "MT1": ["PROD_GS631_MT1_EVT1.0(B0)", "PROD_GS631_MT1_MP"]
    }
    products = ["SS16G", "MU16G", "HY12G", "SS12G"]
    
    tester_clocks = {tester: now - timedelta(days=30) for tester in testers}
    
    for i in range(250):
        tester = np.random.choice(testers)
        op = np.random.choice(ops)
        prog = np.random.choice(programs_map[op])
        prod = np.random.choice(products)
        
        wait_time = timedelta(hours=np.random.uniform(0.5, 3))
        start_time = tester_clocks[tester] + wait_time
        test_duration = timedelta(hours=np.random.uniform(6, 14))
        end_time = start_time + test_duration
        
        if op == "FT1": qty = np.random.randint(3000, 5000)
        elif op == "FTA": qty = np.random.randint(5000, 8000)
        else: qty = np.random.randint(15000, 25000)
        
        first_yield_rate = np.random.uniform(0.85, 0.95)
        final_yield_rate = np.random.uniform(first_yield_rate, 0.99)
        
        first_pass_qty = int(qty * first_yield_rate)
        pass_qty = int(qty * final_yield_rate)
        
        data.append([f"LOT_{i:04d}", prod, op, prog, tester, start_time, end_time, qty, first_pass_qty, pass_qty])
        tester_clocks[tester] = end_time
        
    return pd.DataFrame(data, columns=['LotNo', 'ProductNo', 'OpNo', 'ProgramName', 'Tester', 'CheckInTime', 'CheckOutTime', 'TestQty', 'First Pass Qty', 'PassQty'])

def load_data(file):
    if file is not None:
        try:
            df = pd.read_excel(file, sheet_name="Report")
            df.columns = df.columns.str.strip()
            df['CheckInTime'] = pd.to_datetime(df['CheckInTime'], errors='coerce')
            df['CheckOutTime'] = pd.to_datetime(df['CheckOutTime'], errors='coerce')
            df['TestQty'] = pd.to_numeric(df['TestQty'], errors='coerce').fillna(0)
            df['PassQty'] = pd.to_numeric(df['PassQty'], errors='coerce').fillna(0)
            
            if 'First Pass Qty' in df.columns:
                 df['First Pass Qty'] = pd.to_numeric(df['First Pass Qty'], errors='coerce').fillna(df['PassQty'])
            else:
                 df['First Pass Qty'] = df['PassQty']
                 
            if 'ProductNo' not in df.columns:
                df['ProductNo'] = "Unknown_Product"
            return df
        except Exception as e:
            st.error(f"Failed to load file. Error: {str(e)}")
            return None
    return generate_mock_data()

uploaded_file = st.sidebar.file_uploader("Upload Yield Report", type=["xlsx", "xls", "ods"])
raw_df = load_data(uploaded_file)

if raw_df is None or raw_df.empty:
    st.warning("No data available.")
    st.stop()

global_min_date = raw_df['CheckInTime'].min().date() if pd.notnull(raw_df['CheckInTime'].min()) else datetime.now().date()
global_max_date = raw_df['CheckOutTime'].max().date() if pd.notnull(raw_df['CheckOutTime'].max()) else datetime.now().date()

# ==============================================================================
# --- 1. Main Area: Filters & Help Section ---
# ==============================================================================
with st.expander("ℹ️ Help: Formula & Parameter Definitions"):
    st.markdown("""
    This system employs rigorous Industrial Engineering (IE) logic combined with actual production report data to calculate authentic equipment efficiency. Metric definitions are as follows:

    #### 1. Capacity Metrics (Tester View vs. Product View)
    * **Test Qty:** Total testing actions performed by the testers across all operations.
    * **Active Days:** The exact number of hours a tester spent in the "Testing" state (CheckIn to CheckOut), divided by 24 hours.
    * **Normalized UPD:** Calculated as `Total Qty / Active Days`. This metric converts the actual volume tested over a specific time period into an equivalent 24-hour rate for standardized comparison.

    #### 2. OEE Breakdown (Availability, Performance, Quality)
    * **A (Availability):** Measures how often the tester is actually in production.
      * `Calculation = Active Days / Adjusted Calendar Span Days`.
      *(Note: Full empty days (gaps ≥ 24h) between lots are automatically deducted to ensure fair analysis when filtering specific programs.)*
    * **P (Performance):** Measures if the tester is running at theoretical speed when active.
      * `Theoretical UPH = Theoretical Max UPD / 24`
      * `Actual UPH = Total Test Qty / Total Active Hours`
      * `Calculation = Actual UPH / Theoretical UPH`
    * **Q (Quality):** Testing quality. Includes both First Yield and Final Yield.
      * `First Yield = Total First Pass Qty / Total Test Qty`
      * `Final Yield = Total Final Pass Qty / Total Test Qty`
    * **Overall OEE:**
      * `Calculation = Availability (A) × Performance (P)`

    #### 3. Capacity Planning Metrics
    * **Planned Target UPD:** The safety scheduling baseline preset by Planning or Engineering.
    * **Implied OEE:** `Calculation = Planned Target UPD / Theoretical Max UPD`. Reflects the built-in buffer percentage reserved for setups, maintenance, and re-tests.
    """)

st.sidebar.header("⚙️ Data Cleaning Rules")
min_lot_size = st.sidebar.number_input("Exclude Lots smaller than (Qty)", value=0, step=100)
st.sidebar.divider()

st.markdown("### 🔍 Data Filters")

filter_col1, filter_col2 = st.columns([1, 1])

with filter_col1:
    op_options = sorted(raw_df['OpNo'].dropna().unique().tolist())
    selected_ops = st.multiselect("Select Operation (OpNo)", options=op_options, key="op_select")

filtered_by_op = raw_df[raw_df['OpNo'].isin(selected_ops)] if selected_ops else raw_df
prog_options = sorted(filtered_by_op['ProgramName'].dropna().unique().tolist())

if 'saved_progs' not in st.session_state: st.session_state.saved_progs = []
def update_progs(): st.session_state.saved_progs = st.session_state.prog_select_widget
valid_defaults_prog = [p for p in st.session_state.saved_progs if p in prog_options]

with filter_col2:
    selected_progs = st.multiselect("Select Program", options=prog_options, default=valid_defaults_prog, key="prog_select_widget", on_change=update_progs)

filtered_by_op_prog = filtered_by_op[filtered_by_op['ProgramName'].isin(selected_progs)] if selected_progs else filtered_by_op

if not filtered_by_op_prog.empty:
    curr_min_date = filtered_by_op_prog['CheckInTime'].min().date() if pd.notnull(filtered_by_op_prog['CheckInTime'].min()) else global_min_date
    curr_max_date = filtered_by_op_prog['CheckOutTime'].max().date() if pd.notnull(filtered_by_op_prog['CheckOutTime'].max()) else global_max_date
else:
    curr_min_date, curr_max_date = global_min_date, global_max_date

st.session_state.curr_max_date_ref = curr_max_date
st.session_state.curr_min_date_ref = curr_min_date

curr_selection_hash = hash(str(selected_ops) + str(selected_progs))
if "last_selection_hash" not in st.session_state or st.session_state.last_selection_hash != curr_selection_hash:
    st.session_state.date_picker = (curr_min_date, curr_max_date)
    st.session_state.last_selection_hash = curr_selection_hash

if "date_picker" not in st.session_state:
    st.session_state.date_picker = (curr_min_date, curr_max_date)

def update_date_range(days=None, to_max=False):
    val = st.session_state.date_picker
    if isinstance(val, tuple) and len(val) > 0:
        start_dt = val[0]
    elif isinstance(val, tuple) and len(val) == 0:
        start_dt = st.session_state.curr_min_date_ref
    elif val is None:
        start_dt = st.session_state.curr_min_date_ref
    else:
        start_dt = val
        
    if to_max:
        new_end = st.session_state.curr_max_date_ref
    else:
        new_end = min(start_dt + timedelta(days=days), st.session_state.curr_max_date_ref)
        
    new_end = max(start_dt, new_end)
    st.session_state.date_picker = (start_dt, new_end)

filter_col3, filter_col4 = st.columns([1, 1])

with filter_col3:
    st.markdown("<p style='font-size: 14px; margin-bottom: 2px; color: #31333F;'>Select Date Range (CheckIn/Out Overlap)</p>", unsafe_allow_html=True)
    
    date_range = st.date_input(
        "Date Range", 
        key="date_picker", 
        min_value=global_min_date, 
        max_value=global_max_date,
        label_visibility="collapsed"
    )
    
    if isinstance(date_range, tuple):
        if len(date_range) == 2:
            start_date, end_date = date_range
        elif len(date_range) == 1:
            start_date = end_date = date_range[0]
        else:
            start_date = end_date = curr_min_date
    else:
        start_date = end_date = date_range

    st.markdown("<div style='margin-top: 5px;'></div>", unsafe_allow_html=True)
    c_btn1, c_btn2, c_btn3, c_btn4 = st.columns(4)
    
    c_btn1.button("+ 1 Week", use_container_width=True, on_click=update_date_range, kwargs={"days": 7})
    c_btn2.button("+ 2 Weeks", use_container_width=True, on_click=update_date_range, kwargs={"days": 14})
    c_btn3.button("+ 4 Weeks", use_container_width=True, on_click=update_date_range, kwargs={"days": 28})
    c_btn4.button("Max Range", use_container_width=True, on_click=update_date_range, kwargs={"to_max": True})


mask_stage_3 = (
    (filtered_by_op_prog['CheckInTime'].dt.date <= end_date) & 
    (filtered_by_op_prog['CheckOutTime'].dt.date >= start_date) &
    (filtered_by_op_prog['TestQty'] >= min_lot_size)
)
filtered_by_op_prog_date = filtered_by_op_prog[mask_stage_3]

prod_options = sorted(filtered_by_op_prog_date['ProductNo'].dropna().unique().tolist())

upstream_hash = hash(str(selected_ops) + str(selected_progs) + str(start_date) + str(end_date) + str(min_lot_size))
if "upstream_hash" not in st.session_state or st.session_state.upstream_hash != upstream_hash:
    st.session_state.prod_select_widget = prod_options  
    st.session_state.upstream_hash = upstream_hash      

with filter_col4:
    selected_prods = st.multiselect(
        "Select Product (ProductNo)", 
        options=prod_options, 
        key="prod_select_widget"
    )

st.divider()

# ==============================================================================
# --- 2. Sidebar Settings (Calculator & Targets) ---
# ==============================================================================

st.sidebar.header("🧮 Capacity Calculator")
with st.sidebar.expander("Detailed Parameters", expanded=False):
    calc_lot_size = st.number_input("Lot Size", value=6600, step=100)
    calc_site = st.selectbox("Site", options=[1, 2, 4, 8, 16, 32, 64, 128, 256], index=3)
    calc_test_time = st.number_input("Test Time (s)", value=150.00, step=1.00, format="%.2f")
    calc_fpy = st.number_input("FPY %", value=95.00, step=1.00, format="%.2f")
    calc_oee = st.number_input("OEE %", value=70.00, step=1.00, format="%.2f")
    
    if calc_test_time > 0:
        single_cap = (86400 / calc_test_time) * calc_site * (calc_oee / 100.0) * (calc_fpy / 100.0)
    else:
        single_cap = 0
        
    st.markdown(f"""
        <div style='background-color: #f8f9fa; padding: 15px; border-radius: 8px; margin-top: 10px; border: 1px solid #dee2e6; text-align: center;'>
            <div style='font-size: 11px; color: #6c757d; font-weight: bold; text-transform: uppercase; margin-bottom: 5px; letter-spacing: 0.5px;'>
                Single Capacity (UPD)
            </div>
            <div style='font-size: 26px; color: #1E3A8A; font-weight: bold;'>
                🚀 {single_cap:,.0f} <span style='font-size: 13px; color: #6c757d;'>units/day</span>
            </div>
        </div>
    """, unsafe_allow_html=True)

st.sidebar.divider()

st.sidebar.header("📐 Planning & Targets")
st.sidebar.markdown("<p style='color: #444; font-size: 13px;'>Define specific baselines for EACH selected operation.</p>", unsafe_allow_html=True)

targets = {}
if selected_ops:
    for op in selected_ops:
        st.sidebar.markdown(f"**🔹 Operation: {op}**")
        theo = st.sidebar.number_input(f"Theo Max UPD ({op})", value=4240, step=10, key=f"theo_{op}")
        plan = st.sidebar.number_input(f"Planned Target ({op})", value=2800, step=100, key=f"plan_{op}")
        targets[op] = {'theo': theo, 'plan': plan}
        st.sidebar.write("") 

if not selected_ops or not selected_progs or not selected_prods:
    st.info("👆 Please select **OpNo, ProgramName, and ProductNo** above to generate the report.")
    st.stop()

# ==============================================================================
# --- 3. Data Processing Engine ---
# ==============================================================================
mask_final = (
    filtered_by_op_prog_date['ProductNo'].isin(selected_prods)
)

filtered_df = filtered_by_op_prog_date[mask_final].copy()

if filtered_df.empty:
    st.warning("No production data available for the selected filters and date range.")
    st.stop()

filtered_df['Duration_Hr'] = (filtered_df['CheckOutTime'] - filtered_df['CheckInTime']).dt.total_seconds() / 3600.0

df_sorted = filtered_df.sort_values(by=['Tester', 'OpNo', 'CheckInTime']).copy()
df_sorted['Next_CheckIn'] = df_sorted.groupby(['Tester', 'OpNo'])['CheckInTime'].shift(-1)
df_sorted['Gap_Days'] = (df_sorted['Next_CheckIn'] - df_sorted['CheckOutTime']).dt.total_seconds() / 86400.0
df_sorted['Empty_Days'] = np.where(df_sorted['Gap_Days'] >= 1.0, np.floor(df_sorted['Gap_Days']), 0)
empty_span = df_sorted.groupby(['Tester', 'OpNo'])['Empty_Days'].sum().reset_index()

tester_summary = filtered_df.groupby(['Tester', 'OpNo']).agg(
    Lot_Count=('LotNo', 'nunique'),
    Total_TestQty=('TestQty', 'sum'),
    Total_FirstPassQty=('First Pass Qty', 'sum'),
    Total_PassQty=('PassQty', 'sum'),
    Total_Duration_Hr=('Duration_Hr', 'sum'),
    Min_CheckIn=('CheckInTime', 'min'),
    Max_CheckOut=('CheckOutTime', 'max')
).reset_index()

tester_summary = tester_summary.merge(empty_span, on=['Tester', 'OpNo'], how='left')
tester_summary['Empty_Days'] = tester_summary['Empty_Days'].fillna(0)

tester_summary['Active_Days'] = tester_summary['Total_Duration_Hr'] / 24.0
tester_summary['Avg_Gross_UPD'] = np.where(tester_summary['Active_Days'] > 0, tester_summary['Total_TestQty'] / tester_summary['Active_Days'], 0)
tester_summary['Avg_Net_UPD'] = np.where(tester_summary['Active_Days'] > 0, tester_summary['Total_PassQty'] / tester_summary['Active_Days'], 0)

tester_summary['Raw_Calendar_Days'] = (tester_summary['Max_CheckOut'] - tester_summary['Min_CheckIn']).dt.total_seconds() / 86400.0
tester_summary['Adjusted_Span_Days'] = np.maximum(tester_summary['Raw_Calendar_Days'] - tester_summary['Empty_Days'], tester_summary['Active_Days'])

tester_summary['Availability (A)'] = np.where(tester_summary['Adjusted_Span_Days'] > 0, tester_summary['Active_Days'] / tester_summary['Adjusted_Span_Days'], 0)

tester_summary['Theo_Max_UPD'] = tester_summary['OpNo'].map(lambda x: targets[x]['theo'])
tester_summary['Planned_UPD'] = tester_summary['OpNo'].map(lambda x: targets[x]['plan'])

tester_summary['Actual_UPH'] = np.where(tester_summary['Total_Duration_Hr'] > 0, tester_summary['Total_TestQty'] / tester_summary['Total_Duration_Hr'], 0)
tester_summary['Performance (P)'] = tester_summary['Actual_UPH'] / (tester_summary['Theo_Max_UPD'] / 24.0)

tester_summary['First_Yield'] = np.where(tester_summary['Total_TestQty'] > 0, tester_summary['Total_FirstPassQty'] / tester_summary['Total_TestQty'], 0)
tester_summary['Final_Yield (Q)'] = np.where(tester_summary['Total_TestQty'] > 0, tester_summary['Total_PassQty'] / tester_summary['Total_TestQty'], 0)

tester_summary['Avg_OEE'] = tester_summary['Availability (A)'] * tester_summary['Performance (P)']

tester_summary = tester_summary.sort_values(by=['OpNo', 'Tester'], ascending=True)

# ==============================================================================
# --- 4. Dashboard UI ---
# ==============================================================================
st.markdown("<p style='color: #444; font-size: 14px;'>Select an Operation tab below to view its isolated performance and capacity planning insights.</p>", unsafe_allow_html=True)

# 🌟 V52: 準備 Tabs，將 "Overall Build Yield" 固定放在最右側
tab_names = list(selected_ops) + ["🧬 Overall Build Yield"]
tabs = st.tabs(tab_names)

# --- 處理一般站點的 Tabs ---
for idx, op in enumerate(selected_ops):
    with tabs[idx]:
        st.write("") 
        
        op_df = filtered_df[filtered_df['OpNo'] == op]
        op_summary = tester_summary[tester_summary['OpNo'] == op]
        
        if op_summary.empty:
            st.info(f"No data available for Operation {op} under the current filters.")
            continue
        
        # ---------------------------------------------------------
        # Part A: Overall Performance 
        # ---------------------------------------------------------
        st.markdown(f"#### 📈 Overall Performance ({op})")
        
        total_insertions = int(op_summary['Total_TestQty'].sum())
        total_pass_insertions = int(op_summary['Total_PassQty'].sum())
        avg_step_yield = (total_pass_insertions / total_insertions) * 100 if total_insertions > 0 else 0
        active_testers = op_summary['Tester'].nunique()

        c1, c2, c3, c4 = st.columns(4)
        with c1: st.markdown(f"<div class='kpi-card'><div class='kpi-title'>Test Qty</div><div class='kpi-value'>{total_insertions:,}</div></div>", unsafe_allow_html=True)
        with c2: st.markdown(f"<div class='kpi-card'><div class='kpi-title'>Pass Qty</div><div class='kpi-value'>{total_pass_insertions:,}</div></div>", unsafe_allow_html=True)
        with c3: st.markdown(f"<div class='kpi-card'><div class='kpi-title'>Avg Final Yield</div><div class='kpi-value'>{avg_step_yield:.2f}%</div></div>", unsafe_allow_html=True)
        with c4: st.markdown(f"<div class='kpi-card'><div class='kpi-title'>Active Testers</div><div class='kpi-value'>{active_testers}</div></div>", unsafe_allow_html=True)

        st.divider()

        # ---------------------------------------------------------
        # Part B: Period Performance Summary
        # ---------------------------------------------------------
        st.markdown("#### 1. Period Performance Summary")
        st.markdown("<p style='color: #444; font-size: 14px;'>Note: Normalized UPD is a 24-hour equivalent speed based on actual test time.</p>", unsafe_allow_html=True)
        
        total_op_days = op_summary['Active_Days'].sum()
        global_gross_upd = (total_insertions / total_op_days) if total_op_days > 0 else 0
        global_net_upd = (total_pass_insertions / total_op_days) if total_op_days > 0 else 0
        
        op_theo_val = targets[op]['theo']
        total_calendar_days = op_summary['Adjusted_Span_Days'].sum()
        global_theo_qty = total_calendar_days * op_theo_val
        global_oee = (total_insertions / global_theo_qty) * 100 if global_theo_qty > 0 else 0

        sc1, sc2, sc3 = st.columns(3)
        with sc1: st.markdown(f"<div class='summary-card'><div class='summary-title'>Avg Test Normalized UPD</div><div class='summary-value'>{global_gross_upd:,.0f}</div></div>", unsafe_allow_html=True)
        with sc2: st.markdown(f"<div class='summary-card'><div class='summary-title'>Avg Pass Normalized UPD</div><div class='summary-value'>{global_net_upd:,.0f}</div></div>", unsafe_allow_html=True)
        with sc3: st.markdown(f"<div class='summary-card'><div class='summary-title'>Overall OEE</div><div class='summary-value'>{global_oee:.1f}%</div></div>", unsafe_allow_html=True)

        st.write("") 

        # ---------------------------------------------------------
        # Part C: Capacity Planning Insights
        # ---------------------------------------------------------
        st.markdown("#### 2. Capacity Planning Insights")
        
        op_plan_val = targets[op]['plan']
        buffer_pct = ((global_gross_upd - op_plan_val) / global_gross_upd) * 100 if global_gross_upd > 0 else 0
        implied_oee = (op_plan_val / op_theo_val) * 100 if op_theo_val > 0 else 0

        if global_gross_upd >= op_plan_val:
            insight_text = f"""
            Validation shows the current <span class='insight-highlight'>{op}</span> actual average capacity is approx. <span class='insight-highlight'>{global_gross_upd:,.0f} ea/day</span>, surpassing your planned target of <span class='insight-highlight'>{op_plan_val:,.0f}</span>.<br>
            Using {op_plan_val:,.0f} as your baseline preserves a <span class='insight-highlight'>{buffer_pct:.1f}%</span> capacity buffer.
            """
        else:
            insight_text = f"""
            ⚠️ <b>Notice:</b> The current <span class='insight-highlight'>{op}</span> actual average capacity is approx. <span class='insight-highlight'>{global_gross_upd:,.0f} ea/day</span>, which is <b>below</b> your planned target of <span class='insight-highlight'>{op_plan_val:,.0f}</span>.<br>
            It is recommended to lower the planning baseline or investigate for abnormal downtime.
            """
        st.markdown(f"<div class='insight-box'>{insight_text}</div>", unsafe_allow_html=True)

        insight_data = {
            "Metric": ["Report Avg Normalized UPD", "Planned Target UPD", "Theoretical Max UPD", f"Implied OEE (at {op_plan_val})"],
            "Value": [f"{global_gross_upd:,.0f}", f"{op_plan_val:,.0f}", f"{op_theo_val:,.0f}", f"{implied_oee:.1f}%"]
        }
        st.dataframe(pd.DataFrame(insight_data).T, use_container_width=True, hide_index=True)

        st.write("")

        # ---------------------------------------------------------
        # Part D: Tester Performance Details
        # ---------------------------------------------------------
        st.markdown("#### 3. Tester Performance Details (A/P/Q Breakdown)")
        
        display_df = op_summary[[
            'Tester', 'Lot_Count', 'First_Yield', 'Final_Yield (Q)', 
            'Avg_Gross_UPD', 'Avg_Net_UPD', 'Total_TestQty', 'Total_PassQty', 'Avg_OEE',
            'Active_Days', 'Availability (A)', 'Performance (P)'
        ]].copy()

        display_df = display_df.rename(columns={
            'Tester': 'Tester', 'Lot_Count': 'Lot Count', 
            'First_Yield': 'First Yield', 'Final_Yield (Q)': 'Final Yield (Q)',
            'Avg_Gross_UPD': 'Normalized UPD (Test)', 'Avg_Net_UPD': 'Normalized UPD (Pass)', 
            'Total_TestQty': 'Actual Test Qty', 'Total_PassQty': 'Actual Pass Qty', 
            'Avg_OEE': 'Avg OEE',
            'Active_Days': 'Active Days', 'Availability (A)': 'Availability (A)', 'Performance (P)': 'Performance (P)'
        })

        display_df['First Yield'] = display_df['First Yield'].apply(lambda x: f"{x*100:.2f}%")
        display_df['Final Yield (Q)'] = display_df['Final Yield (Q)'].apply(lambda x: f"{x*100:.2f}%")
        display_df['Normalized UPD (Test)'] = display_df['Normalized UPD (Test)'].apply(lambda x: f"{x:,.0f}")
        display_df['Normalized UPD (Pass)'] = display_df['Normalized UPD (Pass)'].apply(lambda x: f"{x:,.0f}")
        display_df['Actual Test Qty'] = display_df['Actual Test Qty'].apply(lambda x: f"{x:,.0f}")
        display_df['Actual Pass Qty'] = display_df['Actual Pass Qty'].apply(lambda x: f"{x:,.0f}")
        display_df['Avg OEE'] = display_df['Avg OEE'].apply(lambda x: f"{x*100:.1f}%")
        display_df['Active Days'] = display_df['Active Days'].apply(lambda x: f"{x:.2f}")
        display_df['Availability (A)'] = display_df['Availability (A)'].apply(lambda x: f"{x*100:.1f}%")
        display_df['Performance (P)'] = display_df['Performance (P)'].apply(lambda x: f"{x*100:.1f}%")

        def highlight_low_oee(val, threshold):
            try:
                num_val = float(val.strip('%'))
                if num_val < threshold:
                    return 'background-color: #ffebee; color: #dc3545; font-weight: bold;'
            except: pass
            return ''
        
        styled_df = display_df.style\
            .map(lambda x: highlight_low_oee(x, implied_oee), subset=['Avg OEE'])\
            .set_properties(**{'text-align': 'center'})\
            .set_table_styles([
                {'selector': 'th', 'props': [('text-align', 'center')]},
                {'selector': 'td', 'props': [('text-align', 'center')]}
            ])
            
        st.dataframe(styled_df, use_container_width=True, hide_index=True)

        st.write("")

        # ---------------------------------------------------------
        # Part E: Visualizations
        # ---------------------------------------------------------
        st.markdown("#### 4. Visualizations")
        
        plot_df = op_summary.rename(columns={'Avg_Gross_UPD': 'Normalized UPD (Test)', 'Avg_Net_UPD': 'Normalized UPD (Pass)'})
        
        fig1 = px.bar(
            plot_df, 
            x='Tester', 
            y=['Normalized UPD (Test)', 'Normalized UPD (Pass)'], 
            barmode='group',
            title=f"Throughput Comparison & Target Validation ({op})",
            labels={'value': 'Equivalent Rate (UPD)', 'variable': 'Metrics', 'Tester': ''},
            color_discrete_map={'Normalized UPD (Test)': '#1E3A8A', 'Normalized UPD (Pass)': '#28a745'}
        )
        fig1.add_hline(y=op_plan_val, line_dash="dash", line_color="orange", annotation_text="Planned Target", annotation_position="top right")
        fig1.add_hline(y=op_theo_val, line_dash="dot", line_color="red", annotation_text="Theoretical Max", annotation_position="top right")
        
        fig1.update_layout(
            legend_title_text='', 
            margin=dict(t=50, b=20, l=10, r=10), 
            legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
            height=400 
        )
        st.plotly_chart(fig1, use_container_width=True)
        
        st.write("") 
        
        prod_summary = op_df.groupby(['Tester', 'ProductNo'])['TestQty'].sum().reset_index()
        fig2 = px.bar(
            prod_summary, 
            x='Tester', 
            y='TestQty', 
            color='ProductNo',
            title=f"Product Mix Volume Distribution ({op})",
            labels={'TestQty': 'Test Qty', 'Tester': '', 'ProductNo': 'Product'},
            color_discrete_sequence=px.colors.qualitative.Pastel
        )
        
        fig2.update_xaxes(categoryorder='category ascending')
        
        fig2.update_layout(
            margin=dict(t=50, b=20, l=10, r=10), 
            legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
            height=400 
        )
        st.plotly_chart(fig2, use_container_width=True)
        
        st.write("") 

        # ---------------------------------------------------------
        # Part F: Raw Data
        # ---------------------------------------------------------
        st.markdown("#### 5. Raw Data")
        with st.expander(f"Click to view raw data for {op}"):
            st.dataframe(op_df, use_container_width=True, hide_index=True)


# ==============================================================================
# --- 🌟 V52: Overall Build Yield Tab (Interactive Matrix) ---
# ==============================================================================
with tabs[-1]:
    st.write("")
    
    # 標準的 8 個 Build Phase 階段
    build_phases = ["Proto1.0", "Proto1.1", "EVT1.0(A0)", "EVT1.0(B0)", "EVT1.1", "DVT", "PVT", "MP"]
    all_programs = sorted(filtered_df['ProgramName'].dropna().unique().tolist())
    
    # 初始化 Mapping Session State
    if "build_mapping" not in st.session_state:
        st.session_state.build_mapping = {phase: [] for phase in build_phases}
        
        # --- 自動分類腳本 (Smart Auto-Assign) ---
        # 如果程式名稱包含關鍵字，系統自動幫你歸類，省去大量手動時間
        for p in all_programs:
            p_upper = p.upper()
            if "PROTO1.0" in p_upper or "P10" in p_upper: st.session_state.build_mapping["Proto1.0"].append(p)
            elif "PROTO1.1" in p_upper or "P11" in p_upper: st.session_state.build_mapping["Proto1.1"].append(p)
            elif "EVT1.0(A0)" in p_upper or "EVT1.0_A" in p_upper: st.session_state.build_mapping["EVT1.0(A0)"].append(p)
            elif "EVT1.0(B0)" in p_upper or "EVT1.0_B" in p_upper: st.session_state.build_mapping["EVT1.0(B0)"].append(p)
            elif "EVT1.1" in p_upper: st.session_state.build_mapping["EVT1.1"].append(p)
            elif "DVT" in p_upper: st.session_state.build_mapping["DVT"].append(p)
            elif "PVT" in p_upper: st.session_state.build_mapping["PVT"].append(p)
            elif "MP" in p_upper: st.session_state.build_mapping["MP"].append(p)

    # 找出已經被分配的程式
    assigned_progs = []
    for progs in st.session_state.build_mapping.values():
        assigned_progs.extend(progs)
        
    # 尚未分配的程式 (Pool)
    unassigned_pool = [p for p in all_programs if p not in assigned_progs]
    
    # --- UI: 互動式分類器 (Smart Binning UI) ---
    with st.expander("⚙️ Build Phase Mapping Configuration (Interactive Binning)", expanded=True):
        st.markdown("<p style='font-size: 14px; color: #555;'>Assign programs to their corresponding Build Phases. Selected programs will automatically disappear from other menus.</p>", unsafe_allow_html=True)
        
        st.markdown(f"**📦 Unassigned Program Pool ({len(unassigned_pool)})**")
        st.caption(", ".join(unassigned_pool) if unassigned_pool else "🎉 All active programs have been successfully assigned to a Build Phase!")
        st.write("")
        
        # 建立 2 排 4 欄的排版
        cols1 = st.columns(4)
        cols2 = st.columns(4)
        all_cols = cols1 + cols2
        
        for i, phase in enumerate(build_phases):
            with all_cols[i]:
                # 目前已經在這個 Phase 的程式 (需過濾掉可能已不存在於這批資料的歷史垃圾紀錄)
                curr_selected = [p for p in st.session_state.build_mapping[phase] if p in all_programs]
                
                # 選單的選項 = 自己原本的 + 池子裡沒人要的
                options = sorted(list(set(curr_selected + unassigned_pool)))
                
                st.session_state.build_mapping[phase] = st.multiselect(
                    f"📍 {phase}",
                    options=options,
                    default=curr_selected,
                    key=f"map_{phase}"
                )

    # --- Data Engine: 矩陣報表計算 ---
    # 建立反向字典: ProgramName -> Build Phase
    prog_to_build = {}
    for phase, progs in st.session_state.build_mapping.items():
        for p in progs:
            prog_to_build[p] = phase
            
    build_df = filtered_df.copy()
    build_df['Build_Phase'] = build_df['ProgramName'].map(prog_to_build)
    build_df = build_df.dropna(subset=['Build_Phase']) # 濾掉沒被分類的
    
    if build_df.empty:
        st.info("💡 Please assign at least one program to a Build Phase above to generate the Matrix Report.")
    else:
        st.markdown("#### 📊 Build Evolution Matrix (Testing Yield %)")
        st.caption("Note: 'MBU Test CUM yield' is calculated by multiplying the Testing Yield of all listed operations within that phase.")
        
        # 第一步：計算各 Phase & OpNo 的良率
        matrix_sum = build_df.groupby(['OpNo', 'Build_Phase']).agg(
            T_Qty=('TestQty', 'sum'),
            P_Qty=('PassQty', 'sum')
        ).reset_index()
        
        # 避免除以零
        matrix_sum['Yield'] = np.where(matrix_sum['T_Qty'] > 0, (matrix_sum['P_Qty'] / matrix_sum['T_Qty']) * 100, np.nan)
        
        # 第二步：Pivot 樞紐分析 (橫向展開)
        pivot_yield = matrix_sum.pivot(index='OpNo', columns='Build_Phase', values='Yield')
        
        # 強制照我們定義好的時間軸排序 Columns
        exist_phases = [p for p in build_phases if p in pivot_yield.columns]
        pivot_yield = pivot_yield[exist_phases]
        
        # 第三步：計算 CUM Yield (跨站點良率相乘)
        cum_yields = {}
        for col in pivot_yield.columns:
            col_yields = pivot_yield[col].dropna() / 100.0
            if not col_yields.empty:
                cum_yields[col] = col_yields.prod() * 100.0
            else:
                cum_yields[col] = np.nan
                
        cum_row = pd.DataFrame([cum_yields], index=['MBU Test CUM yield'])
        
        # 第四步：合併 CUM row 與 各站點明細，產生最終矩陣
        final_matrix = pd.concat([cum_row, pivot_yield])
        
        # 第五步：視覺化排版 (Pandas Styler)
        def color_yield_matrix(val):
            if pd.isna(val): return ''
            # 大於 95% 淺綠，90~95% 淺黃，低於 90% 淺紅
            color = '#e6f4ea' if val >= 95 else '#fff3cd' if val >= 90 else '#fdecea'
            return f'background-color: {color}; font-weight: bold; color: #333;'

        styled_matrix = final_matrix.style\
            .map(color_yield_matrix)\
            .format("{:.2f}%", na_rep="-")\
            .set_properties(**{'text-align': 'center', 'font-size': '15px'})\
            .set_table_styles([
                {'selector': 'th', 'props': [('text-align', 'center'), ('background-color', '#f0f2f6'), ('color', '#31333F'), ('font-size', '14px')]},
                {'selector': 'th.row_heading', 'props': [('text-align', 'left'), ('min-width', '200px')]} # 左側列名靠左對齊
            ])

        st.dataframe(styled_matrix, use_container_width=True)