import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
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
    testers = [f"HP93K-EXA{i:02d}" for i in range(2, 10)]
    ops = ["FT1", "FTA", "MT1"]
    programs_map = {"FT1": ["PROD_GS631_FT1_REV1"], "FTA": ["PROD_GS631_FTA_REV1"], "MT1": ["PROD_GS631_MT1_REV1"]}
    tester_clocks = {tester: now - timedelta(days=20) for tester in testers}
    
    for i in range(150):
        tester = np.random.choice(testers)
        op = np.random.choice(ops)
        prog = np.random.choice(programs_map[op])
        
        wait_time = timedelta(hours=np.random.uniform(0.5, 3))
        start_time = tester_clocks[tester] + wait_time
        test_duration = timedelta(hours=np.random.uniform(6, 14))
        end_time = start_time + test_duration
        
        if op == "FT1": qty = np.random.randint(3000, 5000)
        elif op == "FTA": qty = np.random.randint(5000, 8000)
        else: qty = np.random.randint(15000, 25000)
        
        yield_rate = np.random.uniform(0.95, 0.99)
        pass_qty = int(qty * yield_rate)
        
        data.append([f"LOT_{i:04d}", op, prog, tester, start_time, end_time, qty, pass_qty])
        tester_clocks[tester] = end_time
        
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

uploaded_file = st.sidebar.file_uploader("Upload Yield Report (Excel)", type=["xlsx", "xls"])
raw_df = load_data(uploaded_file)

if raw_df is None or raw_df.empty:
    st.warning("No data available.")
    st.stop()

# ==============================================================================
# --- 1. Main Area: Filters & (UNABRIDGED) Help Section ---
# ==============================================================================
with st.expander("ℹ️ Help: Formula & Parameter Definitions (公式與參數定義)"):
    st.markdown("""
    This system employs rigorous Industrial Engineering (IE) logic combined with actual production report data to calculate authentic equipment efficiency. Metric definitions are as follows:

    #### 1. Capacity Metrics (Tester View vs. Product View)
    * **Total Insertions (Gross):** Total testing actions performed by the testers across all operations. It evaluates the physical workload on the testers (allows double-counting across ops).
    * **Est. Physical Units:** Estimated actual ICs processed, calculated by de-duplicating `LotNo` and taking the max input quantity across operations.
    * **Active Days:** The exact number of hours a tester spent in the "Testing" state (CheckIn to CheckOut), divided by 24 hours.
    * **Gross UPD (Prorated 24h Rate):** Calculated as `Total TestQty / Active Days`. 
      * *Note on Proration:* This is a **SPEED metric**, not a volume metric. If a tester tests 6,000 units in exactly 12 hours, its 24-hour prorated UPD is 12,000 ea/day.
    * **Global Average UPD (Weighted):** To prevent short-duration outliers from skewing the factory average, the overall operation average is calculated as `Sum of All TestQty / Sum of All Active Days`.

    #### 2. OEE Breakdown (Availability, Performance, Quality)
    * **A (Availability):** Measures how often the tester is actually in production.
      * `Calculation = Active Days / Calendar Span Days (from first CheckIn to last CheckOut)`
    * **P (Performance):** Measures if the tester is running at theoretical speed when active.
      * `Theoretical UPH = Theoretical Max UPD / 24`
      * `Actual UPH = Total TestQty / Total Active Hours`
      * `Calculation = Actual UPH / Theoretical UPH`
    * **Q (Quality):** Testing quality.
      * `Calculation = Total PassQty / Total TestQty`
    * **Overall OEE:**
      * `Calculation = Avg Actual Gross UPD / Theoretical Max UPD`. (Mathematically equivalent to A × P)

    #### 3. Capacity Planning Metrics
    * **Planned Target UPD:** The safety scheduling baseline preset by Planning or Engineering, defined specifically per operation.
    * **Implied OEE:** `Calculation = Planned Target UPD / Theoretical Max UPD`. Reflects the built-in buffer percentage reserved for setups, maintenance, and re-tests for a specific operation.
    """)

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
# --- 2. Sidebar Settings (Dynamic Targets per OpNo) ---
# ==============================================================================
st.sidebar.header("⚙️ Data Cleaning Rules")
min_lot_size = st.sidebar.number_input("Exclude Lots smaller than (Qty)", value=0, step=100)

st.sidebar.divider()
st.sidebar.header("📐 Planning & Targets")
st.sidebar.caption("Define specific baselines for EACH selected operation.")

targets = {}
for op in selected_ops:
    st.sidebar.markdown(f"**🔹 Operation: {op}**")
    theo = st.sidebar.number_input(f"Theo Max UPD (100% OEE)({op})", value=4240, step=10, key=f"theo_{op}")
    plan = st.sidebar.number_input(f"Planned Target UPD ({op})", value=2800, step=100, key=f"plan_{op}")
    targets[op] = {'theo': theo, 'plan': plan}
    st.sidebar.write("") 

# ==============================================================================
# --- 3. Data Processing Engine ---
# ==============================================================================
filtered_df = raw_df[
    (raw_df['OpNo'].isin(selected_ops)) & 
    (raw_df['ProgramName'].isin(selected_progs)) &
    (raw_df['TestQty'] >= min_lot_size)
].copy()

if filtered_df.empty:
    st.warning("No production data available for the selected filters.")
    st.stop()

filtered_df['Duration_Hr'] = (filtered_df['CheckOutTime'] - filtered_df['CheckInTime']).dt.total_seconds() / 3600.0

# Group by Tester AND OpNo
tester_summary = filtered_df.groupby(['Tester', 'OpNo']).agg(
    Lot_Count=('LotNo', 'nunique'),
    Total_TestQty=('TestQty', 'sum'),
    Total_PassQty=('PassQty', 'sum'),
    Total_Duration_Hr=('Duration_Hr', 'sum'),
    Min_CheckIn=('CheckInTime', 'min'),
    Max_CheckOut=('CheckOutTime', 'max')
).reset_index()

tester_summary['Active_Days'] = tester_summary['Total_Duration_Hr'] / 24.0
tester_summary['Avg_Gross_UPD'] = np.where(tester_summary['Active_Days'] > 0, tester_summary['Total_TestQty'] / tester_summary['Active_Days'], 0)
tester_summary['Avg_Net_UPD'] = np.where(tester_summary['Active_Days'] > 0, tester_summary['Total_PassQty'] / tester_summary['Active_Days'], 0)

tester_summary['Calendar_Span_Days'] = (tester_summary['Max_CheckOut'] - tester_summary['Min_CheckIn']).dt.total_seconds() / (24.0 * 3600.0)
tester_summary['Availability (A)'] = np.where(tester_summary['Calendar_Span_Days'] > 0, tester_summary['Active_Days'] / tester_summary['Calendar_Span_Days'], 0)

tester_summary['Theo_Max_UPD'] = tester_summary['OpNo'].map(lambda x: targets[x]['theo'])
tester_summary['Planned_UPD'] = tester_summary['OpNo'].map(lambda x: targets[x]['plan'])

tester_summary['Actual_UPH'] = np.where(tester_summary['Total_Duration_Hr'] > 0, tester_summary['Total_TestQty'] / tester_summary['Total_Duration_Hr'], 0)
tester_summary['Performance (P)'] = tester_summary['Actual_UPH'] / (tester_summary['Theo_Max_UPD'] / 24.0)

tester_summary['Yield (Q)'] = np.where(tester_summary['Total_TestQty'] > 0, tester_summary['Total_PassQty'] / tester_summary['Total_TestQty'], 0)
tester_summary['Avg_OEE'] = tester_summary['Avg_Gross_UPD'] / tester_summary['Theo_Max_UPD']

tester_summary = tester_summary.sort_values(by=['OpNo', 'Tester'], ascending=True)

# ==============================================================================
# --- 4. Dashboard UI (Overall + Tabs by Operation) ---
# ==============================================================================
st.markdown("### 📊 Overall Performance")

# Overall Factory Level Metrics
total_insertions = int(tester_summary['Total_TestQty'].sum())
total_pass_insertions = int(tester_summary['Total_PassQty'].sum())
avg_step_yield = (total_pass_insertions / total_insertions) * 100 if total_insertions > 0 else 0

unique_lots_df = filtered_df.groupby('LotNo').agg(Max_TestQty=('TestQty', 'max')).reset_index()
physical_lots = len(unique_lots_df)
physical_units = int(unique_lots_df['Max_TestQty'].sum())

active_testers = tester_summary['Tester'].nunique()

c1, c2, c3, c4 = st.columns(4)
with c1: 
    st.markdown(f"""<div class='kpi-card'><div class='kpi-title'>Total Insertions (Gross)</div><div class='kpi-value'>{total_insertions:,}</div>
        <div style='font-size: 12px; color: #888; margin-top: 5px;'>Est. Physical Units: ~{physical_units:,}</div></div>""", unsafe_allow_html=True)
with c2: 
    st.markdown(f"""<div class='kpi-card'><div class='kpi-title'>Pass Insertions (Net)</div><div class='kpi-value'>{total_pass_insertions:,}</div>
        <div style='font-size: 12px; color: #888; margin-top: 5px;'>Total Lots: {physical_lots}</div></div>""", unsafe_allow_html=True)
with c3: 
    st.markdown(f"""<div class='kpi-card'><div class='kpi-title'>Avg Step Yield</div><div class='kpi-value'>{avg_step_yield:.2f}%</div>
        <div style='font-size: 12px; color: #888; margin-top: 5px;'>Not Rolled Throughput Yield</div></div>""", unsafe_allow_html=True)
with c4: 
    st.markdown(f"""<div class='kpi-card'><div class='kpi-title'>Active Testers</div><div class='kpi-value'>{active_testers}</div>
        <div style='font-size: 12px; color: #888; margin-top: 5px;'>Operations: {len(selected_ops)}</div></div>""", unsafe_allow_html=True)

st.divider()
st.markdown("## ATE Tester Capacity Analysis & Validation")
st.caption("Each operation has distinct characteristics. Please review the performance by navigating the tabs below.")

# Operation Tabs
tabs = st.tabs(selected_ops)

for idx, op in enumerate(selected_ops):
    with tabs[idx]:
        st.markdown(f"### 📍 Operation: {op}")
        
        op_df = filtered_df[filtered_df['OpNo'] == op]
        op_summary = tester_summary[tester_summary['OpNo'] == op]
        
        if op_summary.empty:
            st.info(f"No data available for Operation {op} under the current filters.")
            continue
        
        op_insertions = int(op_summary['Total_TestQty'].sum())
        op_pass_insertions = int(op_summary['Total_PassQty'].sum())
        
        # --- 1. Period Performance Summary ---
        st.markdown("#### 1. Period Performance Summary")
        st.caption(f"Note: 'Avg Rate' is a 24-hour prorated speed. This represents the weighted average for {op}.")
        
        total_op_days = op_summary['Active_Days'].sum()
        # Weighted Average Calculation
        global_gross_upd = (op_insertions / total_op_days) if total_op_days > 0 else 0
        global_net_upd = (op_pass_insertions / total_op_days) if total_op_days > 0 else 0
        
        op_theo_val = targets[op]['theo']
        global_theo_qty = total_op_days * op_theo_val
        global_oee = (op_insertions / global_theo_qty) * 100 if global_theo_qty > 0 else 0

        sc1, sc2, sc3 = st.columns(3)
        with sc1: st.markdown(f"<div class='summary-card'><div class='summary-title'>Avg Actual Rate (Gross UPD)</div><div class='summary-value'>{global_gross_upd:,.0f}</div></div>", unsafe_allow_html=True)
        with sc2: st.markdown(f"<div class='summary-card'><div class='summary-title'>Avg Effective Rate (Net UPD)</div><div class='summary-value'>{global_net_upd:,.0f}</div></div>", unsafe_allow_html=True)
        with sc3: st.markdown(f"<div class='summary-card'><div class='summary-title'>Overall OEE</div><div class='summary-value'>{global_oee:.1f}%</div></div>", unsafe_allow_html=True)

        st.write("") 

        # --- 2. Capacity Planning Insights ---
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
            "Metric": ["Report Avg Rate (Gross)", "Planned Target UPD", "Theoretical Max UPD", f"Implied OEE (at {op_plan_val})"],
            "Value": [f"{global_gross_upd:,.0f}", f"{op_plan_val:,.0f}", f"{op_theo_val:,.0f}", f"{implied_oee:.1f}%"]
        }
        st.dataframe(pd.DataFrame(insight_data).T, use_container_width=True, hide_index=True)

        st.write("")

        # --- 3. Tester Performance Details ---
        st.markdown("#### 3. Tester Performance Details (A/P/Q Breakdown)")
        
        display_df = op_summary[[
            'Tester', 'Lot_Count', 'Yield (Q)', 
            'Avg_Gross_UPD', 'Avg_Net_UPD', 'Avg_OEE',
            'Active_Days', 'Availability (A)', 'Performance (P)'
        ]].copy()

        display_df = display_df.rename(columns={
            'Tester': 'Tester', 'Lot_Count': 'Lot Count', 'Yield (Q)': 'Yield (Q)',
            'Avg_Gross_UPD': 'Avg Rate (TestQty)', 'Avg_Net_UPD': 'Avg Rate (PassQty)', 'Avg_OEE': 'Avg OEE',
            'Active_Days': 'Active Days', 'Availability (A)': 'Availability (A)', 'Performance (P)': 'Performance (P)'
        })

        display_df['Yield (Q)'] = display_df['Yield (Q)'].apply(lambda x: f"{x*100:.2f}%")
        display_df['Avg Rate (TestQty)'] = display_df['Avg Rate (TestQty)'].apply(lambda x: f"{x:,.0f}")
        display_df['Avg Rate (PassQty)'] = display_df['Avg Rate (PassQty)'].apply(lambda x: f"{x:,.0f}")
        display_df['Avg OEE'] = display_df['Avg OEE'].apply(lambda x: f"{x*100:.1f}%")
        display_df['Active Days'] = display_df['Active Days'].apply(lambda x: f"{x:.2f}")
        display_df['Availability (A)'] = display_df['Availability (A)'].apply(lambda x: f"{x*100:.1f}%")
        display_df['Performance (P)'] = display_df['Performance (P)'].apply(lambda x: f"{x*100:.1f}%")

        st.dataframe(display_df, use_container_width=True, hide_index=True)

        st.write("")

        # --- 4. Chart ---
        st.markdown("#### 4. Tester Throughput Comparison")
        
        fig = px.bar(
            op_summary, 
            x='Tester', 
            y=['Avg_Gross_UPD', 'Avg_Net_UPD'],
            barmode='group',
            labels={'value': 'Prorated Rate (UPD)', 'variable': 'Metrics', 'Tester': ''},
            color_discrete_map={'Avg_Gross_UPD': '#1E3A8A', 'Avg_Net_UPD': '#28a745'}
        )
        
        fig.add_hline(y=op_plan_val, line_dash="dash", line_color="orange", 
                      annotation_text="Planned Target", annotation_position="top right")
        fig.add_hline(y=op_theo_val, line_dash="dot", line_color="red", 
                      annotation_text="Theoretical Max", annotation_position="top right")
        
        fig.update_layout(
            legend_title_text='',
            margin=dict(t=20, b=0, l=0, r=0),
            legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1)
        )
        
        st.plotly_chart(fig, use_container_width=True)
