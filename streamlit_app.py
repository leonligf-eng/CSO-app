import streamlit as st
import math

# --- Page Configuration ---
st.set_page_config(page_title="ATE Capacity Planner", layout="wide")

# --- Custom CSS for Professional UI ---
st.markdown("""
    <style>
    .metric-card {
        background-color: #ffffff;
        border: 1px solid #e0e0e0;
        padding: 15px;
        border-radius: 10px;
        text-align: center;
        box-shadow: 2px 2px 5px rgba(0,0,0,0.05);
    }
    .tester-icon {
        height: 25px;
        width: 25px;
        background-color: #4A90E2;
        border-radius: 50%;
        display: inline-block;
        margin: 4px;
    }
    .section-header {
        background-color: #1E3A8A;
        color: white;
        padding: 10px;
        border-radius: 5px;
        font-weight: bold;
        margin-top: 20px;
        margin-bottom: 10px;
    }
    </style>
""", unsafe_allow_html=True)

st.title("📟 ATE Smart Capacity Allocation Console")

# --- Help & Formulas Section (Excel Documentation) ---
with st.expander("ℹ️ Help: Formula & Parameter Definitions (Excel Reference)"):
    st.markdown("""
    ### 📖 Calculation Logic
    To ensure transparency, the system follows the standard Excel-based production models:
    
    * **TD (Touch Downs):** `Lot Size / Site`
    * **Cycle Time (R0):** `TD * (Test Time / 60)` _(Standard testing time in minutes)_
    * **RT Cycle Time (R1/R2):** `((Lot Size * (1 - FPY%)) / Site * (TT / 60)) + OP Time`
        * *Calculates additional time spent on retests and operator intervention.*
    * **Sum Cycle Time:** `R0 + RT Cycle Time`
    * **Effective Mins/Day:** `1440 * OEE%`
    * **Lots/Day per Tester:** `Effective Mins/Day / Sum Cycle Time`
    * **Required Testers:** `CEILING(Target Lots per Day / Lots per Day per Tester)`
    """)

# --- Sidebar: Global Resources ---
st.sidebar.header("🏢 Global Resources")
total_fleet = st.sidebar.number_input("Total Available ATEs", min_value=1, value=50)
st.sidebar.divider()
st.sidebar.caption("Adjust the 'Target Lots/Day' for each station to see the allocation change.")

# --- Core Calculation Function ---
def calculate_metrics(target_lots, lot_size, site, tt, fpy, oee, op_time):
    # Formulas based on provided Excel
    td = lot_size / site if site > 0 else 0
    r0_mins = td * (tt / 60)
    rt_mins = ((lot_size * (1 - fpy / 100)) / site * (tt / 60)) + op_time if site > 0 else 0
    sum_mins = r0_mins + rt_mins
    
    eff_mins_per_day = 1440 * (oee / 100)
    lot_per_day_tester = eff_mins_per_day / sum_mins if sum_mins > 0 else 0
    
    req_testers = math.ceil(target_lots / lot_per_day_tester) if lot_per_day_tester > 0 else 0
    uph_effective = (lot_size / r0_mins) * 60 * (oee / 100) if r0_mins > 0 else 0
    
    return {
        "req": req_testers,
        "uph": int(uph_effective),
        "cyc": round(sum_mins, 1),
        "cap": lot_per_day_tester
    }

# --- Dashboard Layout ---
st.markdown("<div class='section-header'>Fleet Status Overview</div>", unsafe_allow_html=True)
header_cols = st.columns(4)

# Data storage for summary
stage_results = {}
total_needed = 0

# --- Stations Configuration (FT1, FT2, FT3) ---
st.divider()
input_cols = st.columns(3)
stages = [
    {"name": "FT1", "tt": 144.0, "fpy": 95.0, "target": 12},
    {"name": "FT2", "tt": 24.0, "fpy": 99.9, "target": 25},
    {"name": "FT3", "tt": 45.0, "fpy": 98.0, "target": 0}
]

for i, stage in enumerate(stages):
    with input_cols[i]:
        st.subheader(f"📍 {stage['name']} Station")
        target = st.number_input(f"Target (Lots/day)", value=stage['target'], key=f"t_{i}", help="Target production volume")
        
        with st.expander("Detailed Parameters", expanded=True):
            ls = st.number_input("Lot Size", value=3000, key=f"ls_{i}")
            si = st.selectbox("Site (Parallel)", [4, 8, 16, 32], index=1, key=f"si_{i}")
            tt_val = st.number_input("Test Time (s)", value=stage['tt'], key=f"tt_{i}")
            fpy_val = st.number_input("First Pass Yield %", value=stage['fpy'], key=f"fy_{i}")
            oee_val = st.number_input("OEE %", value=65.0, key=f"oe_{i}")
            op_v = st.number_input("RT OP Time (min)", value=60, key=f"op_{i}", help="Retest setup/operator time")
        
        # Calculation Execution
        res = calculate_metrics(target, ls, si, tt_val, fpy_val, oee_val, op_v)
        stage_results[stage['name']] = res
        total_needed += res["req"]

        # Visual Card for Required Testers
        st.markdown(f"""
            <div class='metric-card'>
                <div style='color: #666; font-size: 12px; font-weight:bold;'>REQUIRED TESTERS</div>
                <div style='font-size: 32px; font-weight: bold; color: #1E3A8A;'>{res['req']} <span style='font-size: 16px;'>Units</span></div>
            </div>
        """, unsafe_allow_html=True)
        
        # Dynamic Tester Icons
        icons = "".join(["<div class='tester-icon'></div>" for _ in range(min(res['req'], 15))])
        if res['req'] > 15: icons += f" <span style='font-size:12px;'>+{res['req']-15} more</span>"
        st.markdown(f"<div style='margin-top:10px; min-height:40px;'>{icons}</div>", unsafe_allow_html=True)
        st.caption(f"Cap: {res['cap']:.2f} Lots/Day/Unit | Total Sum Cycle: {res['cyc']} min")

# --- Global Summary Update ---
with header_cols[0]:
    st.metric("Total Req. Testers", f"{total_needed} Units")
with header_cols[1]:
    remaining = total_fleet - total_needed
    status = "Healthy" if remaining >= 0 else "Shortage"
    st.metric("Resource Status", status, f"{remaining} Units")
with header_cols[2]:
    # Sum of individual station UPHs
    total_uph = sum(res['uph'] for res in stage_results.values())
    st.metric("Total Est. UPH", f"{total_uph:,}")
with header_cols[3]:
    load = (total_needed / total_fleet) * 100 if total_fleet > 0 else 0
    st.metric("Fleet Capacity Load", f"{load:.1f}%")

# --- Capacity Loading Analysis (Visual Progress) ---
st.divider()
st.markdown("### 🛰️ Capacity Loading Analysis")
if total_needed > total_fleet:
    st.error(f"🚨 CRITICAL: Total demand ({total_needed}) exceeds fleet capacity ({total_fleet})!")
else:
    st.progress(int(min(load, 100)), text=f"Utilization: {load:.1f}%")

# Quick Stats Table
df_summary = {
    "Station": list(stage_results.keys()),
    "Required Units": [res['req'] for res in stage_results.values()],
    "Sum Cycle Time (min)": [res['cyc'] for res in stage_results.values()],
    "Est. UPH": [res['uph'] for res in stage_results.values()]
}
st.table(df_summary)