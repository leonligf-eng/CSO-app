import streamlit as st
import math
import pandas as pd
from datetime import date, timedelta
import plotly.graph_objects as go

# --- Page Configuration ---
st.set_page_config(page_title="ATE Capacity Planner", layout="wide", initial_sidebar_state="expanded")

# --- Custom CSS ---
st.markdown("""
    <style>
    /* 調整 Sidebar 寬度，讓地圖有更多空間展開 */
    [data-testid="stSidebar"] {
        min-width: 400px !important;
        max-width: 450px !important;
    }
    
    .metric-card {
        background-color: #ffffff;
        border: 1px solid #e0e0e0;
        padding: 15px;
        border-radius: 10px;
        text-align: center;
        box-shadow: 2px 2px 5px rgba(0,0,0,0.05);
        height: 130px; 
        display: flex;
        flex-direction: column;
        justify-content: center;
    }
    
    .stMultiSelect [data-baseweb="tag"], 
    .stMultiSelect [data-selection="true"] {
        display: none !important;
    }
    .stMultiSelect:has([data-baseweb="tag"]) div[data-baseweb="select"]::before {
        content: "Testers selected as below";
        color: #888;
        font-size: 14px;
        position: absolute;
        left: 10px;
        top: 50%;
        transform: translateY(-50%);
        pointer-events: none;
        z-index: 1;
    }
    .stMultiSelect div[data-baseweb="select"] {
        min-height: 42px !important;
    }
    
    button[kind="primary"] {
        background-color: #6b8cd4 !important; 
        border: 1px solid #4a65a3 !important;
        color: white !important;
        border-radius: 4px !important; 
        padding: 2px 10px !important;
        min-height: 32px !important;
        font-size: 12px !important;
        font-weight: bold !important;
        width: 100% !important;
    }

    [data-testid="stSidebar"] button[kind="primary"] {
        background-color: #6c757d !important; 
        border: 1px solid #5a6268 !important;
    }
    
    .section-header {
        background-color: #1E3A8A;
        color: white;
        padding: 10px;
        border-radius: 5px;
        font-weight: bold;
        margin-top: 20px;
        margin-bottom: 15px; 
    }
    .extra-upd {
        color: #28a745;
        font-weight: bold;
        font-size: 14px;
        margin-top: 5px;
    }
    </style>
""", unsafe_allow_html=True)

st.title("📟 ATE Smart Capacity and Allocation")

# --- 用於移除機台的 Callback 函數 ---
def remove_ate(select_key, ate_id):
    if select_key in st.session_state:
        current_selection = list(st.session_state[select_key])
        if ate_id in current_selection:
            current_selection.remove(ate_id)
            st.session_state[select_key] = current_selection

# --- Help & Formulas Section ---
with st.expander("ℹ️ Help: Formula & Parameter Definitions"):
    st.markdown("""
    ### 📖 Calculation Logic
    To ensure transparency, the system follows the standard Excel-based production models:
    * **TD (Touch Downs):** `Lot Size / Site`
    * **Cycle Time (R0):** `TD * (Test Time / 60)` _(Standard testing time in minutes)_
    * **RT Cycle Time (R1/R2):** `((Lot Size * (1 - FPY%)) / Site * (TT / 60)) + OP Time`
    * **Sum Cycle Time:** `R0 + RT Cycle Time`
    * **Effective Mins/Day:** `1440 * OEE%`
    * **Lots/Day per Tester:** `Effective Mins/Day / Sum Cycle Time`
    * **UPD (Units Per Day):** `Lots/Day per Tester * Lot Size`
    * **Required Testers:** `CEILING(Daily Target / Single Tester UPD)`
    * **Real WIP Flow:** `Single Tester UPD * (Actually Assigned Testers OR Required Testers)`
    * **Flow Logic:** FT2/FT3 Start Date is offset by the first lot completion time of the previous stage.
    """)

# ==============================================================================
# --- (核心邏輯) 三站+Sidebar全連動池管理與防呆初始化 ---
# ==============================================================================
occup_key = "occupied_ates_side"
ft1_key = "select_ft1"
ft2_key = "select_ft2"
ft3_key = "select_ft3"

if occup_key not in st.session_state: st.session_state[occup_key] = []
if ft1_key not in st.session_state: st.session_state[ft1_key] = []
if ft2_key not in st.session_state: st.session_state[ft2_key] = []
if ft3_key not in st.session_state: st.session_state[ft3_key] = []

# --- Sidebar: Project Info ---
st.sidebar.header("📁 Project Info")
project_name = st.sidebar.text_input("Project Name / Code", placeholder="e.g., ZC13", help="Enter the project name for this capacity plan.")
actual_project_name = project_name if project_name.strip() else "Unnamed_Project"
st.sidebar.divider()

# --- Sidebar: Global Resources & Occupied ATEs ---
st.sidebar.header("🏢 Global Resources")

col_res1, col_res2 = st.sidebar.columns(2)
total_fleet_num = col_res1.number_input("Total Slots", min_value=1, value=38, help="Total physical slots in the plant (Phase 3)")
running_fleet_num = col_res2.number_input("Running ATEs", min_value=1, value=24, help="Number of currently active testers available for assignment")

total_ate_pool = [f"ATE{i:02d}" for i in range(1, total_fleet_num + 1)] 
active_ate_pool = [f"ATE{i:02d}" for i in range(1, running_fleet_num + 1)] 

st.sidebar.caption(f"Active Plant IDs: `{active_ate_pool[0]} ~ {active_ate_pool[-1]}`")

for key in [occup_key, ft1_key, ft2_key, ft3_key]:
    st.session_state[key] = [ate for ate in st.session_state[key] if ate in active_ate_pool]

st.sidebar.divider()

st.sidebar.markdown("### 🔒 Exclude Occupied ATEs")

project_selections = []
project_selections.extend(st.session_state[ft1_key])
project_selections.extend(st.session_state[ft2_key])
project_selections.extend(st.session_state[ft3_key])

side_options = [ate for ate in active_ate_pool if ate not in project_selections]
side_options = sorted(list(set(side_options + st.session_state[occup_key])))

occupied_ates = st.sidebar.multiselect(
    "Allocated to Other Projects",
    options=side_options,
    key=occup_key,
    help="Select ATEs that are currently in use by other projects."
)
occupied_ates = sorted(occupied_ates) 

if occupied_ates:
    st.sidebar.caption("Excluded (Unavailable):")
    tag_cols_side = st.sidebar.columns(3) 
    for idx, id in enumerate(occupied_ates):
        with tag_cols_side[idx % 3]:
            st.button(id, key=f"btn_side_{id}", on_click=remove_ate, args=(occup_key, id), type="primary")
    st.sidebar.markdown(f"<div style='color: #6c757d; font-size: 12px; margin-top: 3px;'>Count: {len(occupied_ates)} ATE excluded.</div>", unsafe_allow_html=True)
else:
    st.sidebar.caption("All active plant ATEs available for project.")

net_fleet_num = running_fleet_num - len(occupied_ates)
st.sidebar.write(f"✅ Available ATEs for the Project: **{net_fleet_num}** ATEs")

# --- Core Calculation Function ---
def calculate_metrics(daily_target_units, lot_size, site, tt, fpy, oee):
    td = lot_size / site if site > 0 else 0
    r0_mins = td * (tt / 60)
    rt_mins = ((lot_size * (1 - fpy / 100)) / site * (tt / 60)) if site > 0 else 0
    sum_mins = r0_mins + rt_mins
    eff_mins_per_day = 1440 * (oee / 100)
    units_per_day_tester = (eff_mins_per_day / sum_mins) * lot_size if sum_mins > 0 else 0
    req_testers = math.ceil(daily_target_units / units_per_day_tester) if units_per_day_tester > 0 else 0
    uph_effective = (lot_size / r0_mins) * 60 * (oee / 100) if r0_mins > 0 else 0
    return {"req": req_testers, "uph": int(uph_effective), "upd": int(units_per_day_tester), "cyc": round(sum_mins, 1)}

# --- Dashboard Layout ---
st.markdown("<div class='section-header'>ATE Status Overview</div>", unsafe_allow_html=True)
header_cols = st.columns(4)

stage_results = {}
total_needed = 0
total_real_capacity = 0 
project_assigned_total_ids = [] 
today = date.today()

# --- Stations Configuration ---
st.markdown("<hr style='margin-top: 15px; margin-bottom: 15px; border: none; border-top: 1px solid #e6e6e6;'>", unsafe_allow_html=True)

input_cols = st.columns(3)
stages = [
    {"name": "FT1", "tt": 150.0, "fpy": 95.0, "qty": 36000},
    {"name": "FT2", "tt": 25.0, "fpy": 99.9, "qty": 75000},
    {"name": "FT3", "tt": 25.0, "fpy": 99.9, "qty": 0}
]

global_due_date = None
prev_start = None
prev_cycle_time = 0
prev_out_qty = 0  

for i, stage in enumerate(stages):
    current_ft_key = locals()[f"ft{i+1}_key"] 

    with input_cols[i]:
        is_enabled = True
        if i == 2: 
            col_title, col_toggle = st.columns([0.65, 0.35])
            with col_title:
                st.subheader(f"📍 {stage['name']} Station")
            with col_toggle:
                is_enabled = st.toggle("Enable", value=True, key="ft3_enable")
            
            if not is_enabled:
                st.session_state[current_ft_key] = []
        else:
            st.subheader(f"📍 {stage['name']} Station")

        
        # --- (A) 數量連動 ---
        if i == 0:
            ship_qty = st.number_input(f"Input Quantity (Units)", value=stage['qty'], step=1000, key=f"q_{i}")
            st.caption("(Base Input quantity configuration)")
        else:
            calc_qty = int(prev_out_qty)
            if f"q_{i}" not in st.session_state:
                st.session_state[f"q_{i}"] = calc_qty 
            
            ship_qty = st.number_input(f"Input Quantity (Units)", step=1000, disabled=not is_enabled, key=f"q_{i}")
            st.caption(f"*(Linked Suggestion: {calc_qty:,})*")

        # --- (B) 日期連動 ---
        if i == 0:
            c1, c2 = st.columns(2)
            start_date = c1.date_input(f"Start Date", value=today, key=f"sd_{i}")
            
            if f"dd_{i}" in st.session_state and st.session_state[f"dd_{i}"] < start_date:
                st.session_state[f"dd_{i}"] = start_date
                
            due_date = c2.date_input(f"Due Date", value=start_date + timedelta(days=6), min_value=start_date, key=f"dd_{i}")
            global_due_date = due_date 
            st.caption("*(Base schedule configuration)*")
        else:
            offset_days = math.ceil(prev_cycle_time / 1440)
            calc_start_date = prev_start + timedelta(days=offset_days)
            calc_due_date = global_due_date 
            
            if f"sd_{i}" not in st.session_state:
                st.session_state[f"sd_{i}"] = calc_start_date
            if f"dd_{i}" not in st.session_state:
                st.session_state[f"dd_{i}"] = calc_due_date
            
            d_c1, d_c2 = st.columns(2)
            start_date = d_c1.date_input(f"Start Date", disabled=not is_enabled, key=f"sd_{i}")
            
            if st.session_state[f"dd_{i}"] < start_date:
                st.session_state[f"dd_{i}"] = start_date
                
            due_date = d_c2.date_input(f"Due Date", min_value=start_date, disabled=not is_enabled, key=f"dd_{i}")
            st.caption(f"*(Suggested Start: {calc_start_date.strftime('%m/%d')} | Due: {calc_due_date.strftime('%m/%d')})*")
        
        working_days = max((due_date - start_date).days, 1)
        daily_target_units = ship_qty / working_days

        # --- (C) 詳細參數配置 ---
        with st.expander("Detailed Parameters", expanded=False):
            ls = st.number_input("Lot Size", value=6600, key=f"ls_{i}", disabled=not is_enabled)
            si = st.selectbox("Site", [4, 8, 16, 32], index=1, key=f"si_{i}", disabled=not is_enabled)
            tt_val = st.number_input("Test Time (s)", value=stage['tt'], key=f"tt_{i}", disabled=not is_enabled)
            fpy_val = st.number_input("FPY %", value=stage['fpy'], key=f"fy_{i}", disabled=not is_enabled)
            oee_val = st.number_input("OEE %", value=70.0, key=f"oe_{i}", disabled=not is_enabled)
        
        if is_enabled:
            prev_out_qty = ship_qty * (fpy_val / 100.0)
            res = calculate_metrics(daily_target_units, ls, si, tt_val, fpy_val, oee_val)
            
            st.markdown(f"""
                <div style='background-color: #eef4ff; padding: 16px; border-radius: 8px; margin-bottom: 16px; color: #004280; font-size: 14px;'>
                    <div style='margin-bottom: 12px;'>📅 <b>{working_days} day(s)</b> planning interval</div>
                    <div style='margin-bottom: 12px;'>🎯 Target: <b>{int(daily_target_units):,}</b> units/day</div>
                    <div>🚀 Single Capacity (UPD): <b>{res['upd']:,}</b> units/day</div>
                </div>
            """, unsafe_allow_html=True)

            st.markdown(f"""
                <div class='metric-card'>
                    <div style='color: #666; font-size: 12px; font-weight:bold; text-transform: uppercase;'>REQUIRED</div>
                    <div style='font-size: 32px; font-weight: bold; color: #1E3A8A;'>{res['req']} <span style='font-size: 16px;'>Testers</span></div>
                </div>
            """, unsafe_allow_html=True)
        else:
            prev_out_qty = ship_qty 
            res = {"req": 0, "uph": 0, "upd": 0, "cyc": 0}
            
            st.info("⏸️ **Station Disabled** \n\nNo target or capacity required for this station.  \n&nbsp;")
            st.markdown(f"""
                <div class='metric-card' style='opacity: 0.5; background-color: #f8f9fa;'>
                    <div style='color: #888; font-size: 12px; font-weight:bold; text-transform: uppercase;'>REQUIRED TESTERS</div>
                    <div style='font-size: 32px; font-weight: bold; color: #888;'>0 <span style='font-size: 16px;'>Testers</span></div>
                </div>
            """, unsafe_allow_html=True)

        st.markdown("---")
        
        # --- (D) Assign ATE IDs ---
        all_ft_states = [ft1_key, ft2_key, ft3_key]
        other_selections_for_this_stage = []
        
        other_selections_for_this_stage.extend(occupied_ates)
        
        for k in all_ft_states:
            if k != current_ft_key: 
                other_selections_for_this_stage.extend(st.session_state[k])
        
        options_for_this_stage = [ate for ate in active_ate_pool if ate not in other_selections_for_this_stage]
        options_for_this_stage = sorted(list(set(options_for_this_stage + st.session_state[current_ft_key])))

        assigned_ates = st.multiselect(
            f"Assign ATE for {stage['name']}", 
            options=options_for_this_stage, 
            key=current_ft_key,
            placeholder="Click to select Tester" if is_enabled else "Station Disabled",
            disabled=not is_enabled 
        )
        
        if is_enabled:
            assigned_ates = sorted(assigned_ates) 
            project_assigned_total_ids.extend(assigned_ates)
            
            real_station_upd = res['upd'] * (len(assigned_ates) if assigned_ates else res['req'])
            total_real_capacity += real_station_upd
            
            extra_upd = 0
            if len(assigned_ates) > res['req']:
                extra_upd = (len(assigned_ates) - res['req']) * res['upd']
                st.markdown(f"<div class='extra-upd'>➕ Extra Capacity: +{extra_upd:,} units/day</div>", unsafe_allow_html=True)

            if assigned_ates:
                tag_cols = st.columns(5) 
                for idx, id in enumerate(assigned_ates):
                    with tag_cols[idx % 5]:
                        st.button(id, key=f"btn_{i}_{id}", on_click=remove_ate, args=(current_ft_key, id), type="primary")
                st.markdown(f"<div style='color: #6c757d; font-size: 13px; margin-top: 5px;'>Count: {len(assigned_ates)} ATE selected.</div>", unsafe_allow_html=True)
            else:
                st.markdown("<div style='color: #6c757d; font-size: 13px;'>No units assigned.</div>", unsafe_allow_html=True)

            prev_start = start_date
            prev_cycle_time = res['cyc']

            res.update({
                "daily_target": daily_target_units,
                "real_upd": real_station_upd, 
                "assigned_ids": ", ".join(assigned_ates),
                "assigned_count": len(assigned_ates),
                "ship_qty": int(ship_qty),
                "start_date": start_date,
                "due_date": due_date,
                "days": working_days,
                "extra_upd": extra_upd,
                "oee": oee_val 
            })
            stage_results[stage['name']] = res
            total_needed += res["req"] 
        else:
            prev_start = start_date
            prev_cycle_time = 0
            st.markdown("<div style='color: #aaa; font-size: 13px;'>Station is currently disabled.</div>", unsafe_allow_html=True)

# ==============================================================================
# --- (E) Global Summary Update ---
# ==============================================================================
with header_cols[0]:
    st.markdown(f"""
        <div class='metric-card'>
            <div style='color: #666; font-size: 12px; font-weight:bold; text-transform: uppercase;'>Total Req. Testers (Min)</div>
            <div style='font-size: 32px; font-weight: bold; color: #1E3A8A;'>{total_needed} <span style='font-size: 16px;'>Testers</span></div>
        </div>
    """, unsafe_allow_html=True)

with header_cols[1]:
    remaining_fleet = net_fleet_num - total_needed
    status = "Healthy" if remaining_fleet >= 0 else "Shortage"
    status_color = "#1E3A8A" if remaining_fleet >= 0 else "#dc3545" 
    delta_color = "#28a745" if remaining_fleet >= 0 else "#dc3545"
    sign = "+" if remaining_fleet >= 0 else ""
    st.markdown(f"""
        <div class='metric-card'>
            <div style='color: #666; font-size: 12px; font-weight:bold; text-transform: uppercase;'>Resource Status</div>
            <div style='font-size: 32px; font-weight: bold; color: {status_color};'>{status} <span style='font-size: 14px; color: {delta_color};'>({sign}{remaining_fleet} Testers)</span></div>
        </div>
    """, unsafe_allow_html=True)

with header_cols[2]:
    breakdown_str = " | ".join([f"{k}: {int(v['real_upd']):,}" for k, v in stage_results.items()])
    st.markdown(f"""
        <div class='metric-card'>
            <div style='color: #666; font-size: 12px; font-weight:bold; text-transform: uppercase;'>Total Real Daily WIP Flow</div>
            <div style='font-size: 32px; font-weight: bold; color: #1E3A8A;'>{int(total_real_capacity):,} <span style='font-size: 16px;'>Units/Day</span></div>
            <div style='color: #6c757d; font-size: 11px; font-weight: bold; margin-top: 5px;'>{breakdown_str}</div>
        </div>
    """, unsafe_allow_html=True)

with header_cols[3]:
    if total_needed > 0:
        acc_oee = sum(res['oee'] * res['req'] for res in stage_results.values()) / total_needed
    else:
        acc_oee = 0.0
    st.markdown(f"""
        <div class='metric-card'>
            <div style='color: #666; font-size: 12px; font-weight:bold; text-transform: uppercase;'>Accumulated OEE</div>
            <div style='font-size: 32px; font-weight: bold; color: #1E3A8A;'>{acc_oee:.1f} <span style='font-size: 16px;'>%</span></div>
        </div>
    """, unsafe_allow_html=True)

# ==============================================================================
# --- Capacity Loading Analysis ---
# ==============================================================================
st.divider()
st.markdown("### 🛰️ Capacity Loading Analysis")

num_project_assigned = len(project_assigned_total_ids)
num_occupied = len(occupied_ates)
total_used_in_plant = num_project_assigned + num_occupied

# 計算 Loading
load_project = (num_project_assigned / net_fleet_num) * 100 if net_fleet_num > 0 else 0
load_plant = (total_used_in_plant / running_fleet_num) * 100 if running_fleet_num > 0 else 0

col_load1, col_load2 = st.columns(2)

with col_load1:
    if num_project_assigned > net_fleet_num:
        st.error(f"🚨 Error: Assigned ATEs ({num_project_assigned}) exceeds Net Available ({net_fleet_num})!")
    else:
        st.progress(int(min(load_project, 100)), text=f"📌 Project Utilization: {load_project:.1f}%")
        st.markdown(f"<div style='color: #555; font-size: 13px;'>Calculation: {num_project_assigned} (Project) / {net_fleet_num} (Available).</div>", unsafe_allow_html=True)

with col_load2:
    if total_used_in_plant > running_fleet_num:
        st.error(f"🚨 Error: Total In-Use ATEs ({total_used_in_plant}) exceeds Total Plant Fleet ({running_fleet_num})!")
    else:
        st.progress(int(min(load_plant, 100)), text=f"🏭 Total ATEs Utilization: {load_plant:.1f}%")
        st.markdown(f"<div style='color: #555; font-size: 13px;'>Calculation: {total_used_in_plant} (Project + Occupied) / {running_fleet_num} (ATEs Pool).</div>", unsafe_allow_html=True)


# ==============================================================================
# --- Export Data Table ---
# ==============================================================================
df_summary_data = {
    "Project": [actual_project_name] * len(stage_results.keys()), 
    "Station": list(stage_results.keys()), 
    "Input Qty": [res['ship_qty'] for res in stage_results.values()],
    "Start Date": [res['start_date'].strftime('%Y/%m/%d') for res in stage_results.values()],
    "Due Date": [res['due_date'].strftime('%Y/%m/%d') for res in stage_results.values()],
    "OEE %": [f"{res['oee']:.2f}" for res in stage_results.values()], 
    "Duration (Days)": [res['days'] for res in stage_results.values()],
    "Req. Testers": [res['req'] for res in stage_results.values()],
    "Assigned": [res['assigned_count'] for res in stage_results.values()],
    "Est. UPH": [res['uph'] for res in stage_results.values()],
    "Real UPD (per Station)": [res['real_upd'] for res in stage_results.values()],
    "Extra UPD (Planning)": [res['extra_upd'] for res in stage_results.values()],
    "Assigned Testers (Details)": [res['assigned_ids'] for res in stage_results.values()]
}
df_summary = pd.DataFrame(df_summary_data)
st.table(df_summary)

st.markdown("### 💾 Export Data")
csv = df_summary.to_csv(index=False).encode('utf-8')
st.download_button(
    label="Download Evaluation Result (CSV)",
    data=csv,
    file_name=f'ate_capacity_{actual_project_name}.csv', 
    mime='text/csv',
    use_container_width=True
)

# ==============================================================================
# --- Sidebar: Dynamic Factory Layout (擬真廠區平面圖) ---
# ==============================================================================
st.sidebar.divider()
st.sidebar.markdown("### 🗺️ Live Zion Area Layout (Phase 3)")

fig = go.Figure()

MAP_COORDS = {
    1: (5, 12),  2: (5, 11),  3: (5, 10),  4: (5, 9),
    5: (4, 10),  6: (4, 9),   7: (5, 2),   8: (5, 3),
    9: (4, 2),  10: (4, 3),  11: (5, 5),  12: (5, 6),
    13: (4, 5), 14: (4, 6),  15: (5, 7),  16: (4, 7),
    17: (3, 10), 18: (3, 9),  19: (3, 2),  20: (3, 3),
    21: (2, 5),  22: (2, 6),  23: (3, 5),  24: (3, 6),
    25: (3, 7),  26: (2, 10), 27: (4, 11), 28: (3, 11),
    29: (1, 9),  30: (0, 9),  31: (2, 7),  32: (4, 12),
    33: (2, 2),  34: (2, 3),  35: (1, 2),  36: (1, 3),
    37: (-1, 9), 38: (-2, 9)
}

fig.add_shape(type="rect", x0=-2.5, y0=1.5, x1=5.5, y1=12.5,
              fillcolor="#f8f9fa", line=dict(color="#e9ecef", width=1), layer="below")

x_coords, y_coords, colors, texts, hover_texts = [], [], [], [], []

# 定義狀態顏色
COLOR_AVL = "#28a745"   # 綠色 (可用)
COLOR_OCC = "#dc3545"   # 紅色 (被佔用)
COLOR_FT1 = "#1E3A8A"   # 深藍 (FT1)
COLOR_FT2 = "#6f42c1"   # 紫色 (FT2)
COLOR_FT3 = "#fd7e14"   # 橘色 (FT3)
COLOR_NA  = "#e9ecef"   # 淺灰 (未運作/保留區)

for i in range(1, total_fleet_num + 1):
    if i not in MAP_COORDS: continue
    
    ate = f"ATE{i:02d}"
    x, y = MAP_COORDS[i]
    x_coords.append(x)
    y_coords.append(y)
    
    machine_name = f"93K_{i:02d}"
    if i <= 24:
        machine_name += f" (EXA{i:02d})"
    
    if i > running_fleet_num:
        colors.append(COLOR_NA)
        texts.append(f"<b>{i:02d}</b><br><span style='font-size:8px; color:#aaa;'>N/A</span>")
        hover_texts.append(f"<b>{machine_name}</b><br>Status: Reserved / Offline")
    elif ate in st.session_state[ft1_key]:
        colors.append(COLOR_FT1)
        texts.append(f"<b>{i:02d}</b><br><span style='font-size:8px'>FT1</span>")
        hover_texts.append(f"<b>{machine_name}</b><br>Status: Assigned to FT1")
    elif ate in st.session_state[ft2_key]:
        colors.append(COLOR_FT2)
        texts.append(f"<b>{i:02d}</b><br><span style='font-size:8px'>FT2</span>")
        hover_texts.append(f"<b>{machine_name}</b><br>Status: Assigned to FT2")
    elif ate in st.session_state[ft3_key]:
        colors.append(COLOR_FT3)
        texts.append(f"<b>{i:02d}</b><br><span style='font-size:8px'>FT3</span>")
        hover_texts.append(f"<b>{machine_name}</b><br>Status: Assigned to FT3")
    elif ate in occupied_ates:
        colors.append(COLOR_OCC)
        texts.append(f"<b>{i:02d}</b><br><span style='font-size:8px'>OCC</span>")
        hover_texts.append(f"<b>{machine_name}</b><br>Status: Occupied (Other Project)")
    else:
        colors.append(COLOR_AVL)
        texts.append(f"<b>{i:02d}</b><br><span style='font-size:8px'>AVL</span>")
        hover_texts.append(f"<b>{machine_name}</b><br>Status: Available (Idle)")

fig.add_trace(go.Scatter(
    x=x_coords, y=y_coords, mode='markers+text',
    marker=dict(symbol='square', size=38, color=colors, line=dict(width=1, color="white")),
    text=texts,
    textposition="middle center",
    textfont=dict(color="white", size=11, family="Arial"),
    hoverinfo="text",
    hovertext=hover_texts
))

fig.update_layout(
    xaxis=dict(visible=False, range=[-3, 6]),
    yaxis=dict(visible=False, range=[1, 13]),
    margin=dict(l=0, r=0, t=0, b=0),
    height=480, 
    plot_bgcolor="rgba(0,0,0,0)",
    paper_bgcolor="rgba(0,0,0,0)",
    showlegend=False,
    dragmode=False 
)

st.sidebar.plotly_chart(fig, use_container_width=True, config={'displayModeBar': False})

st.sidebar.markdown(f"""
<div style='display: flex; flex-wrap: wrap; gap: 10px; font-size: 14px; color: #444; margin-top: -10px; margin-left: 40px;'>
    <div style='width: 45%;'><span style='color:{COLOR_AVL}; font-size: 16px;'>■</span> Available</div>
    <div style='width: 45%;'><span style='color:{COLOR_OCC}; font-size: 16px;'>■</span> Occupied</div>
    <div style='width: 45%;'><span style='color:{COLOR_FT1}; font-size: 16px;'>■</span> FT1 Assigned</div>
    <div style='width: 45%;'><span style='color:{COLOR_FT2}; font-size: 16px;'>■</span> FT2 Assigned</div>
    <div style='width: 45%;'><span style='color:{COLOR_FT3}; font-size: 16px;'>■</span> FT3 Assigned</div>
    <div style='width: 45%;'><span style='color:{COLOR_NA}; font-size: 16px;'>■</span> N/A / Offline</div>
</div>
""", unsafe_allow_html=True)