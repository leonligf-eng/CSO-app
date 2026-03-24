import streamlit as st
import math
import pandas as pd
from datetime import date, timedelta

# --- Page Configuration ---
st.set_page_config(page_title="ATE Capacity Planner", layout="wide")

# --- Custom CSS ---
st.markdown("""
    <style>
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
project_name = st.sidebar.text_input("Project Name / Code", placeholder="e.g., MBU_Project", help="Enter the project name for this capacity plan.")
actual_project_name = project_name if project_name.strip() else "Unnamed_Project"
st.sidebar.divider()

# --- Sidebar: Global Resources & Occupied ATEs ---
st.sidebar.header("🏢 Global Resources")
total_fleet_num = st.sidebar.number_input("Total ATEs in Plant", min_value=1, value=24)

total_ate_pool = [f"ATE{i:02d}" for i in range(1, total_fleet_num + 1)]
st.sidebar.caption(f"Plant IDs: `{total_ate_pool[0]} ~ {total_ate_pool[-1]}`")

for key in [occup_key, ft1_key, ft2_key, ft3_key]:
    st.session_state[key] = [ate for ate in st.session_state[key] if ate in total_ate_pool]

st.sidebar.divider()

st.sidebar.markdown("### 🔒 Exclude Occupied ATEs")

project_selections = []
project_selections.extend(st.session_state[ft1_key])
project_selections.extend(st.session_state[ft2_key])
project_selections.extend(st.session_state[ft3_key])

side_options = [ate for ate in total_ate_pool if ate not in project_selections]
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
    st.sidebar.caption("All plant ATEs available for project.")

net_fleet_num = total_fleet_num - len(occupied_ates)
available_for_project_pool = [ate for ate in total_ate_pool if ate not in occupied_ates]

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

        
        # --- (A) 數量連動 (修正：僅首次給定預設值，並開放修改) ---
        if i == 0:
            ship_qty = st.number_input(f"Input Quantity (Units)", value=stage['qty'], step=1000, key=f"q_{i}")
            st.caption("(Base Input quantity configuration)")
        else:
            calc_qty = int(prev_out_qty)
            # 只有初次載入時，將計算建議值寫入 Session State
            if f"q_{i}" not in st.session_state:
                st.session_state[f"q_{i}"] = calc_qty 
            
            # 開放使用者自行輸入 (並根據 Enable 狀態決定是否反灰)
            ship_qty = st.number_input(f"Input Quantity (Units)", step=1000, disabled=not is_enabled, key=f"q_{i}")
            st.caption(f"*(Linked Suggestion: {calc_qty:,})*")

        # --- (B) 日期連動 (修正：僅首次給定預設值，並開放修改) ---
        if i == 0:
            c1, c2 = st.columns(2)
            start_date = c1.date_input(f"Start Date", value=today, key=f"sd_{i}")
            due_date = c2.date_input(f"Due Date", value=start_date + timedelta(days=6), min_value=start_date, key=f"dd_{i}")
            global_due_date = due_date 
            st.caption("*(Base schedule configuration)*")
        else:
            offset_days = math.ceil(prev_cycle_time / 1440)
            calc_start_date = prev_start + timedelta(days=offset_days)
            calc_due_date = global_due_date 
            
            # 只有初次載入時，將計算建議日期寫入 Session State
            if f"sd_{i}" not in st.session_state:
                st.session_state[f"sd_{i}"] = calc_start_date
            if f"dd_{i}" not in st.session_state:
                st.session_state[f"dd_{i}"] = calc_due_date
            
            d_c1, d_c2 = st.columns(2)
            # 開放使用者自行選擇 (並根據 Enable 狀態決定是否反灰)
            start_date = d_c1.date_input(f"Start Date", disabled=not is_enabled, key=f"sd_{i}")
            due_date = d_c2.date_input(f"Due Date", disabled=not is_enabled, key=f"dd_{i}")
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
            
            st.info(f"📅 **{working_days} day(s)** planning interval  \n\n🎯 Target: **{int(daily_target_units):,}** units/day  \n\n🚀 Single Capacity (UPD): **{res['upd']:,}** units/day")

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
        
        options_for_this_stage = [ate for ate in total_ate_pool if ate not in other_selections_for_this_stage]
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
load = (num_project_assigned / net_fleet_num) * 100 if net_fleet_num > 0 else 0

if num_project_assigned > net_fleet_num:
    st.error(f"🚨 Real-time Critical Error: Project has assigned **{num_project_assigned}** ATEs, which exceeds the Net available fleet capacity ({net_fleet_num})!")
else:
    st.progress(int(min(load, 100)), text=f"Planned Loading Utilization: {load:.1f}%")
    st.markdown(f"<div style='color: #555; font-size: 13px;'>Calculation based on: {num_project_assigned} ATEs (Currently Assigned to Project) / {net_fleet_num} ATEs (Available).</div>", unsafe_allow_html=True)


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