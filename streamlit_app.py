import streamlit as st
import math
import pandas as pd

# --- Page Configuration ---
st.set_page_config(page_title="ATE Capacity Planner", layout="wide")

# --- 💡 強效版 CSS：隱藏標籤並在選取後顯示提示文字 ---
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
    
    /* 1. 隱藏所有的紅色選中標籤 (Pills) */
    .stMultiSelect [data-baseweb="tag"], 
    .stMultiSelect [data-selection="true"] {
        display: none !important;
    }

    /* 2. 當有選取機台時，在灰色框框內顯示 "Selected units as below" */
    /* 我們利用 :has 偵測內部是否有標籤元素 */
    .stMultiSelect:has([data-baseweb="tag"]) div[data-baseweb="select"]::before {
        content: "Selected Tester as below";
        color: #888;
        font-size: 14px;
        position: absolute;
        left: 10px;
        top: 50%;
        transform: translateY(-50%);
        pointer-events: none;
        z-index: 1;
    }

    /* 3. 確保輸入框高度固定，不會因為隱藏標籤而留白過多 */
    .stMultiSelect div[data-baseweb="select"] {
        min-height: 42px !important;
    }

    /* 4. 帶編號的小方框樣式 (藍色視覺效果) */
    .tester-box {
        background-color: #4A90E2;
        color: white;
        padding: 4px 8px;
        border-radius: 4px;
        display: inline-block;
        margin: 3px;
        font-size: 12px;
        font-weight: bold;
        min-width: 60px;
        text-align: center;
        border: 1px solid #1E3A8A;
        box-shadow: 1px 1px 3px rgba(0,0,0,0.1);
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
    * **Sum Cycle Time:** `R0 + RT Cycle Time`
    * **Effective Mins/Day:** `1440 * OEE%`
    * **Lots/Day per Tester:** `Effective Mins/Day / Sum Cycle Time`
    * **Required Testers:** `CEILING(Target Lots per Day / Lots per Day per Tester)`
    * **UPD (Units Per Day):** `Lots/Day per Tester * Lot Size`
    """)

# --- Sidebar: Global Resources ---
st.sidebar.header("🏢 Global Resources")
total_fleet_num = st.sidebar.number_input("Total Available ATEs", min_value=1, value=24)

# 功能 1: 固定 Prefix "ATE" 並自動帶出編號清單
ate_pool = [f"ATE{i:02d}" for i in range(1, total_fleet_num + 1)]
st.sidebar.write(f"Generated IDs: `{ate_pool[0]} ~ {ate_pool[-1]}`")
st.sidebar.divider()

# --- Core Calculation Function ---
def calculate_metrics(target_lots, lot_size, site, tt, fpy, oee, op_time):
    td = lot_size / site if site > 0 else 0
    r0_mins = td * (tt / 60)
    rt_mins = ((lot_size * (1 - fpy / 100)) / site * (tt / 60)) + op_time if site > 0 else 0
    sum_mins = r0_mins + rt_mins
    eff_mins_per_day = 1440 * (oee / 100)
    lot_per_day_tester = eff_mins_per_day / sum_mins if sum_mins > 0 else 0
    req_testers = math.ceil(target_lots / lot_per_day_tester) if lot_per_day_tester > 0 else 0
    uph_effective = (lot_size / r0_mins) * 60 * (oee / 100) if r0_mins > 0 else 0
    upd_per_unit = lot_per_day_tester * lot_size
    return {
        "req": req_testers, "uph": int(uph_effective), "upd": int(upd_per_unit),
        "cyc": round(sum_mins, 1), "cap": lot_per_day_tester
    }

# --- Dashboard Layout ---
st.markdown("<div class='section-header'>Fleet Status Overview</div>", unsafe_allow_html=True)
header_cols = st.columns(4)

stage_results = {}
total_needed = 0
all_selected_ates = []

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
        target = st.number_input(f"Target (Lots/day)", value=stage['target'], key=f"t_{i}")
        
        with st.expander("Detailed Parameters", expanded=False):
            ls = st.number_input("Lot Size", value=3000, key=f"ls_{i}")
            si = st.selectbox("Site", [4, 8, 16, 32], index=1, key=f"si_{i}")
            tt_val = st.number_input("Test Time (s)", value=stage['tt'], key=f"tt_{i}")
            fpy_val = st.number_input("FPY %", value=stage['fpy'], key=f"fy_{i}")
            oee_val = st.number_input("OEE %", value=65.0, key=f"oe_{i}")
            op_v = st.number_input("RT OP Time (min)", value=60, key=f"op_{i}")
        
        res = calculate_metrics(target, ls, si, tt_val, fpy_val, oee_val, op_v)
        
        # 視覺卡片
        st.markdown(f"""
            <div class='metric-card'>
                <div style='color: #666; font-size: 12px; font-weight:bold;'>REQUIRED TESTERS</div>
                <div style='font-size: 32px; font-weight: bold; color: #1E3A8A;'>{res['req']} <span style='font-size: 16px;'>Units</span></div>
            </div>
        """, unsafe_allow_html=True)

        st.markdown("---")
        # 下拉指派機台
        assigned_ates = st.multiselect(
            f"Assign ATE IDs for {stage['name']}", 
            options=ate_pool, 
            key=f"select_{i}",
            placeholder="Click to select IDs"
        )
        all_selected_ates.extend(assigned_ates)
        
        # 下方呈現小方塊
        if assigned_ates:
            boxes_html = "".join([f"<div class='tester-box'>{id}</div>" for id in assigned_ates])
            st.markdown(f"<div style='margin-top:10px;'>{boxes_html}</div>", unsafe_allow_html=True)
            st.caption(f"Count: {len(assigned_ates)} ATE selected.")
        else:
            st.caption("No units assigned.")

        res["assigned_ids"] = ", ".join(assigned_ates)
        res["assigned_count"] = len(assigned_ates)
        stage_results[stage['name']] = res
        total_needed += res["req"]

# 檢查重複指派
duplicates = [x for x in set(all_selected_ates) if all_selected_ates.count(x) > 1]
if duplicates:
    st.error(f"🚨 Conflict Detected: {', '.join(duplicates)} assigned to multiple stations!")

# --- Global Summary Update ---
with header_cols[0]:
    st.metric("Total Req. Testers", f"{total_needed} Units")
with header_cols[1]:
    remaining = total_fleet_num - total_needed
    status = "Healthy" if remaining >= 0 else "Shortage"
    st.metric("Resource Status", status, f"{remaining} Units")
with header_cols[2]:
    total_uph = sum(res['uph'] for res in stage_results.values())
    st.metric("Total Est. UPH", f"{total_uph:,}")
with header_cols[3]:
    load = (total_needed / total_fleet_num) * 100 if total_fleet_num > 0 else 0
    st.metric("Fleet Capacity Load", f"{load:.1f}%")

# --- Capacity Loading Analysis ---
st.divider()
st.markdown("### 🛰️ Capacity Loading Analysis")
if total_needed > total_fleet_num:
    st.error(f"🚨 CRITICAL: Total demand ({total_needed}) exceeds fleet capacity ({total_fleet_num})!")
else:
    st.progress(int(min(load, 100)), text=f"Utilization: {load:.1f}%")

# 數據表格
df_summary_data = {
    "Station": list(stage_results.keys()),
    "Required Units": [res['req'] for res in stage_results.values()],
    "Assigned Units": [res['assigned_count'] for res in stage_results.values()],
    "Assigned IDs": [res['assigned_ids'] for res in stage_results.values()],
    "Sum Cycle Time (min)": [res['cyc'] for res in stage_results.values()],
    "Est. UPH": [res['uph'] for res in stage_results.values()],
    "Est. UPD (per Unit)": [res['upd'] for res in stage_results.values()]
}
df_summary = pd.DataFrame(df_summary_data)
st.table(df_summary)

# --- Export Data (原始格式) ---
st.markdown("### 💾 Export Data")
csv = df_summary.to_csv(index=False).encode('utf-8')
st.download_button(
    label="Download Evaluation Result (CSV)",
    data=csv,
    file_name='ate_capacity_report.csv',
    mime='text/csv',
    use_container_width=True
)