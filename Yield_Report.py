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
    tester_clocks = {tester: now - timedelta(days=20) for tester in testers}
    
    for i in range(200):
        tester = np.random.choice(testers)
        wait_time = timedelta(hours=np.random.uniform(0.5, 3))
        start_time = tester_clocks[tester] + wait_time
        test_duration = timedelta(hours=np.random.uniform(6, 14))
        end_time = start_time + test_duration
        qty = np.random.randint(1500, 3500)
        yield_rate = np.random.uniform(0.90, 0.99)
        pass_qty = int(qty * yield_rate)
        
        data.append([f"LOT_{i:04d}", "FT1", "PROD_GS631_ZC13...", tester, start_time, end_time, qty, pass_qty])
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

st.sidebar.divider()
st.sidebar.header("📐 3. Planning & Targets")
st.sidebar.caption("Define baselines for Capacity Planning.")
theo_max_upd = st.sidebar.number_input("Theoretical Max UPD (100% OEE)", value=4240, step=10)
planned_upd = st.sidebar.number_input("Planned Target UPD", value=2800, step=100)

# ==============================================================================
# --- 2. Main Area: Help Section & Filters ---
# ==============================================================================

# 🌟 NEW: 完整的中文定義說明展開區塊
with st.expander("ℹ️ Help: 參數定義與計算公式 (Formula & Definitions)"):
    st.markdown("""
    本系統採用嚴謹的 IE (工業工程) 邏輯，結合產線實際報表數據，推算出最真實的機台效率。各項指標定義如下：

    #### 1. 產能指標 (Capacity Metrics)
    * **Gross UPD (實際吞吐量):** 採用 `TestQty` 計算，代表機台實際測試過的總顆數（包含良品與不良品）。此數據用來評估機台的「純生產速度」。
    * **Net UPD (有效良品量):** 採用 `PassQty` 計算，代表機台產出的有效良品數。
    * **運轉天數 (Active Days):** 機台在選定期間內，處於「正在測試中 (CheckIn ~ CheckOut)」的總時數除以 24 小時。
    * **計算公式 (精準平均):** `Avg UPD = 期間內總顆數 / Active Days`。此算法排除了「零碎未滿一天」造成的平均值誤差。

    #### 2. OEE 設備綜合效率拆解 (Availability, Performance, Quality)
    本系統採用 Top-Down 產出導向與傳統 A/P/Q 雙軌驗證，確保數據準確度。
    * **A (稼動率 / Availability):** 衡量機台有多常在生產。
      * `計算 = 實際運轉天數 (Active Days) / 報表首尾跨越的日曆天數 (Calendar Span)`
    * **P (產能效率 / Performance):** 衡量機台運作時，有沒有達到理論該有的速度。
      * `理論 UPH = Theoretical Max UPD / 24`
      * `實際 UPH = 總測試量 (TestQty) / 實際運作總時數`
      * `計算 = 實際 UPH / 理論 UPH`
    * **Q (良率 / Quality):** 測試品質。
      * `計算 = 總良品數 (PassQty) / 總測試數 (TestQty)`
    * **整體 OEE (Overall OEE):**
      * `計算 = 實際平均 Gross UPD / Theoretical Max UPD`。 (等同於 A × P 綜合表現)

    #### 3. 產能規劃指標 (Capacity Planning)
    * **Planned Target UPD (規劃基準):** 生管或工程師預設的安全排程水準（例如：2,800 ea/day）。
    * **隱含 OEE (Implied OEE):** 該規劃基準佔理論 100% 產能的比例。`計算 = Planned Target UPD / Theoretical Max UPD`。這反映了排程預設保留了多少百分比作為換料、維修、重測的緩衝。
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

tester_summary = filtered_df.groupby('Tester').agg(
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

theo_uph = theo_max_upd / 24.0
tester_summary['Actual_UPH'] = np.where(tester_summary['Total_Duration_Hr'] > 0, tester_summary['Total_TestQty'] / tester_summary['Total_Duration_Hr'], 0)
tester_summary['Performance (P)'] = tester_summary['Actual_UPH'] / theo_uph

tester_summary['Yield (Q)'] = np.where(tester_summary['Total_TestQty'] > 0, tester_summary['Total_PassQty'] / tester_summary['Total_TestQty'], 0)
tester_summary['Avg_OEE'] = tester_summary['Avg_Gross_UPD'] / theo_max_upd
tester_summary = tester_summary.sort_values(by='Tester', ascending=True)

# ==============================================================================
# --- 4. Dashboard UI ---
# ==============================================================================

st.markdown("### 📊 Overall Performance")
total_test = int(tester_summary['Total_TestQty'].sum())
total_pass = int(tester_summary['Total_PassQty'].sum())
overall_yield = (total_pass / total_test) * 100 if total_test > 0 else 0
active_testers = tester_summary['Tester'].nunique()
valid_lots_count = int(tester_summary['Lot_Count'].sum())

c1, c2, c3, c4 = st.columns(4)
with c1: st.markdown(f"<div class='kpi-card'><div class='kpi-title'>Total TestQty (Gross)</div><div class='kpi-value'>{total_test:,}</div></div>", unsafe_allow_html=True)
with c2: st.markdown(f"<div class='kpi-card'><div class='kpi-title'>Total PassQty (Net)</div><div class='kpi-value'>{total_pass:,}</div></div>", unsafe_allow_html=True)
with c3: st.markdown(f"<div class='kpi-card'><div class='kpi-title'>Overall Yield</div><div class='kpi-value'>{overall_yield:.2f}%</div></div>", unsafe_allow_html=True)
with c4: st.markdown(f"<div class='kpi-card'><div class='kpi-title'>Active Testers</div><div class='kpi-value'>{active_testers}</div></div>", unsafe_allow_html=True)

st.divider()

st.markdown("## ATE Tester Capacity Analysis & Validation")
avg_gross_upd = tester_summary['Avg_Gross_UPD'].mean()
avg_net_upd = tester_summary['Avg_Net_UPD'].mean()
avg_oee = tester_summary['Avg_OEE'].mean() * 100

st.markdown("#### 1. Period Performance Summary")
st.caption(f"共計 {valid_lots_count} 個有效批次 (排除 < {min_lot_size} ea)。 此數據採用「有效運作時數」精密折算。")

sc1, sc2, sc3 = st.columns(3)
with sc1: st.markdown(f"<div class='summary-card'><div class='summary-title'>Avg Actual UPD (TestQty)</div><div class='summary-value'>{avg_gross_upd:,.0f}</div></div>", unsafe_allow_html=True)
with sc2: st.markdown(f"<div class='summary-card'><div class='summary-title'>Avg Effective UPD (PassQty)</div><div class='summary-value'>{avg_net_upd:,.0f}</div></div>", unsafe_allow_html=True)
with sc3: st.markdown(f"<div class='summary-card'><div class='summary-title'>Overall OEE</div><div class='summary-value'>{avg_oee:.1f}%</div></div>", unsafe_allow_html=True)

st.markdown("#### 2. Capacity Planning Insights (結論與規劃建議)")
buffer_pct = ((avg_gross_upd - planned_upd) / avg_gross_upd) * 100 if avg_gross_upd > 0 else 0
implied_oee = (planned_upd / theo_max_upd) * 100 if theo_max_upd > 0 else 0

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

insight_data = {
    "指標 (Metric)": ["報表平均 UPD (TestQty)", "您規劃的預估 UPD", "理論最大 UPD", f"隱含 OEE (於 {planned_upd})"],
    "數值 (Value)": [f"{avg_gross_upd:,.0f}", f"{planned_upd:,.0f}", f"{theo_max_upd:,.0f}", f"{implied_oee:.1f}%"],
    "備註 (Note)": [
        "實際產出吞吐量表現",
        "目前設定的排程安全水準",
        "100% OEE, 無停機/無異常",
        "排程預設包含換料、異常之折損空間"
    ]
}
st.dataframe(pd.DataFrame(insight_data), use_container_width=True, hide_index=True)

st.markdown("#### 3. Tester Performance Details (各機台表現明細與 A/P/Q 拆解)")
display_df = tester_summary[[
    'Tester', 'Lot_Count', 'Active_Days', 
    'Availability (A)', 'Performance (P)', 'Yield (Q)', 
    'Avg_Gross_UPD', 'Avg_Net_UPD', 'Avg_OEE'
]].copy()

display_df = display_df.rename(columns={
    'Tester': 'Tester', 'Lot_Count': 'Lot Count', 'Active_Days': 'Active Days (運轉天數)',
    'Availability (A)': 'Availability (稼動率)', 'Performance (P)': 'Performance (效率)', 'Yield (Q)': 'Yield (良率)',
    'Avg_Gross_UPD': 'Avg UPD (TestQty)', 'Avg_Net_UPD': 'Avg UPD (PassQty)', 'Avg_OEE': 'Avg OEE'
})

display_df['Active Days (運轉天數)'] = display_df['Active Days (運轉天數)'].apply(lambda x: f"{x:.1f}")
display_df['Availability (稼動率)'] = display_df['Availability (稼動率)'].apply(lambda x: f"{x*100:.1f}%")
display_df['Performance (效率)'] = display_df['Performance (效率)'].apply(lambda x: f"{x*100:.1f}%")
display_df['Yield (良率)'] = display_df['Yield (良率)'].apply(lambda x: f"{x*100:.1f}%")
display_df['Avg UPD (TestQty)'] = display_df['Avg UPD (TestQty)'].apply(lambda x: f"{x:,.0f}")
display_df['Avg UPD (PassQty)'] = display_df['Avg UPD (PassQty)'].apply(lambda x: f"{x:,.0f}")
display_df['Avg OEE'] = display_df['Avg OEE'].apply(lambda x: f"{x*100:.1f}%")

st.dataframe(display_df, use_container_width=True, hide_index=True)
