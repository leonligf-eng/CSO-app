import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
from datetime import datetime, timedelta, time
import io
import uuid

st.set_page_config(page_title="ATE Capacity & OEE Analyzer", layout="wide")

# --- Custom CSS ---
st.markdown("""
    <style>
    /* 匯入現代字體，優先順序為 Google Sans -> Roboto -> 系統預設 */
    @import url('https://fonts.googleapis.com/css2?family=Roboto:wght@400;500;700;800&display=swap');
    
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
    
    [data-testid="stDataFrame"] th { text-align: center !important; }
    [data-testid="stDataFrame"] td { text-align: center !important; }
    
    div[data-testid="column"] button {
        padding: 0.2rem 0.5rem !important;
        font-size: 0.85rem !important;
        min-height: 32px !important;
        border-radius: 4px !important;
    }

    /* 🌟 選項 A：科技湛藍 (Tech Blue) */
    div.stButton > button {
        background-color: #0ea5e9 !important; /* Sky 500 */
        color: white !important;
        border: none !important;
        border-radius: 8px !important;
        padding: 0.6rem 1.2rem !important;
        font-weight: 600 !important;
        font-size: 15px !important;
        box-shadow: 0 4px 6px -1px rgba(14, 165, 233, 0.3), 0 2px 4px -1px rgba(14, 165, 233, 0.2) !important;
        transition: all 0.2s ease-in-out !important;
    }
    div.stButton > button:hover {
        background-color: #0284c7 !important; /* Sky 600 */
        box-shadow: 0 6px 8px -1px rgba(14, 165, 233, 0.4), 0 4px 6px -1px rgba(14, 165, 233, 0.3) !important;
        transform: translateY(-1px) !important;
    }
    div.stButton > button:active {
        transform: translateY(1px) !important;
        box-shadow: none !important;
    }

    /* 🌟 方案 B：藍天科技風 (Sky Blue / Cloud Native) */
    .custom-matrix-table {
        width: 100%;
        border-collapse: collapse;
        margin: 15px 0;
        font-size: 14px; 
        font-family: 'Google Sans', Roboto, -apple-system, BlinkMacSystemFont, sans-serif;
        table-layout: fixed; 
        word-wrap: break-word; 
        border-radius: 8px;
        overflow: hidden; 
        box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.05), 0 2px 4px -1px rgba(0, 0, 0, 0.03); /* 增加輕微雲端懸浮感 */
    }
    /* 表頭：柔和的雲朵灰藍色 */
    .custom-matrix-table thead th {
        background-color: #F8FAFC; 
        color: #475569; 
        text-align: center;
        padding: 10px 4px; 
        border-bottom: 2px solid #E2E8F0; 
        border-right: 1px solid #E2E8F0;
        font-size: 13px;
        font-weight: 600;
        letter-spacing: 0.5px;
    }
    /* 第一欄：維持乾淨 */
    .custom-matrix-table tbody th {
        background-color: #FFFFFF; 
        color: #334155; 
        font-weight: 600;
        border: 1px solid #E2E8F0; 
        padding: 10px 4px;
        text-align: center;
    }
    /* 資料儲存格 */
    .custom-matrix-table td {
        text-align: center;
        padding: 10px 4px; 
        border: 1px solid #E2E8F0; 
        vertical-align: middle;
        color: #334155;
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
    ops = ["FT1", "FTA", "MT1", "SLT", "QA"]
    programs_map = {
        "FT1": ["PROD_GS631_FT1_Proto1", "PROD_GS631_FT1_EVT1"], 
        "FTA": ["PROD_GS631_FTA_Proto1"], 
        "MT1": ["PROD_GS631_MT1_EVT1"],
        "SLT": ["PROD_GS631_SLT_Proto1", "PROD_GS631_SLT_EVT1"],
        "QA":  ["PROD_GS631_QA_Proto1", "PROD_GS631_QA_EVT1"]
    }
    products = ["SS16G", "MU16G", "HY12G", "SS12G"]
    
    tester_clocks = {tester: now - timedelta(days=30) for tester in testers}
    
    for i in range(350):
        tester = np.random.choice(testers)
        op = np.random.choice(ops)
        prog = np.random.choice(programs_map[op])
        prod = np.random.choice(products)
        
        wait_time = timedelta(hours=np.random.uniform(0.5, 3))
        start_time = tester_clocks[tester] + wait_time
        test_duration = timedelta(hours=np.random.uniform(6, 14))
        end_time = start_time + test_duration
        
        # 精確產生符合真實報表邏輯的資料
        qty = np.random.randint(3000, 8000) # TestQty (MES In)
        
        # 模擬 Input Loss (進站落差)
        input_loss = np.random.randint(0, 10)
        test_in_qty = qty - input_loss
        
        # 模擬機台測試良率
        pass_qty = int(test_in_qty * np.random.uniform(0.85, 0.99))
        
        # 根據真實資料: Test Out Qty 幾乎等於 PassQty (或是比Pass多一點點)
        pass_loss = np.random.randint(0, 5)
        test_out_qty = pass_qty + pass_loss
        
        # 模擬其他的作業損耗 (FailQty)
        handling_loss = np.random.randint(0, 10)
        fail_qty = test_in_qty - test_out_qty - handling_loss
        if fail_qty < 0: fail_qty = 0
        
        first_pass_qty = int(pass_qty * np.random.uniform(0.95, 1.0))
        
        data.append([f"LOT_{i:04d}", prod, op, prog, tester, start_time, end_time, qty, first_pass_qty, pass_qty, test_in_qty, test_out_qty, fail_qty])
        tester_clocks[tester] = end_time
        
    return pd.DataFrame(data, columns=['LotNo', 'ProductNo', 'OpNo', 'ProgramName', 'Tester', 'CheckInTime', 'CheckOutTime', 'TestQty', 'First Pass Qty', 'PassQty', 'TestInQty', 'TestOutQty', 'FailQty'])

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
                
            # 安全防護: 如果上傳的表沒有這些欄位，用標準值補上，避免掛掉
            if 'Test in Qty' in df.columns:
                df['TestInQty'] = pd.to_numeric(df['Test in Qty'], errors='coerce').fillna(df['TestQty'])
            elif 'TestInQty' not in df.columns: 
                df['TestInQty'] = df['TestQty']
                
            if 'Test Out Qty' in df.columns:
                df['TestOutQty'] = pd.to_numeric(df['Test Out Qty'], errors='coerce').fillna(df['PassQty'])
            elif 'TestOutQty' not in df.columns: 
                df['TestOutQty'] = df['PassQty']
                
            if 'FailQty' not in df.columns: 
                df['FailQty'] = df['TestQty'] - df['PassQty']
            else:
                df['FailQty'] = pd.to_numeric(df['FailQty'], errors='coerce').fillna(df['TestQty'] - df['PassQty'])
            
            return df
        except Exception as e:
            st.error(f"Failed to load file. Error: {str(e)}")
            return None
    return generate_mock_data()


# ==============================================================================
# --- 🌟 核心防護網：檔案變更與全面狀態清理 (File Change Detector) ---
# ==============================================================================
uploaded_file = st.sidebar.file_uploader("Upload Yield Report", type=["xlsx", "xls", "ods"])

# 判定目前是真實檔案還是假資料
current_file_name = uploaded_file.name if uploaded_file is not None else "mock_data"

if "last_uploaded_file" not in st.session_state:
    st.session_state.last_uploaded_file = current_file_name
    st.session_state.app_session_id = str(uuid.uuid4())[:8]

# 當使用者上傳新檔案，或是點擊 X 刪除檔案時觸發
if st.session_state.last_uploaded_file != current_file_name:
    st.session_state.last_uploaded_file = current_file_name
    st.session_state.app_session_id = str(uuid.uuid4())[:8]
    
    # 💥 終極大掃除：清除所有可能引發越界、找不到選項的過濾器狀態！
    keys_to_clear = [
        "master_mapping", "saved_progs", "op_select", 
        "prog_select_widget", "prod_select_widget",
        "date_picker", "last_selection_hash",
        "curr_min_date_ref", "curr_max_date_ref",
        "upstream_hash", "binning_op_selector"
    ]
    for key in list(st.session_state.keys()):
        if key in keys_to_clear or key.startswith("bin_ui_") or key.startswith("mapping_uploader_"):
            del st.session_state[key]


# 讀取資料
raw_df = load_data(uploaded_file)

if raw_df is None or raw_df.empty:
    st.warning("No data available.")
    st.stop()

# 取得全局最大最小值
global_min_date = raw_df['CheckInTime'].min().date() if pd.notnull(raw_df['CheckInTime'].min()) else datetime.now().date()
global_max_date = raw_df['CheckOutTime'].max().date() if pd.notnull(raw_df['CheckOutTime'].max()) else datetime.now().date()

st.sidebar.header("⚙️ Data Cleaning Rules")
min_lot_size = st.sidebar.number_input("Exclude Lots smaller than (Qty)", value=0, step=100)
st.sidebar.divider()

# ==============================================================================
# --- 頂層架構切分 (Main Tabs) ---
# ==============================================================================
main_tabs = st.tabs(["📊 OEE Analyzer", "🧬 Overall Build Yield"])

# ==============================================================================
# --- 🧬 Tab 2: Overall Build Yield Tracking ---
# ==============================================================================
with main_tabs[1]:
    st.write("")
    
    clean_build_df = raw_df[raw_df['TestQty'] >= min_lot_size].copy()
    
    if clean_build_df.empty:
        st.warning("No data available after applying Global Data Cleaning rules.")
    else:
        build_phases = ["Proto1.0", "Proto1.1", "EVT1.0(A0)", "EVT1.0(B0)", "EVT1.1", "DVT", "PVT", "MP"]
        all_programs = sorted(clean_build_df['ProgramName'].dropna().unique().tolist())
        
        prog_to_op = clean_build_df.drop_duplicates(subset=['ProgramName']).set_index('ProgramName')['OpNo'].to_dict()
        
        # 🌟 需求 4：自動隱藏 LS 和 T/R 站點
        all_ops_raw = clean_build_df['OpNo'].dropna().unique().tolist()
        available_ops = sorted([op for op in all_ops_raw if not any(x in op.upper() for x in ['LS', 'T/R', 'T\\R'])])
        
        # 初始化 master_mapping
        if "master_mapping" not in st.session_state:
            st.session_state.master_mapping = {phase: [] for phase in build_phases}
            for p in all_programs:
                p_u = p.upper()
                if "PROTO1.0" in p_u or "P10" in p_u: st.session_state.master_mapping["Proto1.0"].append(p)
                elif "PROTO1.1" in p_u or "P11" in p_u: st.session_state.master_mapping["Proto1.1"].append(p)
                elif "EVT1.0(A0)" in p_u or "EVT1.0_A" in p_u: st.session_state.master_mapping["EVT1.0(A0)"].append(p)
                elif "EVT1.0(B0)" in p_u or "EVT1.0_B" in p_u: st.session_state.master_mapping["EVT1.0(B0)"].append(p)
                elif "EVT1.1" in p_u: st.session_state.master_mapping["EVT1.1"].append(p)
                elif "DVT" in p_u: st.session_state.master_mapping["DVT"].append(p)
                elif "PVT" in p_u: st.session_state.master_mapping["PVT"].append(p)
                elif "MP" in p_u: st.session_state.master_mapping["MP"].append(p)

        # 🌟 核心標記：記錄是否剛剛執行過匯入
        if "just_imported" not in st.session_state:
            st.session_state.just_imported = False

        with st.expander("📥 Import Saved Mapping Configuration", expanded=False):
            st.markdown("<p style='font-size: 13px; color: #666;'>Upload a previously exported mapping CSV file to restore your program classifications instantly.</p>", unsafe_allow_html=True)
            
            uploader_key = f"mapping_uploader_{st.session_state.app_session_id}"
            uploaded_mapping = st.file_uploader("Upload build_mapping_config.csv", type=['csv'], key=uploader_key)
            
            if uploaded_mapping is not None:
                try:
                    # 讀取 CSV
                    mapping_df = pd.read_csv(uploaded_mapping)
                    if 'ProgramName' in mapping_df.columns and 'Build_Phase' in mapping_df.columns:
                        # 清空現有狀態，避免舊資料殘留
                        new_mapping = {phase: [] for phase in build_phases}
                        # 依照上傳的設定重新指派
                        for _, row in mapping_df.iterrows():
                            prog = row['ProgramName']
                            phase = row['Build_Phase']
                            if phase in new_mapping and prog in all_programs:
                                new_mapping[phase].append(prog)
                        
                        # 🌟 需求 5：按鈕樣式透過 CSS 注入
                        st.markdown('<div id="btn-apply-mapping"></div>', unsafe_allow_html=True)
                        if st.button("Apply Imported Mapping"):
                            st.session_state.master_mapping = new_mapping
                            # 🌟 強制標記為「剛匯入」狀態，這會保護總帳本在下一次渲染時不被覆寫
                            st.session_state.just_imported = True
                            
                            # 清空所有 UI 狀態，強制重新載入
                            for k in list(st.session_state.keys()):
                                if k.startswith("bin_ui_"):
                                    del st.session_state[k]
                            
                            # 強制將視圖切回 All
                            if "binning_op_selector" in st.session_state:
                                st.session_state.binning_op_selector = "All"
                                
                            st.success("Mapping applied successfully! The matrix has been updated.")
                            st.rerun()
                    else:
                        st.error("Invalid CSV format. Must contain 'ProgramName' and 'Build_Phase' columns.")
                except Exception as e:
                    st.error(f"Error reading file: {e}")

        with st.expander("⚙️ Build Phase Mapping Configuration (Interactive Binning)", expanded=True):
            st.markdown("<p style='font-size: 14px; color: #555;'>Assign programs to their corresponding Build Phases. Use the filter below to classify programs operation by operation.</p>", unsafe_allow_html=True)
            
            def on_op_radio_change():
                for k in list(st.session_state.keys()):
                    if k.startswith("bin_ui_"):
                        del st.session_state[k]
                # 只要使用者手動切換站點，就解除「剛匯入」的保護狀態
                st.session_state.just_imported = False
                        
            st.markdown("##### 🔍 1. Select Operation to Filter Programs")
            selected_binning_op = st.radio(
                "Filter by OpNo:", 
                options=["All"] + available_ops, 
                horizontal=True,
                label_visibility="collapsed",
                key="binning_op_selector",
                on_change=on_op_radio_change
            )
            st.write("")
            
            for phase in build_phases:
                st.session_state.master_mapping[phase] = [p for p in st.session_state.master_mapping[phase] if p in all_programs]
            
            assigned_progs_all = [p for progs in st.session_state.master_mapping.values() for p in progs]
            unassigned_pool_all = [p for p in all_programs if p not in assigned_progs_all]
            
            if selected_binning_op == "All":
                filtered_pool = unassigned_pool_all
            else:
                filtered_pool = [p for p in unassigned_pool_all if prog_to_op.get(p) == selected_binning_op]

            st.markdown(f"##### 📦 2. Unassigned Program Pool ({len(filtered_pool)})")
            if selected_binning_op != "All":
                st.caption(f"Currently showing unassigned programs for **{selected_binning_op}** only.")
            st.caption(", ".join(filtered_pool) if filtered_pool else f"🎉 All active programs for this selection have been assigned!")
            st.write("")
            
            cols1, cols2 = st.columns(4), st.columns(4)
            all_cols = cols1 + cols2
            
            mapping_changed = False
            
            for i, phase in enumerate(build_phases):
                with all_cols[i]:
                    if selected_binning_op == "All":
                        curr_visible_selected = [p for p in st.session_state.master_mapping[phase] if p in all_programs]
                    else:
                        curr_visible_selected = [p for p in st.session_state.master_mapping[phase] if p in all_programs and prog_to_op.get(p) == selected_binning_op]
                    
                    options = sorted(list(set(curr_visible_selected + filtered_pool)))
                    
                    # 💡 確保傳給 multiselect 的預設值，一定要包含在 options 裡面，否則會崩潰
                    safe_default = [x for x in curr_visible_selected if x in options]
                    
                    ui_key = f"bin_ui_{phase}_{selected_binning_op}_{st.session_state.app_session_id}"
                    
                    user_selection = st.multiselect(
                        f"📍 {phase}", 
                        options=options, 
                        default=safe_default, 
                        key=ui_key,
                        help=f"Assign {selected_binning_op if selected_binning_op != 'All' else 'any'} programs to {phase}"
                    )
                    
                    # 🌟 終極防護網：如果狀態是「剛剛匯入」，絕對不允許 UI 將其狀態逆向覆寫回總帳本！
                    if set(user_selection) != set(safe_default) and not st.session_state.just_imported:
                        mapping_changed = True
                        if selected_binning_op == "All":
                            protected_progs = []
                        else:
                            protected_progs = [p for p in st.session_state.master_mapping[phase] if prog_to_op.get(p) != selected_binning_op]
                        
                        st.session_state.master_mapping[phase] = protected_progs + user_selection
            
            # 在這整個區塊渲染完畢後，解除匯入保護鎖，讓後續的正常操作可以生效
            if st.session_state.just_imported:
                st.session_state.just_imported = False
            elif mapping_changed:
                # 只有在非剛匯入且真的有改變的情況下才重新渲染
                for k in list(st.session_state.keys()):
                    if k.startswith("bin_ui_"):
                        del st.session_state[k]
                st.rerun()

            # 🌟 核心功能 2：匯出 Mapping Config
            st.write("")
            st.divider()
            
            # 將目前的 Dictionary 轉換為供下載的 DataFrame
            export_data = []
            for phase, progs in st.session_state.master_mapping.items():
                for prog in progs:
                    export_data.append({"ProgramName": prog, "Build_Phase": phase})
            export_df = pd.DataFrame(export_data)
            
            if not export_df.empty:
                # 轉成 CSV 格式存入記憶體
                csv_buffer = io.StringIO()
                export_df.to_csv(csv_buffer, index=False)
                st.download_button(
                    label="📤 Export Current Mapping Configuration",
                    data=csv_buffer.getvalue(),
                    file_name="build_mapping_config.csv",
                    mime="text/csv",
                    help="Save your program assignments to a CSV file so you can import them later."
                )
            else:
                st.info("Assign some programs first to enable exporting.")

        st.write("")
        
        # ==============================================================================
        # --- 🎨 獨立閾值設定 ---
        # ==============================================================================
        with st.expander("🎨 Matrix Color Thresholds Settings", expanded=False):
            st.markdown("<p style='font-size: 13px; color: #666;'>Set independent visual alerts for each metric.</p>", unsafe_allow_html=True)
            
            t_col1, t_col2, t_col3 = st.columns(3)
            
            with t_col1:
                st.markdown("**Test Yield**")
                ft_g = st.number_input("🟩 Healthy (%)", value=95.0, step=0.5, key="ft_g")
                ft_r = st.number_input("🟥 Critical (%)", value=90.0, step=0.5, key="ft_r")
            
            with t_col2:
                st.markdown("**Operation Yield**")
                op_g = st.number_input("🟩 Healthy (%)", value=99.5, step=0.5, key="op_g")
                op_r = st.number_input("🟥 Critical (%)", value=98.0, step=0.5, key="op_r")
            
            with t_col3:
                st.markdown("**LS Yield**")
                ls_g = st.number_input("🟩 Healthy (%)", value=99.0, step=0.5, key="ls_g")
                ls_r = st.number_input("🟥 Critical (%)", value=95.0, step=0.5, key="ls_r")
            
            if ft_r >= ft_g or op_r >= op_g or ls_r >= ls_g:
                st.error("Critical limits must be lower than Healthy limits.")

        # ==============================================================================
        # --- 🌟 大一統矩陣 (Unified HTML Matrix) ---
        # ==============================================================================
        prog_to_build = {p: phase for phase, progs in st.session_state.master_mapping.items() for p in progs}
        build_df = clean_build_df.copy()
        build_df['Build_Phase'] = build_df['ProgramName'].map(prog_to_build)
        build_df = build_df.dropna(subset=['Build_Phase'])
        
        if build_df.empty:
            st.info("💡 Please assign at least one program to a Build Phase above to generate the Matrix Report.")
        else:
            st.markdown("#### 📊 Comprehensive Build Evolution Matrix")
            
            # --- 排序權重函數 ---
            def get_op_sort_weight(op_name):
                op_upper = op_name.upper()
                if "FT1" in op_upper: return 1
                if "MT1" in op_upper or "SLT" in op_upper: return 2
                if "FTA" in op_upper or "FT2" in op_upper: return 3
                if "LS" in op_upper: return 4
                return 99
            
            build_df['Op_Weight'] = build_df['OpNo'].apply(get_op_sort_weight)
            
            # --- 分類資料 ---
            ft_df = build_df[~build_df['OpNo'].str.contains('LS', case=False, na=False)].copy()
            ls_df = build_df[build_df['OpNo'].str.contains('LS', case=False, na=False)].copy()
            
            # 動態背景色判定
            def get_bg_color(val, g_thresh, r_thresh):
                if pd.isna(val): return ''
                if val >= g_thresh: return '#e6f4ea' # 淺綠
                elif val >= r_thresh: return '#fef08a' # 淺黃 warning
                else: return '#fee2e2' # 淺紅 critical
            
            # 🌟 動態文字顏色判定 (Critical 變深紅)
            def get_text_color(val, r_thresh):
                if pd.isna(val): return '#334155'
                return '#991b1b' if val < r_thresh else '#334155'

            all_used_phases = build_df['Build_Phase'].unique().tolist()
            ordered_phases = [p for p in build_phases if p in all_used_phases]
            
            html_out = '<div style="width: 100%;">'
            html_out += '<table class="custom-matrix-table">'
            html_out += '<thead><tr><th style="width: 10%;">Operation</th>'
            for col in ordered_phases:
                html_out += f'<th>{col}</th>'
            html_out += '</tr></thead><tbody>'

            # 🌟 需求 1：無邊框、實體色塊包覆感 (稍微加深底色界定區塊)
            modern_header_style = "background-color: #E0F2FE; color: #0369A1; font-family: 'Google Sans', Roboto, sans-serif; font-weight: 800; font-size: 16px; text-align: center; padding: 14px; letter-spacing: 0.5px;"

            # ==========================================
            # 區塊 1: Test Yield
            # ==========================================
            if not ft_df.empty:
                html_out += f'<tr><td colspan="{len(ordered_phases) + 1}" style="{modern_header_style}">Test Yield</td></tr>'
                
                ft_sum = ft_df.groupby(['OpNo', 'Build_Phase']).agg(T_Qty=('TestQty', 'sum'), P_Qty=('PassQty', 'sum')).reset_index()
                ft_sum['Yield'] = np.where(ft_sum['T_Qty'] > 0, (ft_sum['P_Qty'] / ft_sum['T_Qty']) * 100, np.nan)
                ft_sum['Op_Weight'] = ft_sum['OpNo'].apply(get_op_sort_weight)
                
                ordered_ops = ft_sum.sort_values('Op_Weight')['OpNo'].unique().tolist()
                
                for op in ordered_ops:
                    html_out += f'<tr><th>{op}</th>'
                    op_data = ft_sum[ft_sum['OpNo'] == op]
                    for phase in ordered_phases:
                        cell_data = op_data[op_data['Build_Phase'] == phase]
                        if cell_data.empty or pd.isna(cell_data['Yield'].values[0]):
                            # 🌟 需求 3：淡化空值符號
                            html_out += '<td><span style="color: #CBD5E1;">-</span></td>'
                        else:
                            y_val = cell_data['Yield'].values[0]
                            t_val = int(cell_data['T_Qty'].values[0])
                            p_val = int(cell_data['P_Qty'].values[0])
                            
                            bg = get_bg_color(y_val, ft_g, ft_r)
                            txt_c = get_text_color(y_val, ft_r)
                            
                            cell_html = f"<div style='color: {txt_c};'><b>{y_val:.2f}%</b><br><span style='font-size: 11px; font-weight: normal; color: #64748B;'>T: {t_val:,} | P: {p_val:,}</span></div>"
                            html_out += f'<td style="background-color: {bg};">{cell_html}</td>'
                    html_out += '</tr>'

            # ==========================================
            # 區塊 2: Operation Yield
            # ==========================================
            if not ft_df.empty:
                html_out += f'<tr><td colspan="{len(ordered_phases) + 1}" style="{modern_header_style}">Operation Yield</td></tr>'
                
                op_sum = ft_df.groupby(['OpNo', 'Build_Phase']).agg(
                    T_Qty=('TestQty', 'sum'),
                    T_In_Qty=('TestInQty', 'sum'),
                    T_Out_Qty=('TestOutQty', 'sum'),
                    P_Qty=('PassQty', 'sum'),
                    F_Qty=('FailQty', 'sum')
                ).reset_index()
                
                op_sum['Op_Weight'] = op_sum['OpNo'].apply(get_op_sort_weight)
                ordered_ops = op_sum.sort_values('Op_Weight')['OpNo'].unique().tolist()

                for op in ordered_ops:
                    html_out += f'<tr><th>{op}</th>'
                    op_data = op_sum[op_sum['OpNo'] == op]
                    for phase in ordered_phases:
                        cell_data = op_data[op_data['Build_Phase'] == phase]
                        if cell_data.empty or cell_data['T_Qty'].values[0] == 0:
                            html_out += '<td><span style="color: #CBD5E1;">-</span></td>'
                        else:
                            t_qty = float(cell_data['T_Qty'].values[0])
                            t_in_qty = float(cell_data['T_In_Qty'].values[0])
                            t_out_qty = float(cell_data['T_Out_Qty'].values[0])
                            p_qty = float(cell_data['P_Qty'].values[0])
                            f_qty = float(cell_data['F_Qty'].values[0])
                            
                            # 🌟 絕對鐵律公式
                            input_loss = max(0, int(t_qty - t_in_qty))
                            pass_loss = max(0, int(t_out_qty - p_qty))
                            fail_loss = max(0, int(t_in_qty - t_out_qty - f_qty))
                            
                            op_loss_total = input_loss + pass_loss
                            op_y_val = (1 - (op_loss_total / t_qty)) * 100 if t_qty > 0 else 0
                            
                            bg = get_bg_color(op_y_val, op_g, op_r)
                            txt_c = get_text_color(op_y_val, op_r)
                            
                            # 🌟 需求 2：次級資訊文字淡化 (#64748B)
                            cell_html = f"<div style='color: {txt_c};'><b>{op_y_val:.2f}%</b><br><div style='font-size: 11px; line-height: 1.3; color: #64748B;'>OP Loss: {op_loss_total:,}<br>(Input: {input_loss:,} | Pass: {pass_loss:,})<br>Fail Loss: {fail_loss:,}</div></div>"
                            html_out += f'<td style="background-color: {bg};">{cell_html}</td>'
                    html_out += '</tr>'

            # ==========================================
            # 區塊 3: Lead Scan Yield
            # ==========================================
            if not ls_df.empty:
                html_out += f'<tr><td colspan="{len(ordered_phases) + 1}" style="{modern_header_style}">LS Yield</td></tr>'
                
                ls_sum = ls_df.groupby(['OpNo', 'Build_Phase']).agg(T_Qty=('TestQty', 'sum'), P_Qty=('PassQty', 'sum')).reset_index()
                ls_sum['Yield'] = np.where(ls_sum['T_Qty'] > 0, (ls_sum['P_Qty'] / ls_sum['T_Qty']) * 100, np.nan)
                ls_sum['Op_Weight'] = ls_sum['OpNo'].apply(get_op_sort_weight)
                
                ordered_ops_ls = ls_sum.sort_values('Op_Weight')['OpNo'].unique().tolist()
                
                for op in ordered_ops_ls:
                    html_out += f'<tr><th>{op}</th>'
                    op_data = ls_sum[ls_sum['OpNo'] == op]
                    for phase in ordered_phases:
                        cell_data = op_data[op_data['Build_Phase'] == phase]
                        if cell_data.empty or pd.isna(cell_data['Yield'].values[0]):
                            html_out += '<td><span style="color: #CBD5E1;">-</span></td>'
                        else:
                            y_val = cell_data['Yield'].values[0]
                            t_val = int(cell_data['T_Qty'].values[0])
                            p_val = int(cell_data['P_Qty'].values[0])
                            
                            bg = get_bg_color(y_val, ls_g, ls_r)
                            txt_c = get_text_color(y_val, ls_r)
                            
                            cell_html = f"<div style='color: {txt_c};'><b>{y_val:.2f}%</b><br><span style='font-size: 11px; font-weight: normal; color: #64748B;'>T: {t_val:,} | P: {p_val:,}</span></div>"
                            html_out += f'<td style="background-color: {bg};">{cell_html}</td>'
                    html_out += '</tr>'
                    
            html_out += '</tbody></table></div>'
            st.write(html_out, unsafe_allow_html=True)


# ==============================================================================
# --- 📊 Tab 1: OEE Analyzer ---
# ==============================================================================
with main_tabs[0]:
    # ==============================================================================
    # --- 1. Main Area: Filters & Help Section ---
    # ==============================================================================
    with st.expander("ℹ️ Help: Formula & Parameter Definitions", expanded=False):
        st.markdown("""
        This system employs rigorous Industrial Engineering (IE) logic combined with actual production report data to calculate authentic equipment capacity and efficiency metrics.

        #### 1. Core OEE Calculation Formulas
        Our system utilizes a **Bottom-Up, Lot-based** tracking methodology. The core formulas driving the OEE metrics are defined as follows:
        
        * **Active Days:** The actual total time the tester spent executing testing operations.
        
          $$ \\text{Active Days} = \\sum (\\text{CheckOutTime} - \\text{CheckInTime}) \\div 24 $$

        * **Adjusted Span Days:** The continuous calendar time span, excluding major idle gaps (no-lot periods exceeding 24 hours).
        
          $$ \\text{Adjusted Span Days} = (\\text{Max CheckOut} - \\text{Min CheckIn}) - (\\text{Idle gaps} \\ge 24h) $$

        * **Availability (A):** Measures the proportion of time the tester is actively in production during the assigned timeframe.
        
          $$ A = \\frac{\\text{Active Days}}{\\text{Adjusted Span Days}} $$

        * **Performance (P):** Measures whether the tester is running at the theoretical speed when active (short stoppages or slow testing will naturally degrade this metric).
        
          $$ P = \\frac{(\\text{Total Test Qty} / \\text{Active Days})}{\\text{Theoretical Max UPD}} $$

        * **Avg OEE:** The authentic efficiency multiplier calculated by our system.
        
          $$ \\text{Avg OEE} = \\text{Availability (A)} \\times \\text{Performance (P)} $$

        ---

        #### 2. Capacity Planning & Throughput Metrics
        * **Normalized UPD (Units Per Day):** Converts the actual volume tested over a specific active period into an equivalent 24-hour rate for standardized comparison across different testers.
        
          $$ \\text{Normalized UPD} = \\frac{\\text{Total Qty}}{\\text{Active Days}} $$

        * **Planned Target UPD:** The safety scheduling baseline preset by Planning or Engineering departments.
        
        * **Implied OEE:** Reflects the built-in buffer percentage reserved for setups, maintenance, and re-tests within the planned schedule.
        
          $$ \\text{Implied OEE} = \\frac{\\text{Planned Target UPD}}{\\text{Theoretical Max UPD}} $$

        ---

        #### 3. Methodology Differences: ATE Smart Capacity vs. OSAT MES
        Understanding the calculation differences between our system and OSAT Daily Reports is crucial for accurate capacity planning.

        | Dimension | Our System (Lot-Based) | Typical OSAT MES (24h-Fixed) |
        | :--- | :--- | :--- |
        | **Time Foundation** | Calculates continuous spans based on actual **Lot Check-In/Out**, seamlessly handling cross-midnight lots. | Relies on a fixed **1440-minute daily window**, forcefully cutting lots at midnight, introducing boundary errors. |
        | **Data Source** | Strictly relies on objective **Machine Logs**, immune to manual reporting delays. | Relies heavily on **Operator Manual Scans** (state codes). Delayed scanning often inflates 'Idle' time. |
        | **OEE Formula** | $$ A \\times P $$ <br> *(Focuses on True ROI of Equipment)* | $$ \\text{Performance} \\times \\text{Rework Eff.} \\times \\text{D\\_O1\\%} $$ <br> *(Focuses on Shop-floor Management)* |

        **Typical OSAT Sub-Formulas for Reference:**
        
        * **Performance (總產出效率):** $$ \\text{Performance} = \\frac{\\text{Test Qty} \\times \\text{Standard Test Time}}{(24\\text{h} - \\text{Idle} - \\text{Other} - \\text{E2}) \\times \\text{Machine Count} \\times 60\\text{mins}} $$

        * **Rework Efficiency (重工效率):** $$ \\text{Rework Efficiency} = \\frac{\\text{Rework Production Time}}{\\text{Total Rework Test Time}} $$

        * **Quality (D_O1%):** $$ \\text{Quality (D\\_O1\\%)} = \\frac{\\text{Pass Qty}}{\\text{Pass Qty} + \\text{Fail Qty}} $$

        **Conclusion:** OSAT reports are designed for shop-floor administrative and human management. Our system's **Avg OEE** reflects the most authentic capacity return and equipment limitations.
        """, unsafe_allow_html=True)

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

    # ⚠️ 這裡完整保留您的 st.stop()
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

    tabs = st.tabs(selected_ops)

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