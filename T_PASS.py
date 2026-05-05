import streamlit as st
import datetime
import uuid

st.set_page_config(page_title="T-PASS (Test Program Approval & Sign-off System)", layout="wide")
#PROVE (Program Release & Offline Validation Engine)
# ==========================================
# 0. UI Theme Override (Vibrant Modern Tech Blue)
# ==========================================
st.markdown("""
<style>
    .stButton > button[kind="primary"] {
        background-color: #3b82f6 !important;
        color: #ffffff !important;
        border: none !important;
        border-radius: 6px !important;
        font-weight: 600 !important;
        letter-spacing: 0.5px !important;
        transition: all 0.2s ease-in-out !important;
        box-shadow: 0 2px 4px rgba(59, 130, 246, 0.2) !important;
    }
    .stButton > button[kind="primary"]:hover {
        background-color: #2563eb !important;
        box-shadow: 0 4px 10px rgba(59, 130, 246, 0.3) !important;
        transform: translateY(-1px) !important;
    }
    .stButton > button[kind="secondary"] {
        border: 1px solid #d1d5db !important;
        border-radius: 6px !important;
        background-color: #f9fafb !important;
        color: #4b5563 !important;
        font-weight: 500 !important;
        transition: all 0.2s ease-in-out !important;
    }
    .stButton > button[kind="secondary"]:hover {
        border-color: #9ca3af !important;
        background-color: #f3f4f6 !important;
    }
</style>
""", unsafe_allow_html=True)

# 💡 定義超緊湊的自定義分隔線，取代原生的 st.divider() 與 st.write("---")
DIVIDER = "<hr style='margin: 12px 0; border: none; border-top: 1px solid #e5e7eb;' />"

# ==========================================
# 1. Core Configuration
# ==========================================
FLOW_CONFIG = {
    "MP (Mass Production Release)": [
        {"title": "1. PDTE Release", "sla_days": 0, "poc": "CSO-PE", "desc": "Release new test program. Set baseline for error-proofing.", "form_type": "init_form"},
        {"title": "2. Offline Trial", "sla_days": 1, "poc": "CSO-OPS", "desc": "Offline CORR validation and Fuse check.", "form_type": "offline_check", "target_offline": 500, "target_fuse": 8},
        {"title": "3. Data Review", "sla_days": 1, "poc": "CSO-PE", "desc": "Review key items and OTP content correctness.", "form_type": "pe_signoff"},
        {"title": "4. Trial Run", "sla_days": 3, "poc": "CSO-OPS", "desc": "Ops handle the FT/SLT offline mode trial run data for MES and checking GSMI, EAC, etc., at KYEC", "form_type": "trial_data", "target_trial": 3000, "trial_mode": "Offline (Using SEN)"},
        {"title": "5. Official Release", "sla_days": 2, "poc": "CSO-PE", "desc": "Official release to OSAT.", "form_type": "final_release"}
    ],
    "NPI - Option 1 (Initial Release)": [
        {"title": "1. PDTE Release", "sla_days": 0, "poc": "CSO-PE", "desc": "Release new test program. Set baseline for error-proofing.", "form_type": "init_form"},
        {"title": "2. Offline Trial", "sla_days": 1, "poc": "CSO-OPS", "desc": "Offline CORR validation and Fuse check.", "form_type": "offline_check", "target_offline": 200, "target_fuse": 2},
        {"title": "3. Data Review", "sla_days": 1, "poc": "CSO-PE", "desc": "Review key items and OTP content correctness.", "form_type": "pe_signoff"},
        {"title": "4. Trial Run", "sla_days": 2, "poc": "CSO-OPS", "desc": "Ops handle the FT/SLT online mode trial run data for MES and checking GSMI, EAC, etc., at KYEC", "form_type": "trial_data", "target_trial": 50, "trial_mode": "Online (No SEN)"},
        {"title": "5. Official Release", "sla_days": 3, "poc": "CSO-PE", "desc": "Official release to OSAT upon team agreement.", "form_type": "final_release"}
    ]
}

selected_flow = "MP (Mass Production Release)"

# ==========================================
# 2. System State Management
# ==========================================
if 'current_phase' not in st.session_state: st.session_state.current_phase = 1
if 'expanded_phase' not in st.session_state: st.session_state.expanded_phase = 1
if 'tp_data' not in st.session_state: st.session_state.tp_data = {}

if 'mes_setup_done' not in st.session_state: st.session_state.mes_setup_done = False
if 'mes_setup_date' not in st.session_state: st.session_state.mes_setup_date = None
if 'mes_comments' not in st.session_state: st.session_state.mes_comments = ""
if 'start_date' not in st.session_state: st.session_state.start_date = None

# ==========================================
# 3. Helper Functions (Dynamic UI & Reports)
# ==========================================
def format_lots_for_report(lots):
    lines = []
    total_qty = 0
    weighted_yield = 0.0
    weighted_tt = 0.0

    for l in lots:
        q = float(l.get('qty') or 0)
        y = float(l.get('yield') or 0.0)
        t = float(l.get('tt') or 0)
        
        total_qty += q
        weighted_yield += q * y
        weighted_tt += q * t

    avg_yield = weighted_yield / total_qty if total_qty > 0 else 0
    avg_tt = weighted_tt / total_qty if total_qty > 0 else 0

    for l in lots:
        lot_info = l.get('lot') if l.get('lot') else "Pending_Lot_Info"
        q = int(l.get('qty') or 0)
        y = float(l.get('yield') or 0.0)
        lines.append(f"---> {lot_info} [{q}ea], Final yield: {y:.2f}%")
        if l.get('failure_summary'):
            lines.append(f"     Final Failure: {l['failure_summary']}")
            
    return "\n".join(lines), int(total_qty), avg_yield, avg_tt

def render_date_timeline(est_date, act_date):
    est_str = est_date.strftime('%m/%d') if est_date else "--/--"
    act_str = act_date.strftime('%m/%d') if act_date else "--/--"
    if act_date:
        act_html = f"<span style='color: {'#d32f2f' if est_date and act_date > est_date else '#2e7d32'}; font-weight: bold;'>Act: {act_str}</span>"
    else:
        act_html = f"<span style='color: #888;'>Act: {act_str}</span>"
    st.markdown(f"<div style='text-align: center; font-size: 12px; margin-top: -10px; margin-bottom: 15px;'><span style='color: gray;'>Est: {est_str}</span> <span style='color: #ccc; margin: 0 4px;'>|</span> {act_html}</div>", unsafe_allow_html=True)

def render_dynamic_programs(is_active):
    if "programs" not in st.session_state.tp_data:
        st.session_state.tp_data["programs"] = [{"id": str(uuid.uuid4()), "stage": "FT1", "revision": "PROD_ZC13_REV06P2", "tt": 45}]
        
    progs = st.session_state.tp_data["programs"]
    for idx, prog in enumerate(progs):
        cols = st.columns([2, 4, 2, 1]) if is_active else st.columns([2, 4, 2])
        with cols[0]: prog["stage"] = st.text_input("Stage (e.g. FT1, FT2)" if idx==0 else "", value=prog["stage"], key=f"p_stg_{prog['id']}", disabled=not is_active, label_visibility="visible" if idx==0 else "collapsed")
        with cols[1]: prog["revision"] = st.text_input("Program Revision" if idx==0 else "", value=prog["revision"], key=f"p_rev_{prog['id']}", disabled=not is_active, label_visibility="visible" if idx==0 else "collapsed")
        with cols[2]: prog["tt"] = st.number_input("Expected TT (s)" if idx==0 else "", value=prog["tt"], key=f"p_tt_{prog['id']}", disabled=not is_active, label_visibility="visible" if idx==0 else "collapsed")
        if is_active:
            with cols[3]:
                if idx == 0: st.markdown("<div style='margin-top: 28px;'></div>", unsafe_allow_html=True)
                if len(progs) > 1 and st.button("🗑️", key=f"del_p_{prog['id']}"):
                    st.session_state.tp_data["programs"] = [p for p in progs if p['id'] != prog['id']]
                    st.rerun()

    if is_active and st.button("➕ Add Program / Stage", key="add_prog"):
        st.session_state.tp_data["programs"].append({"id": str(uuid.uuid4()), "stage": "", "revision": "", "tt": 45})
        st.rerun()

def render_dynamic_lots(lot_key, target_qty, is_active, default_yield=99.0, default_tt=45):
    if lot_key not in st.session_state.tp_data:
        st.session_state.tp_data[lot_key] = [{"id": str(uuid.uuid4()), "lot": "", "qty": target_qty, "yield": default_yield, "tt": default_tt, "failure_summary": ""}]
        
    lots = st.session_state.tp_data[lot_key]
    for idx, lot_data in enumerate(lots):
        col_ratios = [3, 2, 2, 2, 1] if is_active else [3, 2, 2, 2]
        cols = st.columns(col_ratios)
        with cols[0]: lots[idx]["lot"] = st.text_input("Lot Information" if idx==0 else "", value=lot_data["lot"], key=f"{lot_key}_lot_{lot_data['id']}", disabled=not is_active, label_visibility="visible" if idx==0 else "collapsed")
        with cols[1]: lots[idx]["qty"] = st.number_input(f"Qty ({target_qty}u)" if idx==0 else "", value=lot_data["qty"], key=f"{lot_key}_qty_{lot_data['id']}", disabled=not is_active, label_visibility="visible" if idx==0 else "collapsed")
        with cols[2]: lots[idx]["yield"] = st.number_input("Final Yield (%)" if idx==0 else "", value=lot_data["yield"], key=f"{lot_key}_yld_{lot_data['id']}", disabled=not is_active, label_visibility="visible" if idx==0 else "collapsed")
        with cols[3]: lots[idx]["tt"] = st.number_input("Actual TT (s)" if idx==0 else "", value=lot_data["tt"], key=f"{lot_key}_tt_{lot_data['id']}", disabled=not is_active, label_visibility="visible" if idx==0 else "collapsed")
        if is_active:
            with cols[4]:
                if idx == 0: st.markdown("<div style='margin-top: 28px;'></div>", unsafe_allow_html=True)
                if len(lots) > 1 and st.button("🗑️", key=f"del_{lot_key}_{lot_data['id']}"):
                    st.session_state.tp_data[lot_key] = [l for l in lots if l['id'] != lot_data['id']]
                    st.rerun()
        
        lots[idx]["failure_summary"] = st.text_area(
            "📝 Failure Summary / Bin Analysis (Optional)", 
            value=lot_data.get("failure_summary", ""), 
            key=f"{lot_key}_fail_{lot_data['id']}", 
            disabled=not is_active, 
            height=68
        )
        st.write("")

    if is_active and st.button("➕ Add Lot", key=f"add_{lot_key}"):
        st.session_state.tp_data[lot_key].append({"id": str(uuid.uuid4()), "lot": "", "qty": 0, "yield": default_yield, "tt": default_tt, "failure_summary": ""})
        st.rerun()

# ==========================================
# 4. UI Header
# ==========================================
col_title, col_reset = st.columns([5, 1])
with col_title: st.title("🛡️ T-PASS (Test Program Approval & Sign-off System)")
with col_reset:
    st.write("") 
    if st.button("🔄 Reset Demo State", use_container_width=True):
        st.session_state.clear(); st.rerun()
st.markdown("---")

# ==========================================
# 5. Vertical Split Layout & Report Variables
# ==========================================

# 建立全局變數供 Buganizer Report 使用
proj_val = st.session_state.tp_data.get("project_name", "[Project]")
build_val = st.session_state.tp_data.get("build_phase", "[Build Phase]")

programs = st.session_state.tp_data.get("programs", [])
if programs:
    rev_list = [p["revision"] for p in programs if p["revision"]]
    rev_val = "/".join(rev_list) if rev_list else "[Program Revision]"
else:
    rev_val = "[Program Revision]"
    
REPORT_PREFIX = f"Update {proj_val} {build_val} {rev_val}"

col_left, col_right = st.columns([1, 5])

# ----------------------------------------
# [Left Panel]: Parallel Task (MES Setup)
# ----------------------------------------
with col_left:
    st.markdown("#### ⚙️ Parallel Task")
    st.caption("Runs throughout the entire process")
    if not st.session_state.mes_setup_done: st.button("⏳ KYEC MES Setup", key="mes_btn", type="primary", use_container_width=True)
    else: st.button("✅ KYEC MES Setup", key="mes_btn", use_container_width=True)
        
    est_mes = st.session_state.start_date + datetime.timedelta(days=3) if st.session_state.start_date else None
    render_date_timeline(est_mes, st.session_state.mes_setup_date if st.session_state.mes_setup_done else None)
        
    with st.container(border=True, height=687):
        st.markdown("**MES Setup**")
        st.caption("👤 POC: CSO-OPS")
        if not st.session_state.start_date:
            st.info("💡 Please complete the first phase **PDTE Release** to establish the project start date before proceeding with MES setup.")
        elif not st.session_state.mes_setup_done:
            st.markdown("⏳ **Status: Pending**")
            selected_date = st.date_input("📝 Actual Completion Date:", value=None, min_value=st.session_state.start_date, key="mes_date_input")
            comments = st.text_input("Remarks (Optional):", placeholder="e.g., FT1 routing enabled")
            
            err_msg = st.empty() 
            if st.button("✔️ Mark as Ready", use_container_width=True, type="primary"):
                if not selected_date: err_msg.error("⚠️ Please select the actual completion date!")
                else:
                    st.session_state.mes_setup_done = True; st.session_state.mes_setup_date = selected_date; st.session_state.mes_comments = comments; st.rerun()
        else:
            st.markdown("✅ **Status: Ready**")
            st.write(f"**Completion Date:** `{st.session_state.mes_setup_date}`")
            if st.session_state.mes_comments: st.write(f"**Remarks:** {st.session_state.mes_comments}")
            
            st.write("")
            st.markdown("📋 **Copy to Buganizer**")
            mes_date_str = st.session_state.mes_setup_date.strftime('%Y/%m/%d')
            mes_report = f"""{REPORT_PREFIX} KYEC MES Setup.
Status: 🟢 MES Setup Ready

[MES Setup Details]
- Completion Date: {mes_date_str}
- Remarks: {st.session_state.mes_comments if st.session_state.mes_comments else 'None'}"""
            st.code(mes_report, language="text")

# ----------------------------------------
# [Right Panel]: Linear Program Validation SOP
# ----------------------------------------
with col_right:
    st.markdown("#### 🚦 Program Validation SOP")
    st.caption("Proceed linearly according to the process sequence")
    current_template = FLOW_CONFIG[selected_flow]
    
    sop_btn_cols = st.columns(5)
    for i in range(5):
        phase_num = i + 1
        step_title = current_template[i]["title"].split(". ")[1]
        with sop_btn_cols[i]:
            if phase_num < st.session_state.current_phase:
                if st.button(f"✅ {step_title}", key=f"nav_{phase_num}", use_container_width=True): st.session_state.expanded_phase = phase_num
            elif phase_num == st.session_state.current_phase:
                if st.button(f"⏳ {step_title}", key=f"nav_{phase_num}", type="primary", use_container_width=True): st.session_state.expanded_phase = phase_num
            else:
                if st.session_state.current_phase > 5:
                    if st.button(f"✅ {step_title}", key=f"nav_{phase_num}_done", use_container_width=True): st.session_state.expanded_phase = phase_num
                else:
                    st.button(f"🔒 {step_title}", key=f"nav_{phase_num}", disabled=True, use_container_width=True)
                
            if st.session_state.start_date:
                est_date = st.session_state.start_date + datetime.timedelta(days=sum(step['sla_days'] for step in current_template[:i+1]))
                render_date_timeline(est_date, st.session_state.tp_data.get(f'exact_date_phase_{phase_num}'))
            else:
                render_date_timeline(None, None)

    active_idx = min(st.session_state.expanded_phase - 1, len(current_template) - 1)
    active_step_data = current_template[active_idx]
    form_type = active_step_data.get("form_type")
    
    is_active = (st.session_state.expanded_phase == st.session_state.current_phase) and (st.session_state.current_phase <= 5)
    phase_num = active_idx + 1

    with st.container(border=True):
        st.subheader(f"📍 {active_step_data['title']}")
        st.markdown(f"👤 **POC:** `{active_step_data['poc']}` &nbsp;|&nbsp; 📋 **Task:** {active_step_data['desc']}")
        # ✅ 使用超緊湊的分隔線取代原本的 st.divider()
        st.markdown(DIVIDER, unsafe_allow_html=True)
        
        # ==========================================
        # Dynamic Form Rendering
        # ==========================================
        if form_type == "init_form":
            build_options = ["Proto1P0", "Proto1P1", "EVT1P0", "EVT1P1", "DVT", "PVT", "MP"]
            saved_build = st.session_state.tp_data.get("build_phase", "PVT")
            build_idx = build_options.index(saved_build) if saved_build in build_options else 5

            c1, c2 = st.columns(2)
            with c1: project_name = st.text_input("Project", value=st.session_state.tp_data.get("project_name", "MBU(P26)"), disabled=not is_active)
            with c2: build_phase = st.selectbox("Build Stage", options=build_options, index=build_idx, disabled=not is_active)
            
            c3, c4 = st.columns(2)
            with c3: buganizer_link = st.text_input("Buganizer Link", value=st.session_state.tp_data.get("buganizer_link", ""), disabled=not is_active)
            with c4: lhn_link = st.text_input("LHN Link", value=st.session_state.tp_data.get("lhn_link", ""), disabled=not is_active)
            
            st.markdown("**🔹 Program Information**")
            render_dynamic_programs(is_active)
            
            comments = st.text_input("Comment (Special request)", value=st.session_state.tp_data.get("pdte_comments", ""), disabled=not is_active)
            
            # ✅ 使用超緊湊分隔線
            st.markdown(DIVIDER, unsafe_allow_html=True)
            c_date, c_action = st.columns([1, 2])
            with c_date:
                exact_date = st.date_input("📝 Actual Completion Date:", value=st.session_state.tp_data.get(f'exact_date_phase_1', datetime.date.today()), key=f"date_input_1", disabled=not is_active)
            
            if is_active:
                with c_action:
                    st.markdown("<div style='margin-top: 28px;'></div>", unsafe_allow_html=True)
                    err_msg = st.empty() 
                    if st.button("🚀 Submit Settings & Proceed", type="primary", use_container_width=True):
                        if not exact_date: err_msg.error("⚠️ Please select the actual completion date!")
                        else:
                            st.session_state.start_date = exact_date 
                            st.session_state.tp_data.update({
                                f'exact_date_phase_{phase_num}': exact_date, 
                                "project_name": project_name, 
                                "build_phase": build_phase,
                                "buganizer_link": buganizer_link, 
                                "lhn_link": lhn_link, 
                                "pdte_comments": comments
                            })
                            st.session_state.current_phase += 1; st.session_state.expanded_phase += 1; st.rerun()
            else:
                with c_action:
                    st.markdown("<div style='margin-top: 28px;'></div>", unsafe_allow_html=True)
                    st.success("✅ Program initiated successfully, project start date established.")
                
        elif form_type == "offline_check":
            st.markdown("**🔹 CORR Validation (Offline)**")
            render_dynamic_lots("corr_lots", active_step_data.get('target_offline'), is_active, default_yield=99.0)
            corr_data_link = st.text_input("🔗 CORR Data Link (Google Drive / URL)", value=st.session_state.tp_data.get("corr_data_link", ""), disabled=not is_active)

            st.markdown(DIVIDER, unsafe_allow_html=True)
            st.markdown("**🔹 Fuse Validation**")
            render_dynamic_lots("fuse_lots", active_step_data.get('target_fuse'), is_active, default_yield=100.0)
            fuse_data_link = st.text_input("🔗 Fuse Data Link (Google Drive / URL)", value=st.session_state.tp_data.get("fuse_data_link", ""), disabled=not is_active)
                
            st.write("")
            chk1 = st.checkbox("✅ Correlation validation completed", value=not is_active, disabled=not is_active)
            chk2 = st.checkbox("✅ OTP content correctness confirmed", value=not is_active, disabled=not is_active)

            st.markdown(DIVIDER, unsafe_allow_html=True)
            c_date, c_action = st.columns([1, 2])
            with c_date:
                exact_date = st.date_input("📝 Actual Completion Date:", value=st.session_state.tp_data.get(f'exact_date_phase_{phase_num}'), min_value=st.session_state.tp_data.get(f'exact_date_phase_{phase_num-1}'), key=f"date_input_{phase_num}", disabled=not is_active)

            if is_active:
                with c_action:
                    st.markdown("<div style='margin-top: 28px;'></div>", unsafe_allow_html=True)
                    err_msg = st.empty() 
                    if st.button("🚀 Submit Validation Data", type="primary", use_container_width=True):
                        if not exact_date: err_msg.error("⚠️ Please select the actual completion date!")
                        elif not (chk1 and chk2): err_msg.error("⚠️ Please ensure all validation items are checked!")
                        else:
                            st.session_state.tp_data.update({f'exact_date_phase_{phase_num}': exact_date, "corr_data_link": corr_data_link, "fuse_data_link": fuse_data_link})
                            st.session_state.current_phase += 1; st.session_state.expanded_phase += 1; st.rerun()
            else:
                with c_action:
                    st.markdown("<div style='margin-top: 28px;'></div>", unsafe_allow_html=True)
                    st.success("✅ Offline validation data submitted.")
                
                corr_details, corr_qty, corr_yld, corr_tt = format_lots_for_report(st.session_state.tp_data.get("corr_lots", []))
                fuse_details, fuse_qty, fuse_yld, fuse_tt = format_lots_for_report(st.session_state.tp_data.get("fuse_lots", []))
                fuse_status = "PASS" if fuse_yld >= 99.0 else f"{fuse_yld:.2f}%"

                st.markdown("📋 **Copy to Buganizer (Auto-generated Report)**")
                report = f"""{REPORT_PREFIX} Offline Trial data.
Status: 🟡 Offline Validation Data Submitted (Pending PE Review)

1). [Offline] {corr_qty}ea FT1 CORR TP trial run(w/o fuse), Final yield: {corr_yld:.2f}%
-------------------------------------------------------------------------
{corr_details}
-------------------------------------------------------------------------

2). [Offline] {fuse_qty}ea FT1 Prod TP trial run (fusing verification) ---> {fuse_status}
{fuse_details}

[Test Time]
FT1 CORR TP Avg. TT ~= {corr_tt:.1f} sec
FT2 PROD TP Avg. TT ~= {fuse_tt:.1f} sec

[Data Link]
1). CORR Data: {st.session_state.tp_data.get('corr_data_link', 'None')}
2). Fuse Data: {st.session_state.tp_data.get('fuse_data_link', 'None')}

[Next Step]
CSO-PE to review data and sign-off."""
                st.code(report, language="text")

        elif form_type == "pe_signoff":
            review_comment = st.text_input("📝 Comment (Special request) for next phase Trial Run:", value=st.session_state.tp_data.get("review_comment", ""), disabled=not is_active)
            
            if is_active:
                st.write("")
                if 'pe_approved' not in st.session_state: st.session_state.pe_approved = False
                if not st.session_state.pe_approved:
                    if st.button("✔️ CSO-PE Approval Confirmed", type="primary"): st.session_state.pe_approved = True; st.rerun()
                else: 
                    st.success("✅ CSO-PE Approved")
                    st.markdown(DIVIDER, unsafe_allow_html=True)
                    c_date, c_action = st.columns([1, 2])
                    with c_date:
                        exact_date = st.date_input("📝 Actual Completion Date:", value=st.session_state.tp_data.get(f'exact_date_phase_{phase_num}'), min_value=st.session_state.tp_data.get(f'exact_date_phase_{phase_num-1}'), key=f"date_input_{phase_num}")
                    with c_action:
                        st.markdown("<div style='margin-top: 28px;'></div>", unsafe_allow_html=True)
                        err_msg = st.empty() 
                        if st.button("🚀 Sign-off Completed, Proceed", type="primary", use_container_width=True):
                            if not exact_date: err_msg.error("⚠️ Please select the actual completion date!")
                            else:
                                st.session_state.tp_data.update({'review_comment': review_comment, f'exact_date_phase_{phase_num}': exact_date})
                                st.session_state.current_phase += 1; st.session_state.expanded_phase += 1; st.rerun()
            else:
                st.markdown(DIVIDER, unsafe_allow_html=True)
                c_date, c_action = st.columns([1, 2])
                with c_date:
                    exact_date = st.date_input("📝 Actual Completion Date:", value=st.session_state.tp_data.get(f'exact_date_phase_{phase_num}'), disabled=True, key=f"date_input_{phase_num}")
                with c_action:
                    st.markdown("<div style='margin-top: 28px;'></div>", unsafe_allow_html=True)
                    st.success("✅ CSO-PE sign-off completed.")
                
                st.markdown("📋 **Copy to Buganizer (Auto-generated Report)**")
                report = f"""{REPORT_PREFIX} Offline Data Review & Sign-off.
Status: 🟢 Offline Verification Approved by PE

[PE Comment / Request]
{st.session_state.tp_data.get('review_comment', 'None')}

[Next Step]
CSO-OPS to proceed with Trial Run."""
                st.code(report, language="text")

        elif form_type == "trial_data":
            if st.session_state.tp_data.get('review_comment'): st.info(f"💡 **PE Note/Reminder:** {st.session_state.tp_data.get('review_comment')}")
            render_dynamic_lots("trial_lots", active_step_data.get('target_trial'), is_active, default_yield=99.5)
            trial_data_link = st.text_input("🔗 Trial Run Data Link (Google Drive / URL)", value=st.session_state.tp_data.get("trial_data_link", ""), disabled=not is_active)
            
            st.markdown(DIVIDER, unsafe_allow_html=True)
            c_date, c_action = st.columns([1, 2])
            with c_date:
                exact_date = st.date_input("📝 Actual Completion Date:", value=st.session_state.tp_data.get(f'exact_date_phase_{phase_num}'), min_value=st.session_state.tp_data.get(f'exact_date_phase_{phase_num-1}'), key=f"date_input_{phase_num}", disabled=not is_active)

            if is_active:
                with c_action:
                    st.markdown("<div style='margin-top: 28px;'></div>", unsafe_allow_html=True)
                    err_msg = st.empty() 
                    if st.button("🚀 Submit Trial Run Data", type="primary", use_container_width=True):
                        if not exact_date: err_msg.error("⚠️ Please select the actual completion date!")
                        else:
                            st.session_state.tp_data.update({f'exact_date_phase_{phase_num}': exact_date, "trial_data_link": trial_data_link})
                            st.session_state.current_phase += 1; st.session_state.expanded_phase += 1; st.rerun()
            else:
                with c_action:
                    st.markdown("<div style='margin-top: 28px;'></div>", unsafe_allow_html=True)
                    st.success("✅ Trial run data submitted.")
                    
                trial_details, trial_qty, trial_yld, trial_tt = format_lots_for_report(st.session_state.tp_data.get("trial_lots", []))
                st.markdown("📋 **Copy to Buganizer (Auto-generated Report)**")
                report = f"""{REPORT_PREFIX} Trial Run Data Submitted.
Status: 🟡 Waiting for Official Release

1). [{active_step_data.get('trial_mode', 'Offline')}] {trial_qty}ea TP trial run, Final yield: {trial_yld:.2f}%
-------------------------------------------------------------------------
{trial_details}
-------------------------------------------------------------------------

[Test Time]
Trial Run Avg. TT ~= {trial_tt:.1f} sec

[Data Link]
1). Trial Run Data: {st.session_state.tp_data.get('trial_data_link', 'None')}

[Next Step]
CSO-PE to do final data review and proceed with Official Release."""
                st.code(report, language="text")

        elif form_type == "final_release":
            if is_active:
                if not st.session_state.mes_setup_done:
                    st.error("🚫 **Interlock**: Must complete the 'KYEC MES Setup' parallel task on the left before final release!")
                    st.markdown(DIVIDER, unsafe_allow_html=True)
                    c_date, c_action = st.columns([1, 2])
                    with c_date:
                        exact_date = st.date_input("📝 Actual Completion Date:", disabled=True, key=f"date_input_{phase_num}")
                    with c_action:
                        st.markdown("<div style='margin-top: 28px;'></div>", unsafe_allow_html=True)
                        st.button("🚀 Official Release to OSAT", type="primary", disabled=True, use_container_width=True)
                else:
                    st.info("✅ KYEC MES Setup is ready. Proceed with final release.")
                    st.markdown(DIVIDER, unsafe_allow_html=True)
                    c_date, c_action = st.columns([1, 2])
                    with c_date:
                        exact_date = st.date_input("📝 Actual Completion Date:", value=st.session_state.tp_data.get(f'exact_date_phase_{phase_num}'), min_value=st.session_state.tp_data.get(f'exact_date_phase_{phase_num-1}'), key=f"date_input_{phase_num}")
                    
                    with c_action:
                        st.markdown("<div style='margin-top: 28px;'></div>", unsafe_allow_html=True)
                        err_msg = st.empty() 
                        if st.button("🚀 Official Release to OSAT", type="primary", use_container_width=True):
                            if not exact_date: err_msg.error("⚠️ Please select the actual completion date!")
                            else:
                                st.session_state.tp_data[f'exact_date_phase_{phase_num}'] = exact_date
                                st.session_state.current_phase += 1; st.session_state.expanded_phase += 1; st.rerun()
            else:
                st.markdown(DIVIDER, unsafe_allow_html=True)
                c_date, c_action = st.columns([1, 2])
                with c_date:
                    exact_date = st.date_input("📝 Actual Completion Date:", value=st.session_state.tp_data.get(f'exact_date_phase_{phase_num}'), disabled=True, key=f"date_input_{phase_num}")
                with c_action:
                    st.markdown("<div style='margin-top: 28px;'></div>", unsafe_allow_html=True)
                    st.success("🎉 Program officially released.")
                
                st.markdown("📋 **Copy to Buganizer (Auto-generated Report)**")
                mes_date = st.session_state.mes_setup_date.strftime('%Y/%m/%d') if st.session_state.mes_setup_date else 'N/A'
                report = f"""{REPORT_PREFIX} Official Release.
Status: 🟢 CLOSED / RELEASED TO OSAT

[MES Setup Status]
- Setup Completed Date: {mes_date}
- Remarks: {st.session_state.mes_comments if st.session_state.mes_comments else 'None'}

[Conclusion]
Test program validation and KYEC MES setup are fully completed.
Officially released for mass production."""
                st.code(report, language="text")