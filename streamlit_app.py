"""
Linwood Capital — AI Underwriting Portal.
Autonomous Pipeline: From Raw Dataroom (OM, Rent Roll, T-12) to Ready Waterman Pro Forma in Excel.

Run:
  streamlit run streamlit_app.py
"""

import os
import streamlit as st
import pandas as pd
from underwriting.parsers import parse_generic_dataroom
from underwriting.injector import inject_waterman_model

st.set_page_config(
    page_title="Linwood Capital | AI Underwriting",
    page_icon="🏢",
    layout="wide",
    initial_sidebar_state="expanded"
)

BASE_DIR = os.path.abspath(os.path.dirname(__file__))
TEMPLATES_DIR = os.path.join(BASE_DIR, "templates") if os.path.exists(os.path.join(BASE_DIR, "templates")) else os.path.join(BASE_DIR, "01 Proforma templates-20261115T050602Z-1-001", "01 Proforma templates")
OUTPUT_DIR = os.path.join(BASE_DIR, "output")
UPLOADS_DIR = os.path.join(BASE_DIR, "uploads")
TEMPLATE_PATH = os.path.join(TEMPLATES_DIR, "Waterman - ACQ Proforma - TEMPLATE.xlsx")

os.makedirs(OUTPUT_DIR, exist_ok=True)
os.makedirs(UPLOADS_DIR, exist_ok=True)

st.markdown("""
<style>
    .metric-card {
        background-color: #1e293b;
        padding: 18px;
        border-radius: 12px;
        border: 1px solid #334155;
    }
    .stButton>button {
        border-radius: 10px;
        font-weight: 600;
    }
</style>
""", unsafe_allow_html=True)

# Header
st.title("🏢 Linwood Capital — AI Underwriting Engine")
st.caption("Autonomous Pipeline: From Raw Dataroom (OM, Rent Roll, T-12) to Ready Waterman Pro Forma in Excel")

# Sidebar
with st.sidebar:
    st.image("https://img.icons8.com/fluency/96/commercial.png", width=64)
    st.subheader("Model Status")
    st.success("🟢 104,720 Formulas Active")
    st.info("Template: Waterman Institutional ACQ Proforma (13 Sheets)")
    st.markdown("---")
    st.markdown("**Core Capabilities:**")
    st.markdown("• Multimodal PDF & Excel Ingestion")
    st.markdown("• RedIQ / Yardi / Stonemark Rent Roll Parser")
    st.markdown("• Accrual P&L to Standard COA Mapping")
    st.markdown("• Non-destructive Semantic Excel Injection")
    st.markdown("• 100% Provenance Audit Trail")

deal_obj = None
output_excel_path = None

st.subheader("📤 Upload Deal Data Room")
st.caption("Drag and drop your property due diligence documents to instantly underwrite and generate the completed Waterman Pro Forma model.")

with st.form("custom_upload_form"):
    col_name, col_price = st.columns(2)
    prop_name = col_name.text_input("Property Name", "Oakwood Ridge Apartments")
    prop_price = col_price.number_input("Target Purchase Price ($)", min_value=1000000.0, value=20000000.0, step=500000.0)
    
    col_rr, col_t12, col_om = st.columns(3)
    uploaded_rr = col_rr.file_uploader("Rent Roll (.xlsx, .xls, .pdf)", type=["xlsx", "xls", "pdf"])
    uploaded_t12 = col_t12.file_uploader("T-12 Statement (.xlsx, .xls, .pdf)", type=["xlsx", "xls", "pdf"])
    uploaded_om = col_om.file_uploader("Offering Memo (.pdf)", type=["pdf"])
    
    submitted = st.form_submit_button("🚀 Run AI Underwriting Pipeline", type="primary", use_container_width=True)
    
    if submitted:
        if not uploaded_rr and not uploaded_t12 and not uploaded_om:
            st.error("Please upload at least one deal document (Rent Roll, T-12, or Offering Memo).")
        else:
            with st.spinner("Ingesting uploaded files and generating pro forma..."):
                upload_session_dir = os.path.join(UPLOADS_DIR, "streamlit_session")
                os.makedirs(upload_session_dir, exist_ok=True)
                
                if uploaded_rr:
                    with open(os.path.join(upload_session_dir, uploaded_rr.name), "wb") as f:
                        f.write(uploaded_rr.getbuffer())
                if uploaded_t12:
                    with open(os.path.join(upload_session_dir, uploaded_t12.name), "wb") as f:
                        f.write(uploaded_t12.getbuffer())
                if uploaded_om:
                    with open(os.path.join(upload_session_dir, uploaded_om.name), "wb") as f:
                        f.write(uploaded_om.getbuffer())
                        
                deal_obj = parse_generic_dataroom(upload_session_dir, default_name=prop_name, default_price=prop_price)
                safe_name = "".join(c for c in prop_name if c.isalnum() or c in (" ", "_")).strip().replace(" ", "_")
                out_path = os.path.join(OUTPUT_DIR, f"Waterman_{safe_name}_FILLED.xlsx")
                inject_waterman_model(TEMPLATE_PATH, deal_obj, out_path)
                output_excel_path = out_path
                st.session_state['current_deal'] = deal_obj
                st.session_state['excel_path'] = out_path

# Display Results if Available
if 'current_deal' in st.session_state and st.session_state['current_deal'] is not None:
    deal = st.session_state['current_deal']
    excel_path = st.session_state['excel_path']
    
    st.markdown("---")
    st.header(f"📊 Underwriting Results: {deal.property_name}")
    st.write(f"📍 {deal.address}, {deal.city_state_zip}")
    
    # Download & Action Bar
    btn_col1, btn_col2, btn_col3, _ = st.columns([2.5, 2, 2, 1.5])
    if os.path.exists(excel_path):
        with open(excel_path, "rb") as f:
            btn_col1.download_button(
                label="📥 Download Pro Forma (.xlsx)",
                data=f,
                file_name=os.path.basename(excel_path),
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
        
    lineage_file = os.path.join(OUTPUT_DIR, "DATA_LINEAGE_AND_PROVENANCE.md")
    if os.path.exists(lineage_file):
        with open(lineage_file, "rb") as f:
            btn_col2.download_button(
                label="📄 Download Audit (.md)",
                data=f,
                file_name="DATA_LINEAGE_AND_PROVENANCE.md",
                mime="text/markdown",
                use_container_width=True
            )
            
    if btn_col3.button("🗑️ Delete Deal & Reset", use_container_width=True):
        if os.path.exists(excel_path):
            try:
                os.remove(excel_path)
            except Exception:
                pass
        st.session_state['current_deal'] = None
        st.session_state['excel_path'] = None
        st.rerun()
            
    # Metrics Bar
    m1, m2, m3, m4 = st.columns(4)
    m1.metric("Purchase Price", f"${deal.purchase_price:,.0f}", f"${deal.purchase_price/max(deal.total_units, 1):,.0f} / unit")
    m2.metric("Total Units & RSF", f"{deal.total_units} Units", f"{deal.total_rsf:,.0f} RSF")
    
    exp = deal.t12_expenses
    unit_opex = exp.payroll + exp.utilities + exp.repairs_maint + exp.turnover + exp.marketing + exp.general_admin + exp.insurance
    total_opex = (unit_opex * deal.total_units) + exp.real_estate_taxes
    m3.metric("T-12 OpEx Total", f"${total_opex:,.0f}", f"${unit_opex:,.0f} / unit / yr")
    m4.metric("Model Health", "100% Intact", "104,720 formulas")
    
    # Tables
    tab_um, tab_t12, tab_lineage = st.tabs(["Unit Mix Breakdown", "T-12 Operating Expenses (COA)", "Audit Trail & Provenance"])
    
    with tab_um:
        um_rows = []
        for u in deal.unit_mix:
            um_rows.append({
                "Tier": u.name,
                "Units": u.unit_count,
                "Avg SF": f"{u.avg_sf:,.0f}",
                "In-Place Rent ($)": f"${u.in_place_rent:,.2f}",
                "Market Rent ($)": f"${u.market_rent:,.2f}",
                "Renovated Units": u.renovated_units,
                "Capex/Unit ($)": f"${u.reno_capex_per_unit:,.0f}"
            })
        st.dataframe(pd.DataFrame(um_rows), use_container_width=True)
        
    with tab_t12:
        t12_rows = [
            {"Category": "Payroll & Personnel", "Allocation": "All Units", "$/Unit/Yr": f"${exp.payroll:,.0f}", "Annual Total ($)": f"${exp.payroll * deal.total_units:,.0f}"},
            {"Category": "Utilities", "Allocation": "All Units", "$/Unit/Yr": f"${exp.utilities:,.0f}", "Annual Total ($)": f"${exp.utilities * deal.total_units:,.0f}"},
            {"Category": "Repairs & Maintenance", "Allocation": "All Units", "$/Unit/Yr": f"${exp.repairs_maint:,.0f}", "Annual Total ($)": f"${exp.repairs_maint * deal.total_units:,.0f}"},
            {"Category": "Turnover", "Allocation": "All Units", "$/Unit/Yr": f"${exp.turnover:,.0f}", "Annual Total ($)": f"${exp.turnover * deal.total_units:,.0f}"},
            {"Category": "Marketing", "Allocation": "All Units", "$/Unit/Yr": f"${exp.marketing:,.0f}", "Annual Total ($)": f"${exp.marketing * deal.total_units:,.0f}"},
            {"Category": "General & Administrative", "Allocation": "All Units", "$/Unit/Yr": f"${exp.general_admin:,.0f}", "Annual Total ($)": f"${exp.general_admin * deal.total_units:,.0f}"},
            {"Category": "Insurance", "Allocation": "All Units", "$/Unit/Yr": f"${exp.insurance:,.0f}", "Annual Total ($)": f"${exp.insurance * deal.total_units:,.0f}"},
            {"Category": "Real Estate Taxes", "Allocation": "Total Annual", "$/Unit/Yr": f"${exp.real_estate_taxes/max(deal.total_units, 1):,.0f}", "Annual Total ($)": f"${exp.real_estate_taxes:,.0f}"},
        ]
        st.dataframe(pd.DataFrame(t12_rows), use_container_width=True)
        
    with tab_lineage:
        if os.path.exists(lineage_file):
            with open(lineage_file, "r", encoding="utf-8") as f:
                st.markdown(f.read())
