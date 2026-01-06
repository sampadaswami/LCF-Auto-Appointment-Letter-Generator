import os
import io
import time
import zipfile
import re
import tempfile
import subprocess
import pandas as pd
import streamlit as st
from docxtpl import DocxTemplate

# -------------------- PREMIUM CUSTOM CSS --------------------
st.set_page_config(page_title="LCF Auto Appointment Letter Generator", page_icon="🏢", layout="wide")

st.markdown("""
<style>
    .block-container { max-width: 75% !important; padding-top: 1.5rem !important; margin: auto; }
    [data-testid="stMetric"] { background: #ffffff; padding: 25px; border-radius: 12px; border: 1px solid #eef2f6; box-shadow: 0 4px 6px rgba(0,0,0,0.02); text-align: center; }
    .lcf-title { color: #1e4d7b; font-size: 3.2rem !important; font-weight: 800; margin-bottom: 0px !important; padding-bottom: 0px !important; line-height: 1.1; letter-spacing: -1px; }
    .lcf-subtitle { color: #444; font-size: 1.4rem !important; font-weight: 600; margin-top: -5px !important; padding-top: 0px !important; }
    .help-box { background: #f8fbff; padding: 15px; border-radius: 8px; border-left: 5px solid #1e4d7b; margin-bottom: 20px; }
    .col-name { font-size: 0.85rem; color: #555; padding: 2px 0; }
    .footer { text-align: center; color: #aaa; font-size: 13px; margin-top: 60px; padding-top: 20px; border-top: 1px solid #eee; }
</style>
""", unsafe_allow_html=True)

# -------------------- CONFIG --------------------
REQUIRED_COLS = [
    "appointment_date", "employee_name", "employee_first_name", "employee_city", "posting_city", "joining_date", "designation", "center_name", "date_of_birth",
    "basic_monthly", "basic_annual", "hra_monthly", "hra_annual", "special_allowance_monthly", "special_allowance_annual", "mobile_allowance_monthly", "mobile_allowance_annual", "gross_monthly", "gross_annual",
    "epf_monthly", "epf_annual", "pt_monthly", "pt_annual", "total_deduction_monthly", "total_deduction_annual", "net_salary_monthly", "net_salary_annual", "employer_pf_monthly", "employer_pf_annual", "ctc_monthly", "ctc_annual"
]

# -------------------- HEADER --------------------
h_col1, h_col2 = st.columns([1.2, 4], vertical_alignment="center")

with h_col1:
    try:
        st.image("static/lcf_logo.jpg", width=220)
    except:
        st.info("🏢 LCF Logo")

with h_col2:
    st.markdown("<h1 class='lcf-title'>Lighthouse Communities Foundation</h1>", unsafe_allow_html=True)
    st.markdown("<p class='lcf-subtitle'>Auto Appointment Letter Generator</p>", unsafe_allow_html=True)
    st.caption("Designed By: Sampada Swami – HR Analytics Associate")

st.divider()

# -------------------- HELP SECTION --------------------
show_help = st.checkbox("🔍 View Required Data Headers")

if show_help:
    st.markdown("<div class='help-box'><strong>Ensure your Excel includes these exact column headers:</strong></div>", unsafe_allow_html=True)
    grid = st.columns(4)
    for i, col_name in enumerate(REQUIRED_COLS):
        grid[i % 4].markdown(f"<div class='col-name'>{col_name}</div>", unsafe_allow_html=True)
    st.divider()

# -------------------- UPLOAD SECTION --------------------
st.subheader("📤 Upload Files")
up_l, up_r = st.columns(2)
with up_l:
    ex_file = st.file_uploader("Upload Employee Data (Excel)", type=["xlsx"])
with up_r:
    tp_file = st.file_uploader("Upload Word Template (.docx)", type=["docx"])

if ex_file and tp_file:
    df = pd.read_excel(ex_file)
    missing = [c for c in REQUIRED_COLS if c not in df.columns]

    if missing:
        st.error(f"❌ Excel Header Error: Missing headers: {', '.join(missing)}")
    else:
        st.success(f"✅ Ready! {len(df)} records verified.")
        
        st.divider()
        st.subheader("⚙️ Settings")
        c1, c2 = st.columns(2)
        with c1:
            fn_p = st.text_input("Name Format", "{employee_name} Appointment Letter")
        with c2:
            mode = st.radio("Choose Format", ["Both (DOCX+PDF)", "DOCX Only", "PDF Only"], horizontal=True)

        if st.button("Generate Appointment Letters", type="primary", use_container_width=True):
            start_time = time.perf_counter()
            p_bar = st.progress(0)
            status = st.empty()
            success_count = 0
            pdf_success_count = 0
            audit_log = []
            
            with tempfile.TemporaryDirectory() as tmpdir:
                master_tp = os.path.join(tmpdir, "tpl.docx")
                with open(master_tp, "wb") as f: 
                    f.write(tp_file.getbuffer())
                
                dir_docx = os.path.join(tmpdir, "AppointmentLetter_Word")
                dir_pdf = os.path.join(tmpdir, "AppointmentLetter_PDF")
                os.makedirs(dir_docx, exist_ok=True)
                os.makedirs(dir_pdf, exist_ok=True)

                for i, (_, row) in enumerate(df.iterrows(), start=1):
                    try:
                        # Prepare context
                        ctx = {c: str(row[c]).strip() if pd.notna(row[c]) else "" for c in df.columns}
                        
                        # Format dates
                        for d in ["appointment_date", "joining_date", "date_of_birth"]:
                            if ctx.get(d): 
                                ctx[d] = pd.to_datetime(row[d]).strftime("%d-%m-%Y")
                        
                        # Generate safe filename
                        fname = re.sub(r'[<>:"/\\|?*]', '', fn_p.replace("{employee_name}", str(row["employee_name"]))).strip()
                        out_path = os.path.join(dir_docx, f"{fname}.docx")
                        
                        # Create DOCX
                        doc = DocxTemplate(master_tp)
                        doc.render(ctx)
                        doc.save(out_path)
                        
                        success_count += 1
                        status_msg = "Success"
                        
                        # Convert to PDF if requested
                        if "PDF" in mode:
                            try:
                                pdf_path = os.path.join(dir_pdf, f"{fname}.pdf")
                                
                                # Try LibreOffice conversion
                                result = subprocess.run(
                                    ['soffice', '--headless', '--convert-to', 'pdf', '--outdir', dir_pdf, out_path],
                                    capture_output=True,
                                    timeout=30
                                )
                                
                                # Check if PDF was created
                                if os.path.exists(pdf_path):
                                    pdf_success_count += 1
                                    if mode == "PDF Only": 
                                        os.remove(out_path)
                                    status_msg = "Success (PDF Created)"
                                else:
                                    status_msg = "DOCX Created (PDF Failed - LibreOffice may not be installed)"
                                    
                            except subprocess.TimeoutExpired:
                                status_msg = "DOCX Created (PDF Timeout)"
                            except Exception as pdf_err:
                                status_msg = f"DOCX Created (PDF Error: {str(pdf_err)[:50]})"
                        
                        audit_log.append({"Name": row["employee_name"], "Status": status_msg})
                        
                    except Exception as e:
                        audit_log.append({"Name": row["employee_name"], "Status": f"Error: {str(e)[:80]}"})
                    
                    p_bar.progress(i / len(df))
                    status.text(f"Processing... {row['employee_name']}")

                end_time = time.perf_counter()
                time_taken = round(end_time - start_time, 2)

                # Summary Dashboard
                st.subheader("📊 Execution Summary")
                res1, res2, res3, res4 = st.columns(4)
                res1.metric("Rows Processed", len(df))
                res2.metric("DOCX Generated", success_count)
                res3.metric("PDF Generated", pdf_success_count if "PDF" in mode else "N/A")
                res4.metric("Time Taken", f"{time_taken}s")

                # Create ZIP
                zip_io = io.BytesIO()
                with zipfile.ZipFile(zip_io, "w") as zf:
                    # Add DOCX files
                    if os.path.exists(dir_docx):
                        for f in os.listdir(dir_docx):
                            zf.write(os.path.join(dir_docx, f), f"AppointmentLetter_Word/{f}")
                    
                    # Add PDF files
                    if os.path.exists(dir_pdf):
                        for f in os.listdir(dir_pdf):
                            zf.write(os.path.join(dir_pdf, f), f"AppointmentLetter_PDF/{f}")
                    
                    # Add Audit Report
                    audit_io = io.BytesIO()
                    pd.DataFrame(audit_log).to_excel(audit_io, index=False)
                    zf.writestr("Audit_Report.xlsx", audit_io.getvalue())

                zip_io.seek(0)
                st.download_button("⬇ Download ZIP", zip_io.getvalue(), "Appointment_Letters.zip", type="primary", use_container_width=True)

else:
    st.info("👋 Welcome! Please upload your employee data and word template to begin.")

st.markdown("<div class='footer'>© 2026 LCF | Sampada Swami</div>", unsafe_allow_html=True)
