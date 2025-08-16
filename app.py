import streamlit as st
import pandas as pd
import pdfkit
from docx import Document
from docx.enum.section import WD_ORIENT
from docx.shared import Pt, Mm
from num2words import num2words
import os
import zipfile
import tempfile
from jinja2 import Environment, FileSystemLoader
from pypdf import PdfReader, PdfWriter
import numpy as np
import platform
from datetime import datetime
import subprocess
import shutil
import traceback
import argparse
import requests
from xhtml2pdf import pisa
import tarfile
import io

# Page setup and branding
_local_logo = os.path.join(os.getcwd(), "crane_rajkumar.png")
_page_icon = _local_logo if os.path.exists(_local_logo) else "📄"
st.set_page_config(page_title="Bill Generator", page_icon=_page_icon, layout="wide")

def resolve_logo_url() -> str | None:
    candidates = [
        # Prefer local crane_rajkumar.png if present
        _local_logo if os.path.exists(_local_logo) else None,
        "https://raw.githubusercontent.com/CRAJKUMARSINGH/Priyanka_TenderV01/HEAD/logo.png",
        "https://raw.githubusercontent.com/CRAJKUMARSINGH/Priyanka_TenderV01/HEAD/crane_rajkumar.png",
        "https://raw.githubusercontent.com/CRAJKUMARSINGH/Priyanka_TenderV01/HEAD/assets/logo.png",
        "https://raw.githubusercontent.com/CRAJKUMARSINGH/Priyanka_TenderV01/HEAD/assets/logo.jpg",
        "https://raw.githubusercontent.com/CRAJKUMARSINGH/Priyanka_TenderV01/HEAD/images/logo.png",
        "https://raw.githubusercontent.com/CRAJKUMARSINGH/Priyanka_TenderV01/HEAD/images/logo.jpg",
    ]
    for url in candidates:
        try:
            if url is None:
                continue
            if os.path.exists(url):
                return url
            r = requests.get(url, timeout=5)
            if r.status_code == 200 and r.headers.get("content-type", "").startswith("image"):
                return url
        except Exception:
            continue
    return None

_logo_url = resolve_logo_url()
if _logo_url:
    try:
        st.logo(_logo_url, size="large")
    except Exception:
        st.sidebar.image(_logo_url, use_container_width=True)

# Header layout for enhanced appearance
header_cols = st.columns([1, 6])
with header_cols[0]:
    if _logo_url and os.path.exists(_logo_url):
        st.image(_logo_url, use_container_width=True)
    elif _logo_url:
        st.image(_logo_url, width=64)
with header_cols[1]:
    st.markdown("""
    <div style="display:flex; align-items:center; gap:12px;">
      <div>
        <h2 style="margin:0;">Bill Generator</h2>
        <p style="margin:0; color:#666;">A4 documents with professional layout</p>
      </div>
    </div>
    """, unsafe_allow_html=True)

# Debug logging helpers
DEBUG_VERBOSE = os.getenv("BILL_VERBOSE", "0") == "1"

def _show_celebration_once():
    if "_celebrated" not in st.session_state:
        st.session_state["_celebrated"] = True
        try:
            st.balloons()
            st.snow()
        except Exception:
            pass

def _render_landing():
    # Minimal CSS to enhance appearance
    st.markdown(
        """
        <style>
        .hero {
          padding: 28px 24px; border-radius: 14px;
          background: linear-gradient(135deg, #f0f7ff 0%, #ffffff 70%);
          border: 1px solid #e6eef8;
          box-shadow: 0 6px 24px rgba(0,0,0,0.06);
        }
        .hero h1 { margin: 0 0 8px 0; font-size: 28px; }
        .hero p { margin: 0; color: #4a5568; }
        .cta { margin-top: 16px; color: #1f6feb; }
        .tip { font-size: 13px; color: #6b7280; margin-top: 6px; }
        </style>
        """,
        unsafe_allow_html=True,
    )
    hero_cols = st.columns([3, 4])
    with hero_cols[0]:
        st.markdown(
            """
            <div class="hero">
              <h1>Generate professional A4 Bills</h1>
              <p>Uniform 10 mm margins • HTML ↔ PDF ↔ DOCX consistency • LaTeX PDFs</p>
              <div class="tip">Upload your Excel to begin. Configure premium, download merged outputs.</div>
            </div>
            """,
            unsafe_allow_html=True,
        )
    with hero_cols[1]:
        banner_path = os.path.join(os.getcwd(), "landing_page.png")
        if os.path.exists(banner_path):
            st.image(banner_path, use_container_width=True)
        elif _logo_url and os.path.exists(_logo_url):
            st.image(_logo_url, use_container_width=True)
    _show_celebration_once()

def _log_debug(message: str) -> None:
    if DEBUG_VERBOSE:
        try:
            st.write(message)
        except Exception:
            pass

def _log_warn(message: str) -> None:
    if DEBUG_VERBOSE:
        try:
            st.warning(message)
        except Exception:
            pass

def _log_traceback() -> None:
    if DEBUG_VERBOSE:
        try:
            st.write(traceback.format_exc())
        except Exception:
            pass

# Create templates directory if it doesn't exist
os.makedirs("templates", exist_ok=True)

# Set up Jinja2 environment
env = Environment(loader=FileSystemLoader("templates"), cache_size=0)

# Temporary directory
TEMP_DIR = None

def get_temp_dir():
    """Get or create a temporary directory for this session."""
    global TEMP_DIR
    if TEMP_DIR is None or not os.path.exists(TEMP_DIR):
        TEMP_DIR = tempfile.mkdtemp()
    return TEMP_DIR

def ensure_wkhtmltopdf() -> str | None:
    """Ensure wkhtmltopdf is available. If missing, download a static linux tarball and extract to a temp dir. Returns path or None."""
    # 1) Already in PATH?
    path = shutil.which("wkhtmltopdf")
    if path:
        return path
    # 2) Environment override
    env_path = os.getenv("WKHTMLTOPDF_PATH")
    if env_path and os.path.exists(env_path):
        return env_path
    # 3) Try to download a static binary (linux generic amd64)
    # Known packaging release URL (0.12.6-1) with patched Qt
    urls = [
        "https://github.com/wkhtmltopdf/packaging/releases/download/0.12.6-1/wkhtmltox-0.12.6-1_linux-generic-amd64.tar.xz",
        "https://github.com/wkhtmltopdf/packaging/releases/download/0.12.6-1/wkhtmltox-0.12.6-1.centos8.x86_64.rpm.tar.xz"
    ]
    cache_dir = os.path.join(tempfile.gettempdir(), "wkhtmltopdf_bin")
    os.makedirs(cache_dir, exist_ok=True)
    for url in urls:
        try:
            resp = requests.get(url, timeout=30, stream=True)
            if resp.status_code != 200 or len(resp.content) < 1024:
                continue
            # Validate content size to prevent memory exhaustion
            content_length = resp.headers.get('content-length')
            if content_length and int(content_length) > 100 * 1024 * 1024:  # 100MB limit
                _log_warn(f"Download too large: {content_length} bytes")
                continue
            tar_bytes = resp.content
            # Extract from .tar.xz
            with tarfile.open(fileobj=io.BytesIO(tar_bytes), mode="r:xz") as tf:  # type: ignore
                members = tf.getmembers()
                # Look for wkhtmltox/bin/wkhtmltopdf
                target_member = None
                for m in members:
                    if m.name.endswith("/wkhtmltopdf") and "/bin/" in m.name:
                        target_member = m
                        break
                if not target_member:
                    continue
                # Security check: prevent path traversal
                if ".." in target_member.name or target_member.name.startswith("/"):
                    _log_warn(f"Suspicious path in archive: {target_member.name}")
                    continue
                tf.extract(target_member, path=cache_dir)
                extracted_path = os.path.join(cache_dir, target_member.name)
                # Make executable
                os.chmod(extracted_path, 0o755)
                return extracted_path
        except Exception:
            continue
    return None

# Configure wkhtmltopdf
wkhtmltopdf_exe = None
if platform.system() == "Windows":
    wkhtmltopdf_exe = r"C:\\Program Files\\wkhtmltopdf\\bin\\wkhtmltopdf.exe"
else:
    wkhtmltopdf_exe = ensure_wkhtmltopdf() or shutil.which("wkhtmltopdf")

try:
    config = pdfkit.configuration(wkhtmltopdf=wkhtmltopdf_exe) if wkhtmltopdf_exe else pdfkit.configuration()
except Exception:
    config = None

def number_to_words(number):
    try:
        return num2words(int(number), lang="en_IN").title()
    except:
        return str(number)

def process_bill(ws_wo, ws_bq, ws_extra, premium_percent, premium_type):
    _log_debug("Starting process_bill")
    
    # Initialize data structures
    data = {
        'header': [],
        'items': [],
        'work_order_items': [],
        'extra_items': [],
        'totals': {},
        'premium_percent': premium_percent,
        'premium_type': premium_type,
        'certificates': {},
        'deductions': {},
        'calculations': {}
    }
    
    from datetime import datetime, date

    # Header (A1:G19) only — matching actual data range
    header_data = ws_wo.iloc[:19, :7].replace(np.nan, "").values.tolist()

    # Ensure all dates are formatted as date-only strings
    for i in range(len(header_data)):
        for j in range(len(header_data[i])):
            val = header_data[i][j]
            if isinstance(val, (pd.Timestamp, datetime, date)):
                header_data[i][j] = val.strftime("%d-%m-%Y")

    # Extract key header information
    data['header'] = header_data
    data['contractor_name'] = str(header_data[5][0]) if len(header_data) > 5 else ""
    data['work_name'] = str(header_data[7][0]) if len(header_data) > 7 else ""
    data['agreement_no'] = str(header_data[12][0]) if len(header_data) > 12 else ""
    data['work_order_amount'] = str(header_data[18][0]) if len(header_data) > 18 else "0"
    
    # Extract dates from header
    data['commencement_date'] = str(header_data[13][0]) if len(header_data) > 13 else ""
    data['completion_date'] = str(header_data[15][0]) if len(header_data) > 15 else ""
    data['actual_completion_date'] = str(header_data[16][0]) if len(header_data) > 16 else ""

    # Process Work Order items
    work_order_total = 0
    last_row_wo = ws_wo.shape[0]
    for i in range(21, last_row_wo):
        if i >= ws_bq.shape[0]:
            continue
            
        qty_raw = ws_bq.iloc[i, 3] if i < ws_bq.shape[0] and pd.notnull(ws_bq.iloc[i, 3]) else 0
        rate_raw = ws_wo.iloc[i, 4] if pd.notnull(ws_wo.iloc[i, 4]) else 0

        qty = 0
        if isinstance(qty_raw, (int, float)):
            qty = float(qty_raw)
        elif isinstance(qty_raw, str):
            cleaned_qty = qty_raw.strip().replace(',', '').replace(' ', '')
            try:
                qty = float(cleaned_qty)
            except ValueError:
                qty = 0

        rate = 0
        if isinstance(rate_raw, (int, float)):
            rate = float(rate_raw)
        elif isinstance(rate_raw, str):
            cleaned_rate = rate_raw.strip().replace(',', '').replace(' ', '')
            try:
                rate = float(cleaned_rate)
            except ValueError:
                rate = 0

        amount = round(qty * rate) if qty and rate else 0
        work_order_total += amount

        item = {
            "serial_no": str(ws_wo.iloc[i, 0]) if pd.notnull(ws_wo.iloc[i, 0]) else "",
            "description": str(ws_wo.iloc[i, 1]) if pd.notnull(ws_wo.iloc[i, 1]) else "",
            "unit": str(ws_wo.iloc[i, 2]) if pd.notnull(ws_wo.iloc[i, 2]) else "",
            "quantity": qty,
            "rate": rate,
            "amount": amount,
            "remark": str(ws_wo.iloc[i, 6]) if pd.notnull(ws_wo.iloc[i, 6]) else ""
        }
        
        data['work_order_items'].append(item)
        data['items'].append(item)

    # Process Extra Items
    extra_items_total = 0
    last_row_extra = ws_extra.shape[0]
    for j in range(6, last_row_extra):
        qty_raw = ws_extra.iloc[j, 3] if pd.notnull(ws_extra.iloc[j, 3]) else 0
        rate_raw = ws_extra.iloc[j, 5] if pd.notnull(ws_extra.iloc[j, 5]) else 0

        qty = 0
        if isinstance(qty_raw, (int, float)):
            qty = float(qty_raw)
        elif isinstance(qty_raw, str):
            cleaned_qty = qty_raw.strip().replace(',', '').replace(' ', '')
            try:
                qty = float(cleaned_qty)
            except ValueError:
                qty = 0

        rate = 0
        if isinstance(rate_raw, (int, float)):
            rate = float(rate_raw)
        elif isinstance(rate_raw, str):
            cleaned_rate = rate_raw.strip().replace(',', '').replace(' ', '')
            try:
                rate = float(cleaned_rate)
            except ValueError:
                rate = 0

        amount = round(qty * rate) if qty and rate else 0
        extra_items_total += amount

        extra_item = {
            "serial_no": str(ws_extra.iloc[j, 0]) if pd.notnull(ws_extra.iloc[j, 0]) else "",
            "bsr_no": str(ws_extra.iloc[j, 1]) if pd.notnull(ws_extra.iloc[j, 1]) else "",
            "description": str(ws_extra.iloc[j, 2]) if pd.notnull(ws_extra.iloc[j, 2]) else "",
            "quantity": qty,
            "unit": str(ws_extra.iloc[j, 4]) if pd.notnull(ws_extra.iloc[j, 4]) else "",
            "rate": rate,
            "amount": amount,
            "remarks": str(ws_extra.iloc[j, 7]) if pd.notnull(ws_extra.iloc[j, 7]) else ""
        }
        
        data['extra_items'].append(extra_item)

    # Calculate totals and deductions
    grand_total = work_order_total + extra_items_total
    
    # Premium calculation
    if premium_type == "Addition":
        premium_amount = grand_total * (premium_percent / 100)
        total_with_premium = grand_total + premium_amount
    else:
        premium_amount = grand_total * (premium_percent / 100)
        total_with_premium = grand_total - premium_amount
    
    # Calculate deductions
    sd_amount = int(total_with_premium * 0.10)  # 10% Security Deposit
    it_amount = int(total_with_premium * 0.02)  # 2% Income Tax
    gst_amount = int(total_with_premium * 0.02)  # 2% GST
    lc_amount = int(total_with_premium * 0.01)   # 1% Labour Cess
    
    total_deductions = sd_amount + it_amount + gst_amount + lc_amount
    net_payable = int(total_with_premium - total_deductions)
    
    # Store calculations
    data['totals'] = {
        'work_order_total': int(work_order_total),
        'extra_items_total': int(extra_items_total),
        'grand_total': int(grand_total),
        'premium_amount': int(premium_amount),
        'total_with_premium': int(total_with_premium),
        'net_payable': net_payable
    }
    
    data['deductions'] = {
        'sd_amount': sd_amount,
        'it_amount': it_amount,
        'gst_amount': gst_amount,
        'lc_amount': lc_amount,
        'total_deductions': total_deductions
    }
    
    data['calculations'] = {
        'payable_amount': net_payable,
        'amount_words': number_to_words(net_payable)
    }
    
    # Certificate data
    data['certificates'] = {
        'measurement_officer': "Site Engineer",
        'measurement_date': datetime.now().strftime("%d-%m-%Y"),
        'measurement_book_page': "04-20",
        'measurement_book_no': "887",
        'preparing_officer_name': "Site Engineer",
        'preparing_officer_date': datetime.now().strftime("%d-%m-%Y"),
        'authorizing_officer_name': "Executive Engineer",
        'authorizing_officer_date': datetime.now().strftime("%d-%m-%Y")
    }
    
    # Note sheet specific data
    work_order_amt = int(data.get('work_order_amount', '0').replace(',', '') if isinstance(data.get('work_order_amount', '0'), str) else data.get('work_order_amount', 0))
    progress_percentage = (total_with_premium / work_order_amt * 100) if work_order_amt > 0 else 0
    extra_item_percentage = (extra_items_total / work_order_amt * 100) if work_order_amt > 0 else 0
    
    data['note_sheet'] = {
        'progress_percentage': progress_percentage,
        'extra_item_status': "Yes" if extra_items_total > 0 else "No",
        'extra_item_percentage': extra_item_percentage,
        'approval_status': "under 5%, approval of the same is to be granted by this office" if extra_item_percentage < 5 else "more than 5% and Approval of the Deviation Case is required from the Superintending Engineer",
        'work_completion_note': "Work was completed in time.",
        'deviation_note': f"Requisite Deviation Statement is enclosed. The Overall Excess is {'more than 5% and Approval of the Deviation Case is required from the Superintending Engineer, PWD Electrical Circle, Udaipur' if progress_percentage > 105 else 'under acceptable limits'}."
    }
    
    _log_debug(f"Processed bill data: {len(data['items'])} items, total: {data['totals']['net_payable']}")
    return data

def generate_html_from_template(template_name, data):
    """Generate HTML content from template"""
    try:
        template = env.get_template(template_name)
        return template.render(data=data)
    except Exception as e:
        _log_warn(f"Error generating HTML from template {template_name}: {str(e)}")
        return f"<html><body><h1>Error in {template_name}</h1><p>{str(e)}</p></body></html>"

def create_pdf_from_html(html_content, filename):
    """Create PDF from HTML content"""
    try:
        if config is None:
            return None
        
        options = {
            'page-size': 'A4',
            'margin-top': '10mm',
            'margin-right': '10mm',
            'margin-bottom': '10mm',
            'margin-left': '10mm',
            'encoding': "UTF-8",
            'no-outline': None,
            'enable-local-file-access': None
        }
        
        temp_dir = get_temp_dir()
        pdf_path = os.path.join(temp_dir, filename)
        
        pdfkit.from_string(html_content, pdf_path, options=options, configuration=config)
        
        if os.path.exists(pdf_path):
            with open(pdf_path, 'rb') as f:
                return f.read()
    except Exception as e:
        _log_warn(f"Error creating PDF: {str(e)}")
        _log_traceback()
    return None

def create_docx_from_html(html_content, filename):
    """Create DOCX from HTML content (basic conversion)"""
    try:
        doc = Document()
        
        # Set margins to 10mm (converted to inches: 10mm ≈ 0.394 inches)
        sections = doc.sections
        for section in sections:
            section.top_margin = Mm(10)
            section.bottom_margin = Mm(10)
            section.left_margin = Mm(10)
            section.right_margin = Mm(10)
        
        # Simple HTML to DOCX conversion (basic text extraction)
        import re
        text_content = re.sub('<[^<]+?>', '', html_content)
        lines = text_content.split('\n')
        
        for line in lines:
            if line.strip():
                doc.add_paragraph(line.strip())
        
        temp_dir = get_temp_dir()
        docx_path = os.path.join(temp_dir, filename)
        doc.save(docx_path)
        
        if os.path.exists(docx_path):
            with open(docx_path, 'rb') as f:
                return f.read()
    except Exception as e:
        _log_warn(f"Error creating DOCX: {str(e)}")
        _log_traceback()
    return None

def main():
    st.title("Government Bill Generator")
    
    # File upload
    uploaded_file = st.file_uploader("Upload Excel file", type=['xlsx', 'xls'])
    
    if uploaded_file is None:
        _render_landing()
        return
    
    # Premium settings
    col1, col2 = st.columns(2)
    with col1:
        premium_percent = st.number_input("Premium Percentage", value=11.25, min_value=0.0, max_value=100.0)
    with col2:
        premium_type = st.selectbox("Premium Type", ["Addition", "Deduction"])
    
    if st.button("Generate Bill"):
        try:
            # Read Excel file
            excel_file = pd.ExcelFile(uploaded_file)
            sheet_names = excel_file.sheet_names
            
            # Read required sheets
            ws_wo = pd.read_excel(uploaded_file, sheet_name=sheet_names[0]) if len(sheet_names) > 0 else pd.DataFrame()
            ws_bq = pd.read_excel(uploaded_file, sheet_name=sheet_names[1]) if len(sheet_names) > 1 else ws_wo.copy()
            ws_extra = pd.read_excel(uploaded_file, sheet_name=sheet_names[2]) if len(sheet_names) > 2 else pd.DataFrame()
            
            # Process the bill
            with st.spinner("Processing bill data..."):
                data = process_bill(ws_wo, ws_bq, ws_extra, premium_percent, premium_type)
            
            # Generate documents
            templates = [
                ("first_page.html", "First_Page"),
                ("certificate_ii.html", "Certificate_II"),
                ("certificate_iii.html", "Certificate_III"),
                ("deviation_statement.html", "Deviation_Statement"),
                ("extra_items.html", "Extra_Items"),
                ("note_sheet.html", "Note_Sheet")
            ]
            
            st.success("Bill processed successfully!")
            
            # Create download section
            st.subheader("Download Generated Documents")
            
            for template_name, doc_name in templates:
                col1, col2 = st.columns(2)
                
                with col1:
                    if st.button(f"Generate PDF - {doc_name}"):
                        with st.spinner(f"Generating {doc_name} PDF..."):
                            html_content = generate_html_from_template(template_name, data)
                            pdf_content = create_pdf_from_html(html_content, f"{doc_name}.pdf")
                            
                            if pdf_content:
                                st.download_button(
                                    label=f"Download {doc_name} PDF",
                                    data=pdf_content,
                                    file_name=f"{doc_name}.pdf",
                                    mime="application/pdf"
                                )
                            else:
                                st.error(f"Failed to generate {doc_name} PDF")
                
                with col2:
                    if st.button(f"Generate DOCX - {doc_name}"):
                        with st.spinner(f"Generating {doc_name} DOCX..."):
                            html_content = generate_html_from_template(template_name, data)
                            docx_content = create_docx_from_html(html_content, f"{doc_name}.docx")
                            
                            if docx_content:
                                st.download_button(
                                    label=f"Download {doc_name} DOCX",
                                    data=docx_content,
                                    file_name=f"{doc_name}.docx",
                                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                                )
                            else:
                                st.error(f"Failed to generate {doc_name} DOCX")
            
            # Display summary
            with st.expander("Bill Summary", expanded=True):
                st.write(f"**Contractor:** {data.get('contractor_name', 'N/A')}")
                st.write(f"**Work:** {data.get('work_name', 'N/A')}")
                st.write(f"**Agreement No:** {data.get('agreement_no', 'N/A')}")
                st.write(f"**Work Order Amount:** ₹{data['totals']['work_order_total']:,}")
                st.write(f"**Extra Items:** ₹{data['totals']['extra_items_total']:,}")
                st.write(f"**Premium ({premium_percent}% {premium_type}):** ₹{data['totals']['premium_amount']:,}")
                st.write(f"**Total with Premium:** ₹{data['totals']['total_with_premium']:,}")
                st.write(f"**Total Deductions:** ₹{data['deductions']['total_deductions']:,}")
                st.write(f"**Net Payable:** ₹{data['totals']['net_payable']:,}")
                st.write(f"**Amount in Words:** {data['calculations']['amount_words']}")
        
        except Exception as e:
            st.error(f"Error processing file: {str(e)}")
            _log_traceback()

if __name__ == "__main__":
    main()
