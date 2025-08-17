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
    # 3) Download and extract a Linux static binary
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
    first_page_data = {"header": [], "items": [], "totals": {}}
    last_page_data = {"payable_amount": 0, "amount_words": ""}
    deviation_data = {"items": [], "summary": {}}
    extra_items_data = {"items": []}
    note_sheet_data = {"notes": []}
    
    from datetime import datetime, date

    # Header (A1:G19) only — matching actual data range
    header_data = ws_wo.iloc[:19, :7].replace(np.nan, "").values.tolist()

    # Ensure all dates are formatted as date-only strings
    for i in range(len(header_data)):
        for j in range(len(header_data[i])):
            val = header_data[i][j]
            if isinstance(val, (pd.Timestamp, datetime, date)):
                header_data[i][j] = val.strftime("%d-%m-%Y")

    # Assign to first page
    first_page_data["header"] = header_data
    
    # Work Order items
    last_row_wo = ws_wo.shape[0]
    for i in range(21, last_row_wo):
        qty_raw = ws_bq.iloc[i, 3] if i < ws_bq.shape[0] and pd.notnull(ws_bq.iloc[i, 3]) else 0
        rate_raw = ws_wo.iloc[i, 4] if pd.notnull(ws_wo.iloc[i, 4]) else None

        qty = 0
        if isinstance(qty_raw, (int, float)):
            qty = float(qty_raw)
        elif isinstance(qty_raw, str):
            cleaned_qty = qty_raw.strip().replace(',', '').replace(' ', '')
            try:
                qty = float(cleaned_qty)
            except ValueError:
                _log_warn(f"Skipping invalid quantity at Bill Quantity row {i+1}: '{qty_raw}'")
                qty = 0

        rate = 0
        if isinstance(rate_raw, (int, float)):
            rate = float(rate_raw)
        elif isinstance(rate_raw, str):
            cleaned_rate = rate_raw.strip().replace(',', '').replace(' ', '')
            try:
                rate = float(cleaned_rate)
            except ValueError:
                _log_warn(f"Skipping invalid rate at Work Order row {i+1}: '{rate_raw}'")
                rate = 0

        item = {
            "serial_no": str(ws_wo.iloc[i, 0]) if pd.notnull(ws_wo.iloc[i, 0]) else "",
            "description": str(ws_wo.iloc[i, 1]) if pd.notnull(ws_wo.iloc[i, 1]) else "",
            "unit": str(ws_wo.iloc[i, 2]) if pd.notnull(ws_wo.iloc[i, 2]) else "",
            "quantity": qty,
            "rate": rate,
            "remark": str(ws_wo.iloc[i, 6]) if pd.notnull(ws_wo.iloc[i, 6]) else "",
            "amount": round(qty * rate) if qty and rate else 0,
            "is_divider": False
        }
        first_page_data["items"].append(item)

    # Extra Items divider
    first_page_data["items"].append({
        "description": "Extra Items (With Premium)",
        "bold": True,
        "underline": True,
        "amount": 0,
        "quantity": 0,
        "rate": 0,
        "serial_no": "",
        "unit": "",
        "remark": "",
        "is_divider": True
    })

    # Extra Items
    last_row_extra = ws_extra.shape[0]
    for j in range(6, last_row_extra):
        qty_raw = ws_extra.iloc[j, 3] if pd.notnull(ws_extra.iloc[j, 3]) else 0
        rate_raw = ws_extra.iloc[j, 5] if pd.notnull(ws_extra.iloc[j, 5]) else None

        qty = 0
        if isinstance(qty_raw, (int, float)):
            qty = float(qty_raw)
        elif isinstance(qty_raw, str):
            cleaned_qty = qty_raw.strip().replace(',', '').replace(' ', '')
            try:
                qty = float(cleaned_qty)
            except ValueError:
                _log_warn(f"Skipping invalid quantity at Extra Items row {j+1}: '{qty_raw}'")
                qty = 0

        rate = 0
        if isinstance(rate_raw, (int, float)):
            rate = float(rate_raw)
        elif isinstance(rate_raw, str):
            cleaned_rate = rate_raw.strip().replace(',', '').replace(' ', '')
            try:
                rate = float(cleaned_rate)
            except ValueError:
                _log_warn(f"Skipping invalid rate at Extra Items row {j+1}: '{rate_raw}'")
                rate = 0

        item = {
            "serial_no": str(ws_extra.iloc[j, 0]) if pd.notnull(ws_extra.iloc[j, 0]) else "",
            "description": str(ws_extra.iloc[j, 2]) if pd.notnull(ws_extra.iloc[j, 2]) else "",
            "unit": str(ws_extra.iloc[j, 4]) if pd.notnull(ws_extra.iloc[j, 4]) else "",
            "quantity": qty,
            "rate": rate,
            "remark": str(ws_extra.iloc[j, 1]) if pd.notnull(ws_extra.iloc[j, 1]) else "",
            "amount": round(qty * rate) if qty and rate else 0,
            "is_divider": False
        }
        first_page_data["items"].append(item)
        extra_items_data["items"].append(item.copy())  # Copy for standalone Extra Items

    # Totals
    data_items = [item for item in first_page_data["items"] if not item.get("is_divider", False)]
    total_amount = round(sum(item.get("amount", 0) for item in data_items))
    premium_amount = round(total_amount * (premium_percent / 100) if premium_type == "above" else -total_amount * (premium_percent / 100))
    payable_amount = round(total_amount + premium_amount)

    first_page_data["totals"] = {
        "grand_total": total_amount,
        "premium": {"percent": premium_percent / 100, "type": premium_type, "amount": premium_amount},
        "payable": payable_amount
    }

    try:
        extra_items_start = next(i for i, item in enumerate(first_page_data["items"]) if item.get("description") == "Extra Items (With Premium)")
        extra_items = [item for item in first_page_data["items"][extra_items_start + 1:] if not item.get("is_divider", False)]
        extra_items_sum = round(sum(item.get("amount", 0) for item in extra_items))
        extra_items_premium = round(extra_items_sum * (premium_percent / 100) if premium_type == "above" else -extra_items_sum * (premium_percent / 100))
        first_page_data["totals"]["extra_items_sum"] = extra_items_sum + extra_items_premium
    except StopIteration:
        first_page_data["totals"]["extra_items_sum"] = 0

    # Last Page
    last_page_data = {"payable_amount": payable_amount, "amount_words": number_to_words(payable_amount)}

    # Deviation Statement
    work_order_total = 0
    executed_total = 0
    overall_excess = 0
    overall_saving = 0
    for i in range(21, last_row_wo):
        _log_debug(f"Processing deviation row {i+1}: wo_qty={ws_wo.iloc[i, 3]}, wo_rate={ws_wo.iloc[i, 4]}, bq_qty={ws_bq.iloc[i, 3] if i < ws_bq.shape[0] else 'N/A'}")
        qty_wo_raw = ws_wo.iloc[i, 3] if pd.notnull(ws_wo.iloc[i, 3]) else 0
        rate_raw = ws_wo.iloc[i, 4] if pd.notnull(ws_wo.iloc[i, 4]) else None
        qty_bill_raw = ws_bq.iloc[i, 3] if i < ws_bq.shape[0] and pd.notnull(ws_bq.iloc[i, 3]) else 0

        qty_wo = 0
        if isinstance(qty_wo_raw, (int, float)):
            qty_wo = float(qty_wo_raw)
        elif isinstance(qty_wo_raw, str):
            cleaned_qty_wo = qty_wo_raw.strip().replace(',', '').replace(' ', '')
            try:
                qty_wo = float(cleaned_qty_wo)
            except ValueError:
                _log_warn(f"Skipping invalid qty_wo at row {i+1}: '{qty_wo_raw}'")
                qty_wo = 0

        rate = 0
        if isinstance(rate_raw, (int, float)):
            rate = float(rate_raw)
        elif isinstance(rate_raw, str):
            cleaned_rate = rate_raw.strip().replace(',', '').replace(' ', '')
            try:
                rate = float(cleaned_rate)
            except ValueError:
                _log_warn(f"Skipping invalid rate at row {i+1}: '{rate_raw}'")
                rate = 0

        qty_bill = 0
        if isinstance(qty_bill_raw, (int, float)):
            qty_bill = float(qty_bill_raw)
        elif isinstance(qty_bill_raw, str):
            cleaned_qty_bill = qty_bill_raw.strip().replace(',', '').replace(' ', '')
            try:
                qty_bill = float(cleaned_qty_bill)
            except ValueError:
                _log_warn(f"Skipping invalid qty_bill at row {i+1}: '{qty_bill_raw}'")
                qty_bill = 0

        amt_wo = round(qty_wo * rate)
        amt_bill = round(qty_bill * rate)
        excess_qty = qty_bill - qty_wo if qty_bill > qty_wo else 0
        excess_amt = round(excess_qty * rate)
        saving_qty = qty_wo - qty_bill if qty_wo > qty_bill else 0
        saving_amt = round(saving_qty * rate)

        work_order_total += amt_wo
        executed_total += amt_bill
        overall_excess += excess_amt
        overall_saving += saving_amt

        item = {
            "serial_no": str(ws_wo.iloc[i, 0]) if pd.notnull(ws_wo.iloc[i, 0]) else "",
            "description": str(ws_wo.iloc[i, 1]) if pd.notnull(ws_wo.iloc[i, 1]) else "",
            "unit": str(ws_wo.iloc[i, 2]) if pd.notnull(ws_wo.iloc[i, 2]) else "",
            "qty_wo": qty_wo,
            "rate": rate,
            "amt_wo": amt_wo,
            "qty_bill": qty_bill,
            "amt_bill": amt_bill,
            "excess_qty": excess_qty,
            "excess_amt": excess_amt,
            "saving_qty": saving_qty,
            "saving_amt": saving_amt,
            "remark": str(ws_wo.iloc[i, 6]) if pd.notnull(ws_wo.iloc[i, 6]) else ""
        }
        deviation_data["items"].append(item)

    net_difference = round(executed_total - work_order_total)
    
    deviation_data["summary"] = {
        "work_order_total": round(work_order_total),
        "executed_total": round(executed_total),
        "overall_excess": round(overall_excess),
        "overall_saving": round(overall_saving),
        "premium": {"percent": premium_percent / 100, "type": premium_type},
        "net_difference": round(net_difference)
    }

    _log_debug("Prepared first_page_data items")
    _log_debug("Prepared extra_items_data items")
    _log_debug("Prepared deviation_data items")
    return first_page_data, last_page_data, deviation_data, extra_items_data, note_sheet_data

def generate_bill_notes(payable_amount, work_order_amount, extra_item_amount):
    percentage_work_done = float(payable_amount / work_order_amount * 100) if work_order_amount > 0 else 0
    serial_number = 1
    note = []
    note.append(f"{serial_number}. The work has been completed {percentage_work_done:.2f}% of the Work Order Amount.")
    serial_number += 1
    if percentage_work_done < 90:
        note.append(f"{serial_number}. The execution of work at final stage is less than 90%...")
        serial_number += 1
    elif percentage_work_done > 100 and percentage_work_done <= 105:
        note.append(f"{serial_number}. Requisite Deviation Statement is enclosed...")
        serial_number += 1
    elif percentage_work_done > 105:
        note.append(f"{serial_number}. Requisite Deviation Statement is enclosed...")
        serial_number += 1
    note.append(f"{serial_number}. Quality Control (QC) test reports attached.")
    serial_number += 1
    if extra_item_amount > 0:
        extra_item_percentage = float(extra_item_amount / work_order_amount * 100) if work_order_amount > 0 else 0
        if extra_item_percentage > 5:
            note.append(f"{serial_number}. The amount of Extra items is Rs. {extra_item_amount}...")
        else:
            note.append(f"{serial_number}. The amount of Extra items is Rs. {extra_item_amount}...")
        serial_number += 1
    note.append(f"{serial_number}. Please peruse above details for necessary decision-making.")
    note.append("")
    note.append("                                Premlata Jain")
    note.append("                               AAO- As Auditor")
    return {"notes": note}

def generate_pdf(sheet_name, data, orientation, output_path):
    _log_debug(f"Generating PDF for {sheet_name}")
    try:
        template = env.get_template(f"{sheet_name.lower().replace(' ', '_')}.html")
        html_content = template.render(data=data)
        # Save HTML alongside PDF for consistency checks/comparison
        try:
            html_path = output_path[:-4] + ".html" if output_path.lower().endswith(".pdf") else output_path + ".html"
            with open(html_path, "w", encoding="utf-8") as f:
                f.write(html_content)
        except Exception:
            pass
        options = {
            "page-size": "A4",
            "orientation": orientation,
            "margin-top": "10mm",
            "margin-bottom": "10mm",
            "margin-left": "10mm",
            "margin-right": "10mm",
            "print-media-type": None,
            "enable-local-file-access": None,
            "disable-smart-shrinking": None,
            "zoom": "1",
            "dpi": 300,
        }
        if config is None:
            _log_warn("wkhtmltopdf missing; using xhtml2pdf fallback.")
            try:
                # Adjust layout for xhtml2pdf to avoid narrow content due to mm widths
                fallback_html = html_content
                for mm in ("190mm", "277mm"):
                    fallback_html = fallback_html.replace(f"width: {mm}", "width: 100%")
                    fallback_html = fallback_html.replace(f"max-width: {mm}", "width: 100%")
                # Remove centered margin that can introduce side gaps
                fallback_html = fallback_html.replace("margin: 0 auto;", "margin: 0;")
                default_css = '@page { size: A4 %s; margin: 10mm; }' % ("landscape" if orientation=="landscape" else "portrait")
                with open(output_path, "wb") as pdf_file:
                    pisa.CreatePDF(src=fallback_html, dest=pdf_file, default_css=default_css)
            except Exception:
                _log_traceback()
        else:
            pdfkit.from_string(
                html_content,
                output_path,
                configuration=config,
                options=options
            )
        _log_debug(f"Finished PDF for {sheet_name}")
    except Exception as e:
        st.error(f"Error generating PDF for {sheet_name}: {str(e)}")
        _log_traceback()
        raise

def generate_latex_pdf(template_name, data, output_path):
    """Generate PDF from LaTeX template"""
    try:
        # Check if pdflatex is available
        if not shutil.which("pdflatex"):
            _log_warn("pdflatex not available for LaTeX PDF generation")
            return False
            
        template = env.get_template(f"{template_name}.tex")
        latex_content = template.render(data=data)
        
        temp_dir = get_temp_dir()
        tex_path = os.path.join(temp_dir, f"{template_name}.tex")
        
        with open(tex_path, "w", encoding="utf-8") as f:
            f.write(latex_content)
        
        # Run pdflatex
        result = subprocess.run(
            ["pdflatex", "-output-directory", temp_dir, tex_path],
            capture_output=True,
            text=True,
            timeout=30
        )
        
        if result.returncode == 0:
            pdf_source = os.path.join(temp_dir, f"{template_name}.pdf")
            if os.path.exists(pdf_source):
                shutil.copy2(pdf_source, output_path)
                return True
        else:
            _log_warn(f"LaTeX compilation failed: {result.stderr}")
            
    except Exception as e:
        _log_warn(f"LaTeX PDF generation error: {e}")
    
    return False

def create_combined_zip(templates_data):
    """Create a zip file with all generated documents"""
    temp_dir = get_temp_dir()
    zip_path = os.path.join(temp_dir, "Bill_Documents.zip")
    
    templates = [
        ("first_page", "First_Page"),
        ("certificate_ii", "Certificate_II"),
        ("certificate_iii", "Certificate_III"),
        ("deviation_statement", "Deviation_Statement"),
        ("extra_items", "Extra_Items"),
        ("note_sheet", "Note_Sheet")
    ]
    
    with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zip_file:
        for template_name, doc_name in templates:
            try:
                # Generate HTML
                html_template = env.get_template(f"{template_name}.html")
                html_content = html_template.render(data=templates_data[template_name])
                zip_file.writestr(f"{doc_name}.html", html_content)
                
                # Generate HTML-based PDF
                if config:
                    try:
                        pdf_path = os.path.join(temp_dir, f"{doc_name}_HTML.pdf")
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
                        pdfkit.from_string(html_content, pdf_path, options=options, configuration=config)
                        if os.path.exists(pdf_path):
                            with open(pdf_path, 'rb') as f:
                                zip_file.writestr(f"{doc_name}_HTML.pdf", f.read())
                    except Exception as e:
                        _log_warn(f"HTML PDF generation failed for {doc_name}: {e}")
                
                # Generate LaTeX-based PDF if available
                latex_pdf_path = os.path.join(temp_dir, f"{doc_name}_LaTeX.pdf")
                if generate_latex_pdf(template_name, templates_data[template_name], latex_pdf_path):
                    with open(latex_pdf_path, 'rb') as f:
                        zip_file.writestr(f"{doc_name}_LaTeX.pdf", f.read())
                
                # Generate LaTeX source
                if os.path.exists(f"templates/{template_name}.tex"):
                    latex_template = env.get_template(f"{template_name}.tex")
                    latex_content = latex_template.render(data=templates_data[template_name])
                    zip_file.writestr(f"{doc_name}.tex", latex_content)
                    
            except Exception as e:
                _log_warn(f"Error generating {doc_name}: {e}")
    
    return zip_path if os.path.exists(zip_path) else None

def main():
    st.title("Government Bill Generator")
    
    # File upload
    uploaded_file = st.file_uploader("Choose an Excel file", type=['xlsx', 'xls'])
    
    if uploaded_file is None:
        _render_landing()
        return
    
    # Premium settings
    col1, col2 = st.columns(2)
    with col1:
        premium_percent = st.number_input("Premium Percentage", value=11.25, min_value=0.0, max_value=100.0)
    with col2:
        premium_type = st.selectbox("Premium Type", ["Addition", "Deduction"])
    
    # Generate bill and capture results into session state
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
                first_page_data, last_page_data, deviation_data, extra_items_data, note_sheet_data = process_bill(ws_wo, ws_bq, ws_extra, premium_percent, premium_type)
            
            st.success("Bill processed successfully!")
            
<<<<<<< HEAD
            # Convert data for templates and store in session
            templates_data = {
                "first_page": {"bill_items": first_page_data["items"], "header": first_page_data["header"], "totals": first_page_data["totals"]},
                "certificate_ii": {"measurement_officer": "Site Engineer", "measurement_date": datetime.now().strftime("%d-%m-%Y"), "measurement_book_page": "04-20", "measurement_book_no": "887", "officer_name": "Site Engineer", "officer_designation": "Assistant Engineer", "authorising_officer_name": "Executive Engineer", "authorising_officer_designation": "Executive Engineer"},
                "certificate_iii": {"totals": first_page_data["totals"], "payable_words": last_page_data["amount_words"], "current_date": datetime.now()},
                "deviation_statement": deviation_data,
                "extra_items": extra_items_data,
                "note_sheet": generate_bill_notes(first_page_data["totals"]["payable"], 1000000, sum(item.get("amount", 0) for item in extra_items_data["items"]))
            }
            
            st.session_state["templates_data"] = templates_data
            
=======
            # Compute derived totals and deductions used by templates
            # Identify main vs extra items
            items = first_page_data["items"]
            divider_index = None
            for idx, it in enumerate(items):
                if it.get("is_divider") and it.get("description") == "Extra Items (With Premium)":
                    divider_index = idx
                    break
            main_items = items[:divider_index] if divider_index is not None else [it for it in items if not it.get("is_divider", False)]
            extra_items_only = items[divider_index + 1:] if divider_index is not None else []
            
            # Base totals
            bill_total = round(sum(it.get("amount", 0) for it in main_items))
            extra_items_base = round(sum(it.get("amount", 0) for it in extra_items_only))
            
            # Premium
            premium_fraction = float(premium_percent) / 100.0
            is_addition = str(premium_type).lower().startswith("a")  # Addition vs Deduction
            signed = 1 if is_addition else -1
            bill_premium = round(bill_total * premium_fraction * signed)
            extra_premium = round(extra_items_base * premium_fraction * signed)
            
            bill_grand_total = bill_total + bill_premium
            extra_items_total = extra_items_base + extra_premium
            total_with_premium = bill_grand_total + extra_items_total
            
            # Deductions (default policy)
            sd_pct, it_pct, gst_pct, lc_pct = 0.10, 0.02, 0.02, 0.01
            sd_amount = round(total_with_premium * sd_pct)
            it_amount = round(total_with_premium * it_pct)
            gst_amount = round(total_with_premium * gst_pct)
            lc_amount = round(total_with_premium * lc_pct)
            total_deductions = sd_amount + it_amount + gst_amount + lc_amount
            net_payable = max(total_with_premium - total_deductions, 0)
            
            expanded_totals = dict(first_page_data["totals"])
            expanded_totals.update({
                # Derived rollups used by several templates
                "total_with_premium": total_with_premium,
                "extra_items_total": extra_items_total,
                "net_payable": net_payable,
                # Duplicate deductions under totals for LaTeX templates expecting totals.sd_amount
                "sd_amount": sd_amount,
                "it_amount": it_amount,
                "gst_amount": gst_amount,
                "lc_amount": lc_amount,
                "total_deductions": total_deductions,
            })
            
            deductions = {
                "sd_amount": sd_amount,
                "it_amount": it_amount,
                "gst_amount": gst_amount,
                "lc_amount": lc_amount,
                "total_deductions": total_deductions,
            }
            
            # Placeholders for metadata required by some templates
            work_order_amount = 1000000
            progress_percentage = float(total_with_premium / work_order_amount * 100) if work_order_amount > 0 else 0.0
            extra_item_percentage = float(extra_items_total / work_order_amount * 100) if work_order_amount > 0 else 0.0
            note_sheet_meta = {
                "progress_percentage": progress_percentage,
                "deviation_note": "Requisite Deviation Statement is enclosed where applicable.",
                "work_completion_note": "Work executed as per specifications and contract.",
                "extra_item_percentage": extra_item_percentage,
                "approval_status": "Approval required" if extra_item_percentage > 5 else "Within limit",
                "extra_item_status": "Yes" if extra_items_total > 0 else "No",
            }
            
            templates_data = {
                # Keep existing shape for first_page; include expanded totals
                "first_page": {
                    "bill_items": first_page_data["items"],
                    "header": first_page_data["header"],
                    "totals": expanded_totals,
                    # Provide fields used by first_page.html and .tex
                    "bill_total": bill_total,
                    "bill_premium": bill_premium,
                    "bill_grand_total": bill_grand_total,
                    "extra_items": extra_items_data.get("items", []),
                    "extra_items_base": extra_items_base,
                    "extra_premium": extra_premium,
                    "extra_items_sum": extra_items_total,
                    "tender_premium_percent": premium_fraction,
                },
                "certificate_ii": {
                    "measurement_officer": "Site Engineer",
                    "measurement_date": datetime.now().strftime("%d-%m-%Y"),
                    "measurement_book_page": "04-20",
                    "measurement_book_no": "887",
                    "officer_name": "Site Engineer",
                    "officer_designation": "Assistant Engineer",
                    "authorising_officer_name": "Executive Engineer",
                    "authorising_officer_designation": "Executive Engineer"
                },
                "certificate_iii": {
                    "totals": expanded_totals,
                    "deductions": deductions,
                    "calculations": {"amount_words": last_page_data["amount_words"]},
                    "payable_words": last_page_data["amount_words"],
                    "current_date": datetime.now()
                },
                "deviation_statement": deviation_data,
                "extra_items": extra_items_data,
                "note_sheet": {
                    "agreement_no": "",
                    "work_name": "",
                    "contractor_name": "",
                    "commencement_date": "",
                    "completion_date": "",
                    "actual_completion_date": "",
                    "work_order_amount": work_order_amount,
                    "totals": expanded_totals,
                    "deductions": deductions,
                    "note_sheet": note_sheet_meta,
                    "current_date": datetime.now().strftime("%d-%m-%Y"),
                },
            }
            
            st.session_state["templates_data"] = templates_data
            
>>>>>>> 3987d94bf8947dde557f9dbc0125c7ece8ab0f3f
            # Clear previously generated byte caches to avoid stale downloads
            keys_to_delete = [
                key for key in list(st.session_state.keys())
                if key == "zip_bytes" or key.startswith("html_bytes_") or key.startswith("pdf_html_bytes_") or key.startswith("pdf_latex_bytes_")
            ]
            for key in keys_to_delete:
                del st.session_state[key]
        except Exception as e:
            st.error(f"Error processing file: {str(e)}")
            _log_traceback()
    
    # If we have templates data, show download actions (persist across reruns)
    if "templates_data" in st.session_state:
        templates_data = st.session_state["templates_data"]
        
        # Combined download section (build once, then persistent download button)
        st.subheader("Download All Documents")
        col1, col2 = st.columns(2)
        with col1:
            if st.button("📦 Build Complete Package (ZIP)", use_container_width=True, key="build_zip"):
                with st.spinner("Creating complete document package..."):
                    zip_path = create_combined_zip(templates_data)
                    if zip_path and os.path.exists(zip_path):
                        with open(zip_path, 'rb') as f:
                            st.session_state["zip_bytes"] = f.read()
                        st.success("✅ Package ready. Use the download button below.")
                    else:
                        st.error("Failed to create document package")
            
            if "zip_bytes" in st.session_state:
                st.download_button(
                    label="📥 Download ZIP Package",
                    data=st.session_state["zip_bytes"],
                    file_name="Bill_Documents_Complete.zip",
                    mime="application/zip",
                    use_container_width=True,
                    key="zip_download"
                )
        with col2:
            st.info("📋 **Package Contents:**\n- HTML templates\n- PDF from HTML\n- PDF from LaTeX\n- LaTeX source files")
        
        # Individual document generation (persist buttons)
        st.subheader("Individual Documents")
        templates = [
            ("first_page", "First_Page"),
            ("certificate_ii", "Certificate_II"),
            ("certificate_iii", "Certificate_III"),
            ("deviation_statement", "Deviation_Statement"),
            ("extra_items", "Extra_Items"),
            ("note_sheet", "Note_Sheet")
        ]
        
        for template_name, doc_name in templates:
            with st.expander(f"📄 {doc_name}", expanded=False):
                col1, col2, col3 = st.columns(3)
                
                # Direct HTML download (fast; no pre-generation button)
                with col1:
                    try:
                        template = env.get_template(f"{template_name}.html")
                        html_content = template.render(data=templates_data[template_name])
                        st.download_button(
                            label=f"📥 Download {doc_name}.html",
                            data=html_content,
                            file_name=f"{doc_name}.html",
                            mime="text/html",
                            key=f"html_dl_{template_name}"
                        )
                    except Exception as e:
                        st.error(f"Error: {str(e)}")
                
                # PDF from HTML (generate -> then persistent download button)
                with col2:
                    if config:
                        if st.button(f"Generate PDF (HTML) - {doc_name}", key=f"gen_pdf_html_{template_name}"):
                            try:
                                with st.spinner(f"Generating {doc_name} PDF from HTML..."):
<<<<<<< HEAD
                                    # Generate HTML content
                                    template = env.get_template(f"{template_name}.html")
                                    html_content = template.render(data=templates_data[template_name])
                                    
                                    # Create a temporary directory for the PDF
                                    with tempfile.TemporaryDirectory() as temp_dir:
                                        pdf_path = os.path.join(temp_dir, f"{doc_name}.pdf")
                                        
                                        # Generate PDF
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
                                        
                                        try:
                                            # Generate PDF to a file
                                            pdfkit.from_string(html_content, pdf_path, options=options, configuration=config)
                                            
                                            # Read the generated PDF into memory
                                            with open(pdf_path, 'rb') as f:
                                                # Store PDF bytes in session state
                                                st.session_state[f"pdf_html_bytes_{template_name}"] = f.read()
                                                
                                            st.success("PDF generated successfully!")
                                            
                                        except Exception as e:
                                            st.error(f"Error generating PDF: {str(e)}")
                                            _log_traceback()
                                            
                            except Exception as e:
                                st.error(f"Error: {str(e)}")
                                _log_traceback()
                    else:
                        st.error("PDF generation not available (wkhtmltopdf not configured)")
                    
                    # Show download button if PDF is ready
=======
                                    template = env.get_template(f"{template_name}.html")
                                    html_content = template.render(data=templates_data[template_name])
                                    temp_dir = get_temp_dir()
                                    pdf_path = os.path.join(temp_dir, f"{doc_name}.pdf")
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
                                    pdfkit.from_string(html_content, pdf_path, options=options, configuration=config)
                                    if os.path.exists(pdf_path):
                                        with open(pdf_path, 'rb') as f:
                                            st.session_state[f"pdf_html_bytes_{template_name}"] = f.read()
                                        st.success("PDF ready. Use the download button below.")
                                    else:
                                        st.error(f"Failed to generate {doc_name} PDF")
                            except Exception as e:
                                st.error(f"Error: {str(e)}")
                    else:
                        st.error("PDF generation not available")
                    
>>>>>>> 3987d94bf8947dde557f9dbc0125c7ece8ab0f3f
                    if f"pdf_html_bytes_{template_name}" in st.session_state:
                        st.download_button(
                            label=f"📥 Download {doc_name}_HTML.pdf",
                            data=st.session_state[f"pdf_html_bytes_{template_name}"],
                            file_name=f"{doc_name}_HTML.pdf",
                            mime="application/pdf",
<<<<<<< HEAD
                            key=f"pdf_html_dl_{template_name}",
                            use_container_width=True
=======
                            key=f"pdf_html_dl_{template_name}"
>>>>>>> 3987d94bf8947dde557f9dbc0125c7ece8ab0f3f
                        )
                
                # PDF from LaTeX (generate -> then persistent download button)
                with col3:
                    if st.button(f"Generate PDF (LaTeX) - {doc_name}", key=f"gen_pdf_latex_{template_name}"):
                        try:
                            with st.spinner(f"Generating LaTeX {doc_name} PDF..."):
                                temp_dir = get_temp_dir()
                                pdf_path = os.path.join(temp_dir, f"{doc_name}_LaTeX.pdf")
                                if generate_latex_pdf(template_name, templates_data[template_name], pdf_path):
                                    with open(pdf_path, 'rb') as f:
                                        st.session_state[f"pdf_latex_bytes_{template_name}"] = f.read()
                                    st.success("LaTeX PDF ready. Use the download button below.")
                                else:
                                    st.error("LaTeX PDF generation failed - pdflatex not available")
                        except Exception as e:
                            st.error(f"Error: {str(e)}")
                    
                    if f"pdf_latex_bytes_{template_name}" in st.session_state:
                        st.download_button(
                            label=f"📥 Download {doc_name}_LaTeX.pdf",
                            data=st.session_state[f"pdf_latex_bytes_{template_name}"],
                            file_name=f"{doc_name}_LaTeX.pdf",
                            mime="application/pdf",
                            key=f"pdf_latex_dl_{template_name}"
                        )
        
        # Display summary (based on templates_data)
        with st.expander("Bill Summary", expanded=True):
            totals = templates_data['first_page']['totals']
            st.write(f"**Total Amount:** ₹{totals.get('grand_total', 0):,}")
            premium = totals.get('premium', {})
            premium_amount = premium.get('amount', 0)
            st.write(f"**Premium ({premium_percent}% {premium_type}):** ₹{premium_amount:,}")
            st.write(f"**Payable Amount:** ₹{totals.get('payable', 0):,}")
            st.write(f"**Amount in Words:** {templates_data['certificate_iii'].get('payable_words', '')}")

if __name__ == "__main__":
    main()