import streamlit as st
import pandas as pd
import pdfkit
from num2words import num2words
import os
import zipfile
import tempfile
from jinja2 import Environment, FileSystemLoader
import platform
from datetime import datetime
import subprocess
import shutil
import traceback
import requests
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
            if r.status_code == 200 and r.headers.get("content-type", "").startswith(
                "image"
            ):
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
    st.markdown(
        """
    <div style="display:flex; align-items:center; gap:12px;">
      <div>
        <h2 style="margin:0;">Bill Generator</h2>
        <p style="margin:0; color:#666;">A4 documents with professional layout</p>
      </div>
    </div>
    """,
        unsafe_allow_html=True,
    )

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


def _log_error(message: str) -> None:
    """Log error messages"""
    if DEBUG_VERBOSE:
        try:
            st.error(message)
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
        "https://github.com/wkhtmltopdf/packaging/releases/download/0.12.6-1/wkhtmltox-0.12.6-1.centos8.x86_64.rpm.tar.xz",
    ]
    cache_dir = os.path.join(tempfile.gettempdir(), "wkhtmltopdf_bin")
    os.makedirs(cache_dir, exist_ok=True)
    for url in urls:
        try:
            resp = requests.get(url, timeout=30, stream=True)
            if resp.status_code != 200 or len(resp.content) < 1024:
                continue
            # Validate content size to prevent memory exhaustion
            content_length = resp.headers.get("content-length")
            if (
                content_length and int(content_length) > 100 * 1024 * 1024
            ):  # 100MB limit
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
    config = (
        pdfkit.configuration(wkhtmltopdf=wkhtmltopdf_exe)
        if wkhtmltopdf_exe
        else pdfkit.configuration()
    )
except Exception:
    config = None


def number_to_words(number):
    try:
        return num2words(int(number), lang="en_IN").title()
    except (ValueError, TypeError, AttributeError) as e:
        _log_error(f"Error converting number to words: {e}")
        return str(number)


def process_excel(file, premium_percent, premium_type, has_headers=True):
    df = pd.read_excel(file, header=0 if has_headers else None)

    templates_data = {}

    # Detect sections dynamically
    header_start = (
        df[
            df.iloc[:, 0].str.contains(
                "Name of Contractor|Name of Contractor or supplier", na=False
            )
        ].index[0]
        if has_headers
        else 0
    )
    bill_items_start = (
        df[
            df.iloc[:, 0].str.contains("Bill Items|Item No", na=False, case=False)
        ].index[0]
        + 1
        if has_headers
        else 13
    )
    extra_items_start = (
        df[df.iloc[:, 0].str.contains("Extra Items", na=False, case=False)].index[0] + 1
        if has_headers
        else find_extra_start(df, bill_items_start)
    )

    # Map Header
    header_df = df.iloc[header_start : header_start + 12, 0:2].fillna("N/A")
    header_dict = dict(zip(header_df.iloc[:, 0], header_df.iloc[:, 1]))
    templates_data["first_page"] = {
        "header": header_dict,
        "name_of_firm": header_dict.get("Name of Contractor or supplier", "N/A"),
        "name_of_work": header_dict.get("Name of Work", "N/A"),
        "work_order_amount": float(
            header_dict.get("WORK ORDER AMOUNT RS.", 0)
            or header_dict.get("WORK ORDER AMOUNT RS.", "")
            or 0
        ),
        "last_bill_amount": float(
            header_dict.get("Last Bill Amount", 0)
            or header_dict.get("No. and date of the last bill", "").split()[-1]
            if " " in header_dict.get("No. and date of the last bill", "")
            else 0
        ),
    }

    # Map Bill Items with validation
    bill_items_cols = (
        [
            "Unit",
            "Quantity Since",
            "Quantity Upto",
            "Item No.",
            "Description",
            "Rate",
            "Amount Upto",
            "Amount Since",
            "Remark",
        ]
        if has_headers
        else range(9)
    )
    bill_items_df = df.iloc[
        bill_items_start : bill_items_start + len(df) - bill_items_start,
        : max(len(bill_items_cols), df.shape[1]),
    ].dropna(how="all", subset=bill_items_cols)
    if not bill_items_df.empty:
        bill_items = bill_items_df.to_dict(orient="records")
        templates_data["first_page"]["bill_items"] = [
            {
                "unit": str(item.get(bill_items_cols[0], "")).strip(),
                "quantity_since": float(
                    str(item.get(bill_items_cols[1], "0")).strip() or 0
                ),
                "quantity_upto": float(
                    str(item.get(bill_items_cols[2], "0")).strip() or 0
                ),
                "serial_no": str(item.get(bill_items_cols[3], "")).strip(),
                "description": str(item.get(bill_items_cols[4], "")).strip(),
                "rate": float(str(item.get(bill_items_cols[5], "0")).strip() or 0),
                "amount_upto": float(
                    str(item.get(bill_items_cols[6], "0")).strip() or 0
                ),
                "amount_since": float(
                    str(item.get(bill_items_cols[7], "0")).strip() or 0
                ),
                "remark": str(item.get(bill_items_cols[8], "")).strip(),
            }
            for item in bill_items
            if any(item.get(col, "") != "" for col in bill_items_cols)
        ]
    else:
        templates_data["first_page"]["bill_items"] = []

    # Map Extra Items with validation
    extra_items_cols = (
        [
            "Unit",
            "Item No.",
            "Description",
            "Quantity Since",
            "Quantity Upto",
            "Rate",
            "Amount Upto",
            "Amount Since",
            "Remark",
        ]
        if has_headers
        else range(9)
    )
    extra_items_df = df.iloc[
        extra_items_start:, : max(len(extra_items_cols), df.shape[1])
    ].dropna(how="all", subset=extra_items_cols)
    if not extra_items_df.empty:
        extra_items = extra_items_df.to_dict(orient="records")
        templates_data["first_page"]["extra_items"] = [
            {
                "unit": str(item.get(extra_items_cols[0], "")).strip(),
                "serial_no": str(item.get(extra_items_cols[1], "")).strip(),
                "description": str(item.get(extra_items_cols[2], "")).strip(),
                "quantity_since": float(
                    str(item.get(extra_items_cols[3], "0")).strip() or 0
                ),
                "quantity_upto": float(
                    str(item.get(extra_items_cols[4], "0")).strip() or 0
                ),
                "rate": float(str(item.get(extra_items_cols[5], "0")).strip() or 0),
                "amount_upto": float(
                    str(item.get(extra_items_cols[6], "0")).strip() or 0
                ),
                "amount_since": float(
                    str(item.get(extra_items_cols[7], "0")).strip() or 0
                ),
                "remark": str(item.get(extra_items_cols[8], "")).strip(),
            }
            for item in extra_items
            if any(item.get(col, "") != "" for col in extra_items_cols)
        ]
    else:
        templates_data["first_page"]["extra_items"] = []

    # Calculate Totals with validation
    bill_total = (
        sum(item["amount_upto"] for item in templates_data["first_page"]["bill_items"])
        if templates_data["first_page"]["bill_items"]
        else 0
    )
    extra_items_total = (
        sum(item["amount_upto"] for item in templates_data["first_page"]["extra_items"])
        if templates_data["first_page"]["extra_items"]
        else 0
    )
    tender_premium_percent = premium_percent / 100 if premium_percent else 0
    is_addition = premium_type.lower() == "add"
    premium_amount = (
        (bill_total + extra_items_total) * tender_premium_percent
        if is_addition
        else -(bill_total + extra_items_total) * tender_premium_percent
    )
    grand_total = bill_total + extra_items_total + premium_amount
    net_payable = max(
        0, grand_total - templates_data["first_page"]["last_bill_amount"]
    )  # Ensure non-negative

    # Populate totals with new structure
    templates_data["first_page"]["totals"] = {
        "bill_total": bill_total,
        "bill_premium": premium_amount if is_addition else 0,
        "grand_total": grand_total,
        "extra_items_base": extra_items_total,
        "extra_premium": 0,  # Not in instructions, but keeping for backward compatibility
        "extra_items_total": extra_items_total,
        "work_order_total": templates_data["first_page"]["work_order_amount"],
        "last_bill_amount": templates_data["first_page"]["last_bill_amount"],
        "payable": net_payable,
        "premium": {
            "percent": tender_premium_percent,
            "type": "Add" if is_addition else "Deduct",
            "amount": premium_amount,
        },
    }

    # For backward compatibility, also keep the flat structure
    templates_data["first_page"].update(
        {
            "bill_total": bill_total,
            "bill_premium": premium_amount if is_addition else 0,
            "bill_grand_total": bill_total + (premium_amount if is_addition else 0),
            "extra_items_base": extra_items_total,
            "extra_premium": 0,  # Not in instructions, but keeping for backward compatibility
            "extra_items_total": extra_items_total,
            "total_with_premium": grand_total,
            "work_order_amount": templates_data["first_page"]["work_order_amount"],
            "last_bill_amount": templates_data["first_page"]["last_bill_amount"],
            "net_payable": net_payable,
            "premium_percent": premium_percent,
            "premium_type": "Add" if is_addition else "Deduct",
            "tender_premium_percent": tender_premium_percent,
        }
    )

    return templates_data


def find_extra_start(df, bill_start):
    for row in range(bill_start, len(df)):
        if pd.isna(df.iloc[row, 0]):
            return row + 1
    return len(df)  # Default to end if no blank found


def generate_bill_notes(payable_amount, work_order_amount, extra_item_amount):
    percentage_work_done = (
        float(payable_amount / work_order_amount * 100) if work_order_amount > 0 else 0
    )
    serial_number = 1
    note = []
    note.append(
        f"{serial_number}. The work has been completed {percentage_work_done:.2f}% of the Work Order Amount."
    )
    serial_number += 1
    if percentage_work_done < 90:
        note.append(
            f"{serial_number}. The execution of work at final stage is less than 90%..."
        )
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
        extra_item_percentage = (
            float(extra_item_amount / work_order_amount * 100)
            if work_order_amount > 0
            else 0
        )
        if extra_item_percentage > 5:
            note.append(
                f"{serial_number}. The amount of Extra items is Rs. {extra_item_amount}..."
            )
        else:
            note.append(
                f"{serial_number}. The amount of Extra items is Rs. {extra_item_amount}..."
            )
        serial_number += 1
    note.append(
        f"{serial_number}. Please peruse above details for necessary decision-making."
    )
    note.append("")
    note.append("                                Premlata Jain")
    note.append("                               AAO- As Auditor")
    return {"notes": note}


def generate_pdf(template_name, data, output_path):
    """Generate a PDF from a template with the given data.

    Args:
        template_name (str): Name of the template file (without .html extension)
        data (dict): Data to render in the template
        output_path (str): Path to save the generated PDF
    """
    try:
        # Precompute deviation percentage if this is a deviation statement
        if template_name == "deviation_statement":
            total_with_premium = data.get("totals", {}).get("total_with_premium", 0)
            work_order_total = (
                data.get("totals", {}).get("work_order_total", 0) or 1
            )  # Avoid division by zero

            # Calculate and store the deviation percentage
            deviation = (
                (total_with_premium - work_order_total) / work_order_total
            ) * 100
            data["totals"]["deviation_percentage"] = abs(deviation)

        # Render the template
        template = env.get_template(f"{template_name}.html")
        html_content = template.render(data=data)

        # Generate PDF
        pdfkit.from_string(
            html_content,
            output_path,
            configuration=config,
            options={
                "encoding": "UTF-8",
                "enable-local-file-access": None,
                "quiet": "",
            },
        )
        return True
    except Exception as e:
        _log_error(f"Error generating {template_name}.pdf: {str(e)}")
        return False


def generate_latex_pdf(template_name, data, output_path):
    """Generate PDF from LaTeX template"""
    try:
        # Precompute deviation percentage if this is a deviation statement
        if template_name == "deviation_statement":
            total_with_premium = data.get("totals", {}).get("total_with_premium", 0)
            work_order_total = (
                data.get("totals", {}).get("work_order_total", 0) or 1
            )  # Avoid division by zero

            # Calculate and store the deviation percentage
            deviation = (
                (total_with_premium - work_order_total) / work_order_total
            ) * 100
            data["totals"]["deviation_percentage"] = abs(deviation)

        # Render the LaTeX template
        template = env.get_template(f"{template_name}.tex")
        tex_content = template.render(data=data)

        # Create a temporary directory for compilation
        with tempfile.TemporaryDirectory() as temp_dir:
            # Write the .tex file
            tex_path = os.path.join(temp_dir, f"{template_name}.tex")
            with open(tex_path, "w", encoding="utf-8") as f:
                f.write(tex_content)

            # Compile the LaTeX to PDF
            try:
                subprocess.run(
                    [
                        "pdflatex",
                        "-interaction=nonstopmode",
                        f"-output-directory={temp_dir}",
                        tex_path,
                    ],
                    check=True,
                    capture_output=True,
                    text=True,
                )

                # Copy the generated PDF to the output path
                pdf_path = os.path.splitext(tex_path)[0] + ".pdf"
                if os.path.exists(pdf_path):
                    shutil.copy2(pdf_path, output_path)
                    return True
                else:
                    _log_error(f"PDF generation failed: {pdf_path} not found")
                    return False

            except subprocess.CalledProcessError as e:
                _log_error(f"LaTeX compilation failed: {e.stderr}")
                return False

    except Exception as e:
        _log_error(f"Error generating {template_name}.pdf: {str(e)}")
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
        ("note_sheet", "Note_Sheet"),
    ]

    with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as zip_file:
        for template_name, doc_name in templates:
            try:
                # Generate HTML
                html_template = env.get_template(f"{template_name}.html")
                html_content = html_template.render(data=templates_data[template_name])
                zip_file.writestr(f"{doc_name}.html", html_content)

                # Generate HTML-based PDF
                if config:
                    try:
                        pdf_path = os.path.join(temp_dir, f"{doc_name}.pdf")
                        options = {
                            "page-size": "A4",
                            "margin-top": "10mm",
                            "margin-right": "10mm",
                            "margin-bottom": "10mm",
                            "margin-left": "10mm",
                            "encoding": "UTF-8",
                            "no-outline": None,
                            "enable-local-file-access": None,
                        }
                        pdfkit.from_string(
                            html_content,
                            pdf_path,
                            options=options,
                            configuration=config,
                        )
                        if os.path.exists(pdf_path):
                            with open(pdf_path, "rb") as f:
                                zip_file.writestr(f"{doc_name}_HTML.pdf", f.read())
                    except Exception as e:
                        _log_warn(f"HTML PDF generation failed for {doc_name}: {e}")

                # Generate LaTeX-based PDF if available
                latex_pdf_path = os.path.join(temp_dir, f"{doc_name}_LaTeX.pdf")
                if generate_latex_pdf(
                    template_name, templates_data[template_name], latex_pdf_path
                ):
                    with open(latex_pdf_path, "rb") as f:
                        zip_file.writestr(f"{doc_name}_LaTeX.pdf", f.read())

                # Generate LaTeX source
                if os.path.exists(f"templates/{template_name}.tex"):
                    latex_template = env.get_template(f"{template_name}.tex")
                    latex_content = latex_template.render(
                        data=templates_data[template_name]
                    )
                    zip_file.writestr(f"{doc_name}.tex", latex_content)

            except Exception as e:
                _log_warn(f"Error generating {doc_name}: {e}")

    return zip_path if os.path.exists(zip_path) else None


def main():
    st.title("Government Bill Generator")

    # File upload
    uploaded_file = st.file_uploader("Choose an Excel file", type=["xlsx", "xls"])

    if uploaded_file is None:
        _render_landing()
        return

    # Premium settings
    col1, col2 = st.columns(2)
    with col1:
        premium_percent = st.number_input(
            "Tender Premium (%)", value=2.5, min_value=0.0, max_value=100.0, step=0.1
        )
    with col2:
        premium_type = st.selectbox("Premium Type", ["Add", "Deduct"])

    # Add last bill amount input (currently not used in processing)
    st.number_input(
        "Last Bill Amount",
        min_value=0.0,
        value=0.0,
        step=0.01,
        format="%.2f",
        help="Enter the amount from the previous bill",
    )

    # Generate bill and capture results into session state
    if st.button("Generate Bill"):
        try:
            # Process the Excel file
            with st.spinner("Processing Excel file..."):
                templates_data = process_excel(
                    uploaded_file, premium_percent, premium_type
                )

            # Store templates data in session
            st.session_state["templates_data"] = templates_data

            # Display success message
            st.success("Bill processed successfully!")

            # Debug: Display the full templates_data to verify structure
            if DEBUG_VERBOSE:
                st.json(templates_data)

            # Access the totals directly as per instructions
            try:
                totals = templates_data["first_page"]["totals"]

                # Display key financial information
                st.subheader("Bill Summary")
                col1, col2, col3 = st.columns(3)

                with col1:
                    st.metric("Bill Total", f"₹{totals.get('bill_total', 0):,.2f}")
                    st.metric(
                        "Extra Items Base", f"₹{totals.get('extra_items_base', 0):,.2f}"
                    )

                with col2:
                    premium = totals.get("premium", {})
                    st.metric(
                        f"Premium ({premium.get('percent', 0) * 100}% {premium.get('type', 'N/A')})",
                        f"₹{premium.get('amount', 0):,.2f}",
                    )
                    st.metric(
                        "Work Order Total", f"₹{totals.get('work_order_total', 0):,.2f}"
                    )

                with col3:
                    st.metric(
                        "Grand Total",
                        f"₹{totals.get('grand_total', 0):,.2f}",
                        delta=f"₹{totals.get('payable', 0):,.2f} after last bill",
                    )

                # Add download buttons for different templates
                st.subheader("Download Documents")
                col1, col2, col3 = st.columns(3)

                with col1:
                    if st.button("Generate PDF"):
                        # Add PDF generation logic here
                        pass

                with col2:
                    if st.button("Generate DOCX"):
                        # Add DOCX generation logic here
                        pass

                with col3:
                    if st.button("Download All"):
                        # Add zip file generation logic here
                        pass

            except KeyError as e:
                st.error(
                    f"KeyError: {e}. The processed data is missing expected fields."
                )
                if DEBUG_VERBOSE:
                    st.json(templates_data)

        except KeyError as e:
            st.error(
                f"KeyError: {e}. Ensure the Excel file has the correct format and all required fields."
            )
            if DEBUG_VERBOSE:
                st.exception(e)
        except Exception as e:
            st.error(f"Error processing file: {str(e)}")
            if DEBUG_VERBOSE:
                st.exception(e)


if __name__ == "__main__":
    main()

