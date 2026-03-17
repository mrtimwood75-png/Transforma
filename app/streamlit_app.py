# streamlit_app.py

import streamlit as st
from pathlib import Path

from core_logic import (
    prepare_preview_rows,
    convert_uploaded_files,
    default_template_path,
    detect_report_format,
    parse_order_bytes,
    parse_uploaded_order_file,
    get_uploaded_name,
)

st.set_page_config(
    page_title="BoConcept Order Converter",
    page_icon="■",
    layout="wide",
)

st.markdown(
    """
    <style>
    #MainMenu {visibility:hidden;}
    footer {visibility:hidden;}
    header {visibility:hidden;}

    .stApp {
        background-color: #f6f6f4;
    }

    .block-container {
        max-width: 1220px;
        padding-top: 28px;
        padding-bottom: 40px;
    }

    .hero {
        background: #111111;
        border-radius: 18px;
        padding: 24px 28px;
        margin-bottom: 24px;
    }

    .hero-title {
        color: white;
        font-size: 2.2rem;
        font-weight: 700;
        line-height: 1.1;
        margin: 0;
    }

    .hero-sub {
        color: #d6d6d6;
        font-size: 1rem;
        margin-top: 8px;
    }

    .note-box {
        background: #ffffff;
        border: 1px solid #dadada;
        border-radius: 14px;
        padding: 16px 18px;
        margin-bottom: 18px;
        color: #111111;
    }

    .side-card {
        background: #ffffff;
        border: 1px solid #dadada;
        border-radius: 14px;
        padding: 18px 20px;
        color: #111111;
    }

    .section-title {
        font-size: 1rem;
        font-weight: 700;
        color: #111111;
        margin-bottom: 8px;
    }

    .preview-wrap {
        background: #ffffff;
        border: 1px solid #dadada;
        border-radius: 14px;
        padding: 18px 18px 8px 18px;
    }

    .stButton > button {
        background: #111111 !important;
        color: white !important;
        border: 1px solid #111111 !important;
        border-radius: 10px !important;
        font-weight: 600 !important;
        min-height: 44px !important;
    }

    .stDownloadButton > button {
        background: #111111 !important;
        color: white !important;
        border: 1px solid #111111 !important;
        border-radius: 10px !important;
        font-weight: 600 !important;
        min-height: 44px !important;
    }

    div[data-testid="stTextInput"] input {
        background: white !important;
    }

    div[data-testid="stFileUploader"] section {
        background: white !important;
        border-radius: 12px !important;
    }

    .small-muted {
        color: #666666;
        font-size: 0.92rem;
    }
    </style>
    """,
    unsafe_allow_html=True,
)

if "rows" not in st.session_state:
    st.session_state.rows = []
if "workbook_bytes" not in st.session_state:
    st.session_state.workbook_bytes = None
if "output_name" not in st.session_state:
    st.session_state.output_name = "converted_sales_orders.xlsx"

logo_candidates = [
    Path("files/BCLOGO.jpg"),
    Path("files/BCLOGO.png"),
    Path("Files/BCLOGO.jpg"),
    Path("Files/BCLOGO.png"),
]

logo_path = None
for candidate in logo_candidates:
    if candidate.exists():
        logo_path = candidate
        break

hero_left, hero_right = st.columns([1, 5], vertical_alignment="center")

with hero_left:
    if logo_path:
        st.image(str(logo_path), width=130)

with hero_right:
    st.markdown(
        """
        <div class="hero">
            <div class="hero-title">BoConcept Order Converter</div>
            <div class="hero-sub">Convert ASCII reports and XLS reports into the delivery import workbook.</div>
        </div>
        """,
        unsafe_allow_html=True,
    )

st.markdown(
    """
    <div class="note-box">
        <strong>Accepted input:</strong> ASCII text reports in the existing supported formats, plus the supported
        <strong>XLS delivery report</strong>.
    </div>
    """,
    unsafe_allow_html=True,
)

left, right = st.columns([2.3, 1], gap="large")

with left:
    st.markdown('<div class="section-title">Upload files</div>', unsafe_allow_html=True)

    text_files = st.file_uploader(
        "Upload ASCII files",
        type=["txt"],
        accept_multiple_files=True,
        key="text_reports",
    )

    xls_files = st.file_uploader(
        "Import XLS",
        type=["xls"],
        accept_multiple_files=True,
        key="xls_reports",
    )

    uploaded_files = []
    if text_files:
        uploaded_files.extend(text_files)
    if xls_files:
        uploaded_files.extend(xls_files)

    st.markdown("<div style='height:10px'></div>", unsafe_allow_html=True)

    b1, b2, b3 = st.columns([1, 1, 2])

    with b1:
        load_preview = st.button("Load Preview", use_container_width=True)

    with b2:
        create_file = st.button("Create Workbook", use_container_width=True)

    with b3:
        st.session_state.output_name = st.text_input(
            "Output file name",
            value=st.session_state.output_name,
        )

with right:
    st.markdown(
        """
        <div class="side-card">
            <div class="section-title">Instructions</div>
            <div class="small-muted">
                1. Upload ASCII and/or XLS file(s)<br><br>
                2. Click <strong>Load Preview</strong><br><br>
                3. Review the parsed rows<br><br>
                4. Click <strong>Create Workbook</strong><br><br>
                5. Download the Excel file
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

template_path = default_template_path()
template_bytes = None
if template_path:
    with open(template_path, "rb") as f:
        template_bytes = f.read()

if load_preview:
    if not uploaded_files:
        st.error("Upload at least one file.")
    else:
        try:
            total_orders = 0
            total_items = 0
            formats = []

            for f in uploaded_files:
                name = get_uploaded_name(f).lower()

                if name.endswith(".xls"):
                    fmt = "xls"
                    parsed = parse_uploaded_order_file(f)
                else:
                    text = f.getvalue().decode(errors="ignore")
                    fmt = detect_report_format(text)
                    parsed = parse_order_bytes(f.getvalue())

                formats.append(fmt)
                total_orders += len(parsed)
                total_items += sum(len(items) for _, items in parsed)

            st.success(f"Format detected: {', '.join(sorted(set(formats)))}")
            st.success(f"Orders detected: {total_orders} | Items detected: {total_items}")

            st.session_state.rows = prepare_preview_rows(uploaded_files)
            st.session_state.workbook_bytes = None

        except Exception as e:
            st.error(f"Error: {e}")

if create_file:
    if not template_bytes:
        st.error("Template not found in /files folder.")
    elif not uploaded_files:
        st.error("Upload at least one file.")
    else:
        try:
            file_bytes = convert_uploaded_files(
                uploaded_files,
                template_bytes
            )
            st.session_state.workbook_bytes = file_bytes
            st.success("Workbook created successfully.")
        except Exception as e:
            st.error(f"Export error: {e}")

if st.session_state.workbook_bytes:
    st.download_button(
        "Download Excel Workbook",
        data=st.session_state.workbook_bytes,
        file_name=st.session_state.output_name,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

st.markdown("<div style='height:20px'></div>", unsafe_allow_html=True)

st.markdown('<div class="preview-wrap">', unsafe_allow_html=True)
st.subheader("Preview")

if st.session_state.rows:
    preview_rows = st.session_state.rows[:200]

    for r in preview_rows:
        st.write(
            f"**{r['sales order number']}**  |  "
            f"{r['sku number']}  |  "
            f"{r['product description']}  |  "
            f"Qty {r['quantity']}"
        )

    if len(st.session_state.rows) > 200:
        st.caption(f"Showing first 200 of {len(st.session_state.rows)} rows")
else:
    st.write("No data loaded")

st.markdown("</div>", unsafe_allow_html=True)
