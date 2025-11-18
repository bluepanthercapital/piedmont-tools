import streamlit as st
import pandas as pd
from io import BytesIO
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Pt


# Helper to clean NaN -> ""
def fmt(val):
    if pd.isna(val):
        return ""
    return str(val)


# Columns we absolutely need (Kiosk handled separately because of naming issues)
REQUIRED_COLUMNS = [
    "Property Name",
    "Store ID (NSA Unique ID)",
    "Address",
    "City",
    "St",
    "Zip",
    "Size",
    "Gen",
    "Indoor/ Outdoor",
    "Config",
    "Locker Name",
    "Contact Name",
    "Contact Phone #",
    "PO for Invoice",
]

st.set_page_config(page_title="Amazon Locker Sheet Generator", layout="centered")

st.title("Amazon Locker Sheet Generator")

st.write(
    "Upload the Amazon Locker Excel file, select a Locker Name, "
    "and generate a Word document with the property and locker details."
)

uploaded_file = st.file_uploader(
    "Upload Excel file (.xlsx)", type=["xlsx"]
)

if uploaded_file is not None:
    # Read first sheet by default
    try:
        xls = pd.ExcelFile(uploaded_file)
        sheet_name = xls.sheet_names[0]
        df = pd.read_excel(xls, sheet_name=sheet_name)

    except Exception as e:
        st.error(f"Error reading Excel file: {e}")
        st.stop()

    # Check that required columns exist (excluding Kiosk for now)
    missing = [c for c in REQUIRED_COLUMNS if c not in df.columns]
    if missing:
        st.error(
            "The uploaded file is missing these columns: "
            + ", ".join(missing)
        )
        st.stop()

    # Handle Kiosk column which may be named "Kiosk " (with space) or "Kiosk"
    kiosk_col = None
    if "Kiosk " in df.columns:
        kiosk_col = "Kiosk "
    elif "Kiosk" in df.columns:
        kiosk_col = "Kiosk"
    else:
        st.error(
            "The uploaded file is missing the 'Kiosk' column "
            "(expected header 'Kiosk' or 'Kiosk ')."
        )
        st.stop()

    # Drop rows where Locker Name is missing
    locker_series = df["Locker Name"].dropna().astype(str)
    locker_options = sorted(locker_series.unique())

    if not locker_options:
        st.error("No Locker Name values found in the file.")
        st.stop()

    selected_locker = st.selectbox(
        "Select Locker Name",
        options=locker_options,
        index=0,
    )

    # Filter rows for selected locker
    matching_rows = df[df["Locker Name"].astype(str) == selected_locker]

    st.subheader("Row preview")
    preview_cols = [
        "Locker Name",
        kiosk_col,
        "Property Name",
        "Store ID (NSA Unique ID)",
        "Address",
        "City",
        "St",
        "Zip",
        "Size",
        "Gen",
        "Indoor/ Outdoor",
        "Config",
        "Contact Name",
        "Contact Phone #",
        "PO for Invoice",
    ]
    st.dataframe(matching_rows[preview_cols])

    st.write(
        "If there are multiple rows with the same Locker Name, "
        "the first one will be used for the Word document for now."
    )

    if st.button("Generate Word document"):
        if matching_rows.empty:
            st.error("No row found for that Locker Name.")
            st.stop()

        row = matching_rows.iloc[0]

        # ---- Title block: Locker Name + Kiosk ----
        doc = Document()

        title_para = doc.add_paragraph()
        title_para.alignment = WD_ALIGN_PARAGRAPH.CENT_
