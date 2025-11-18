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
        title_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
        title_run = title_para.add_run(
            f"{fmt(row['Locker Name'])}  (Kiosk {fmt(row[kiosk_col])})"
        )
        title_run.bold = True
        title_run.font.size = Pt(18)

        doc.add_paragraph()  # spacer

        # ---- Table 1: Property block ----
        table1 = doc.add_table(rows=2, cols=6)
        table1.style = "Table Grid"   # borders

        hdr_cells = table1.rows[0].cells
        hdr_cells[0].text = "Property Name"
        hdr_cells[1].text = "Store ID (NSA Unique ID)"
        hdr_cells[2].text = "Address"
        hdr_cells[3].text = "City"
        hdr_cells[4].text = "St"
        hdr_cells[5].text = "Zip"

        val_cells = table1.rows[1].cells
        val_cells[0].text = fmt(row["Property Name"])
        val_cells[1].text = fmt(row["Store ID (NSA Unique ID)"])
        val_cells[2].text = fmt(row["Address"])
        val_cells[3].text = fmt(row["City"])
        val_cells[4].text = fmt(row["St"])
        val_cells[5].text = fmt(row["Zip"])

        doc.add_paragraph()  # spacer

        # ---- Table 2: Locker and contact block ----
        table2 = doc.add_table(rows=2, cols=8)
        table2.style = "Table Grid"

        hdr2 = table2.rows[0].cells
        hdr2[0].text = "Size"
        hdr2[1].text = "Gen"
        hdr2[2].text = "Indoor/ Outdoor"
        hdr2[3].text = "Config"
        hdr2[4].text = "Locker Name"
        hdr2[5].text = "Contact Name"
        hdr2[6].text = "Contact Phone #"
        hdr2[7].text = "PO for Invoice"

        val2 = table2.rows[1].cells
        val2[0].text = fmt(row["Size"])
        val2[1].text = fmt(row["Gen"])
        val2[2].text = fmt(row["Indoor/ Outdoor"])
        val2[3].text = fmt(row["Config"])
        val2[4].text = fmt(row["Locker Name"])
        val2[5].text = fmt(row["Contact Name"])
        val2[6].text = fmt(row["Contact Phone #"])
        val2[7].text = fmt(row["PO for Invoice"])

        # Save to memory buffer
        buffer = BytesIO()
        doc.save(buffer)
        buffer.seek(0)

        file_name = f"{fmt(row['Locker Name'])}_Kiosk_{fmt(row[kiosk_col])}.docx"

        st.download_button(
            label="Download Word document",
            data=buffer,
            file_name=file_name,
            mime=(
                "application/"
                "vnd.openxmlformats-officedocument.wordprocessingml.document"
            ),
        )
