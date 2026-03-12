import re
from io import BytesIO

import streamlit as st
from pypdf import PdfReader
from openpyxl import Workbook
from openpyxl.utils import get_column_letter

st.set_page_config(page_title="PO PDF to Excel", page_icon="📄", layout="wide")
st.title("📄 Purchase Order PDFs → One Excel")
st.caption("Upload multiple PO PDFs, extract line items, and download one combined XLSX.")

ITEM_CODE_RE = re.compile(r"^[A-Z0-9][A-Z0-9\-]{3,}$")
NUM_RE = re.compile(r"^\d+(?:\.\d+)?$")


def extract_text_from_pdf(file) -> str:
    reader = PdfReader(file)
    parts = []
    for page in reader.pages:
        parts.append(page.extract_text() or "")
    return "\n".join(parts)


def parse_po_number(text: str) -> str:
    m = re.search(r"(\d{10})\s*\n\s*PURCHASE ORDER NUMBER", text, re.IGNORECASE)
    if m:
        return m.group(1)
    m = re.search(r"PURCHASE ORDER NUMBER\s*\n\s*(\d{6,12})", text, re.IGNORECASE)
    if m:
        return m.group(1)
    m = re.search(r"\b\d{10}\b", text)
    return m.group(0) if m else ""


def parse_po_date(text: str) -> str:
    m = re.search(r"PO Date:\s*([0-3]?\d/[01]?\d/\d{4})", text, re.IGNORECASE)
    if m:
        return m.group(1)
    m = re.search(r"\b([0-3]?\d/[01]?\d/\d{4})\b", text)
    return m.group(1) if m else ""


def parse_rows(text: str):
    rows = []
    lines = [ln.strip() for ln in text.splitlines() if ln.strip()]

    i = 0
    while i < len(lines):
        token = lines[i]
        if ITEM_CODE_RE.match(token) and i + 6 < len(lines):
            desc = lines[i + 1]
            qty, price, texcl, tax, tincl = lines[i + 2:i + 7]
            if all(NUM_RE.match(x) for x in [qty, price, texcl, tax, tincl]):
                rows.append({
                    "item_number": token,
                    "description": desc,
                    "quantity": float(qty),
                    "price": float(price),
                    "total_excl": float(texcl),
                    "tax": float(tax),
                    "total_incl": float(tincl),
                })
                i += 7
                continue
        i += 1

    return rows


def build_excel(all_rows) -> bytes:
    wb = Workbook()
    ws = wb.active
    ws.title = "PO Data"

    headers = [
        "Source PDF",
        "Purchase Order Number",
        "PO Date",
        "Item Number",
        "Description",
        "Quantity",
        "Price",
        "Total Excl",
        "Tax",
        "Total Incl",
    ]
    ws.append(headers)

    for r in all_rows:
        ws.append([
            r["source_pdf"],
            r["po_number"],
            r["po_date"],
            r["item_number"],
            r["description"],
            r["quantity"],
            r["price"],
            r["total_excl"],
            r["tax"],
            r["total_incl"],
        ])

    widths = [26, 22, 12, 16, 50, 10, 10, 12, 10, 12]
    for idx, width in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(idx)].width = width

    for row in ws.iter_rows(min_row=2, min_col=6, max_col=10):
        for cell in row:
            cell.number_format = "0.00"

    bio = BytesIO()
    wb.save(bio)
    bio.seek(0)
    return bio.getvalue()


files = st.file_uploader(
    "Upload one or more purchase-order PDFs",
    type=["pdf"],
    accept_multiple_files=True,
)

if st.button("Process", type="primary"):
    if not files:
        st.warning("Please upload at least one PDF.")
    else:
        all_rows = []
        warnings = []

        for f in files:
            try:
                text = extract_text_from_pdf(f)
                po_number = parse_po_number(text)
                po_date = parse_po_date(text)
                rows = parse_rows(text)

                if not rows:
                    warnings.append(f"No item rows detected in {f.name}")
                    continue

                for row in rows:
                    row["source_pdf"] = f.name
                    row["po_number"] = po_number
                    row["po_date"] = po_date
                    all_rows.append(row)
            except Exception as e:
                warnings.append(f"Failed parsing {f.name}: {e}")

        if not all_rows:
            st.error("Nothing could be parsed from uploaded PDFs.")
            if warnings:
                st.write("Warnings:")
                st.write("\n".join([f"- {w}" for w in warnings]))
        else:
            xlsx = build_excel(all_rows)
            st.success(f"Processed {len(files)} file(s), extracted {len(all_rows)} line item(s).")
            if warnings:
                with st.expander("Warnings"):
                    st.write("\n".join([f"- {w}" for w in warnings]))

            st.download_button(
                label="Download combined Excel",
                data=xlsx,
                file_name="purchase_orders_combined.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
