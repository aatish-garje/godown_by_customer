import streamlit as st
import pandas as pd
from io import BytesIO
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.styles import Font, PatternFill, Border, Side

# Function to process Excel
def split_excel_by_customer(uploaded_file):
    # Read file
    df = pd.read_excel(uploaded_file, sheet_name=0, dtype=str)

    # Rename Item Code -> ZFI Part No No
    if "Item Code" in df.columns:
        df = df.rename(columns={"Item Code": "ZFI Part No No"})

    # Ensure blank column at index 0 (internal name _blank_col)
    if "_blank_col" not in df.columns:
        df.insert(0, "_blank_col", "")

    # Ensure Customer Part No exists at index 1
    if "Customer Part No" not in df.columns:
        df.insert(1, "Customer Part No", "")

    required_cols = ["_blank_col", "Customer Part No", "ZFI Part No No", "Item Desc",
                     "Inv No", "Inv Date", "Qty", "Amount"]

    # Keep only required cols + Code for grouping
    available_cols = [c for c in required_cols if c in df.columns]
    df = df[available_cols + ["Code"]]

    # Convert Qty & Amount to numbers
    if "Qty" in df.columns:
        df["Qty"] = pd.to_numeric(df["Qty"], errors="coerce").fillna(0).astype(float)
    if "Amount" in df.columns:
        df["Amount"] = pd.to_numeric(df["Amount"], errors="coerce").fillna(0).astype(float)

    # Create new workbook
    wb = Workbook()
    wb.remove(wb.active)

    # Define styles
    header_font = Font(bold=True, color="FFFFFF")
    header_fill = PatternFill("solid", fgColor="4F81BD")
    thin_border = Border(
        left=Side(style="thin"), right=Side(style="thin"),
        top=Side(style="thin"), bottom=Side(style="thin")
    )

    # Create sheets by customer code
    for code, group in df.groupby("Code"):
        ws = wb.create_sheet(title=str(code))

        # Drop Code column for final export
        group = group[available_cols]

        # Write rows
        rows = list(dataframe_to_rows(group, index=False, header=True))

        # Make the first header (blank col) actually empty
        rows[0][0] = None

        for r_idx, row in enumerate(rows, 1):
            ws.append(row)
            for c_idx, cell in enumerate(ws[r_idx], 1):
                if c_idx == 1:  # Blank column
                    continue

                # Formatting per column
                if r_idx > 1:
                    if ws.cell(1, c_idx).value == "Customer Part No":
                        cell.number_format = "@"
                    elif ws.cell(1, c_idx).value == "Qty":
                        try:
                            cell.value = float(cell.value)
                            cell.number_format = "0"
                        except:
                            pass
                    elif ws.cell(1, c_idx).value == "Amount":
                        try:
                            cell.value = float(cell.value)
                            cell.number_format = "#,##0.00"
                        except:
                            pass

                # Borders + header styles
                cell.border = thin_border
                if r_idx == 1 and c_idx > 1:  # Header row (except blank col)
                    cell.font = header_font
                    cell.fill = header_fill

        # Auto-adjust column widths
        for col in ws.columns:
            col_letter = col[0].column_letter
            if col_letter == "A":  # Blank column default width
                ws.column_dimensions[col_letter].width = 3
                continue
            max_length = 0
            for cell in col:
                try:
                    if cell.value is not None:
                        max_length = max(max_length, len(str(cell.value)))
                except:
                    pass
            ws.column_dimensions[col_letter].width = max_length + 2

    # Save to BytesIO
    output = BytesIO()
    wb.save(output)
    output.seek(0)
    return output


# ---------------- STREAMLIT APP ----------------
st.title("📊 Split Excel by Customer Code")

uploaded_file = st.file_uploader("Upload your Excel file", type=["xlsx"])

if uploaded_file:
    st.success("✅ File uploaded successfully!")

    if st.button("Process File"):
        processed_file = split_excel_by_customer(uploaded_file)

        st.download_button(
            label="📥 Download Formatted Excel",
            data=processed_file,
            file_name="Split_By_Customer_Formatted.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
