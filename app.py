from io import BytesIO

import streamlit as st #hosting
from openpyxl import load_workbook #manipulating excel files
from openpyxl.styles import PatternFill
from openpyxl.worksheet.datavalidation import DataValidation

def InsertNewColumn(workbook):
    worksheet = workbook.active
    worksheet.insert_cols(2)  # Insert a new column at index 2
    worksheet.cell(row=1, column=2, value="Standaard Object")

    last_row = max(worksheet.max_row, 2)

    dropdown = DataValidation(
        type="list",
        formula1='"Object A,Object B,Object C"',
        allow_blank=False,
    )
    worksheet.add_data_validation(dropdown)
    dropdown.add(f"B2:B{last_row}")

    for row in range(2, last_row + 1):
        worksheet.cell(row=row, column=2, value="Object A")

def MakeThirdColumnRed(workbook):
    worksheet = workbook.active
    red_fill = PatternFill(fill_type="solid", fgColor="FF0000")
    blue_fill = PatternFill(fill_type="solid", fgColor="0000FF")

    for row in range(2, worksheet.max_row + 1):
        if row % 2 == 0:
            worksheet.cell(row=row, column=3).fill = red_fill
        else:
            worksheet.cell(row=row, column=3).fill = blue_fill

st.title("Excel processor")
st.subheader("version 1.003")

uploaded_file = st.file_uploader("Choose an Excel file", type=["xlsx"])

if uploaded_file is not None:
    workbook = load_workbook(uploaded_file)

    # Add your Excel modifications here.
    InsertNewColumn(workbook)
    MakeThirdColumnRed(workbook)

    output = BytesIO()
    workbook.save(output)

    st.download_button(
        "Download modified workbook",
        data=output.getvalue(),
        file_name="modified.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )