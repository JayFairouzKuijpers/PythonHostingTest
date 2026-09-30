from io import BytesIO

import streamlit as st #hosting
from openpyxl import load_workbook #manipulating excel files
from openpyxl.styles import PatternFill


def MakeSecondColumnRed(workbook):
    worksheet = workbook.active
    red_fill = PatternFill(fill_type="solid", fgColor="FF0000")
    blue_fill = PatternFill(fill_type="solid", fgColor="0000FF")

    for row in range(2, worksheet.max_row + 1):
        if row % 2 == 0:
            worksheet.cell(row=row, column=2).fill = red_fill
        else:
            worksheet.cell(row=row, column=2).fill = blue_fill

st.title("Excel processor")
st.subheader("version 1.002")

uploaded_file = st.file_uploader("Choose an Excel file", type=["xlsx"])

if uploaded_file is not None:
    workbook = load_workbook(uploaded_file)

    # Add your Excel modifications here.
    MakeSecondColumnRed(workbook)

    output = BytesIO()
    workbook.save(output)

    st.download_button(
        "Download modified workbook",
        data=output.getvalue(),
        file_name="modified.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )