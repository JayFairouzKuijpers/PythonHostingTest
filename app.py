from io import BytesIO

import streamlit as st #hosting
from openpyxl import load_workbook #manipulating excel files

st.title("Excel Modifier")

uploaded_file = st.file_uploader("Choose an Excel file", type=["xlsx"])

if uploaded_file is not None:
    workbook = load_workbook(uploaded_file)

    # Add your Excel modifications here.

    output = BytesIO()
    workbook.save(output)

    st.download_button(
        "Download modified workbook",
        data=output.getvalue(),
        file_name="modified.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )