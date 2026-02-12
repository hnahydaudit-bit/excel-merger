import streamlit as st
import pandas as pd
from datetime import datetime
from io import BytesIO

st.title("Excel Merger")

uploaded_files = st.file_uploader(
    "Upload Excel files",
    type=["xls", "xlsx"],
    accept_multiple_files=True
)

if uploaded_files:
    all_data = []

    for file in uploaded_files:
        try:
            if file.name.endswith(".xls"):
                df = pd.read_excel(file, engine="xlrd")
            else:
                df = pd.read_excel(file, engine="openpyxl")
        except:
            df = pd.read_html(file)[0]

        # Remove last row if mostly empty (merged/summary row)
        last_row = df.tail(1)
        if last_row.isnull().sum(axis=1).values[0] > (len(df.columns) // 2):
            df = df.iloc[:-1]

        all_data.append(df)

    # Combine files
    final_df = pd.concat(all_data, ignore_index=True)

    # Convert Column C to Month
    date_col = final_df.columns[2]

    def convert_month(x):
        try:
            return datetime.strptime(str(x), "%d %b %Y").strftime("%b-%y")
        except:
            return ""

    final_df["Month"] = final_df[date_col].apply(convert_month)

    st.success("Files merged successfully!")
    st.dataframe(final_df)

    # Download button
    output = BytesIO()
    final_df.to_excel(output, index=False, engine="openpyxl")
    output.seek(0)

    st.download_button(
        label="Download Consolidated Excel",
        data=output,
        file_name="Consolidated Excel.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
