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
        df = None
        file_name = file.name

        try:
            if file_name.endswith(".xlsx"):
                df = pd.read_excel(file, engine="openpyxl")

            elif file_name.endswith(".xls"):
                try:
                    df = pd.read_excel(file, engine="xlrd")
                except:
                    file.seek(0)
                    tables = pd.read_html(file)
                    df = tables[0]

        except Exception:
            st.warning(f"Could not read file: {file_name}")
            continue

        if df is None or df.empty:
            st.warning(f"No usable data in file: {file_name}")
            continue

        # Remove last row if mostly empty
        last_row = df.tail(1)
        if last_row.isnull().sum(axis=1).values[0] > (len(df.columns) // 2):
            df = df.iloc[:-1]

        st.write(f"Rows read from {file_name}: {len(df)}")

        all_data.append(df)

    if not all_data:
        st.error("No valid Excel files could be read.")
    else:
        # Combine all files
        final_df = pd.concat(all_data, ignore_index=True)

        # Reset index and rebuild S No.
        final_df.reset_index(drop=True, inplace=True)
        final_df.iloc[:, 0] = range(1, len(final_df) + 1)

        # -----------------------------
        # Convert Column C to Month
        # -----------------------------
        date_col = final_df.columns[2]

        def convert_month(x):
            try:
                return datetime.strptime(str(x), "%d %b %Y").strftime("%b-%y")
            except:
                return ""

        final_df["Month"] = final_df[date_col].apply(convert_month)

        # -----------------------------
        # Convert Columns E and F to numeric
        # -----------------------------
        col_e = final_df.columns[4]
        col_f = final_df.columns[5]

        for col in [col_e, col_f]:
            final_df[col] = (
                final_df[col]
                .astype(str)
                .str.replace(",", "", regex=False)
                .str.replace(" ", "", regex=False)
            )
            final_df[col] = pd.to_numeric(final_df[col], errors="coerce").fillna(0)

        # -----------------------------
        # Sort by Financial Year Month Order
        # -----------------------------
        fy_order = {
            "Apr": 1, "May": 2, "Jun": 3, "Jul": 4, "Aug": 5, "Sep": 6,
            "Oct": 7, "Nov": 8, "Dec": 9, "Jan": 10, "Feb": 11, "Mar": 12
        }

        def get_fy_order(month_str):
            try:
                mon = month_str.split("-")[0]
                return fy_order.get(mon, 99)
            except:
                return 99

        final_df["FY_Order"] = final_df["Month"].apply(get_fy_order)
        final_df = final_df.sort_values(by="FY_Order").drop(columns=["FY_Order"])
        final_df.reset_index(drop=True, inplace=True)

        # Rebuild S No. after sorting
        final_df.iloc[:, 0] = range(1, len(final_df) + 1)

        st.success(f"Total rows after merge: {len(final_df)}")
        st.dataframe(final_df)

        # -----------------------------
        # Create Excel file
        # -----------------------------
        def create_excel(df):
            buffer = BytesIO()
            df.to_excel(buffer, index=False, engine="openpyxl")
            buffer.seek(0)
            return buffer

        excel_file = create_excel(final_df)

        st.download_button(
            label="Download Consolidated Excel",
            data=excel_file,
            file_name="Consolidated Excel.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
