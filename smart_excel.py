import streamlit as st
import pandas as pd
from io import BytesIO

# Title
st.title("Excel Matcher")

# File uploaders
file1 = st.file_uploader("Please upload the file with no rates")
file2 = st.file_uploader("Please uplaod the file with rates")

# Initialize dataframes
df1 = None
df2 = None

# Sheet selection
if file1 and file2:
    # Read Excel files
    excel_file1 = pd.ExcelFile(file1)
    excel_file2 = pd.ExcelFile(file2)

    # Get sheet names
    sheet_names1 = excel_file1.sheet_names
    sheet_names2 = excel_file2.sheet_names

    # Select sheet numbers
    sheet_number1 = st.selectbox("Select sheet number for File 1", range(1, len(sheet_names1) + 1))
    sheet_number2 = st.selectbox("Select sheet number for File 2", range(1, len(sheet_names2) + 1))

    # Read selected sheets
    df1 = excel_file1.parse(sheet_names1[sheet_number1 - 1])
    df2 = excel_file2.parse(sheet_names2[sheet_number2 - 1])

    # Display column names
    st.write("Columns in File 1:", df1.columns)
    st.write("Columns in File 2:", df2.columns)

# Column name input
if df1 is not None:
    column_name = st.selectbox("Select column to match", df1.columns)
else:
    column_name = ""

# Button to generate output
if st.button("Generate Output"):
    # Check if dataframes are loaded
    if df1 is not None and df2 is not None:
        # Trim whitespace from column names
        df1.columns = df1.columns.str.strip()
        df2.columns = df2.columns.str.strip()

        # Remove duplicates from both files based on the matching column
        df1 = df1.drop_duplicates(subset=column_name, keep='first')
        df2 = df2.drop_duplicates(subset=column_name, keep='first')

        # Merge dataframes on selected column, prioritizing Rate from File 2
        merged_df = pd.merge(df1, df2, on=column_name, how='left', suffixes=('_x', '_y'))

        # Fill Rate from File 2 where available
        merged_df['Rate'] = merged_df['Rate_y'].fillna(merged_df['Rate_x'])

        # Drop unnecessary columns
        merged_df = merged_df.drop(['Rate_x', 'Rate_y'], axis=1)

        # Create output Excel file
        buffer = BytesIO()
        merged_df.to_excel(buffer, index=False)
        buffer.seek(0)

        # Download output file
        st.download_button(
            label="Download Output Excel",
            data=buffer,
            file_name="output.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
    else:
        st.error("Please upload both Excel files.")