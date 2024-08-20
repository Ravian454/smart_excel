import streamlit as st
import pandas as pd
from io import BytesIO

# Title
st.title("Excel Tools")

# Add a sidebar with buttons to switch between features
with st.sidebar:
    st.title("Features")
    if st.button("Update Rate with Percentage"):
        st.session_state.feature = "update_rate"
    if st.button("Excel Matcher"):
        st.session_state.feature = "excel_matcher"

# Main screen content based on selected feature
if "feature" not in st.session_state:
    st.session_state.feature = "excel_matcher"

if st.session_state.feature == "update_rate":
    # Percentage update feature content
    file3 = st.file_uploader("Upload Excel file to update Rate")
    if file3:
        excel_file = pd.ExcelFile(file3)
        sheet_names = excel_file.sheet_names
        
        # Display sheet names for selection
        selected_sheet = st.selectbox("Select Sheet Number", 
                                      options=range(1, len(sheet_names) + 1), 
                                      format_func=lambda x: sheet_names[x-1])
        
        # Read selected sheet
        df3 = excel_file.parse(sheet_names[selected_sheet - 1], header=0)
        
        # Display column names
        st.write("Columns in the uploaded file:")
        st.write(df3.columns)
        
        percentage = st.number_input("Enter percentage (%)", min_value=0, max_value=100)
        if st.button("Update Rate"):
            # Update Rate logic
            df3['Rate'] = df3['Rate'] + (df3['Rate'] * (percentage / 100))
            buffer = BytesIO()
            df3.to_excel(buffer, index=False)
            buffer.seek(0)
            st.download_button(
                label="Download Updated Excel",
                data=buffer,
                file_name="updated_file.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
    else:
        st.error("Please upload an Excel file")
elif st.session_state.feature == "excel_matcher":

    # File uploaders
    file1 = st.file_uploader("Please upload the file with no rates", type=['xlsx', 'csv', 'xls'])
    file2 = st.file_uploader("Please upload the file with rates", type=['xlsx', 'csv', 'xls'])

    # Initialize dataframes
    df1 = None
    df2 = None

    # Sheet selection
    if file1 and file2:
        # Read Excel files
        if file1.type in ['application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', 'application/vnd.ms-excel']:
            excel_file1 = pd.ExcelFile(file1)
            excel_file2 = pd.ExcelFile(file2)
        else:
            excel_file1 = pd.read_csv(file1)
            excel_file2 = pd.read_csv(file2)

        # Get sheet names
        if file1.type in ['application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', 'application/vnd.ms-excel']:
            sheet_names1 = excel_file1.sheet_names
            sheet_names2 = excel_file2.sheet_names

            # Select sheet numbers
            sheet_number1 = st.selectbox("Select sheet number for File 1", range(1, len(sheet_names1) + 1))
            sheet_number2 = st.selectbox("Select sheet number for File 2", range(1, len(sheet_names2) + 1))

            # Read selected sheets
            df1 = excel_file1.parse(sheet_names1[sheet_number1 - 1])
            df2 = excel_file2.parse(sheet_names2[sheet_number2 - 1])
        else:
            df1 = excel_file1
            df2 = excel_file2

        # Display column names
        st.write("Columns in File 1:", df1.columns)
        st.write("Columns in File 2:", df2.columns)

    # Column name input
    if df1 is not None:
        column_name1 = st.selectbox("Select first column to match", df1.columns)
        column_name2 = st.selectbox("Select second column to match (optional)", df1.columns)

        # Button to generate output
        if st.button("Generate Output"):
            # Check if dataframes are loaded
            if df1 is not None and df2 is not None:
                # Trim whitespace from column names
                df1.columns = df1.columns.str.strip()
                df2.columns = df2.columns.str.strip()

                # Remove duplicates from both files based on the matching column
                df1 = df1.drop_duplicates(subset=column_name1, keep='first')
                df2 = df2.drop_duplicates(subset=column_name1, keep='first')

                # Merge dataframes on selected columns
                if column_name1 != column_name2:
                    merged_df = pd.merge(df1, df2, on=[column_name1], how='left', suffixes=('_x', '_y'))
                else:
                    merged_df = pd.merge(df1, df2, on=column_name1, how='left', suffixes=('_x', '_y'))

                # Fill Rate from File 2 where available
                merged_df['Rate'] = merged_df['Rate_y'].fillna(merged_df['Rate_x'])

                # Drop unnecessary columns
                merged_df = merged_df.drop(['Rate_x', 'Rate_y'], axis=1, errors='ignore')

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