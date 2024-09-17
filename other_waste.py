import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO

# Streamlit interface
st.title('Vessel Waste Data Mapping Tool')

# Upload Excel file button for Waste Tracker and Template
uploaded_file = st.file_uploader("Upload Vessel Waste Tracker Excel File", type=["xlsx"])
template_file = st.file_uploader("Upload Template Excel File", type=["xlsx"])

if uploaded_file and template_file:
    # Load the template file to extract the correct headers
    template_df = pd.read_excel(template_file, sheet_name=0)
    correct_headers = list(template_df.columns)

    # Load the waste tracker file with specified sheets
    sheets_to_extract = [
        'TBC BADRINATH', 'TBC KAILASH', 'SSL KRISHNA', 'SSL VISAKHAPATNAM',
        'SSL BRAMHAPUTRA', 'SSL MUMBAI', 'SSL GANGA', 'SSL BHARAT',
        'SSL SABARIMALAI', 'SSL GUJARAT', 'SSL DELHI', 'SSL KAVERI',
        'SSL GODAVARI', 'SSL THAMIRABARANI'
    ]
    
    all_sheets = pd.read_excel(uploaded_file, sheet_name=sheets_to_extract, header=[0, 1, 2])

    # Function to match and reorder columns
    def match_and_reorder_columns(df, correct_headers):
        # Find common columns between DataFrame and template
        matching_columns = [col for col in df.columns if col in correct_headers]
        
        # Filter the DataFrame to keep only the matching columns
        df = df[matching_columns]
        
        # Reorder the columns based on the template
        df = df.reindex(columns=correct_headers)
        
        return df

    # Function to extract and unpivot garbage data
    def extract_and_unpivot_garbage_data(sheets_dict, garbage_type, correct_headers):
        combined_data = []
        
        for sheet_name, vessel_df in sheets_dict.items():
            date_column = pd.to_datetime(vessel_df.iloc[:, 0], errors='coerce')
            first_valid_index = date_column.first_valid_index()

            if first_valid_index is None:
                continue  # Skip sheets without valid dates

            date_column = date_column.loc[first_valid_index:].reset_index(drop=True)
            data_rows = vessel_df.loc[first_valid_index:].reset_index(drop=True)

            garbage_columns = [
                col for col in data_rows.columns
                if col[0] == 'Garbage Record Book' and col[1].strip() == garbage_type
            ]

            if not garbage_columns:
                st.write(f"No '{garbage_type}' data found in sheet {sheet_name}")
                continue

            for col in garbage_columns:
                sub_section = col[2]
                temp_df = pd.DataFrame({
                    'Date': date_column,
                    'Sub Section': sub_section,
                    'Amount': data_rows[col].reset_index(drop=True),
                    'Sheet Name': sheet_name
                })
                combined_data.append(temp_df)

        final_df = pd.concat(combined_data, ignore_index=True)
        final_df.rename(columns={
            "Date": "Res_Date",
            "Sub Section": "Source Sub Type",
            "Amount": "Activity",
            "Sheet Name": "Facility"
        }, inplace=True)

        # Add additional columns
        final_df['CF Standard'] = "IPCCC"
        final_df['Activity Unit'] = "m3"
        final_df['Gas'] = "CO2"

        # Match and reorder the columns with the template
        final_df = match_and_reorder_columns(final_df, correct_headers)
        
        return final_df

    # Extract data for different garbage types
    garbage_incinerated_df = extract_and_unpivot_garbage_data(all_sheets, 'Garbage Incinerated', correct_headers)
    garbage_landed_ashore_df = extract_and_unpivot_garbage_data(all_sheets, 'Garbage Landed Ashore', correct_headers)
    garbage_generated_df = extract_and_unpivot_garbage_data(all_sheets, 'Garbage Generated', correct_headers)

    # Function to convert DataFrame to Excel format and return a BytesIO object
    def convert_df_to_excel(df):
        output = BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            df.to_excel(writer, index=False)
        return output.getvalue()

    # Display DataFrames and provide download buttons
    st.subheader('Garbage Incinerated Data')
    st.dataframe(garbage_incinerated_df)
    st.download_button(
        label="Download Garbage Incinerated Data",
        data=convert_df_to_excel(garbage_incinerated_df),
        file_name='Garbage_Incinerated_Data.xlsx'
    )

    st.subheader('Garbage Landed Ashore Data')
    st.dataframe(garbage_landed_ashore_df)
    st.download_button(
        label="Download Garbage Landed Ashore Data",
        data=convert_df_to_excel(garbage_landed_ashore_df),
        file_name='Garbage_Landed_Ashore_Data.xlsx'
    )

    st.subheader('Garbage Generated Data')
    st.dataframe(garbage_generated_df)
    st.download_button(
        label="Download Garbage Generated Data",
        data=convert_df_to_excel(garbage_generated_df),
        file_name='Garbage_Generated_Data.xlsx'
    )
