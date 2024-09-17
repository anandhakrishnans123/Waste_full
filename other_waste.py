import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO

# Streamlit interface
st.title('Vessel Waste Data Mapping Tool')

# Upload Excel file button for Waste Tracker and Template
uploaded_file = st.file_uploader("Upload Vessel Waste Tracker Excel File", type=["xlsx"])
template_file_path = "Waste-Sample.xlsx"

if uploaded_file:
    # Load the template file to extract the correct headers
    template_df = pd.read_excel(template_file_path, sheet_name=0)
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
        matching_columns = [col for col in df.columns if col in correct_headers]
        df = df[matching_columns]
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
        final_df['Res_Date'] = pd.to_datetime(final_df['Res_Date']).dt.date

        # Match and reorder the columns with the template
        final_df = match_and_reorder_columns(final_df, correct_headers)
        final_df.dropna(subset=["Res_Date"], inplace=True)
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

    # Oil Generated Processing Function
    def process_oil_generated_file(uploaded_file, correct_headers):
        dfs = pd.read_excel(uploaded_file, header=[0, 1], sheet_name=None)
        processed_dfs = []

        for sheet_name, df in dfs.items():
            df = df.sort_index(axis=1)
            df.columns = df.columns.droplevel(0)
            df_selected = df.loc[2:13, :]
            df_selected = df_selected.iloc[:, -8:]
            df_selected['Sheet Name'] = sheet_name
            processed_dfs.append(df_selected)

        combined_df = pd.concat(processed_dfs, ignore_index=True)
        combined_df = combined_df.dropna(thresh=5)

        columns_to_unpivot = [col for col in combined_df.columns if col != 'Month']
        if 'Month' not in combined_df.columns:
            st.warning("'Month' column not found. Please check the input data.")
            return None

        unpivoted_df = pd.melt(
            combined_df,
            id_vars=['Sheet Name', 'Month'],
            value_vars=columns_to_unpivot,
            var_name='Waste Type',
            value_name='Waste Amount'
        )

        column_mapping = {
            'Waste Type': 'Source Sub Type',
            'Sheet Name': 'Facility',
            'Month': 'Res_Date',
            'Waste Amount': 'Activity'
        }

        unpivoted_df = unpivoted_df.dropna(subset="Waste Amount")
        unpivoted_df.rename(columns=column_mapping, inplace=True)
        unpivoted_df['CF Standard'] = "IPCCC"
        unpivoted_df['Activity Unit'] = "m3"
        unpivoted_df['Gas'] = "CO2"
        unpivoted_df = match_and_reorder_columns(unpivoted_df, correct_headers)

        return unpivoted_df

    # Function to convert DataFrame to Excel and return as a downloadable file
    def to_excel(df):
        output = BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            df.to_excel(writer, index=False)
        processed_data = output.getvalue()
        return processed_data

    # Oil Generated Section in Streamlit App
    st.subheader('Oil Generated Data')
    
    # Process the uploaded file for oil generated
    unpivoted_oil_df = process_oil_generated_file(uploaded_file, correct_headers)
    
    # Display the processed DataFrame in the Streamlit app
    st.dataframe(unpivoted_oil_df)

    # Convert the processed DataFrame to Excel and provide download button
    processed_oil_data = to_excel(unpivoted_oil_df)
    st.download_button(
        label="Download Processed Oil Generated Data",
        data=processed_oil_data,
        file_name='Processed_Oil_Generated.xlsx',
        mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    )
