import streamlit as st
import pandas as pd
import numpy as np
import random
from io import BytesIO
from PIL import Image
import streamlit as st
import pandas as pd
from io import BytesIO
from PIL import Image as PILImage
import base64
# Streamlit interface

st.title('Vessel Waste Data Mapping Tool')

# Upload Excel file button
uploaded_file = st.file_uploader("Upload Vessel Waste Tracker Excel File", type=["xlsx"])

if uploaded_file:
    # Load the Excel file from the upload
    sheets_to_extract = [
        'TBC BADRINATH', 'TBC KAILASH', 'SSL KRISHNA', 'SSL VISAKHAPATNAM',
        'SSL BRAMHAPUTRA', 'SSL MUMBAI', 'SSL GANGA', 'SSL BHARAT',
        'SSL SABARIMALAI', 'SSL GUJARAT', 'SSL DELHI', 'SSL KAVERI',
        'SSL GODAVARI', 'SSL THAMIRABARANI'
    ]
    
    # Load the specified sheets into a dictionary of DataFrames
    all_sheets = pd.read_excel(uploaded_file, sheet_name=sheets_to_extract, header=[0, 1, 2])

    # Function to extract and unpivot garbage data
    def extract_and_unpivot_garbage_data(sheets_dict, garbage_type):
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
        final_df.rename(columns={"Date":"Res_Date","Sub Section":"Source Sub Type","Amount":"Activity","Sheet Name":"Facility"},inplace=True)
        final_df['CF Standard'] = "IPCCC"
        final_df['Activity Unit'] = "m3"
        final_df['Gas'] = "CO2"
        final_df.replace({"m3", "Total"}, np.nan, inplace=True)
        final_df.dropna(subset=["Date"], inplace=True)
        final_df.drop_duplicates(inplace=True)
        return final_df

    # Extract data for different garbage types
    garbage_incinerated_df = extract_and_unpivot_garbage_data(all_sheets, 'Garbage Incinerated')
    garbage_landed_ashore_df = extract_and_unpivot_garbage_data(all_sheets, 'Garbage Landed Ashore')
    garbage_generated_df = extract_and_unpivot_garbage_data(all_sheets, 'Garbage Generated')

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

# Mapping part of the code can be added here
#-----------------------------------------------------------------------------------------------------------------------------
def process_file(uploaded_file):
    # Load the uploaded Excel file
    dfs = pd.read_excel(uploaded_file, header=[0, 1], sheet_name=None)
    
    # Process each sheet separately
    processed_dfs = []
    for sheet_name, df in dfs.items():
        # Sort the MultiIndex to avoid UnsortedIndexError
        df = df.sort_index(axis=1)

        # Resetting the index to handle MultiIndex columns correctly
        df.columns = df.columns.droplevel(0)

        # Slicing the DataFrame to select specific rows (June 2023 to March 2024)
        df_selected = df.loc[2:13, :]

        # Selecting the last 8 columns (adjust as needed)
        df_selected = df_selected.iloc[:, -8:]

        # Adding a column for the sheet name
        df_selected['Sheet Name'] = sheet_name

        # Append the processed DataFrame to the list
        processed_dfs.append(df_selected)

    # Combine all DataFrames into a single DataFrame
    combined_df = pd.concat(processed_dfs, ignore_index=True)
    combined_df = combined_df.dropna(thresh=5)

    # Debug: Print the columns of combined_df
    print("Combined DataFrame columns:", combined_df.columns)

    # Assuming the 'Month' is already a column and unpivot the remaining columns
    # Adjust the column name if necessary
    columns_to_unpivot = [col for col in combined_df.columns if col != 'Month']  # Adjust if 'Month' has a different name

    if 'Month' not in combined_df.columns:
        st.warning("'Month' column not found. Please check the input data.")
        return None

    unpivoted_df = pd.melt(
        combined_df,
        id_vars=['Sheet Name', 'Month'],  # Keep 'Sheet Name' and 'Month' as identifier columns
        value_vars=columns_to_unpivot,    # Columns to unpivot
        var_name='Waste Type',            # New column for the type of waste or data type
        value_name='Waste Amount'         # New column for the values
    )

    # Rename columns and format the DataFrame
    column_mapping = {
        'Waste Type': 'Source Sub Type',
        'Sheet Name': 'Facility',
        'Month': 'Res_Date',
        'Waste Amount': 'Activity'
    }

    unpivoted_df = unpivoted_df.dropna(subset="Waste Amount")
    unpivoted_df.rename(columns=column_mapping, inplace=True)

    # Add the new columns as requested
    unpivoted_df['CF Standard'] = "IPCCC"
    unpivoted_df['Activity Unit'] = "m3"
    unpivoted_df['Gas'] = "CO2"

    # Reorder columns
    new_order = ['Res_Date', 'Facility', 'Source Sub Type', 'Activity', 'Activity Unit', 'CF Standard', 'Gas']
    unpivoted_df = unpivoted_df[new_order]

    return unpivoted_df

# Function to convert DataFrame to Excel and return as a downloadable file
def to_excel(df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False)
    processed_data = output.getvalue()
    return processed_data

# Function to resize and return an image

if uploaded_file is not None:
    # Process the uploaded file
    st.subheader('Oil Generated')
    unpivoted_df = process_file(uploaded_file)
    
    # Display success message and show the processed DataFrame

    st.dataframe(unpivoted_df)

    # Button to download the processed file
    processed_data = to_excel(unpivoted_df)
    st.download_button(
        label="Download Processed Excel",
        data=processed_data,
        file_name='Processed_Oil_Generated.xlsx',
        mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    )
