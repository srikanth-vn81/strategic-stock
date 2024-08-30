import pandas as pd
import re
import streamlit as st
from io import BytesIO

# Set the page layout to wide mode
st.set_page_config(layout="wide")

# Custom CSS to style the app with a professional font
st.markdown("""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Roboto:wght@400;500;700&display=swap');
    
    html, body, [class*="css"]  {
        font-family: 'Roboto', sans-serif;
    }
    
    .main {
        background-color: #e8f5e9;  /* Light green background for better contrast */
    }
    
    .css-1d391kg {
        background-color: #e8f5e9;  /* Same light green background for sidebar */
    }
    
    .stButton>button {
        background-color: #4CAF50;
        color: white;
        padding: 10px 20px;
        border: none;
        border-radius: 4px;
        cursor: pointer;
        font-size: 16px;
    }
    
    .stButton>button:hover {
        background-color: #45a049;
    }
    
    .stFileUploader>div {
        background-color: #ffffff;
        padding: 20px;
        border-radius: 4px;
        border: 1px solid #cccccc;
        box-shadow: 2px 2px 5px rgba(0,0,0,0.1);
    }
    
    .stTextInput>div>input {
        border-radius: 4px;
        border: 1px solid #cccccc;
    }
    
    .stSidebar h2 {
        color: #4CAF50;
    }
    
    .stSidebar h3 {
        color: #4CAF50;
    }
    
    h1, h2, h3, h4, h5, h6, p, div {
        font-family: 'Roboto', sans-serif;
    }
    </style>
    """, unsafe_allow_html=True)

# Streamlit interface for uploading files and processing data
st.markdown("<h1 style='color:#4CAF50;'>📦 Strategic Stock</h1>", unsafe_allow_html=True)

# Sidebar for file uploads with icons
st.sidebar.markdown("<h2>📁 File Uploads</h2>", unsafe_allow_html=True)

# Upload Norm Sensing Excel file
norm_sensing_file = st.sidebar.file_uploader("Upload Norm Sensing Excel File", type=["xlsx"], key="norm_sensing_file")

# Upload RM Macro Excel file
rm_macro_file = st.sidebar.file_uploader("Upload RM Macro Excel File", type=["xlsx"], key="rm_macro_file")

# Display all unique programs in the data and let the user select multiple
if norm_sensing_file:
    df = pd.read_excel(norm_sensing_file, sheet_name=0)
    available_programs = df['Program'].unique()
    selected_programs = st.multiselect("Select Program(s)", options=available_programs, default=available_programs)
else:
    selected_programs = []

# Move the 'Run' button to the main page
if st.button("Run"):
    if norm_sensing_file and rm_macro_file:
        st.success("Both files uploaded successfully. Processing the data...")

        # Filter the DataFrame based on the selected programs
        filtered_df = df[df['Program'].isin(selected_programs)]

        # Load the RM Macro data from the first sheet
        df1 = pd.read_excel(rm_macro_file, sheet_name=0)

        # Filter the DataFrame where 'PROC_GRP' is 'ELS' or 'LAC'
        df1 = df1[df1['PROC_GRP'].isin(['ELS', 'LAC'])].copy()

        # Replace NaN values in the column 'l' with 0 and extract numerical values
        df1.loc[:, 'l'] = df1['l'].fillna('0')
        df1.loc[:, 'l'] = df1['l'].apply(lambda x: re.search(r'\d+', str(x)).group() if re.search(r'\d+', str(x)) else '0')

        # Perform group by operation and calculate the mean of CONSUMPTION
        grouped_norm = df1.groupby(['l', 'GMT colour', 'PROC_GRP'])['CONSUMPTION'].mean().reset_index()

        # Convert left_on and right_on fields to string and remove extra spaces
        filtered_df.loc[:, 'Style'] = filtered_df['Style'].astype(str).str.strip()
        filtered_df.loc[:, 'GMT Color'] = filtered_df['GMT Color'].astype(str).str.strip()

        grouped_norm.loc[:, 'l'] = grouped_norm['l'].astype(str).str.strip()
        grouped_norm.loc[:, 'GMT colour'] = grouped_norm['GMT colour'].astype(str).str.strip()

        # Perform the left join on the specified columns
        merged_df = pd.merge(filtered_df, grouped_norm, how='left', left_on=['Style', 'GMT Color'], right_on=['l', 'GMT colour'])

        # Force conversion to datetime and handle errors
        date_columns = ['Start Date', 'End Date', 'Ramp up date', 'Ramp down date']

        for col in date_columns:
            if col in merged_df.columns:
                merged_df[col] = pd.to_datetime(merged_df[col], errors='coerce')

        # Optional: Format the datetime columns to short date format (YYYY-MM-DD)
        for col in date_columns:
            if col in merged_df.columns:
                merged_df[col] = merged_df[col].dt.strftime('%Y-%m-%d')

        # Add calculated fields
        merged_df['No of Pieces'] = merged_df['Concluded Norms - Post discussion'] / merged_df['CF']
        merged_df['Requirement'] = merged_df['No of Pieces'] * merged_df['CONSUMPTION']

        # Display the final DataFrame with calculated fields
        st.write("Final Merged DataFrame with Calculated Fields:")
        st.dataframe(merged_df)

        # Save the final DataFrame to an in-memory buffer
        output = BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            merged_df.to_excel(writer, index=False)

        # Ensure the buffer's pointer is at the start so it can be read from the beginning
        output.seek(0)

        # Provide the download link directly after processing
        st.download_button(
            label="Download Merged Data as Excel",
            data=output,
            file_name="strategicstock.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    else:
        st.info("Please upload both the Norm Sensing and RM Macro Excel files to proceed.")
