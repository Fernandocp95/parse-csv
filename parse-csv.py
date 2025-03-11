import pandas as pd
import os
import re
import streamlit as st

# Check for openpyxl
try:
    import openpyxl
except ImportError:
    st.error("Missing dependency: 'openpyxl'. Please install it using `pip install openpyxl`.")
    st.stop()

def extract_month_from_filename(filename):
    """
    Extracts the month from the filename if present.
    """
    months = {
        "january": "January", "february": "February", "march": "March", "april": "April",
        "may": "May", "june": "June", "july": "July", "august": "August",
        "september": "September", "october": "October", "november": "November", "december": "December"
    }
    filename = filename.lower()
    for eng in months.keys():
        if eng in filename:
            return months[eng]
    return None  # If no month is found in the filename

def standardize_columns(df, column_mapping, filename):
    """
    Renames columns based on the mapping dictionary and adds missing columns.
    """
    df = df.rename(columns=column_mapping)
    for col in column_mapping.values():
        if col not in df.columns:
            df[col] = None  # Add missing columns with empty values
    
    # Add mandatory new columns
    df["Published Date"] = extract_month_from_filename(filename)
    df["Owner"] = None
    df["Project"] = None
    df["Expiry Warranty"] = "1 year"  # Default value
    
    return df[list(column_mapping.values()) + ["Published Date", "Owner", "Project", "Expiry Warranty"]]  # Reorder columns

def process_files(uploaded_files, column_mapping, output_file="output.xlsx"):
    """
    Processes uploaded Excel files.
    """
    if not uploaded_files:
        st.warning("No files uploaded. Please upload Excel files.")
        return
    
    standardized_dfs = []
    
    for uploaded_file in uploaded_files:
        df = pd.read_excel(uploaded_file, engine="openpyxl")  # Use openpyxl explicitly
        df = standardize_columns(df, column_mapping, uploaded_file.name)
        standardized_dfs.append(df)
    
    final_df = pd.concat(standardized_dfs, ignore_index=True)
    final_df.to_excel(output_file, index=False, engine="openpyxl")
    st.success(f"Unified file saved as {output_file}")
    
    # Provide a download link
    with open(output_file, "rb") as f:
        st.download_button("Download Processed File", f, file_name=output_file)

# Streamlit UI
st.title("Excel File Standardization Tool")
st.write("Upload your Excel files to standardize the column structure.")

# File uploader
uploaded_files = st.file_uploader("Choose Excel files", type=["xlsx"], accept_multiple_files=True)

# Define a column mapping dictionary
column_mapping = {
    "Site": "Site",
    "Domain": "Site",
    "Link origin": "Site",
    "Market": "Market",
    "DA": "DA",
    "DR": "DR",
    "Traffic": "Traffic",
    "Price": "Price",
    "Price €": "Price",
    "Link costs": "Price",
    "Status": "Status",
    "LL approved": "Status",
    "Publish": "Publish",
    "Anchor Text": "Anchor Text",
    "Link text": "Anchor Text",
    "Target URL": "Target URL",
    "Link target": "Target URL",
    "Live URL": "Live URL / Published Link",
    "Published Link": "Live URL / Published Link",
    "Project": "Project",
    "Date": "Published Date"
}

if st.button("Process Files"):
    process_files(uploaded_files, column_mapping)
