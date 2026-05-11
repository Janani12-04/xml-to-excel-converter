import streamlit as st
import pandas as pd
import xml.etree.ElementTree as ET
from io import BytesIO

# Page Configuration
st.set_page_config(page_title="XML to Excel Converter")

# Title
st.title("XML to Excel Converter")

# File Upload
uploaded_file = st.file_uploader(
    "Upload XML File",
    type=["xml"]
)

# Process XML
if uploaded_file is not None:

    try:
        # Read XML File
        tree = ET.parse(uploaded_file)
        root = tree.getroot()

        data = []

        # Extract Data
        for child in root:

            row_data = {}

            for element in child:
                row_data[element.tag] = element.text

            data.append(row_data)

        # Convert to DataFrame
        df = pd.DataFrame(data)

        # Success Message
        st.success("XML File Converted Successfully!")

        # Show Data
        st.subheader("Preview")
        st.dataframe(df)

        # Create Excel File
        output = BytesIO()

        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df.to_excel(
                writer,
                index=False,
                sheet_name='Converted_Data'
            )

        excel_data = output.getvalue()

        # Download Button
        st.download_button(
            label="Download Excel File",
            data=excel_data,
            file_name="converted_file.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.error(f"Error: {e}")