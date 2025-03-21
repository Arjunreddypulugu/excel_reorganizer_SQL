import streamlit as st
import pandas as pd
import os
from collections import defaultdict
from io import BytesIO

# Define the path to AMI.xlsx dynamically
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
AMI_PATH = os.path.join(BASE_DIR, "AMI.xlsx")

# Debug: Check if AMI.xlsx exists
if not os.path.exists(AMI_PATH):
    st.error(f"❌ AMI.xlsx not found at: {AMI_PATH}")
    st.stop()  # Stop execution if file is missing
else:
    st.write(f"✅ AMI.xlsx found at: {AMI_PATH}")

def process_excel(uploaded_file, ami_path):
    input_df = pd.read_excel(uploaded_file, sheet_name="FCC Placer MSW - Spares List")
    ami_df = pd.read_excel(ami_path, sheet_name="EquipmentDB")

    # Extract necessary columns
    input_df = input_df[['Serial', 'Total qty', 'Spare qty', 'Item no.', 'Description', 'Unit Price ($)']]
    ami_df = ami_df[['SerialNumber', 'Model', 'EquipmentType']]
    ami_df.dropna(subset=['SerialNumber'], inplace=True)

    # Map Serial → Model & EquipmentType
    serial_to_model = {
        row['SerialNumber']: (row['Model'] if pd.notna(row['Model']) else "MODEL MISSING", 
                              row['EquipmentType'] if pd.notna(row['EquipmentType']) else "EQUIPMENT TYPE MISSING")
        for _, row in ami_df.iterrows()
    }

    # Data aggregation
    output_rows = []
    model_spares = defaultdict(lambda: defaultdict(lambda: {
        "Total qty": 0,
        "Spare qty": 0,
        "Description": "",
        "Unit Price ($)": None
    }))

    last_serial = None
    current_model = None
    current_equipment = None

    for _, row in input_df.iterrows():
        serial = row['Serial']

        # Detect new machine (change in serial)
        if serial != last_serial:
            last_serial = serial
            current_model, current_equipment = serial_to_model.get(serial, ("MODEL MISSING", "EQUIPMENT TYPE MISSING"))
            continue  # Skip the header row itself

        if current_model:
            item_no = row['Item no.']
            description = row['Description']
            unit_price = row['Unit Price ($)']
            total_qty = pd.to_numeric(row['Total qty'], errors='coerce') or 0
            spare_qty = pd.to_numeric(row['Spare qty'], errors='coerce') or 0

            if pd.notna(item_no) and pd.notna(description) and str(item_no).strip().upper() != "TBD":
                part = model_spares[(current_equipment, current_model)][item_no]
                part['Total qty'] += total_qty
                part['Spare qty'] += spare_qty
                part['Description'] = description
                part['Unit Price ($)'] = unit_price

    # Output formatting with 'Equipment Type' and 'Model' columns added
    for (equipment, model), parts in sorted(model_spares.items()):
        output_rows.append([equipment, model, '', '', '', ''])  # Equipment Type & Model in first columns
        for item_no, data in sorted(parts.items(), key=lambda x: x[1]['Description']):  # Sort spares by description
            output_rows.append([
                '', '',  # Empty Equipment Type & Model columns for spare rows
                data['Total qty'],
                data['Spare qty'],
                item_no,
                data['Description'],
                data['Unit Price ($)']
            ])

    # Save output
    output_df = pd.DataFrame(output_rows, columns=['Equipment Type', 'Model', 'Total qty', 'Spare qty', 'Item no.', 'Description', 'Unit Price ($)'])
    output_buffer = BytesIO()
    with pd.ExcelWriter(output_buffer, engine='openpyxl') as writer:
        output_df.to_excel(writer, index=False, sheet_name='Formatted Output')

    return output_buffer

# Streamlit App UI
st.title("Excel Reorganizer")
st.write("Upload an input Excel file to generate a formatted output.")

uploaded_file = st.file_uploader("Choose an Excel file", type=["xlsx"])

if uploaded_file:
    st.write("✅ File uploaded successfully!")
    
    # Process the uploaded file
    output_buffer = process_excel(uploaded_file, AMI_PATH)
    
    # Provide download link
    st.download_button(
        label="📥 Download Processed Excel",
        data=output_buffer.getvalue(),
        file_name="processed_output.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
