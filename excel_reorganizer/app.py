import streamlit as st
import pandas as pd
from collections import defaultdict
from io import BytesIO
import os

def process_excel(uploaded_file, ami_path):
    excel_input = pd.ExcelFile(uploaded_file)
    ami_df = pd.read_excel(ami_path, sheet_name="EquipmentDB")
    ami_df = ami_df[['SerialNumber', 'Model', 'EquipmentType']]
    ami_df.dropna(subset=['SerialNumber'], inplace=True)

    # Create Serial → Model and EquipmentType mappings
    serial_to_model = {}
    serial_to_type = {}
    for _, row in ami_df.iterrows():
        serial = row['SerialNumber']
        model = row['Model'] if pd.notna(row['Model']) else "MODEL MISSING"
        equip_type = row['EquipmentType'] if pd.notna(row['EquipmentType']) else "TYPE MISSING"
        serial_to_model[serial] = model
        serial_to_type[serial] = equip_type

    model_to_type = {model: serial_to_type[serial] for serial, model in serial_to_model.items()}

    combined_spares = defaultdict(lambda: defaultdict(lambda: {
        'Description': '',
        'Unit Price ($)': None,
        'Total qty': 0,
        'Spare qty': 0
    }))

    # Process each sheet
    for sheet_name in excel_input.sheet_names:
        input_df = pd.read_excel(excel_input, sheet_name=sheet_name)
        expected_cols = ['Serial', 'Total qty', 'Spare qty', 'Item no.', 'Description', 'Unit Price ($)']
        input_df.columns = [col.strip() for col in input_df.columns]

        if not set(expected_cols).issubset(set(input_df.columns)):
            continue  # Skip this sheet if essential columns are missing

        input_df = input_df[expected_cols]
        last_serial = None
        current_model = None

        for _, row in input_df.iterrows():
            serial = row['Serial']
            if serial != last_serial:
                last_serial = serial
                current_model = serial_to_model.get(serial, "MODEL MISSING")
                continue  # Skip header row

            if current_model:
                item_no = row['Item no.']
                description = row['Description']
                unit_price = row['Unit Price ($)']
                total_qty = pd.to_numeric(row['Total qty'], errors='coerce') or 0
                spare_qty = pd.to_numeric(row['Spare qty'], errors='coerce') or 0

                if pd.notna(item_no) and str(item_no).strip().upper() != 'TBD' and pd.notna(description):
                    item_str = str(item_no).strip()
                    sp = combined_spares[current_model][item_str]
                    sp['Description'] = description
                    sp['Unit Price ($)'] = unit_price
                    sp['Total qty'] += total_qty
                    sp['Spare qty'] += spare_qty

    # Build output
    output_rows = []
    grouped_models = sorted(
        [(model_to_type.get(model, "TYPE MISSING"), model) for model in combined_spares.keys()],
        key=lambda x: (x[0], x[1])
    )

    for equip_type, model in grouped_models:
        output_rows.append([equip_type, model, '', '', '', '', ''])
        parts = combined_spares[model]

        # Sort spares inside the model by Description
        for item_no, part in sorted(parts.items(), key=lambda x: x[1]['Description']):
            output_rows.append([
                '', '', part['Total qty'], part['Spare qty'],
                item_no, part['Description'], part['Unit Price ($)']
            ])

    output_df = pd.DataFrame(output_rows, columns=[
        'Equipment Type', 'Model', 'Total qty', 'Spare qty', 'Item no.', 'Description', 'Unit Price ($)'
    ])

    # Return Excel output
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        output_df.to_excel(writer, index=False, sheet_name='Formatted Output')
    output.seek(0)
    return output


# ===== Streamlit Interface =====
st.title("🔧 Spare Parts Formatter")

uploaded_file = st.file_uploader("Upload the input Excel file", type=["xlsx"])

if uploaded_file:
    with st.spinner("Processing..."):
        ami_path = os.path.join(os.path.dirname(__file__), "AMI.xlsx")
        output_excel = process_excel(uploaded_file, ami_path)

    st.success("✅ File processed successfully!")
    st.download_button(
        label="📥 Download Output Excel",
        data=output_excel,
        file_name="formatted_output.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
