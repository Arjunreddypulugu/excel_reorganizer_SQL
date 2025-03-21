import pandas as pd
from collections import defaultdict

def process_excel(input_path, ami_path, output_path):
    # Load input and reference Excel files
    input_df = pd.read_excel(input_path, sheet_name="FCC Placer MSW - Spares List")
    ami_df = pd.read_excel(ami_path, sheet_name="EquipmentDB")

    # Extract necessary columns
    input_df = input_df[['Serial', 'Total qty', 'Spare qty', 'Item no.', 'Description', 'Unit Price ($)']]
    ami_df = ami_df[['SerialNumber', 'Model']]
    ami_df.dropna(subset=['SerialNumber'], inplace=True)

    # Map Serial → Model
    serial_to_model = {
        row['SerialNumber']: row['Model'] if pd.notna(row['Model']) else "MODEL MISSING"
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

    for _, row in input_df.iterrows():
        serial = row['Serial']

        # Detect new machine (change in serial)
        if serial != last_serial:
            last_serial = serial
            current_model = serial_to_model.get(serial, "MODEL MISSING")
            continue  # Skip the header row itself

        if current_model:
            item_no = row['Item no.']
            description = row['Description']
            unit_price = row['Unit Price ($)']
            total_qty = pd.to_numeric(row['Total qty'], errors='coerce') or 0
            spare_qty = pd.to_numeric(row['Spare qty'], errors='coerce') or 0

            if pd.notna(item_no) and pd.notna(description):
                part = model_spares[current_model][item_no]
                part['Total qty'] += total_qty
                part['Spare qty'] += spare_qty
                part['Description'] = description
                part['Unit Price ($)'] = unit_price

    # Output formatting with 'Model' column added
    for model, parts in model_spares.items():
        output_rows.append([model, '', '', '', '', ''])  # Model in first column
        for item_no, data in parts.items():
            output_rows.append([
                '',  # Model column is empty for spare rows
                data['Total qty'],
                data['Spare qty'],
                item_no,
                data['Description'],
                data['Unit Price ($)']
            ])

    # Save output
    output_df = pd.DataFrame(output_rows, columns=['Model', 'Total qty', 'Spare qty', 'Item no.', 'Description', 'Unit Price ($)'])
    with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
        output_df.to_excel(writer, index=False, sheet_name='Formatted Output')
        worksheet = writer.sheets['Formatted Output']
        for col in worksheet.columns:
            max_length = 0
            col_letter = col[0].column_letter
            for cell in col:
                try:
                    if cell.value:
                        max_length = max(max_length, len(str(cell.value)))
                except:
                    pass
            worksheet.column_dimensions[col_letter].width = max_length + 2

    print(f"✅ Output saved to: {output_path}")

# ==== Modify these paths as needed ====
output_generated_path = r"C:\Users\APulugu\OneDrive - VAN DYK BALER\Desktop\generated_output.xlsx"
process_excel(
    r"C:\Users\APulugu\OneDrive - VAN DYK BALER\Desktop\input.xlsx",
    r"C:\Users\APulugu\OneDrive - VAN DYK BALER\Desktop\AMI.xlsx",
    output_generated_path
)
