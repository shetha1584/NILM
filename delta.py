import pandas as pd

# Load original Excel file
file_path = "energy_meter_9_June15_16.xlsx"
excel_file = pd.ExcelFile(file_path)

# Create a new Excel file with 'delta' column added to each sheet
output_path = "energy_meter_with_delta.xlsx"

with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
    for sheet_name in excel_file.sheet_names:
        # Read and clean column names
        df = excel_file.parse(sheet_name)
        df.columns = df.columns.str.strip()

        # Check and compute delta if 'kWh' column exists
        if 'kWh' not in df.columns:
            raise ValueError(f"'kWh' column not found in sheet '{sheet_name}'")

        df['delta'] = df['kWh'].diff().fillna(0)

        # Write modified data to output file
        df.to_excel(writer, sheet_name=sheet_name, index=False)

print(f"Delta-added Excel saved to: {output_path}")
