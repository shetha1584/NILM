import pandas as pd

# Load the original Excel file
file_path = "energy_meter_9_June15_16.xlsx"
excel_file = pd.ExcelFile(file_path)

# Output file path
output_path = "energy_meter_with_nonzero_delta.xlsx"

with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
    for sheet_name in excel_file.sheet_names:
        # Read and clean the sheet
        df = excel_file.parse(sheet_name)
        df.columns = df.columns.str.strip()

        # Ensure 'kWh' column exists
        if 'kWh' not in df.columns:
            raise ValueError(f"'kWh' column not found in sheet '{sheet_name}'")

        # Compute delta
        df['delta'] = df['kWh'].diff().fillna(0)

        # Remove rows where delta == 0
        df_filtered = df[df['delta'] != 0]

        # Write to new Excel file
        df_filtered.to_excel(writer, sheet_name=sheet_name, index=False)

print(f"Filtered Excel saved to: {output_path}")
