import pandas as pd

# Read the Excel file
file_path = 'energy_meter_1_June15_16.xlsx'
all_sheets = pd.read_excel(file_path, sheet_name=None)  # Reads all sheets into a dictionary

# Combine all sheets into one DataFrame
combined_df = pd.concat(all_sheets.values(), ignore_index=True)

# Write the combined data to a new Excel file
output_file = 'combined_energy_meter_data.xlsx'
combined_df.to_excel(output_file, index=False, sheet_name='Combined Data')

print(f"All sheets have been combined and saved to {output_file}")