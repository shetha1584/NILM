import pandas as pd

# Load the Excel file
df = pd.read_excel('machine capacity.xlsx')

# Strip column names of spaces and uppercase them for consistency
df.columns = df.columns.str.strip().str.upper()

# Add a 'QUANTITY' column with default value 1 (or customize below)
df['QUANTITY'] = 1

# Optional: Customize quantities based on known machines
# (uncomment and modify as needed)
# df.loc[df['MACHINE NAME'].str.contains('TUBE LIGHT', case=False), 'QUANTITY'] = 45
# df.loc[df['MACHINE NAME'].str.contains('COMPUTER', case=False), 'QUANTITY'] = 10
# df.loc[df['MACHINE NAME'].str.contains('FAN', case=False), 'QUANTITY'] = 6

# Save to new Excel file
df.to_excel('machine_capacity_with_quantity.xlsx', index=False)

print("✅ 'QUANTITY' column added and file saved as 'machine_capacity_with_quantity.xlsx'")
