import pandas as pd
from itertools import combinations

# Load the Excel files
machine_df = pd.read_excel("final_machine_capacity.xlsx")
energy_xls = pd.ExcelFile("energy_meter_with_nonzero_delta.xlsx")  # Load all sheets

# Extract machine names and capacities
machines = machine_df[['machine_name', 'machine_capacity_(kw)']].dropna()
machine_list = list(zip(machines['machine_name'], machines['machine_capacity_(kw)']))

# Tolerance level
tolerance = 0.5

# Initialize result list
all_results = []

# Loop through each sheet
for sheet_name in energy_xls.sheet_names:
    energy_df = energy_xls.parse(sheet_name)

    for _, row in energy_df[['timestamp', 'total_kW']].dropna().iterrows():
        timestamp = row['timestamp']
        target_kw = row['total_kW']
        
        valid_combinations = []

        for r in range(1, len(machine_list) + 1):
            for combo in combinations(machine_list, r):
                combo_kw = sum([cap for _, cap in combo])
                if abs(combo_kw - target_kw) <= tolerance:
                    valid_combinations.append([name for name, _ in combo])
        
        all_results.append({
            'sheet': sheet_name,
            'timestamp': timestamp,
            'total_kw': target_kw,
            'number_of_combinations': len(valid_combinations),
            'combinations': str(valid_combinations)
        })

# Convert all results to DataFrame
final_df = pd.DataFrame(all_results)

# ➕ Calculate grand total
grand_total = final_df['number_of_combinations'].sum()

# Create a new row with only the grand total
total_row = pd.DataFrame([{
    'sheet': 'ALL',
    'timestamp': 'TOTAL',
    'total_kw': '',
    'number_of_combinations': grand_total,
    'combinations': 'TOTAL'
}])

# Append this total row to the DataFrame
final_df = pd.concat([final_df, total_row], ignore_index=True)

# Save to Excel
final_df.to_excel("machine_combinations_output.xlsx", index=False)

print("✅ File 'machine_combinations_output.xlsx' has been created with a single total row at the end.")
