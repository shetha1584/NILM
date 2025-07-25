import pandas as pd
from itertools import combinations
import math

# Load Excel files
machine_df = pd.read_excel("final_machine_capacity_with_pf.xlsx")
energy_xls = pd.ExcelFile("energy_meter_with_nonzero_delta.xlsx")

# Extract machine data
machines = machine_df[['machine_name', 'machine_capacity_(kw)', 'Estimated PF']].dropna()
machine_list = list(zip(machines['machine_name'], machines['machine_capacity_(kw)'], machines['Estimated PF']))

# Tolerances
kw_tolerance = 0.5
pf_tolerance = 0.58

# Function to calculate combined PF
def calculate_combined_pf(combo):
    """
    combo: list of (name, kW, PF)
    """
    total_kw = 0
    total_kvar = 0
    for _, kw, pf in combo:
        kva = kw / pf
        kvar = math.sqrt(kva**2 - kw**2)
        total_kw += kw
        total_kvar += kvar
    total_kva = math.sqrt(total_kw**2 + total_kvar**2)
    return total_kw / total_kva if total_kva != 0 else 0

# Collect results
all_results = []

# Loop through each sheet
for sheet_name in energy_xls.sheet_names:
    energy_df = energy_xls.parse(sheet_name)

    if not {'timestamp', 'total_kW', 'avg_power_factor'}.issubset(energy_df.columns):
        continue

    for _, row in energy_df[['timestamp', 'total_kW', 'avg_power_factor']].dropna().iterrows():
        timestamp = row['timestamp']
        target_kw = row['total_kW']
        target_pf = row['avg_power_factor']

        valid_combinations = []

        for r in range(1, len(machine_list) + 1):
            for combo in combinations(machine_list, r):
                combo_kw = sum([cap for _, cap, _ in combo])  # Check kW first
                if abs(combo_kw - target_kw) <= kw_tolerance:
                    combo_pf = calculate_combined_pf(combo)  # Now calculate PF
                    if abs(combo_pf - target_pf) <= pf_tolerance:
                        valid_combinations.append([name for name, _, _ in combo])

        all_results.append({
            'sheet': sheet_name,
            'timestamp': timestamp,
            'total_kw': target_kw,
            'avg_power_factor': target_pf,
            'number_of_combinations': len(valid_combinations),
            'combinations': str(valid_combinations)
        })

# Save results
final_df = pd.DataFrame(all_results)
final_df.to_excel("machine_combinations_pf_filtered.xlsx", index=False)

print("✅ File 'machine_combinations_pf_filtered.xlsx' created with kW & PF filters.")
