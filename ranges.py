import pandas as pd

# Load the Excel file
file_path = "energy_meter_1_June15_16.xlsx"
excel_file = pd.ExcelFile(file_path)

# To collect all non-zero deltas from all sheets
all_deltas = []

# First pass: compute delta values
for sheet_name in excel_file.sheet_names:
    df = excel_file.parse(sheet_name)
    df.columns = df.columns.str.strip()

    if 'kWh' in df.columns:
        delta = df['kWh'].diff().fillna(0)
        all_deltas.extend(delta[delta > 0])  # only meaningful deltas

# Compute dynamic threshold from all non-zero deltas
threshold = pd.Series(all_deltas).mean()
print(f"\n▶ Threshold for 'Active' based on mean delta: {threshold:.4f} kWh")

# Prepare output Excel with activity classification
output_path = "energy_meter_with_activity_labels.xlsx"

with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
    for sheet_name in excel_file.sheet_names:
        df = excel_file.parse(sheet_name)
        df.columns = df.columns.str.strip()

        if 'kWh' not in df.columns:
            raise ValueError(f"'kWh' column not found in sheet '{sheet_name}'")

        # Compute delta and classify activity
        df['delta'] = df['kWh'].diff().fillna(0)
        df['activity'] = df['delta'].apply(lambda x: 'Active' if x > threshold else 'Non-Active')

        # Print value range of Active and Non-Active hours
        active_range = df[df['activity'] == 'Active']['delta']
        nonactive_range = df[df['activity'] == 'Non-Active']['delta']

        print(f"\n📄 Sheet: {sheet_name}")
        print(f"  Active delta range: {active_range.min():.2f} to {active_range.max():.2f} kWh")
        print(f"  Non-Active delta range: {nonactive_range.min():.2f} to {nonactive_range.max():.2f} kWh")

        # Save to new sheet
        df.to_excel(writer, sheet_name=sheet_name, index=False)

print(f"\n✅ Saved Excel with activity labels to: {output_path}")
