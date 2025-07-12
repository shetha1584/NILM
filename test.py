import pandas as pd

# Load the existing file
df = pd.read_excel("combinations_with_0.9kw_tolerance.xlsx")

# Convert total_kw to numeric (force errors to NaN), then remove rows where it's 0 or NaN
df['total_kw'] = pd.to_numeric(df['total_kw'], errors='coerce')
df_cleaned = df[df['total_kw'].fillna(0) != 0]

# Save cleaned version
df_cleaned.to_excel("combinations_with_0.9kw_tolerance_cleaned.xlsx", index=False)

print(f"✅ Cleaned file saved with {len(df_cleaned)} rows (removed {len(df) - len(df_cleaned)} rows with total_kw = 0 or invalid).")
