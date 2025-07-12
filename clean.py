import pandas as pd

# Load machine capacity Excel
file_path = "terracemachine.xlsx"
df = pd.read_excel(file_path)

# Standardize column names
df.columns = df.columns.str.strip().str.lower().str.replace(' ', '_')

# Keep only relevant columns
df_clean = df[['s.no', 'machine_name', 'machine_capacity_(kw)']]

# Drop rows with missing capacity values
df_clean = df_clean.dropna(subset=['machine_capacity_(kw)'])

# Save the cleaned file
df_clean.to_excel("final_machine_capacity.xlsx", index=False)

print("✅ Final machine data saved to: final_machine_capacity.xlsx")
