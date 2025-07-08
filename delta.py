import pandas as pd

# Load the Excel file
df = pd.read_excel("output_with_delta.xlsx")  # Replace with your actual file name

# Calculate Delta
df["Delta"] = df["kWh"].diff().fillna(0)

# Remove rows where Delta == 0
df = df[df["Delta"] != 0]

# Save to new file
df.to_excel("output_without_zero_delta.xlsx", index=False)

print("✅ Rows with Delta = 0 removed. File saved as 'output_without_zero_delta.xlsx'")

