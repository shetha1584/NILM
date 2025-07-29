import pandas as pd
import ast

# Load Excel file
file_path = "machine_combinations_pf_filtered_2.xlsx"
df = pd.read_excel(file_path, sheet_name="Sheet1")

# Convert 'combinations' from string to list (if needed)
df['combinations'] = df['combinations'].apply(ast.literal_eval)

# Count how many rows have each number of combinations
combination_count_distribution = df['number_of_combinations'].value_counts().sort_index()

# Print results
print("Number of rows for each combination count:")
for combos, rows in combination_count_distribution.items():
    print(f"{combos} combinations → {rows} rows")
