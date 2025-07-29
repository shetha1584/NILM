import pandas as pd

# Load both Excel files
output_df = pd.read_excel("machine_combinations_output.xlsx")
filtered_df = pd.read_excel("machine_combinations_pf_filtered_2.xlsx")

# Pick the right column
col = "number_of_combinations"

# Calculate totals
total_output = output_df[col].sum()
total_filtered = filtered_df[col].sum()

# Calculate decrease and percentage
decrease = total_output - total_filtered
percentage_reduction = (decrease / total_output) * 100 if total_output > 0 else 0

print("Total before (output):", total_output)
print("Total after (filtered):", total_filtered)
print("Reduction:", decrease)
print("Percentage reduction: {:.2f}%".format(percentage_reduction))
