import pandas as pd

# --- STEP 1: Load the Data File ---
df = pd.read_excel('output_without_zero_delta.xlsx')

# --- STEP 2: Clean Column Names and Check for Required Fields ---
df.columns = df.columns.str.strip().str.lower()

# Ensure required columns are present
required_cols = ['timestamp', 'total_kw', 'avg_current_value']
missing = [col for col in required_cols if col not in df.columns]
if missing:
    raise ValueError(f"❌ Missing columns in the file: {missing}")

# --- STEP 2.5: Fix Negative total_kw Values (e.g., -5 becomes 5) ---
df['total_kw'] = df['total_kw'].abs()

# --- STEP 3: Convert Timestamp and Sort ---
df['timestamp'] = pd.to_datetime(df['timestamp'])
df.sort_values(by='timestamp', inplace=True)

# --- STEP 4: Classify Into Active and Inactive Based on Current Median ---
threshold = df['avg_current_value'].median()
print(f"\nMedian avg_current_value used as threshold: {threshold:.2f}")

df['activity'] = df['avg_current_value'].apply(
    lambda x: 'active' if x >= threshold else 'inactive'
)

# --- STEP 5: Split Active and Inactive Data ---
active_df = df[df['activity'] == 'active']
inactive_df = df[df['activity'] == 'inactive']

# --- STEP 6: Print Summary Stats Function ---
def print_range_stats(label, data, column):
    print(f"\n{label.upper()} — {column}:")
    print(f"  Min:     {data[column].min():.2f}")
    print(f"  Max:     {data[column].max():.2f}")
    print(f"  Mean:    {data[column].mean():.2f}")
    print(f"  Median:  {data[column].median():.2f}")
    print(f"  Std Dev: {data[column].std():.2f}")
    print(f"  Count:   {len(data)} rows")

# --- STEP 7: Display Stats ---
print_range_stats("Active", active_df, "total_kw")
print_range_stats("Inactive", inactive_df, "total_kw")
print_range_stats("Active", active_df, "avg_current_value")
print_range_stats("Inactive", inactive_df, "avg_current_value")

# --- STEP 8: Save Output with Activity Column ---
df.to_excel('output_with_activity_tag.xlsx', index=False)
print("\n✅ Saved updated file as 'output_with_activity_tag.xlsx'")
