import pandas as pd
from itertools import combinations
from collections import defaultdict
import time
from multiprocessing import Pool, cpu_count
import ast

def pre_filter_machines(target_kw, tolerance, machine_data):
    """Filter machines within range of target_kw ± tolerance"""
    max_possible = target_kw + tolerance
    min_possible = target_kw - tolerance
    return [(kw, name) for kw, name in machine_data if kw <= max_possible and kw >= min_possible / 7]

def find_small_combinations(capacities_names, target, tolerance, max_size=4):
    """Find 1-4 machine combinations with capacities close to target"""
    results = []
    capacities, names = zip(*capacities_names) if capacities_names else ([], [])
    
    for r in range(1, min(max_size, len(capacities_names)) + 1):
        if sum(capacities[:r]) < target - tolerance:
            continue
        if sum(capacities[-r:]) > target + tolerance:
            continue
            
        for combo in combinations(capacities_names, r):
            combo_sum = sum(kw for kw, name in combo)
            if abs(combo_sum - target) <= tolerance:
                results.append([name for kw, name in combo])
    return results

def find_large_combinations_dp(capacities_names, target, tolerance, min_size=5, max_size=7):
    """Find 5-7 machine combinations using dynamic programming"""
    if not capacities_names:
        return []
    
    dp = [defaultdict(list) for _ in range(max_size + 1)]
    dp[0][0] = [[]]
    results = []
    
    for cap, name in capacities_names:
        for size in range(min(max_size, len(capacities_names)), min_size - 1, -1):
            for s in list(dp[size - 1]):
                new_sum = s + cap
                if new_sum > target + tolerance:
                    continue
                for combo in dp[size - 1][s]:
                    new_combo = combo + [name]
                    dp[size][new_sum].append(new_combo)
                    if abs(new_sum - target) <= tolerance:
                        results.append(new_combo)
    return results

def find_all_valid_combinations(target_kw, machine_data, tolerance=0.9):
    """Return all valid machine combinations (1 to 7 machines)"""
    capacities_names = sorted(
        pre_filter_machines(target_kw, tolerance, machine_data),
        key=lambda x: -x[0]
    )
    
    if not capacities_names:
        return []
    
    small_combos = find_small_combinations(capacities_names, target_kw, tolerance)
    large_combos = []
    if len(capacities_names) >= 5:
        large_combos = find_large_combinations_dp(capacities_names, target_kw, tolerance)
    
    return small_combos + large_combos

def process_chunk(args):
    """Process a data chunk in parallel"""
    chunk, machine_data = args
    results = []
    for _, row in chunk.iterrows():
        combos = find_all_valid_combinations(row['total_kw'], machine_data, tolerance=0.5)
        results.append({
            'timestamp': row['timestamp'],
            'total_kw': row['total_kw'],
            'valid_combinations_count': len(combos),
            'valid_combinations_sets': str(combos) if combos else "[]"
        })
    return results

def main():
    start_time = time.time()
    
    # Load machine capacity data
    capacity_df = pd.read_excel("final_machine_capacity.xlsx")
    machine_data = list(zip(
        capacity_df['machine_capacity_(kw)'],
        capacity_df['machine_name']
    ))
    
    # Load energy meter data from all sheets
    energy_sheets = pd.read_excel("energy_meter_with_nonzero_delta.xlsx", sheet_name=None)
    energy_df = pd.concat(energy_sheets.values(), ignore_index=True)
    energy_df.columns = energy_df.columns.str.strip().str.lower().str.replace(' ', '_')

    energy_df = energy_df[energy_df['total_kw'].fillna(0) != 0]
    
    # Prepare parallel processing
    num_cores = min(cpu_count(), 8)
    chunk_size = len(energy_df) // num_cores + 1
    chunks = [energy_df.iloc[i:i + chunk_size].copy() for i in range(0, len(energy_df), chunk_size)]
    
    print(f"🔄 Processing {len(energy_df)} timestamps from {len(energy_sheets)} sheets with ±0.9kW tolerance...")
    
    with Pool(num_cores) as pool:
        results = pool.map(process_chunk, [(chunk, machine_data) for chunk in chunks])
    
    # Flatten results
    output_df = pd.DataFrame([r for chunk in results for r in chunk])
    output_df['valid_combinations_sets'] = output_df['valid_combinations_sets'].apply(ast.literal_eval)
    
    # Save results
    output_df.to_excel("combinations_with_0.9kw_tolerance.xlsx", 
                       columns=['timestamp', 'total_kw', 'valid_combinations_count', 'valid_combinations_sets'],
                       index=False)
    
    print(f"\n✅ Completed in {time.time() - start_time:.2f} seconds")
    print(f"Total combinations found: {output_df['valid_combinations_count'].sum()}")
    print("📊 Sample output:")
    print(output_df.head(2).to_string())

if __name__ == "__main__":
    main()
