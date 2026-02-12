# Description:
# This script combines Excel files and uses a hybrid approach to find duplicates.
#
# Deduplication Logic:
# 1. Spatial Clustering: Groups records within a 200-meter radius.
# 2. Hybrid Address Matching: Within each cluster, it requires an EXACT match
#    on the street NUMBER and a FUZZY match on the street NAME and store name.
# ---------------------------------------------------------------------------

# --- REQUIRED LIBRARIES ---
# You must install these libraries before running the script.
# Open your command prompt or terminal and run:
# pip install pandas openpyxl geopy thefuzz python-Levenshtein

import pandas as pd  # For reading Excel file
import os 
import glob  # For searching through all subfolders
from geopy.distance import great_circle  # calculates the projected distance 
from thefuzz import fuzz # calculates a similarity score between strings
from itertools import combinations # generate all possible pairs of items from a list
import numpy as np # Numerical operations
import re

# --- CONFIGURATION ---
DISTANCE_THRESHOLD_METERS = 200
FUZZY_NAME_THRESHOLD = 85      # Similarity score (0-100) for store names
FUZZY_ADDRESS_NAME_THRESHOLD = 80 # Similarity score (0-100) for street names

# --- To solve the problem of inconsistent column names across different files ---
def find_and_standardize_columns(df):
    """Finds and renames key columns to standard names for consistent processing."""
    column_potentials = {
        'std_lat': ['latitude', 'location_y', 'lat'],
        'std_lon': ['longitude', 'location_x', 'lon', 'long'],
        'std_name': ['store_name', 'business_name', 'name'],
        'std_address': ['store_street_address', 'address', 'location_address', 'street']
    }
    rename_map = {}
    column_map_lower = {col.lower().strip(): col for col in df.columns}
    for standard_name, potential_names in column_potentials.items():
        for name in potential_names:
            if name in column_map_lower:
                rename_map[column_map_lower[name]] = standard_name
                break
    df.rename(columns=rename_map, inplace=True)
    return df

# check numbers at the beginning
def split_address(address):
    """
    Splits an address string into its leading number and the rest of the string.
    Returns (None, original_address) if no leading number is found.
    """
    if not isinstance(address, str):
        return None, address
    
    match = re.match(r'^(\d+)\s+(.*)', address)
    if match:
        return match.group(1), match.group(2) # (number, street_name)
    else:
        return None, address # (None, full_address_as_name)

def process_and_combine_excels(root_directory):
    """Main function to find, combine, and process Excel files using hybrid deduplication."""
    if not os.path.isdir(root_directory):
        print(f"Error: Root directory not found at '{root_directory}'")
        return

    excel_files = glob.glob(os.path.join(root_directory, '**', '*.xlsx'), recursive=True)
    excel_files.extend(glob.glob(os.path.join(root_directory, '**', '*.xls'), recursive=True))

    if not excel_files:
        print(f"No Excel files found in '{root_directory}'.")
        return

    print(f"Found {len(excel_files)} Excel files.")
    
    all_dataframes = []
    for file_path in excel_files:
        filename = os.path.basename(file_path)
        print(f"Reading and standardizing file: {filename}...")
        try:
            df = pd.read_excel(file_path)
            df['Source File'] = filename
            df = find_and_standardize_columns(df)
            all_dataframes.append(df)
        except Exception as e:
            print(f"  Could not read file {filename}. Error: {e}")

    if not all_dataframes:
        print("No dataframes created. Halting.")
        return
        
    print("\nCombining all files...")
    combined_df = pd.concat(all_dataframes, ignore_index=True, sort=False).copy()

    # --- 1. PRE-PROCESSING & COLUMN VALIDATION ---
    required_cols = ['std_lat', 'std_lon', 'std_name', 'std_address']
    if not all(col in combined_df.columns for col in required_cols):
        print("\nError: Could not find all required columns (lat, lon, name, address) after standardization.")
        return

    # Clean and type data
    #  1) converts latitude and longitude to numbers
    combined_df['std_lat'] = pd.to_numeric(combined_df['std_lat'], errors='coerce')
    combined_df['std_lon'] = pd.to_numeric(combined_df['std_lon'], errors='coerce')
    # 2) removes any rows that are missing coordinates
    combined_df.dropna(subset=['std_lat', 'std_lon'], inplace=True)
    combined_df.reset_index(drop=True, inplace=True)
    combined_df['std_name'] = combined_df['std_name'].astype(str).str.lower().str.strip()
    combined_df['std_address'] = combined_df['std_address'].astype(str).str.lower().str.strip()

    # Create new columns for the split address parts
    address_parts = combined_df['std_address'].apply(split_address)
    combined_df['Extracted_Street_Number'] = address_parts.apply(lambda x: x[0])
    combined_df['Extracted_Street_Name'] = address_parts.apply(lambda x: x[1])

    # --- 2. SPATIAL CLUSTERING ---
    # Groups records within 200m
    print(f"\nStep 1: Grouping records within {DISTANCE_THRESHOLD_METERS}m...")
    coords = combined_df[['std_lat', 'std_lon']].to_numpy()
    visited = np.zeros(len(coords), dtype=bool)
    cluster_id_counter = 0
    combined_df['Cluster_ID'] = -1
    for i in range(len(coords)):
        if not visited[i]:
            cluster_id_counter += 1
            q = [i]
            visited[i] = True
            head = 0
            while head < len(q):
                p1_idx = q[head]; head += 1
                for j in range(i + 1, len(coords)):
                    if not visited[j]:
                        if great_circle(coords[p1_idx], coords[j]).meters <= DISTANCE_THRESHOLD_METERS:
                            visited[j] = True
                            q.append(j)
            combined_df.loc[q, 'Cluster_ID'] = cluster_id_counter
    print(f"Found {cluster_id_counter} spatial clusters.")

    # --- 3. HYBRID MATCHING WITHIN CLUSTERS ---
    # compares records that are in the same spatial cluster
    print("\nStep 2: Performing hybrid matching (exact number + fuzzy name/street) within clusters...")
    combined_df['Duplicate_Group_ID'] = -1
    combined_df['Name_Score'] = 0
    combined_df['Address_Score'] = 0
    duplicate_group_counter = 0
    
    for cluster_id in combined_df['Cluster_ID'].unique():
        if cluster_id == -1: continue
        cluster_df_indices = combined_df[combined_df['Cluster_ID'] == cluster_id].index
        
        for idx1, idx2 in combinations(cluster_df_indices, 2):
            num1 = combined_df.at[idx1, 'Extracted_Street_Number']
            num2 = combined_df.at[idx2, 'Extracted_Street_Number']

            # HARD CHECK: Street numbers must exist and be identical.
            if num1 is not None and num1 == num2:
                name1 = combined_df.at[idx1, 'std_name']
                name2 = combined_df.at[idx2, 'std_name']
                street1 = combined_df.at[idx1, 'Extracted_Street_Name']
                street2 = combined_df.at[idx2, 'Extracted_Street_Name']
                
                # FUZZY CHECKS
                name_score = fuzz.token_set_ratio(name1, name2)
                addr_score = fuzz.token_set_ratio(street1, street2)
                
                if name_score >= FUZZY_NAME_THRESHOLD and addr_score >= FUZZY_ADDRESS_NAME_THRESHOLD:
                    # This pair is a likely duplicate. Group them.
                    group1_id = combined_df.at[idx1, 'Duplicate_Group_ID']
                    group2_id = combined_df.at[idx2, 'Duplicate_Group_ID']

                    if group1_id == -1 and group2_id == -1:
                        duplicate_group_counter += 1
                        combined_df.loc[[idx1, idx2], 'Duplicate_Group_ID'] = duplicate_group_counter
                    elif group1_id != -1 and group2_id == -1:
                        combined_df.at[idx2, 'Duplicate_Group_ID'] = group1_id
                    elif group1_id == -1 and group2_id != -1:
                        combined_df.at[idx1, 'Duplicate_Group_ID'] = group2_id
                    elif group1_id != group2_id:
                        combined_df.loc[combined_df['Duplicate_Group_ID'] == group2_id, 'Duplicate_Group_ID'] = group1_id
                        
                    # Store scores for review
                    combined_df.at[idx1, 'Name_Score'] = max(combined_df.at[idx1, 'Name_Score'], name_score)
                    combined_df.at[idx2, 'Name_Score'] = max(combined_df.at[idx2, 'Name_Score'], name_score)
                    combined_df.at[idx1, 'Address_Score'] = max(combined_df.at[idx1, 'Address_Score'], addr_score)
                    combined_df.at[idx2, 'Address_Score'] = max(combined_df.at[idx2, 'Address_Score'], addr_score)

    print(f"Identified {duplicate_group_counter} potential duplicate groups.")

    # --- 4. SEPARATE AND SAVE FILES ---
    print("\nStep 3: Separating and saving output files...")
    
    duplicates_comparison_df = combined_df[combined_df['Duplicate_Group_ID'] != -1].copy()
    duplicates_comparison_df.sort_values(by=['Duplicate_Group_ID', 'Source File'], inplace=True)

    unique_df = combined_df.drop_duplicates(subset=['Duplicate_Group_ID'], keep='first').copy()
    
    is_first_occurrence = ~duplicates_comparison_df.duplicated(subset=['Duplicate_Group_ID'], keep='first')
    duplicates_comparison_df['Comparison Status'] = 'Duplicate (Removed)'
    duplicates_comparison_df.loc[is_first_occurrence, 'Comparison Status'] = 'Original (Kept)'
    
    # Drop helper columns from the final unique dataset
    cols_to_drop = ['Cluster_ID', 'Duplicate_Group_ID', 'Name_Score', 'Address_Score',
                    'Extracted_Street_Number', 'Extracted_Street_Name']
    unique_df.drop(columns=cols_to_drop, inplace=True)
    
    output_parent_dir = os.path.dirname(root_directory)
    unique_output_filename = os.path.join(output_parent_dir, 'Consolidated_Unique_Data11.xlsx')
    duplicates_output_filename = os.path.join(output_parent_dir, 'Removed_Duplicates_Comparison11.xlsx')

    try:
        print(f"\nSaving {len(unique_df)} unique records to '{os.path.basename(unique_output_filename)}'...")
        unique_df.to_excel(unique_output_filename, index=False)
        
        if not duplicates_comparison_df.empty:
            print(f"Saving comparison file with {len(duplicates_comparison_df)} records to '{os.path.basename(duplicates_output_filename)}'...")
            duplicates_comparison_df.to_excel(duplicates_output_filename, index=False)
        else:
            print("No duplicate records found to save.")

        print(f"\nProcessing complete! Files saved in:\n{output_parent_dir}")
    except Exception as e:
        print(f"\nError saving the final files: {e}")

# --- SCRIPT EXECUTION ---
if __name__ == "__main__":
    # IMPORTANT: Update this path to the top-level directory you want to search.
    target_root_directory = r"D:\Madusha_NIH\clinical\california"
    
    process_and_combine_excels(target_root_directory)
