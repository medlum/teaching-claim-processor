import pandas as pd
import numpy as np

def merge_with_partial_match(filtered_df, lookup_df):
    """
    Merges two DataFrames based on partial name matching and catalog/remarks matching.
    Args:
        filtered_df (pd.DataFrame): DataFrame from filtered_results.xlsx
        lookup_df (pd.DataFrame): DataFrame from position_program_id_lookup.xlsx
    Returns:
        tuple: (merged_df, unmatched_df, unmatched_count)
            - merged_df (pd.DataFrame): Merged DataFrame with required columns
            - unmatched_df (pd.DataFrame): DataFrame containing unmatched rows
            - unmatched_count (int): Count of unmatched rows
    """
    # Create a copy to avoid modifying original DataFrames
    filtered = filtered_df.copy()
    lookup = lookup_df.copy()
    # Preprocess names for comparison - convert to uppercase and split into words
    filtered['name_words'] = filtered['Name'].str.upper().str.split()
    lookup['lookup_words'] = lookup['Full Legal Name'].str.upper().str.split()
    # Function to check if names are a partial match
    def is_partial_match(name1, name2, min_common_tokens=2):
        # Convert to sets
        set1 = set(name1)
        set2 = set(name2)
        # Handle empty sets
        if not set1 or not set2:
            return False
        # Check if there are enough common tokens
        common = set1.intersection(set2)
        return len(common) >= min_common_tokens
    
    # Improved function to check if catalog number and requester remarks have a partial match
    def is_catalog_match(catalog, remarks):
        # Handle NaN values
        if pd.isna(catalog) or pd.isna(remarks):
            return False
        # Convert to strings and strip whitespace
        catalog_str = str(catalog).strip().replace(' ', '')
        remarks_str = str(remarks).strip()
        # Check if one string contains the other
        if (catalog_str in remarks_str) or (remarks_str in catalog_str):
            return True
        # Extract only alphabetic characters from catalog_str
        catalog_alpha = ''.join([c for c in catalog_str if c.isalpha()])
        if catalog_alpha and (catalog_alpha in remarks_str):
            return True
        # Additional check: see if any word in catalog_str is in remarks_str
        catalog_words = catalog_str.split('_')
        for word in catalog_words:
            word_alpha = ''.join([c for c in word if c.isalpha()])
            if word_alpha and word_alpha in remarks_str:
                return True
        return False
    # Create empty result DataFrame
    result_columns = [
        'Empl ID',
        'Full Legal Name',
        'Time entry code',
        'Date',  # Will be filled in Step 3
        'Start Time',
        'End Time',
        'Position ID',
        'Program ID',
        'Comment',
        'Day',
        'Catalog Nbr',
        'Name'
    ]
    result_df = pd.DataFrame(columns=result_columns)
    # Track unmatched rows
    unmatched_count = 0
    unmatched_rows = []  # Store unmatched rows here
    # Iterate through each row in filtered DataFrame
    for _, filt_row in filtered.iterrows():
        filt_name = filt_row['name_words']
        catalog_nbr = filt_row['Catalog Nbr'] if 'Catalog Nbr' in filt_row else None
        match_found = False
        # Iterate through each row in lookup DataFrame
        for _, lookup_row in lookup.iterrows():
            lookup_name = lookup_row['lookup_words']
            requester_remarks = lookup_row['Requester Remarks'] if 'Requester Remarks' in lookup_row else None
            # Check for partial name match
            name_match = is_partial_match(filt_name, lookup_name)
            # Check for catalog/remarks match
            catalog_match = is_catalog_match(catalog_nbr, requester_remarks)
            # If both conditions are met, create a new row
            if name_match and catalog_match:
                # Create new row with required data
                new_row = {
                    'Empl ID': lookup_row['Empl ID'],
                    'Full Legal Name': lookup_row['Full Legal Name'],
                    'Time entry code': lookup_row['Time entry code'],
                    'Date': None,  # Will be filled in Step 3
                    'Start Time': filt_row['Start Time'],
                    'End Time': filt_row['End Time'],
                    'Position ID': lookup_row['Position ID'],
                    'Program ID': lookup_row['Program ID'],
                    'Comment': lookup_row['Requester Remarks'],
                    'Day': filt_row['Day'],
                    'Catalog Nbr': filt_row['Catalog Nbr'],
                    'Name': filt_row['Name']
                }
                # Append to result
                result_df = pd.concat(
                    [result_df, pd.DataFrame([new_row])], ignore_index=True)
                match_found = True
                break  # Stop after first match
        if not match_found:
            unmatched_count += 1
            # Store the original row without the added 'name_words' column
            unmatched_rows.append(filt_row.drop('name_words'))
    # Create DataFrame from unmatched rows
    unmatched_df = pd.DataFrame(unmatched_rows)
    
    # Return all three values: merged DataFrame, unmatched DataFrame, and unmatched count
    return result_df, unmatched_df, unmatched_count

# Read the Excel files
filtered_df = pd.read_excel('processed_data/filtered_results.xlsx')
# test file: position_program_id_lookup.xlsx
lookup_df = pd.read_excel('subset_data/all_hiring_form.xlsx')
# Print column names for debugging
print("Columns in filtered DataFrame:", filtered_df.columns.tolist())
print("Columns in lookup DataFrame:", lookup_df.columns.tolist())
# Perform the merge with partial matching
result_df, unmatched_df, unmatched_count = merge_with_partial_match(filtered_df, lookup_df)
# Save the result
result_df.to_excel('processed_data/merged_output.xlsx', index=False)
# Display the first few rows
print("\nMerged result sample:")
print(result_df.head())

# Print statistics and sample of unmatched rows
print(f"\nProcessed {len(filtered_df)} rows from filtered DataFrame")
print(f"Found matches for {len(result_df)} rows")
print(f"Unmatched rows: {unmatched_count}")
if unmatched_count > 0:
    print("\nSample of unmatched rows:")
    print(unmatched_df.head())
    # Save unmatched rows to Excel
    unmatched_df.to_excel(
        'processed_data/unmatched_rows.xlsx', index=False)
    print("\nUnmatched rows saved to 'processed_data/unmatched_rows.xlsx'")