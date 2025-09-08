import pandas as pd
import numpy as np


def merge_with_partial_match(expanded_df, lookup_df):
    """
    Merges two DataFrames based on partial name matching and catalog/remarks matching.

    Args:
        expanded_df (pd.DataFrame): DataFrame from expanded_with_dates.xlsx
        lookup_df (pd.DataFrame): DataFrame from position_program_id_lookup.xlsx

    Returns:
        pd.DataFrame: Merged DataFrame with required columns
    """
    # Create a copy to avoid modifying original DataFrames
    expanded = expanded_df.copy()
    lookup = lookup_df.copy()

    # Rename the column in lookup DataFrame to match desired output name
    lookup = lookup.rename(columns={'Worker* (Emp ID)': 'Empl ID'})

    # Preprocess names for comparison - convert to uppercase and split into words
    expanded['name_words'] = expanded['Name'].str.upper().str.split()
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

    # Function to check if catalog number and requester remarks have a partial match

    def is_catalog_match(catalog, remarks):
        # Handle NaN values
        if pd.isna(catalog) or pd.isna(remarks):
            return False

        # Convert to strings and strip whitespace
        catalog_str = str(catalog).strip()
        remarks_str = str(remarks).strip()

        print(
            f"Comparing catalog: '{catalog_str}' with remarks: '{remarks_str}'")

        # Check if one string contains the other
        return (catalog_str in remarks_str) or (remarks_str in catalog_str)

    # Create empty result DataFrame
    result_columns = [
        'Empl ID',
        'Full Legal Name',
        'Time entry code',
        'Date',
        'Start Time',
        'End Time',
        'Position ID',
        'Program ID'
    ]
    result_df = pd.DataFrame(columns=result_columns)

    # Track unmatched rows
    unmatched_count = 0

    # Iterate through each row in expanded DataFrame
    for _, exp_row in expanded.iterrows():
        exp_name = exp_row['name_words']
        catalog_nbr = exp_row['Catalog Nbr'] if 'Catalog Nbr' in exp_row else None
        match_found = False

        # Iterate through each row in lookup DataFrame
        for _, lookup_row in lookup.iterrows():
            lookup_name = lookup_row['lookup_words']
            requester_remarks = lookup_row['Requester Remarks'] if 'Requester Remarks' in lookup_row else None

            # Check for partial name match
            name_match = is_partial_match(exp_name, lookup_name)

            # Check for catalog/remarks match
            catalog_match = is_catalog_match(catalog_nbr, requester_remarks)

            # If both conditions are met, create a new row
            if name_match and catalog_match:
                # Create new row with required data
                new_row = {
                    'Empl ID': lookup_row['Empl ID'],
                    'Full Legal Name': lookup_row['Full Legal Name'],
                    'Time entry code': lookup_row['Time entry code'],
                    'Date': exp_row['Date'],
                    'Start Time': exp_row['Start Time'],
                    'End Time': exp_row['End Time'],
                    'Position ID': lookup_row['Position ID'],
                    'Program ID': lookup_row['Program ID']
                }

                # Append to result
                result_df = pd.concat(
                    [result_df, pd.DataFrame([new_row])], ignore_index=True)
                match_found = True
                break  # Stop after first match

        if not match_found:
            unmatched_count += 1

    print(f"Processed {len(expanded)} rows from expanded DataFrame")
    print(f"Found matches for {len(result_df)} rows")
    print(f"Unmatched rows: {unmatched_count}")

    return result_df


# Read the Excel files
expanded_df = pd.read_excel('processed_data/expanded_with_dates.xlsx')
lookup_df = pd.read_excel('subset_data/position_program_id_lookup.xlsx')

# Print column names for debugging
print("Columns in expanded DataFrame:", expanded_df.columns.tolist())
print("Columns in lookup DataFrame:", lookup_df.columns.tolist())

# Perform the merge with partial matching
result_df = merge_with_partial_match(expanded_df, lookup_df)

# Save the result
result_df.to_excel('processed_data/merged_output.xlsx', index=False)

# Display the first few rows
print(result_df.head())
