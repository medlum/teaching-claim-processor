import pandas as pd
import numpy as np


def merge_with_partial_match(expanded_df, lookup_df):
    """
    Merges two DataFrames based on partial name matching.

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
    def is_partial_match(name1, name2):
        # Convert to sets of words
        set1 = set(name1)
        set2 = set(name2)

        # Check if one set is a subset of the other
        return set1.issubset(set2) or set2.issubset(set1)

    # Create empty result DataFrame
    result_columns = [
        'Empl ID',
        'Full Legal Name',
        'Time entry code',  # Corrected column name
        'Date',
        'Start Time',
        'End Time',
        'Position ID',
        'Program ID'  # Corrected column name
    ]
    result_df = pd.DataFrame(columns=result_columns)

    # Track unmatched rows
    unmatched_count = 0

    # Iterate through each row in expanded DataFrame
    for _, exp_row in expanded.iterrows():
        exp_name = exp_row['name_words']
        match_found = False

        # Iterate through each row in lookup DataFrame
        for _, lookup_row in lookup.iterrows():
            lookup_name = lookup_row['lookup_words']

            # Check for partial match
            if is_partial_match(exp_name, lookup_name):
                # Create new row with required data
                new_row = {
                    'Empl ID': lookup_row['Empl ID'],
                    'Full Legal Name': lookup_row['Full Legal Name'],
                    # Corrected column name
                    'Time entry code': lookup_row['Time entry code'],
                    'Date': exp_row['Date'],
                    'Start Time': exp_row['Start Time'],
                    'End Time': exp_row['End Time'],
                    'Position ID': lookup_row['Position ID'],
                    # Corrected column name
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
expanded_df = pd.read_excel('expanded_with_dates.xlsx')
lookup_df = pd.read_excel('position_program_id_lookup.xlsx')

# Print column names for debugging
print("Columns in expanded DataFrame:", expanded_df.columns.tolist())
print("Columns in lookup DataFrame:", lookup_df.columns.tolist())

# Perform the merge with partial matching
result_df = merge_with_partial_match(expanded_df, lookup_df)

# Save the result
result_df.to_excel('merged_output.xlsx', index=False)

# Display the first few rows
print(result_df.head())
