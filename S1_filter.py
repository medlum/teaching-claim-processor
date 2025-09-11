import pandas as pd

# Read the Excel file
df = pd.read_excel('subset_data/all_asrq180.xlsx')

# Filter rows containing "@adj.np.edu.sg" in Email (case insensitive)
email_filtered = df[df['Email'].str.contains(
    '@adj.np.edu.sg', case=False, na=False)].copy()


# Define function to check if Class Section has ≤2 letters


def has_max_two_letters(value):
    if pd.isna(value):
        return False  # Handle missing values
    # Extract only alphabetic characters
    letters = ''.join([char for char in str(value) if char.isalpha()])
    return len(letters) <= 2

# Function to expand rows with multiple values in Day column


def expand_day_column(df):
    # Create a list to store the expanded rows
    expanded_rows = []
    #  Count rows with more than a single value: MON - SAT
    multidays_count = 0
    # Iterate over each row in the input DataFrame
    for _, row in df.iterrows():
        # Get the Day value and split it into individual days
        day_value = str(row['Day'])
        days = day_value.split()

        # If there's only one day, just add the row as is
        if len(days) <= 1:
            expanded_rows.append(row.to_dict())
        else:
            # For each day, create a copy of the row with that single day
            for day in days:
                new_row = row.copy()
                new_row['Day'] = day
                expanded_rows.append(new_row.to_dict())
            multidays_count += 1

    # Create a new DataFrame from the expanded rows
    expanded_df = pd.DataFrame(expanded_rows)
    return multidays_count, expanded_df


# New function to format time columns
def format_time_columns(df, columns):
    for col in columns:
        # Convert to datetime, then format as HH:MM:SS string
        df.loc[:, col] = pd.to_datetime(
            df[col], format='%H:%M:%S', errors='coerce').dt.strftime('%H:%M:%S')

    return df


excluded_sections = ['TSP1', 'WSP1']  # User-specified list

email_filtered.loc[:, 'ExcludeFromFilter'] = email_filtered['Class Section'].isin(
    excluded_sections)


filtered_df = email_filtered[(email_filtered['Class Section'].apply(has_max_two_letters)) |
                             (email_filtered['ExcludeFromFilter'])
                             ]


# Apply the Class Section filter
# filtered_df = email_filtered[email_filtered['Class Section'].apply(
#    has_max_two_letters)]

# Format 'Start Time' and 'End Time' columns
formatted_time_df = format_time_columns(
    filtered_df, ['Start Time', 'End Time'])

# Apply the expansion function to the filtered DataFrame
multidays_count, expanded_df = expand_day_column(formatted_time_df)

print(multidays_count)
# Display the result
print(expanded_df)

# Save to new Excel file
expanded_df.to_excel('processed_data/filtered_results.xlsx', index=False)
