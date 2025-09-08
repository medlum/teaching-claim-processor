import pandas as pd
import datetime


def create_weekday_date_dict(start_date_str, end_date_str):
    """
    Creates a dictionary mapping weekdays (Mon-Sat) to dates within a specified range.

    Args:
        start_date_str (str): Start date in format "DD Month YYYY" (e.g., "21 April 2025")
        end_date_str (str): End date in format "DD Month YYYY" (e.g., "23 August 2025")

    Returns:
        dict: Dictionary with keys 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat' 
              and values as lists of dates in "YYYY-MM-DD" format
    """
    # Parse input dates
    start_date = datetime.datetime.strptime(start_date_str, "%d %B %Y").date()
    end_date = datetime.datetime.strptime(end_date_str, "%d %B %Y").date()

    # Initialize weekday dictionary
    weekday_dict = {
        'Mon': [], 'Tue': [], 'Wed': [],
        'Thu': [], 'Fri': [], 'Sat': []
    }

    # Iterate through each date in the range
    current_date = start_date
    while current_date <= end_date:
        # Get weekday (0=Monday, 6=Sunday)
        weekday_num = current_date.weekday()

        # Skip Sundays (6)
        if weekday_num < 6:
            # Map weekday number to dictionary key
            weekday_key = ['Mon', 'Tue', 'Wed',
                           'Thu', 'Fri', 'Sat'][weekday_num]

            # Add date in YYYY-MM-DD format
            weekday_dict[weekday_key].append(current_date.strftime("%Y-%m-%d"))

        # Move to next day
        current_date += datetime.timedelta(days=1)

    return weekday_dict


def expand_df_with_dates(filtered_df, start_date_str, end_date_str):
    """
    Expands DataFrame by duplicating rows for each date in the 'Day' column.
    Skips rows with invalid day entries.
    Places 'Date' column immediately after 'Day' column.

    Args:
        filtered_df (pd.DataFrame): Input DataFrame with 'Day' column
        start_date_str (str): Start date in format "DD Month YYYY"
        end_date_str (str): End date in format "DD Month YYYY"

    Returns:
        pd.DataFrame: Expanded DataFrame with 'Date' column after 'Day' column
    """
    # Get weekday-date mapping
    weekday_date_dict = create_weekday_date_dict(start_date_str, end_date_str)

    # Define valid weekdays
    valid_days = {'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat'}

    # Create a list to store expanded rows
    expanded_rows = []
    skipped_rows = 0

    # Process each row in the filtered DataFrame
    for _, row in filtered_df.iterrows():
        # Get the Day value and split into individual days
        day_str = str(row['Day']).strip()

        # Skip if empty
        if not day_str:
            skipped_rows += 1
            continue

        # Split the day string and clean each part
        day_parts = [part.strip() for part in day_str.split() if part.strip()]

        # Check each part to see if it's a valid weekday
        day_keys = []
        skip_row = False

        for part in day_parts:
            # Capitalize first letter only (e.g., "MON" -> "Mon", "tue" -> "Tue")
            standardized = part.capitalize()

            # Check if it's a valid weekday
            if standardized in valid_days:
                # Add to day_keys if not already present (to avoid duplicates)
                if standardized not in day_keys:
                    day_keys.append(standardized)
            else:
                # Skip this row if any part is invalid
                skip_row = True
                break

        if skip_row:
            skipped_rows += 1
            continue

        # For each valid day key, create a row for each date
        for day_key in day_keys:
            for date in weekday_date_dict[day_key]:
                # Create a copy of the row
                new_row = row.copy()
                # Add the Date column
                new_row['Date'] = date
                # Add to expanded rows
                expanded_rows.append(new_row)

    # Create the expanded DataFrame
    expanded_df = pd.DataFrame(expanded_rows)

    # Reorder columns to put 'Date' immediately after 'Day'
    if 'Day' in expanded_df.columns and 'Date' in expanded_df.columns:
        # Get list of columns
        cols = list(expanded_df.columns)

        # Find the index of 'Day' column
        day_index = cols.index('Day')

        # Remove 'Date' from its current position
        cols.remove('Date')

        # Insert 'Date' right after 'Day'
        cols.insert(day_index + 1, 'Date')

        # Reorder the DataFrame columns
        expanded_df = expanded_df[cols]

    # Reset index to avoid duplicate indices
    expanded_df.reset_index(drop=True, inplace=True)

    # Print information about skipped rows
    print(f"Processed {len(filtered_df)} rows")
    print(f"Skipped {skipped_rows} rows with invalid day entries")
    print(f"Expanded to {len(expanded_df)} rows")

    return expanded_df


# Example usage:
# Assuming filtered_df is your DataFrame from previous steps
start_date = "21 April 2025"
end_date = "23 August 2025"

# Apply the function to expand the DataFrame
filtered_df = pd.read_excel('processed_data/filtered_results.xlsx')
expanded_df = expand_df_with_dates(filtered_df, start_date, end_date)

# Display the result
print(expanded_df.head())

# Save to new Excel file
expanded_df.to_excel('processed_data/expanded_with_dates.xlsx', index=False)
