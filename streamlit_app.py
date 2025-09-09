import streamlit as st
import pandas as pd
import datetime
import numpy as np
from io import BytesIO

# Set page configuration
st.set_page_config(
    page_title="Teaching Claim Process",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Define step names
STEP_NAMES = ['Clean Data', 'Merge Headers',
              'Date Transform']

# Initialize session state variables
if 'step1_data' not in st.session_state:
    st.session_state.step1_data = None
if 'step2_data' not in st.session_state:
    st.session_state.step2_data = None
if 'step3_data' not in st.session_state:
    st.session_state.step3_data = None
if 'current_step' not in st.session_state:
    st.session_state.current_step = 0  # Using 0-based index for step names

# Define processing functions


def filter_data(df):
    """Step 1: Filter data by adjunct and remove duplicates in ARSQ180"""
    # Filter rows containing "@adj.np.edu.sg" in Email (case insensitive)
    email_filtered = df[df['Email'].str.contains(
        '@adj.np.edu.sg', case=False, na=False)]

    # Define function to check if Class Section has ≤2 letters
    def has_max_two_letters(value):
        if pd.isna(value):
            return False  # Handle missing values
        # Extract only alphabetic characters
        letters = ''.join([char for char in str(value) if char.isalpha()])
        return len(letters) <= 2

    # Apply the Class Section filter
    filtered_df = email_filtered[email_filtered['Class Section'].apply(
        has_max_two_letters)]

    return filtered_df


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


def create_weekday_date_dict(start_date_str, end_date_str):
    """Creates a dictionary mapping weekdays to dates within a specified range"""
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


def expand_df_with_dates(merged_df, start_date_str, end_date_str):
    """Step 3: Expand DataFrame with dates"""
    # Get weekday-date mapping
    weekday_date_dict = create_weekday_date_dict(start_date_str, end_date_str)

    # Define valid weekdays
    valid_days = {'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat'}

    # Create a list to store expanded rows
    expanded_rows = []
    skipped_rows = 0

    # Process each row in the merged DataFrame
    for _, row in merged_df.iterrows():
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

    return expanded_df, skipped_rows

# Helper function to convert DataFrame to Excel for download


def to_excel(df):
    output = BytesIO()
    writer = pd.ExcelWriter(output, engine='xlsxwriter')
    df.to_excel(writer, index=False, sheet_name='Sheet1')
    writer.close()
    processed_data = output.getvalue()
    return processed_data


# Sidebar for navigation
st.sidebar.title("Workflow Steps")
st.sidebar.markdown("### Current Step")

# Create a radio button with step names
selected_step_name = st.sidebar.radio(
    "**Select Step**",
    STEP_NAMES,
    index=st.session_state.current_step
)

# Update current_step based on the selected step name
st.session_state.current_step = STEP_NAMES.index(selected_step_name)

# Main content area
st.title("Teaching Claim Workflow")
st.markdown("### Process your data in three simple steps")

# Step 1: Filter Data
if st.session_state.current_step == 0:  # Filter ASRQ180
    st.header("Step 1: Filter ASRQ180")
    st.markdown("""
    **Objective**: \n
        - Filter data by adjunct (@ADJ.NP.EDU.SG)
        - Remove duplicates (class section with 3 letters) 
        - Expand rows (Day column with more than a weekday).
    
    **Instructions**: 
    1. Upload a ASRQ180 file.
    2. The app will filter, remove and expand rows and display results.
    3. Download the filtered data and proceed to Step 2.
    """)

    # File upload
    uploaded_file = st.file_uploader(
        "**Upload ASRQ180 (xlsx file)**", type=["xlsx"])

    if uploaded_file is not None:
        # Read the Excel file
        df = pd.read_excel(uploaded_file)

        # Process the data
        with st.spinner("Processing your data..."):
            filtered_df = filter_data(df)
            # Apply the expansion function to the filtered DataFrame
            multiday, filtered_df = expand_day_column(filtered_df)
            st.session_state.step1_data = filtered_df

        # Display results
        st.subheader("Filtered Data Results")
        st.write(f"Original rows: {len(df)}")
        st.write(f"Filtered rows: {len(filtered_df)}")
        st.write(f"Expanded rows with multiple DAY: {multiday}")
        st.dataframe(filtered_df)

        # Download button
        st.download_button(
            label="Download Filtered Data",
            data=to_excel(filtered_df),
            file_name="filtered_results.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

        # Proceed to next step
        if st.button("Proceed to Step 2"):
            st.session_state.current_step = 1
            st.rerun()
    else:
        st.info("Please upload your main data file to begin processing.")

# Step 2: Merge Data
elif st.session_state.current_step == 1:  # Merge ASRQ180 headers with Hiring Form
    st.header("Step 2: Merge ASRQ180 with Hiring Form")
    st.markdown("""
    **Objective**: Partial match 'Name' column from ASRQ180 with Hiring Form 'Full Name' Column.
    
    **Instructions**: 
    1. Upload hiring form
    2. The app will merge the filtered data with the lookup data
    3. Download the merged data and proceed to Step 3
    """)

    # Check if we have data from Step 1
    if st.session_state.step1_data is None:
        st.warning("Please complete Step 1 first.")
        if st.button("Go to Step 1"):
            st.session_state.current_step = 0
            st.rerun()
    else:
        # File upload
        uploaded_file = st.file_uploader(
            "Upload your lookup data file (Excel format)", type=["xlsx"])

        if uploaded_file is not None:
            # Read the Excel file
            lookup_df = pd.read_excel(uploaded_file)

            # Display column names for debugging
            st.subheader("Lookup File Columns")
            st.write("Column names in the uploaded lookup file:")
            st.code(", ".join(lookup_df.columns.tolist()))

            # Check for required columns
            required_columns = ['Full Legal Name', 'Empl ID', 'Time entry code',
                                'Position ID', 'Program ID', 'Requester Remarks']
            missing_columns = [
                col for col in required_columns if col not in lookup_df.columns]

            if missing_columns:
                st.error(
                    f"The lookup file is missing the following required columns: {', '.join(missing_columns)}")
                st.info(
                    "Please check your file and ensure it contains all required columns.")
            else:
                # Process the data
                with st.spinner("Merging data..."):
                    merged_df, unmatched_df, unmatched_count = merge_with_partial_match(
                        st.session_state.step1_data, lookup_df)
                    st.session_state.step2_data = merged_df

                st.subheader("Merge Results Summary")
                st.write(f"Filtered rows: {len(st.session_state.step1_data)}")
                st.write(f"Merged rows: {len(merged_df)}")
                st.dataframe(merged_df)

                # Display unmatched rows instead of merged data
                st.subheader("Unmatched Rows")
                if unmatched_count > 0:
                    # st.write(f"Unmatched rows: {unmatched_count}")
                    st.write(
                        f"Displaying {unmatched_count} unmatched rows:")
                    st.dataframe(unmatched_df)

                # Download button
                st.download_button(
                    label="Download Merged Data",
                    data=to_excel(merged_df),
                    file_name="merged_output.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

                # Proceed to next step
                if st.button("Proceed to Step 3"):
                    st.session_state.current_step = 2
                    st.rerun()
        else:
            st.info("Please upload your lookup data file.")

# Step 3: Expand with Dates
elif st.session_state.current_step == 2:  # Expand rows with date range
    st.header("Step 3: Expand rows with date range")
    st.markdown("""
    **Objective**: Create a set of date range with Day of Week and map to ARSQ180.
    
    **Instructions**: 
    1. Set the date range parameters
    2. The app will expand the merged data with the date range
    3. Download the final expanded data
    """)

    # Check if we have data from Step 2
    if st.session_state.step2_data is None:
        st.warning("Please complete Step 2 first.")
        if st.button("Go to Step 2"):
            st.session_state.current_step = 1
            st.rerun()
    else:

        # Date range inputs
        st.subheader("Date Range Parameters")
        col1, col2 = st.columns(2)

        with col1:
            start_date_obj = st.date_input(
                "Start Date",
                value=datetime.date(2025, 4, 21),
                min_value=datetime.date(2020, 1, 1),
                max_value=datetime.date(2030, 12, 31)
            )
            start_date = start_date_obj.strftime("%d %B %Y")

        with col2:
            end_date_obj = st.date_input(
                "End Date",
                value=datetime.date(2025, 8, 23),
                min_value=datetime.date(2020, 1, 1),
                max_value=datetime.date(2030, 12, 31)
            )
            end_date = end_date_obj.strftime("%d %B %Y")

        # Process the data
        if st.button("Expand Data"):
            with st.spinner("Expanding data with dates..."):
                expanded_df, skipped_rows = expand_df_with_dates(
                    st.session_state.step2_data, start_date, end_date
                )
                st.session_state.step3_data = expanded_df

            # Display results
            st.subheader("Expanded Data")
            st.write(f"Merged rows: {len(st.session_state.step2_data)}")
            st.write(f"Expanded rows: {len(expanded_df)}")
            st.write(f"Skipped rows: {skipped_rows}")
            st.dataframe(expanded_df)

            # Download button
            st.download_button(
                label="Download Expanded Data",
                data=to_excel(expanded_df),
                file_name="expanded_with_dates.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            # Success message
            st.success("Processing complete! You can download your final data.")
        else:
            st.info("Click 'Expand Data' to process with the selected date range.")

# Footer
st.sidebar.markdown("---")
st.sidebar.markdown("### Instructions")
st.sidebar.markdown("""
1. Complete each step in order
2. Upload the required files
3. Review the results
4. Download the processed data
5. Proceed to the next step
""")

st.sidebar.divider()

st.sidebar.markdown("### About")
st.sidebar.markdown("""
This app processes adjunct lecturer data through three steps:
1. Filtering
2. Merging
3. Date expansion
""")
