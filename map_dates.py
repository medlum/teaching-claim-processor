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


def map_dates_to_weeks(schedule_dict):
    week_mapping = {}

    for day, dates in schedule_dict.items():
        for index, date in enumerate(dates):
            week_number = f"Week {index + 1}"
            week_mapping[date] = week_number

    return week_mapping


# Example usage:
start_date = "21 April 2025"
end_date = "23 August 2025"
result = create_weekday_date_dict(start_date, end_date)

# Print the result
for day, dates in result.items():
    print(f"{day}: {dates[:5]}...")  # Show first 5 dates for each day

# Print week number 1
week_mapping = map_dates_to_weeks(result)
print(week_mapping)
# Example output:
for date, week in sorted(week_mapping.items()):
    print(f"{date}: {week}")
