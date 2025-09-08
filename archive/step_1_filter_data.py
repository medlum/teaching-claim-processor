import pandas as pd

# Read the Excel file
df = pd.read_excel('subset_data/test_data.xlsx')

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

# Display the result
print(filtered_df)

# Save to new Excel file
filtered_df.to_excel('processed_data/filtered_results.xlsx', index=False)
