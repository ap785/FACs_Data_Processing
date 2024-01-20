import pandas as pd
import re


# Read the Excel file
file_path = 'FACs_test.xls'
df = pd.read_excel(file_path)

# Initialize an empty dictionary to store the final result
result_dict = {}

# Function to clean column names for valid Excel sheet names
def clean_sheet_name(name):
    return re.sub(r'[\[\]:*?/\\]', '', name)[:31]  # Limit to 31 characters

# Iterate through each column and create separate DataFrames
for col in df.columns:
    # Slice the column into sets of 10 values
    sliced_data = [df[col][i:i+10].reset_index(drop=True) for i in range(0, len(df), 10) if i + 10 <= len(df)]

    # Find the maximum length of the sliced data
    max_len = max(len(slice_data) for slice_data in sliced_data)

    # Pad the shorter slices with NaN values to make them of equal length
    sliced_data = [slice_data.reindex(range(max_len)).fillna(method='ffill').fillna(method='bfill') for slice_data in sliced_data]

    # Concatenate consecutive slices horizontally
    col_slices = pd.concat(sliced_data, axis=1)

    # Rename columns based on the index
    col_slices.columns = [f'{col}_{i+1}' for i in range(col_slices.shape[1])]

    # Clean the column name for valid sheet name
    sheet_name = clean_sheet_name(col)

    # Add the DataFrame to the dictionary
    result_dict[sheet_name] = col_slices

# Export each DataFrame to an Excel file
with pd.ExcelWriter('output_file.xlsx') as writer:
    for sheet_name, col_df in result_dict.items():
        col_df.to_excel(writer, sheet_name=sheet_name, index=False)
          
# Print the DataFrame for each column
for col, col_df in result_dict.items():
    print(f"DataFrame for column '{col}':")
    print(col_df)
    print("\n")

