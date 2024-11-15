import pandas as pd
import time
import os

# Define the path to the Excel file
file_path = 'Raw Data/2015.xlsx'  # Update this with the path to your file
file_name = os.path.splitext(os.path.basename(file_path))[0]  # Extracts "Number" from e.g. "2015.xlsx"

# Load the Excel file
df = pd.read_excel(file_path)

# Count the number of rows with any data in them
num_rows_with_data = df.dropna(how='all').shape[0] + 1

# Count the number of columns with any data in them, excluding completely empty columns
num_columns_with_data = df.dropna(axis=1, how='all').shape[1]

print(f'The file has {num_rows_with_data} row(s) and {num_columns_with_data} column(s) with data')
time.sleep(1)

while True:
    question = input('Do you wish to split the data with unique values in rows or columns? (R or C)?')
    if question.upper() == 'C':
        # Loop over each column up to `num_columns_with_data` and create a DataFrame for each with "Key", "Value", and "Original Row" columns
        for col in range(num_columns_with_data):
            # Get the column values up to `num_rows_with_data` and drop any NaN values
            column_data = df.iloc[:num_rows_with_data, col].dropna()

            # Extract unique values along with their first occurrence row index
            unique_values = column_data.unique().tolist()
            original_row_indices = [column_data[column_data == value].index[0] + 2 for value in
                                    unique_values]  # +1 for 1-based indexing, +2 if header
            new_row_indices = original_row_indices - original_row_indices[0]

            # Create a DataFrame with "Key", "Value", and "Original Row" columns
            output_df = pd.DataFrame({
                "Key": range(1, len(unique_values) + 1),  # Key column from 1 to number of unique values
                "Value": unique_values,  # Value column with unique values
                "Original Row": original_row_indices,  # Original row numbers for each unique value
                "New Row": new_row_indices  # New row numbers for each unique value
            })

            # Save each DataFrame to a separate Excel file with the modified name
            output_file_path = f"{file_name}_unique_values_column_{col + 1}.xlsx"
            output_df.to_excel(output_file_path, index=False)

            print(f"Saved {len(unique_values)} unique values for Column {col + 1} in '{output_file_path}'")
        break

    elif question.upper() == 'R':
        print('Not complete yet ...')
        break

    else:
        time.sleep(1)
        print('Invalid input, please choose Rows (R) or Coloumns (C).')