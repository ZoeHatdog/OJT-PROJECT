import pandas as pd

# Replace 'file_path.xlsx' with the path to your Excel file
file_path = 'BACKEND\excel_test.xlsx'

# Read the Excel file
df = pd.read_excel(file_path)

# Find the row with the word 'total'
total_row = df[df.apply(lambda row: row.astype(str).str.contains('Total Cost of Raw Materials - Pins', case=False).any(), axis=1)]

if not total_row.empty:
    total_row_index = total_row.index[0]
    total_row_values = df.iloc[total_row_index]

    # Find the number nearest to 'total' on the same row
    nearest_number = None
    for value in total_row_values:
        if isinstance(value, (int, float)) and not pd.isna(value):
            nearest_number = value
            break

    if nearest_number is not None:
        print("Number nearest to 'total':", nearest_number)
    else:
        print("No number found near 'total' in the same row.")
else:
    print("Word 'total' not found in the Excel file.")
