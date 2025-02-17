import pandas as pd

# Load the two Excel files into pandas DataFrames
file1= 'C:/Vivek/pre.xlsx' 
file2= 'C:/Vivek/post.xlsx'

df1 = pd.read_excel(file1)
df2 = pd.read_excel(file2)

# Ensure both DataFrames have the same columns
if not df1.columns.equals(df2.columns):
    raise ValueError("The columns of both Excel files are different!")

# Prepare a list to store the differences
differences = []

# Compare rows in both DataFrames
min_length = min(len(df1), len(df2))
max_length = max(len(df1), len(df2))

# Compare matching rows in both files
for idx in range(min_length):
    row1 = df1.iloc[idx]
    row2 = df2.iloc[idx]
    
    for col in df1.columns:
        if row1[col] != row2[col]:
            differences.append({
                'Row Index': idx + 1,  # 1-based index for human readability
                'Column': col,
                'File 1 Value': row1[col],
                'File 2 Value': row2[col]
            })

# Handle additional rows in df1 if df1 has more rows
if len(df1) > len(df2):
    for idx in range(min_length, len(df1)):
        differences.append({
            'Row Index': idx + 1,
            'Column': 'Missing row',  # No matching row in df2
            'File 1 Value': 'Row Exist',  # Show row from file1
            'File 2 Value': 'Missing row'  # No matching row in file2
        })

# Handle additional rows in df2 if df2 has more rows
if len(df2) > len(df1):
    for idx in range(min_length, len(df2)):
        differences.append({
            'Row Index': idx + 1,
            'Column': 'Missing row',  # No matching row in df1
            'File 1 Value': 'Missing row',  # No matching row in file1
            'File 2 Value': 'Row exist'  # Show row from file2
        })

# If there are no differences
if not differences:
    print("No differences found between the files.")
else:
    # Convert the list of differences into a DataFrame
    differences_df = pd.DataFrame(differences)

    # Save the difference report to an Excel file
    differences_report_path = 'C:/Vivek/difference_report.xlsx'
    differences_df.to_excel(differences_report_path, index=False)
    print(f'Difference report saved to {differences_report_path}')  