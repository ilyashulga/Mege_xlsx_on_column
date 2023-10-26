import pandas as pd

def merge_xlsx_files(file1, file2, output_file):
    # Load the two Excel files into pandas DataFrames
    df1 = pd.read_excel(file1, engine='openpyxl', skiprows=4)
    df2 = pd.read_excel(file2, engine='openpyxl', skiprows=4)
    
    # Merge the two DataFrames based on "Serial Number"
    merged_df = pd.merge(df1, df2, on="Serial Number", how="inner")
    
    # Save the merged DataFrame to an Excel file
    merged_df.to_excel(output_file, index=False, engine='openpyxl')
    
    # Print the number of matched rows
    print(f"Number of matched rows: {len(merged_df)}")

# Paths to the Excel files
file1_path = "Files/Session.xlsx"
file2_path = "Files/Server.xlsx"
output_path = "Files/Merged.xlsx"

merge_xlsx_files(file1_path, file2_path, output_path)
