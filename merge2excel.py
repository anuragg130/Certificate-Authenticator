import pandas as pd

def merge_excel_files(file1, file2, common_column):
    # Read both Excel files
    df1 = pd.read_excel(file1)
    df2 = pd.read_excel(file2)

    # Merge the data frames based on the common_column
    merged_df = pd.merge(df1, df2, on=common_column, how='inner')

    return merged_df

if __name__ == "__main__":
    file1_path = "D:\EY\Cerificate Scanner\output42.xlsx"
    file2_path = "D:\EY\Cerificate Scanner\outputqr.xlsx"
    common_column_name = "File Name"

    merged_data = merge_excel_files(file1_path, file2_path, common_column_name)

    # Save the merged data to a new Excel file
    output_file = "D:\EY\Cerificate Scanner/final_output.xlsx"
    merged_data.to_excel(output_file, index=False, engine='openpyxl')

    print("Files merged and saved successfully.")
