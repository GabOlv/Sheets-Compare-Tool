from tkinter import filedialog as files
import pandas as pd

def open_file():
    try:
        file_path_1 = files.askopenfilename(title="Select Excel file", filetypes=[("Excel files", "*.xlsx;*.xls"), ("All files", "*.*")])
        print("File 1: ", file_path_1)
        sheet_name1 = input("Enter the sheet name for the first file (press Enter for default): ")
        if sheet_name1 == '':
            sheet_name1 = None  # Default sheet name
        print("Selected sheet for File 1:", sheet_name1)

        file_path_2 = files.askopenfilename(title="Select Excel file", filetypes=[("Excel files", "*.xlsx;*.xls"), ("All files", "*.*")])
        print("File 2: ", file_path_2)
        sheet_name2 = input("Enter the sheet name for the second file (press Enter for default): ")
        if sheet_name2 == '':
            sheet_name2 = None  # Default sheet name
        print("Selected sheet for File 2:", sheet_name2)

        if file_path_1 and file_path_2:
            load_files(file_path_1, sheet_name1, file_path_2, sheet_name2)
        else:
            print("Error")
    except Exception as e:
        print(e)

def load_files(path1, sheet1, path2, sheet2):
    try:
        if sheet1 is None: # Case the sheet1 is the default value
            df1 = pd.read_excel(path1)
        else:
            df1 = pd.read_excel(path1, sheet_name=sheet1)
        
        if sheet2 is None: # Case the sheet1 is the default value
            df2 = pd.read_excel(path2)
        else:
            df2 = pd.read_excel(path2, sheet_name=sheet2)
        compare_excel(df1, df2)
    except Exception as e:
        print("Cannot read the files. Error:", e)

def compare_excel(df1, df2):
    # Merge DataFrames on all columns
    merged_df = pd.merge(df1, df2, how='outer', indicator=True)

    # Filter rows where the indicator column is not both
    differences = merged_df[merged_df['_merge'] != 'both']

    # Display the differences
    if not differences.empty:
        print("\nDifferences between DataFrames:")
        for row in differences.index:
            file1_values = df1.loc[row].to_dict()
            file2_values = df2.loc[row].to_dict()

            print(f"\nRow {row + 2} -")
            print("  File 1:", ', '.join(f'[{col} | {val}]' for col, val in file1_values.items()))
            print("  File 2:", ', '.join(f'[{col} | {val}]' for col, val in file2_values.items()))
    else:
        print("\nNo differences found.")

open_file()
i = input("Digite qualquer botão para finalizar...")