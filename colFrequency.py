import os
import glob
import pandas as pd
import csv

# specify the path to the parent directory
parent_dir_path = '/Users/sasha/desktop/bayanat-data'
column_freq_dict = {}

# loop through each child directory in the parent directory
for child_dir in os.listdir(parent_dir_path):
    # check if the child is a directory
    if os.path.isdir(os.path.join(parent_dir_path, child_dir)):
        # get a list of all excel files in the child directory
        excel_files = glob.glob(os.path.join(parent_dir_path, child_dir, "*.xlsx"))
        # loop through each excel file in the child directory
        for excel_file in excel_files:
            # check if the file name starts with "(metadata)"
            if not excel_file.startswith(os.path.join(parent_dir_path, child_dir, "(metadata)")):
                excel_name = os.path.basename(excel_file)
                df = pd.read_excel(excel_file, sheet_name=None)
                # loop through each sheet in the Excel file
                for sheet_name, sheet_df in df.items():
                    # loop through each column in the sheet dataframe
                    for col in sheet_df.columns:
                        # check if the column is of string type
                        if isinstance(sheet_df[col].iloc[0], str):
                            # check if the column name is already in the dictionary
                            col_lower = col.lower()
                            if col_lower not in column_freq_dict:
                                # add the column name to the dictionary for the first time
                                column_freq_dict[col_lower] = {"Frequency": 1, "Excel File": excel_name, "Top Values": []}
                                # get the three most common values in the column
                                top_values = sheet_df[col].value_counts().iloc[:3].index.tolist()
                                column_freq_dict[col_lower]["Top Values"] = top_values
                            else:
                                # increment the column frequency
                                column_freq_dict[col_lower]["Frequency"] += 1

# write the final dictionary to a CSV file
with open("column_freq_1.csv", "w") as csv_file:
    writer = csv.writer(csv_file)
    writer.writerow(["Column Name", "Frequency", "Excel File", "Most Occurring Values"])
    for key, value in column_freq_dict.items():
        writer.writerow([key, value["Frequency"], str(value["Excel File"]), ', '.join(map(str, value["Top Values"]))])

