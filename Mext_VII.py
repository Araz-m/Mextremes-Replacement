import os
import pandas as pd

# Define the directory containing the output files
output_dir = r"C:\Users\araz_\Desktop\GITHUB\Mextremes\DLC12"

# Dictionary to store extreme values for each parameter across files
extreme_values = {}

# Loop through each file in the output directory
for filename in os.listdir(output_dir):
    if filename.endswith(".out"):
        filepath = os.path.join(output_dir, filename)
        
        # Read the file, skipping the first 6 rows of comments
        data = pd.read_csv(filepath, skiprows=6, delimiter='\t')
        
        # Assume the first row after skipping contains the parameter names, second row contains units
        parameter_names = data.columns
        units = data.iloc[0]
        data = data.iloc[2:].reset_index(drop=True)  # Get actual values, starting from row 9
        
        # Find min and max values for each parameter in the current file
        for column in parameter_names:
            if column not in extreme_values:
                extreme_values[column] = {
                    "Min Values": [],
                    "Max Values": [],
                    "Unit": units[column]
                }
            
            min_value = data[column].min()
            max_value = data[column].max()
            
            # Append the min and max values along with the corresponding file name
            extreme_values[column]["Min Values"].append((min_value, filename))
            extreme_values[column]["Max Values"].append((max_value, filename))

# Prepare the DataFrame layout for Excel output
header_columns = ["Parameter", "Type", "File Name"] + list(extreme_values.keys())
rows = [[""] * len(header_columns) for _ in range(3)]
rows[0][3:] = list(extreme_values.keys())  # Parameter names in row 1
rows[1][3:] = [extreme_values[param]["Unit"] for param in extreme_values]  # Units in row 2

# Create rows for min and max values for each parameter
for parameter, values in extreme_values.items():
    for min_val, min_file in values["Min Values"]:
        row = [parameter, "min", min_file] + [""] * (len(header_columns) - 3)
        col_index = header_columns.index(parameter)
        row[col_index] = min_val
        rows.append(row)
    
    for max_val, max_file in values["Max Values"]:
        row = [parameter, "max", max_file] + [""] * (len(header_columns) - 3)
        col_index = header_columns.index(parameter)
        row[col_index] = max_val
        rows.append(row)

# Create DataFrame and output to Excel
final_output_df = pd.DataFrame(rows, columns=header_columns)
output_excel_path = r"C:\Users\araz_\Desktop\GITHUB\Mextremes\extreme_values_output.xlsx"
final_output_df.to_excel(output_excel_path, index=False)

print(f"Extreme values have been saved to {output_excel_path}")
