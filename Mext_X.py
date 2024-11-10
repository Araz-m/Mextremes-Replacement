import os
import pandas as pd
import warnings

# Suppress FutureWarnings 
warnings.simplefilter(action='ignore', category=FutureWarning)

# Define the directory containing the output files
output_dir = r"C:\Users\araz_\Desktop\GITHUB\Mextremes\DATA"

# Dictionary to store global extreme values for each parameter across files
extreme_values = {}

# Loop through each file in the output directory without verbose output
for filename in os.listdir(output_dir):
    if filename.endswith(".out"):
        filepath = os.path.join(output_dir, filename)
        
        # Read the file, skipping the first 6 rows of comments, with flexible whitespace delimiter
        data = pd.read_csv(filepath, skiprows=6, sep=r'\s+', header=0)
        
        # Assume the first row after skipping contains the parameter names, and second row contains units
        parameter_names = data.columns
        units = data.iloc[0]  # Units in row 8
        data = data.iloc[1:].reset_index(drop=True)  # Get actual values starting from row 9

        # Convert data to numeric, with non-numeric entries handled as NaN
        data = data.apply(pd.to_numeric, errors='coerce')

        # Find min and max values for each parameter in the current file
        for column in parameter_names:
            # Calculate min and max for the current file
            min_value = data[column].min()
            max_value = data[column].max()
            
            # Initialize storage for each parameter if not already present
            if column not in extreme_values:
                extreme_values[column] = {
                    "Min": {"value": min_value, "file": filename},
                    "Max": {"value": max_value, "file": filename},
                    "Unit": units[column]
                }
            else:
                # Update global min if the current min is lower
                if min_value < extreme_values[column]["Min"]["value"]:
                    extreme_values[column]["Min"] = {"value": min_value, "file": filename}
                
                # Update global max if the current max is higher
                if max_value > extreme_values[column]["Max"]["value"]:
                    extreme_values[column]["Max"] = {"value": max_value, "file": filename}

# Prepare the DataFrame layout for Excel output
header_columns = ["Parameter", "Type", "File Name"] + list(extreme_values.keys())
rows = [[""] * len(header_columns) for _ in range(3)]
rows[0][3:] = list(extreme_values.keys())  # Parameter names in row 1
rows[1][3:] = [extreme_values[param]["Unit"] for param in extreme_values]  # Units in row 2

# Create rows for the min and max values for each parameter, formatted to 3 decimal places
for parameter, values in extreme_values.items():
    # Add the minimum row
    min_row = [parameter, "Minimum", values["Min"]["file"]] + [""] * (len(header_columns) - 3)
    min_row[header_columns.index(parameter)] = f"{values['Min']['value']:.3e}"
    rows.append(min_row)
    
    # Add the maximum row
    max_row = [parameter, "Maximum", values["Max"]["file"]] + [""] * (len(header_columns) - 3)
    max_row[header_columns.index(parameter)] = f"{values['Max']['value']:.3e}"
    rows.append(max_row)

# Create DataFrame and output to Excel with scientific notation and 3 decimal places
final_output_df = pd.DataFrame(rows, columns=header_columns)
output_excel_path = r"C:\Users\araz_\Desktop\GITHUB\Mextremes\MEX_File\extreme_values_output.xlsx"

# Writing to Excel with precision formatted to scientific notation
with pd.ExcelWriter(output_excel_path) as writer:
    final_output_df.to_excel(writer, index=False, float_format="%.3e")

print(f"Extreme values have been saved to {output_excel_path}")
