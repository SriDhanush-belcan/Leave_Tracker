import pandas as pd
import tkinter as tk
from tkinter import filedialog
from openpyxl.styles import Font, Alignment

# Create a Tkinter window
root = tk.Tk()
root.withdraw()  # Hide the main window

import pandas as pd
from openpyxl.styles import Font, Alignment

# Specify the file paths for the Excel files
file_path1 = 'Deltek_Output.xlsx'  # Replace 'path_to_file1.xlsx' with your file path
file_path2 = 'PaySquare_Output.xlsx'  # Replace 'path_to_file2.xlsx' with your file path

# Read the Excel files
df1 = pd.read_excel(file_path1)
df2 = pd.read_excel(file_path2)

# Rest of your script remains unchanged...


# Group by 'Employee ID' in the first file and calculate total hours
grouped1 = df1.groupby('Employee ID').agg({'Employee Name': 'first', 'Hours': 'sum'})

# Group by 'Emp Code' in the second file and calculate total hours
grouped2 = df2.groupby('Emp Code').agg({'Employee Name': 'first', 'Total No of Hours': 'sum'})

# Compare total hours between the two files
mismatched_hours = []
for emp_id, (name1, hours1) in grouped1.iterrows():
    if emp_id in grouped2.index:
        data2 = grouped2.loc[emp_id]
        if len(data2) == 2:  # Check if data2 contains two elements
            name2, hours2 = data2
            absolute_diff = abs(hours1 - hours2)  # Absolute difference between hours
            if hours1 != hours2:
                mismatched_hours.append((emp_id, name1, hours1, hours2, absolute_diff, name2))
        else:
            print(f"Unexpected data structure for employee ID: {emp_id}")

# Create a DataFrame for mismatched data
columns = ['Employee ID', 'Employee Name in Deltek', 'Total Hours in Deltek', 'Total Hours in Paysquare', 'Hours Difference', 'Employee Name in Paysquare']
mismatched_df = pd.DataFrame(mismatched_hours, columns=columns)

# Save the mismatched data to an Excel file
excel_output_file = 'Mismatched_data.xlsx'

with pd.ExcelWriter(excel_output_file, engine='openpyxl') as writer:
    mismatched_df.to_excel(writer, index=False)
    workbook = writer.book
    worksheet = writer.sheets['Sheet1']

    # Adjust column widths based on the length of the content in each cell
    for column in worksheet.columns:
        max_length = 0
        column = column[0].column_letter  # Get the column name
        for cell in worksheet[column]:
            try:
                if len(str(cell.value)) > max_length:
                    max_length = len(cell.value)
            except:
                pass
        adjusted_width = (max_length + 2) * 1.2  # Adjusted width based on content length
        worksheet.column_dimensions[column].width = adjusted_width

    # Format the 'Absolute Difference' column
    absolute_diff_column = worksheet['E']  # Change 'E' to the appropriate column letter
    font = Font(bold=True)
    alignment = Alignment(horizontal='center')
    for cell in absolute_diff_column:
        cell.font = font
        cell.alignment = alignment

print(f"Excel file '{excel_output_file}' created.")

# Close the Tkinter window
root.destroy()
