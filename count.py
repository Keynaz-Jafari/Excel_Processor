from openpyxl import load_workbook

# Load the modified Excel file
file_path = 'knoush.xlsx'
workbook = load_workbook(filename=file_path)
sheet = workbook.active

# Dictionary to store all names and their counts in each drugstore
name_counts = {}


# Iterate through rows to collect all names and their counts
for row_index in range(3, sheet.max_row + 1):
    for cell in sheet[row_index][2:]:  # Start from column index 2 (C column)
        name = cell.value
        if name:
         
            # Increment the count for the current name in the corresponding drugstore
            if cell.column == 3 or cell.column ==4:
                drugstore_name = sheet.cell(row=1, column=3).value                
            if cell.column == 5 or cell.column == 6 or cell.column ==7:
                drugstore_name = sheet.cell(row=1, column=5).value
            if cell.column == 8 or cell.column ==9 :
                drugstore_name = sheet.cell(row=1, column=8).value
            if cell.column == 10 or cell.column ==11 or cell.column ==12:
                drugstore_name = sheet.cell(row=1, column=10).value
            if cell.column == 13 or cell.column ==14 or cell.column ==15:
                drugstore_name = sheet.cell(row=1, column=13).value
            if cell.column == 16 or cell.column ==17 or cell.column ==18:
                drugstore_name = sheet.cell(row=1, column=16).value
            if cell.column == 19:
                drugstore_name = sheet.cell(row=1, column=19).value
            if cell.column == 20:
                drugstore_name = sheet.cell(row=1, column=20).value

            if name not in name_counts:
                name_counts[name] = {drugstore_name: 1}
            else:
                if drugstore_name not in name_counts[name]:
                    name_counts[name][drugstore_name] = 1
                else:
                    name_counts[name][drugstore_name] += 1

# Write the report to a text file
report_file_path = 'name_counts_report.txt'
with open(report_file_path, 'w') as report_file:
    num =1
    for name, counts in name_counts.items():
        report_file.write(str(num)+'. '+f"Name: {name}\n")
        
        for drugstore, count in counts.items():
            report_file.write(f"{drugstore}\t{count}\n")
        report_file.write('\n')
        num+=1
print(f"Name counts report saved as {report_file_path}")
