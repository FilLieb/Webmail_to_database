import openpyxl
import os


# Open existing workbook and select the active worksheet
workbook = openpyxl.load_workbook('ColumnLogbook.xlsx')
sheet = workbook.active


Body = [Name Patricia
Date 2024-06-28
Column used SEC-17
# runs 6
End pressure 0.5
Flow rate 1
Column cleaned ? no
solution equilibrated Ethanol
Errors/Comments Great]


sheet.append(['Subject', 'From', 'Date', 'Body'])

# Save the workbook
workbook.save('ColumnLogbook.xlsx')