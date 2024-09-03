import openpyxl

# Body of the email as a string
Body = "Name Patricia Brown Date 2024-06-28 Column used SEC-17 runs 6 End pressure 0.5 Flow rate 1 Column cleaned ? no solution equilibrated Ethanol solution Errors/Comments Great it was a pleasure working with you"

#functions to receive information
def find_between(text, first, last ):
    try:
        start = text.index( first ) + len( first )
        end = text.index( last, start )
        return text[start:end]
    except ValueError:
        return ""

def substring_after(text, after):
    return text.partition(after)[2]

# create variables with info
name = find_between(Body, "Name ", "Date")
date = find_between(Body, "Date ", "Column")
column = find_between(Body, "used ", "runs")
runs = find_between(Body, "runs ", "End")
pressure = find_between(Body, "pressure ", "Flow")
flow = find_between(Body, "rate ", "Column")
clean = find_between(Body, "? ", "solution")
solution = find_between(Body, "equilibrated ", "Errors/Comments")
comments = substring_after(Body, "Errors/Comments ")

# add variable to data frame
new_data = [[name, date, column, runs, pressure, flow, clean, solution, comments]]


# Open existing workbook and select the active worksheet
workbook = openpyxl.load_workbook('ColumnLogbook_QR.xlsx')
sheet = workbook.active

# Append new data
for row in new_data:
    sheet.append(row)

# Save the workbook
workbook.save('ColumnLogbook_QR.xlsx')
