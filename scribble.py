import openpyxl
import os


# Open existing workbook and select the active worksheet
workbook = openpyxl.load_workbook('ColumnLogbook.xlsx')
sheet = workbook.active


Body = "Name Patricia Brown Date 2024-06-28 Column used SEC-17 runs 6 End pressure 0.5 Flow rate 1 Column cleaned ? no solution equilibrated Ethanol solution Errors/Comments Great it was a pleasure working with you"

words = Body.split()

def nextword(target, source):
    for i, w in enumerate(source):
        if w == target:
            return source[i+1]

def find_between(Body, first, last ):
    try:
        start = Body.index( first ) + len( first )
        end = Body.index( last, start )
        return Body[start:end]
    except ValueError:
        return ""

def substring_after(s, delim):
    return s.partition(delim)[2]


name = find_between(Body, "Name ", "Date")

date = nextword("Date", words)

column = nextword("used", words)

runs = nextword("runs", words)

pressure = nextword("pressure", words)

flow = nextword("rate", words)

clean = nextword("?", words)

solution = find_between(Body, "equilibrated ", "Errors/Comments")

comments = substring_after(Body, "Errors/Comments ")

print(name)
print(date)
print(column)
print(runs)
print(pressure)
print(flow)
print(clean)
print(solution)
print(comments)



#sheet.append(['Subject', 'From', 'Date', 'Body'])

# Save the workbook
#workbook.save('ColumnLogbook.xlsx')