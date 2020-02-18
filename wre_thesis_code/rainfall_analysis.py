# from dust i have come, dust i will be

import xlrd
import xlwt

location = "E:\Programming\experimenting-api\wre_thesis_code\SW132.xlsx"

wb = xlrd.open_workbook(location)
sheet = wb.sheet_by_index(0)

# extracting number of rows
row = sheet.nrows

# extracting number of columns
col = sheet.ncols

output = xlwt.Workbook()
output_sheet = output.add_sheet("rainfall")

'''
task :
find average rainfall for each month from 1981-2010 (mathematical average)
'''

output_sheet.write(0, 0, "District")
output_sheet.write(0, 1, "Year")
output_sheet.write(0, 2, "Month")
output_sheet.write(0, 3, "Avg Rainfall")

last_row_written = 0

months = {
    'jan': 0,
    'feb': 1,
    'mar': 2,
    'apr': 3,
    'may': 4,
    'jun': 5,
    'jul': 6,
    'aug': 7,
    'sep': 8,
    'oct': 9,
    'nov': 10,
    'dec': 11
}

data = [[0 for j in range(12)] for i in range(3000)]
day_count = [[0 for j in range(12)] for i in range(3000)]

for i in range(1, row):
    dt = sheet.cell_value(1, 3)
    d, m, y = dt.split('-')

    year = int(y)
    month = months[m]

    if year < 1981 or year > 2010:
        continue

    rainfall = float(sheet.cell_value(i, 4))

    data[year][month] += rainfall
    day_count[year][month] += 1

for j in range(12):
    for k in range(1981, 2011, 1):
        val = data[k][j]

        if day_count[k][j]:
            val /= day_count[k][j]

        val = round(val, 3)

        last_row_written += 1

        output_sheet.write(last_row_written, 1, str(k))
        output_sheet.write(last_row_written, 2, str(j + 1))
        output_sheet.write(last_row_written, 3, str(val))

#output.save("Rainfall_avg.xls")