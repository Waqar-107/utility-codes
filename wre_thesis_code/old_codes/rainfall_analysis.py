# from dust i have come, dust i will be

import xlrd
import xlwt

location = "E:\Programming\experimenting-api\wre_thesis_code\ActualData\SW132.xlsx"

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

output_sheet.write(0, 0, "Year")
output_sheet.write(0, 1, "Month")
output_sheet.write(0, 2, "Avg Rainfall")

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
day_count = [[0 for y in range(12)] for x in range(3000)]


for i in range(1, row):
    dt = sheet.cell_value(i, 3)
    d, m, y = dt.split('-')

    year = int(y)
    month = months[m]

    if year < 1981 or year > 2010:
        continue

    rainfall = float(sheet.cell_value(i, 4))

    #print(year, month, m, months[m], dt, i)

    data[year][month] += rainfall
    day_count[year][month] += 1


for j in range(12):
    for k in range(1981, 2011, 1):
        val = data[k][j]

        if day_count[k][j]:
            val /= day_count[k][j]

        val = round(val, 3)

        last_row_written += 1

        output_sheet.write(last_row_written, 0, k)
        output_sheet.write(last_row_written, 1, (j + 1))
        output_sheet.write(last_row_written, 2, val)

output.save("Rainfall_avg.xls")

# -------------------------------------------------------------------

output2 = xlwt.Workbook()
output_sheet2 = output2.add_sheet("10_year_rainfall")

output_sheet2.write(0, 0, "Year-Range")
output_sheet2.write(0, 1, "Month")
output_sheet2.write(0, 2, "10Y Avg Rainfall")

y_st = [1981, 1991, 2001]
y_end = [1991, 2001, 2011]

last_row_written = 0

for k in range(3):
    for j in range(12):
        val = 0.0
        for i in range(y_st[k], y_end[k], 1):
            val += data[i][j]

        val /= 10.0
        val = round(val, 3)

        last_row_written += 1

        output_sheet2.write(last_row_written, 0, str(y_st[k]) + '-' + str(y_end[k]))
        output_sheet2.write(last_row_written, 1, str(j + 1))
        output_sheet2.write(last_row_written, 2, str(val))

output2.save("10_year_rainfall_avg.xls")