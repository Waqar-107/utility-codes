# from dust i have come, dust i will be

import xlrd
import xlwt
from datetime import datetime

location = "E:\Programming\experimenting-api\wre_thesis_code\ActualData\Wind.xlsx"

wb = xlrd.open_workbook(location)
sheet = wb.sheet_by_index(0)

# extracting number of rows
row = sheet.nrows

# extracting number of columns
col = sheet.ncols

'''
task : for station id => 11313 
find max wind for each month from 1981-2010
'''

output = xlwt.Workbook()
output_sheet = output.add_sheet("max_wind")

output_sheet.write(0, 0, "Year")
output_sheet.write(0, 1, "Month")
output_sheet.write(0, 2, "Value")

data = [[0.0 for j in range(12)] for i in range(3000)]

for i in range(1, row):
    stationId = int(sheet.cell_value(i, 0))
    if stationId != 11313:
        continue

    excel_date = int(sheet.cell_value(i, 2))
    dt = datetime.fromordinal(datetime(1900, 1, 1).toordinal() + excel_date - 2)

    dt = str(dt)
    dt = dt.split(' ')
    y, m, d = dt[0].split('-')

    y = int(y)
    m = int(m) - 1

    val = sheet.cell_value(i, 4)
    try:
        l = len(val)
        # print(val, l)
    except TypeError:
        data[y][m] = max(data[y][m], val)

last_row_written = 0

for i in range(1981, 2011):
    for j in range(12):
        last_row_written += 1

        val = round(data[i][j], 3)

        output_sheet.write(last_row_written, 0, i)
        output_sheet.write(last_row_written, 1, j + 1)
        output_sheet.write(last_row_written, 2, val)

output.save("wind_max.xls")

# ----------------------------------------

output2 = xlwt.Workbook()
output_sheet2 = output2.add_sheet("max_wind_10y")

output_sheet2.write(0, 0, "Year-Range")
output_sheet2.write(0, 1, "Month")
output_sheet2.write(0, 2, "Value")

st = [1981, 1991, 2001]
en = [1991, 2001, 2011]

last_row_written = 0
for i in range(3):
    for j in range(12):
        val = 0.0
        for k in range(st[i], en[i], 1):
            val += data[k][j]

        val /= 12
        val = round(val, 3)

        last_row_written += 1

        output_sheet2.write(last_row_written, 0, str(st[i]) + '-' + str(en[i] - 1))
        output_sheet2.write(last_row_written, 1, j + 1)
        output_sheet2.write(last_row_written, 2, val)

output2.save("wind_max_10y_avg.xls")