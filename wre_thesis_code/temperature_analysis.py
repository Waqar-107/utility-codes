# from dust i have come, dust i will be

import xlrd
import xlwt
from datetime import datetime

location = "E:\Programming\experimenting-api\wre_thesis_code\TEMPERATUREDATA.xlsx"

wb = xlrd.open_workbook(location)
sheet = wb.sheet_by_index(0)

# extracting number of rows
row = sheet.nrows

# extracting number of columns
col = sheet.ncols

'''
task : for station id => 11313 
find max and min average temperature for each month from 1981-2010 (mathematical average)
then find avg of january from 1981-1990
'''
# from 1981 to 2010, 1-12 (jan-dec)
max_temp = {}
min_temp = {}
mx_cnt = {}
mn_cnt = {}

# initialize the dictionary
for i in range(1981, 2011):
    mx_cnt[i] = {}
    mn_cnt[i] = {}
    max_temp[i] = {}
    min_temp[i] = {}

    for j in range(1, 13):
        mx_cnt[i][j] = 0
        mn_cnt[i][j] = 0

        max_temp[i][j] = 0
        min_temp[i][j] = 0

mx_ignored = 0
mn_ignored = 0

for i in range(1, row):
    station_code = sheet.cell_value(i, 0)
    if station_code != "11313":
        continue

    # col1 : date, col3: max col5: min
    excel_date = int(sheet.cell_value(i, 1))
    dt = datetime.fromordinal(datetime(1900, 1, 1).toordinal() + excel_date - 2)

    dt = str(dt)
    dt = dt.split(' ')
    y, m, d = dt[0].split('-')

    y = int(y)
    m = int(m)

    if y > 2010:
        continue

    try:
        max_temp[y][m] += float(sheet.cell_value(i, 3))
        mx_cnt[y][m] += 1
    except ValueError:
        mx_ignored += 1

    try:
        min_temp[y][m] += float(sheet.cell_value(i, 5))
        mn_cnt[y][m] += 1
    except ValueError:
        mn_ignored += 1

output = xlwt.Workbook()
output_sheet = output.add_sheet("temperature_avg")

months = ["January", "February", "March", "April", "May", "June",
          "July", "August", "September", "October", "November", "December"]

output_sheet.write(0, 0, "Year")
output_sheet.write(0, 1, "Month")
output_sheet.write(0, 2, "Max Temperature Avg")
output_sheet.write(0, 3, "Min Temperature Avg")

r = 1

# determine avg
for i in range(1981, 2011):
    for j in range(1, 13):
        if mx_cnt[i][j]:
            max_temp[i][j] /= mx_cnt[i][j]

        if mn_cnt[i][j]:
            min_temp[i][j] /= mn_cnt[i][j]

        output_sheet.write(r, 0, str(i))
        output_sheet.write(r, 1, months[j - 1])
        output_sheet.write(r, 2, round(max_temp[i][j], 3))
        output_sheet.write(r, 3, round(min_temp[i][j], 3))

        r += 1

output.save("temperature_avg.xls")