# from dust i have come, dust i will be

import xlrd
import xlwt

startFrom = 9
skipZero = True

location = "ActualData\MAXTEMPERATURE.xls"

wb = xlrd.open_workbook(location)
sheet = wb.sheet_by_index(0)

# extracting number of rows
row = sheet.nrows

# extracting number of columns
col = sheet.ncols

output = xlwt.Workbook()
output_sheet = output.add_sheet("newSheet")

'''
task : for district -> Comilla, Feni, Chandpur
from 1981 to 2010 => for each month => divide the month in 3 parts (each of 10 days)
'''

req_district = ['Comilla', 'Feni', 'Chandpur']

output_sheet.write(0, 0, "District")
output_sheet.write(0, 1, "Year")
output_sheet.write(0, 2, "Month")
output_sheet.write(0, 3, "Decade 1")
output_sheet.write(0, 4, "Decade 2")
output_sheet.write(0, 5, "Decade 3")

last_row_written = 0
for i in range(startFrom, row, 1):
    district = sheet.cell_value(i, 0)

    if district not in req_district:
        continue

    year = int(sheet.cell_value(i, 1))
    month = int(sheet.cell_value(i, 2))

    if year < 1981 or year > 2010:
        continue

    # if anything other than numeric value found then assign -1
    data = [0] * 31
    for j in range(31):
        try:
            data[j] = float(sheet.cell_value(i, j + 3))
        except ValueError:
            data[j] = -1

    st = [0, 10, 20]
    en = [10, 20, 31]

    last_row_written += 1
    output_sheet.write(last_row_written, 0, district)
    output_sheet.write(last_row_written, 1, year)
    output_sheet.write(last_row_written, 2, month)

    for j in range(3):
        cnt = 0
        total = 0

        for k in range(st[j], en[j], 1):
            if data[j] == -1:
                continue
            if data[j] > 0:
                total += data[j]
                cnt += 1
            elif data[j] == 0 and not skipZero:
                cnt += 1
        try:
            total /= cnt
        except ZeroDivisionError:
            total = 0

        output_sheet.write(last_row_written, 3 + j, round(total, 3))

output.save("output1.xls")
