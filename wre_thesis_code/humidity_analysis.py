# from dust i have come, dust i will be

import xlrd
import xlwt

location = "E:\Programming\experimenting-api\wre_thesis_code\RELATIVEHUMIDITY.xls"

wb = xlrd.open_workbook(location)
sheet = wb.sheet_by_index(0)

# extracting number of rows
row = sheet.nrows

# extracting number of columns
col = sheet.ncols

output = xlwt.Workbook()
output_sheet = output.add_sheet("humidity_avg")

'''
task : for district -> Comilla, Feni, Chandpur
find average humidity for each month from 1981-2010 (mathematical average)
'''

req_district = ['Comilla', 'Feni', 'Chandpur']

output_sheet.write(0, 0, "District")
output_sheet.write(0, 1, "Year")
output_sheet.write(0, 2, "Month")
output_sheet.write(0, 3, "Avg humidity")

last_row_written = 0

for i in range(8, row):
    district = sheet.cell_value(i, 0)

    if district not in req_district:
        continue

    year = int(sheet.cell_value(i, 1))
    month = int(sheet.cell_value(i, 2))

    if year < 1981 or year > 2010:
        continue

    cnt = 0
    humidity = 0.0
    for j in range(3, 33):
        val = sheet.cell_value(i, j)

        length = -1
        try:
            length = len(val)
        except TypeError:
            cnt += 1
            humidity += val

    if cnt > 0:
        humidity /= cnt

    humidity = round(humidity, 3)

    last_row_written += 1
    output_sheet.write(last_row_written, 0, district)
    output_sheet.write(last_row_written, 1, year)
    output_sheet.write(last_row_written, 2, month)
    output_sheet.write(last_row_written, 3, str(humidity))

output.save("humidity_avg.xls")