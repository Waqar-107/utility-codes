# from dust i have come, dust i will be

import xlrd
import xlwt

location = "output1.xls"

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
for each month, for each decade, avg of 30y
'''

data = [[[0 for k in range(3)] for j in range(12)] for i in range(3)]

output_sheet.write(0, 0, "District")
output_sheet.write(0, 1, "Month")
output_sheet.write(0, 2, "Decade 1")
output_sheet.write(0, 3, "Decade 2")
output_sheet.write(0, 4, "Decade 3")

req_district = ['Comilla', 'Feni', 'Chandpur']

row_last_written = 0

for i in range(1, row, 1):
    district = sheet.cell_value(i, 0)
    month = int(sheet.cell_value(i, 2))

    d = -1
    for j in range(3):
        if district == req_district[j]:
            d = j
            break

    for j in range(3):
        val = float(sheet.cell_value(i, j + 3))

        data[d][month - 1][j] += val

for i in range(3):
    for j in range(12):
        row_last_written += 1
        output_sheet.write(row_last_written, 0, req_district[i])
        output_sheet.write(row_last_written, 1, j + 1)

        for k in range(3):
            val = data[i][j][k] / 30
            output_sheet.write(row_last_written, k + 2, val)

output.save("output2.xls")