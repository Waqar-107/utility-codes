# from dust i have come, dust i will be

import xlrd
import xlwt

startFrom = 9
skipZero = False

location = "ActualData\SW132.xlsx"

wb = xlrd.open_workbook(location)
sheet = wb.sheet_by_index(0)

# extracting number of rows
row = sheet.nrows

# extracting number of columns
col = sheet.ncols

output = xlwt.Workbook()
output_sheet = output.add_sheet("newSheet")

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

data = [[[0 for k in range(3)] for j in range(12)] for i in range(3000)]
cnt = [[[0 for k in range(3)] for j in range(12)] for i in range(3000)]

for i in range(1, 2, 1):
    y, m, d = sheet.cell_value(i, 3).split('-')
    val = sheet.cell_value(i, 4)

    y = int(y)
    m = int(months[m])
    d = int(d)

    try:
        val = float(val)
    except ValueError:
        val = -1

    if val == -1:
        continue

    if 10 <= d <= 1:
        data[y][m][0] += val
        cnt[y][m][0] += 1

    elif 11 <= d <= 20:
        data[y][m][1] += val
        cnt[y][m][1] += 1

    else:
        data[y][m][2] += val
        cnt[y][m][2] += 1

output_sheet.write(0, 0, "District")
output_sheet.write(0, 1, "Year")
output_sheet.write(0, 2, "Month")
output_sheet.write(0, 3, "Decade 1")
output_sheet.write(0, 4, "Decade 2")
output_sheet.write(0, 5, "Decade 3")

last_row_written = 0
for i in range(1981, 2011, 1):
    for j in range(12):
        last_row_written += 1

        output_sheet.write(last_row_written, 0, "Brahmanbaria")
        output_sheet.write(last_row_written, 1, i)
        output_sheet.write(last_row_written, 2, j + 1)

        for k in range(3):
            if cnt[i][j][k]:
                val = data[i][j][k] / cnt[i][j][k]
            else:
                val = 0

            output_sheet.write(last_row_written, k + 3, val)

output.save("output1.xls")