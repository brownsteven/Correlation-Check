from openpyxl import load_workbook
import csv


def similarCheck():
    wb = load_workbook(filename='test input file.xlsx')
    ws = wb['Sheet1']

    with open('Test output.csv', 'w', newline='') as outcsv:
        writer = csv.writer(outcsv)
        writer.writerow(['Percentage Match', 'Gene (A)', 'Assay ID (A)',
                         'Gene (B)', 'Assay ID (B)'])
        for row in ws.rows:
            col = 0
            for cell in row:
                col += 1
                if cell.value == 1:
                    break
                elif type(cell.value) == float:
                    # if cell.value > .9 and cell.value < 1:
                    if 0.9 <= cell.value < 1:
                        writer.writerow([cell.value, row[0].value, row[1].value,
                                         ws.cell(row=1, column=col).value, ws.cell(row=2, column=col).value])

                        print('Percentage:', cell.value, 'Gene:',
                              row[0].value, ws.cell(row=1, column=col).value,
                              'AssayID:', row[1].value,
                              ws.cell(row=2, column=col).value)


similarCheck()
