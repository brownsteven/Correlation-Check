from openpyxl import load_workbook


def similarCheck():
    wb = load_workbook(filename='test input file.xlsx')
    ws = wb['Sheet1']

    for row in ws.rows:
        c = 0
        # Look up how to set end points or max cell here. Maybe only do the bottom half
        for cell in row:
            c += 1
            if type(cell.value) == float:
                if cell.value > .9 and cell.value < 1:
                    # Make a .csv file with the following:
                    """with open('Test output.csv', 'w') as outcsv:
                        writer = csv.writer(outcsv)
                        writer.writerow(['Percentage', 'Gene', 'AssayID'])"""

                    print('Percentage:', cell.value, 'Gene:',
                          row[0].value, ws.cell(row=1, column=c).value,
                          'AssayID:', row[1].value,
                          ws.cell(row=2, column=c).value)


similarCheck()
