# This is a sample Python script.

# Press Shift+F10 to execute it or replace it with your code.
# Press Double Shift to search everywhere for classes, files, tool windows, actions, and settings.
import xlrd
import xlwt


def print_hi(name):
    # Use a breakpoint in the code line below to debug your script.
    print(f'Hi, {name}')  # Press Ctrl+F8 to toggle the breakpoint.


def readxl():
    file = r"C:\Users\Administrator\Desktop\hero.xls"
    data = xlrd.open_workbook(file)
    table = data.sheet_by_index(0)
    rows = table.nrows
    cols = table.ncols
    for row in range(rows):
        for col in range(cols):
            if col == cols - 1:
                if col != 0:
                    print(end=' ')
                print(table.cell_value(row, col))
            else:
                print(table.cell_value(row, col), end='')

    # print(value)
    # print(value1)


def wrExcel():
    file = r"C:\Users\Administrator\Desktop\result.xls"
    data = xlwt.Workbook(encoding='utf-8')
    sheet = data.add_sheet("sheet1")
    sheet.write(0, 0, "哈哈")
    data.save(file)


# Press the green button in the gutter to run the script.
if __name__ == '__main__':
    print_hi('PyCharm')
    readxl()
    wrExcel()

# See PyCharm help at https://www.jetbrains.com/help/pycharm/
