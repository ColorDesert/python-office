import xlwt
import xlrd

from pathlib import Path, PurePath

# 需要合并的Excel的文件夹
src_path = r'D:\PythonCode\part1\工资单\工资单.xlsx'

dst_path = r'D:\PythonCode\part1\工资单\调查结果.xlsx'

def split_excel_write(filename, data):
    workbook = xlwt.Workbook(encoding='utf-8')
    sheet = workbook.add_sheet("本月工资")

    row = 0
    for line in data:

        col = 0
        for cell in line:
            sheet.write(row, col, cell)
            col += 1
            workbook.save(PurePath(dst_path).with_name(filename).with_suffix('.xlsx'))
        row += 1

def split_excel_read():
    book = xlrd.open_workbook(src_path)
    table = book.sheet_by_index(0)
    rows = table.nrows
    # 获取工资单的表头
    salary_header = table.row_values(rowx=0, start_colx=0, end_colx=None)
    for row in range(1, rows):
        contents = [salary_header]

        content = table.row_values(rowx=row, start_colx=0, end_colx=None)
        contents.append(content)
        print(contents)
        split_excel_write(content[1], contents)


split_excel_read()

