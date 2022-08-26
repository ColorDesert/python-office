import xlwt
import xlrd

from pathlib import Path, PurePath

# 需要合并的Excel的文件夹
src_path = f'D:\PythonCode\part1\调查问卷'

dst_path = f'D:\PythonCode\part1\dresult\调查结果.xlsx'

path = Path(src_path)

files = [x for x in path.iterdir() if PurePath(x).match('*.xls')]
# 存放合并的内容
contents = []


def merge_excel_read():
    # 读取数据
    for file in files:
        fileName = file.stem
        book = xlrd.open_workbook(file)
        table = book.sheet_by_index(0)

        # 取得每一题的答案
        answer1 = table.cell_value(4, 4)
        answer2 = table.cell_value(10, 4)
        temp = f'{fileName},{answer1},{answer2}'
        print(temp)
        contents.append(temp.split(','))


def merge_excel_write():
    col = 0
    row = 0
    workbook = xlwt.Workbook(encoding='utf-8')
    sheet = workbook.add_sheet("统计结果")
    titles = ['员工姓名', '第一题', '第二题']
    for title in titles:
        sheet.write(row, col, title)
        col += 1

    for line in contents:
        row += 1
        col = 0
        for cell in line:
            sheet.write(row, col, cell)
            col += 1
    f = open(dst_path, 'w+b')
    workbook.save(f)


merge_excel_read()
merge_excel_write()
