# -*- coding:utf-8 -*-
import xlrd


def main():
    data = xlrd.open_workbook('./联系人.xls')
    sheet_yh2 = data.sheet_by_name('银行2')
    if sheet_yh2.cell(3, 5).ctype == 2:
        print(sheet_yh2.cell(3, 5).value)  # 133111.0
        num_value = int(sheet_yh2.cell(3, 5).value)
        print(num_value)  # 133111


if __name__ == '__main__':
    main()