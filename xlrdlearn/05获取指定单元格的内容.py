# -*- coding:utf-8 -*-
import xlrd


def main():
    data = xlrd.open_workbook('./联系人.xls')
    # 根据sheet 的名称获取sheet对象
    sheet1 = data.sheet_by_name('银行2')
    # 获取第2行第1列的信息(默认从0，0开始算的)
    print(sheet1.cell(1, 0))
    print(sheet1.col_values(1, 0))
    print(sheet1.row(1)[0].value)


if __name__ == '__main__':
    main()