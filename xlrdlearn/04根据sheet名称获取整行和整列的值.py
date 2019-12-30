# -*- coding:utf-8 -*-
import xlrd


def main():
    data = xlrd.open_workbook('./联系人.xls')
    # 根据sheet名称获取sheet对象
    sheet_yh2 = data.sheet_by_name('银行2')
    # 根据sheet_yh2对象的方法去获取具体行和具体列
    print(sheet_yh2.row_values(3))
    print(sheet_yh2.col_values(3))


if __name__ == '__main__':
    main()