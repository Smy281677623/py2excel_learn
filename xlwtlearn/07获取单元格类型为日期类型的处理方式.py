# -*- coding:utf-8 -*-
from datetime import datetime,date
import xlrd


def main():
    data = xlrd.open_workbook('./联系人.xls')
    sheet_yh2 = data.sheet_by_name('银行2')
    if sheet_yh2.cell(3, 6).ctype == 3:
        print(sheet_yh2.cell(3, 6).value)  # 41463.0
        date_value = xlrd.xldate_as_tuple(sheet_yh2.cell(3, 6).value, data.datemode)
        print(date_value)  # (2013, 7, 8, 0, 0, 0)
        print(date(*date_value[:3]))  # 2013-07-08
        print(date(*date_value[:3]).strftime('%Y/%m/%d'))  # 2013/07/08


if __name__ == '__main__':
    main()