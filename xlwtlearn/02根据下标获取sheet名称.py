# -*- coding:utf-8 -*-
import xlrd


def main():
    data = xlrd.open_workbook('./联系人.xls')
    sheet_list = data.sheet_names()
    # print(sheet_list[0])
    # print(sheet_list[1])
    for i in sheet_list:
        print(i)


if __name__ == '__main__':
    main()