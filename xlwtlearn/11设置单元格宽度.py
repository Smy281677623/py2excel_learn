#!/usr/bin/env python
# -*- coding:utf-8 -*-
# Author:Smy
import xlwt


def main():
    workbook = xlwt.Workbook()
    worksheet = workbook.add_sheet('My Sheet')
    worksheet.write(0, 0, 'My Cell Contents')

    # 设置单元格宽度
    worksheet.col(0).width = 3333
    workbook.save('cell_width.xls')


if __name__ == '__main__':
    main()
