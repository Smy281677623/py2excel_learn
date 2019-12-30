#!/usr/bin/env python
# -*- coding:utf-8 -*-
# Author:Smy
import xlwt
import datetime


def main():
    workbook = xlwt.Workbook()
    worksheet = workbook.add_sheet('My Sheet')
    style = xlwt.XFStyle()
    style.num_format_str = 'M/D/YY'  # Other options: D-MMM-YY, D-MMM, MMM-YY, h:mm, h:mm:ss, h:mm, h:mm:ss, M/D/YY h:mm, mm:ss, [h]:mm:ss, mm:ss.0
    worksheet.write(0, 0, datetime.datetime.now(), style)
    workbook.save('Excel_Workbook.xls')


if __name__ == '__main__':
    main()
