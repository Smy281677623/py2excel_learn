#!/usr/bin/env python
# -*- coding:utf-8 -*-
# Author:Smy
import xlwt

def main():
    workbook = xlwt.Workbook()
    worksheet = workbook.add_sheet('My Sheet')
    alignment = xlwt.Alignment()  # Create Alignment
    alignment.horz = xlwt.Alignment.HORZ_CENTER  # May be: HORZ_GENERAL, HORZ_LEFT, HORZ_CENTER, HORZ_RIGHT, HORZ_FILLED, HORZ_JUSTIFIED, HORZ_CENTER_ACROSS_SEL, HORZ_DISTRIBUTED
    alignment.vert = xlwt.Alignment.VERT_CENTER  # May be: VERT_TOP, VERT_CENTER, VERT_BOTTOM, VERT_JUSTIFIED, VERT_DISTRIBUTED
    style = xlwt.XFStyle()  # Create Style
    style.alignment = alignment  # Add Alignment to Style
    worksheet.write(0, 0, 'Cell Contents', style)
    workbook.save('Excel_Workbook.xls')


if __name__ == '__main__':
    main()
