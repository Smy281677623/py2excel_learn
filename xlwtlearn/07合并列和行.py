#!/usr/bin/env python
# -*- coding:utf-8 -*-
# Author:Smy
import xlwt


def main():
    workbook = xlwt.Workbook()
    worksheet = workbook.add_sheet('My Sheet')
    worksheet.write_merge(0, 0, 0, 3, 'First Merge')  # Merges row 0's columns 0 through 3.
    font = xlwt.Font()  # Create Font
    font.bold = True  # Set font to Bold
    style = xlwt.XFStyle()  # Create Style
    style.font = font  # Add Bold Font to Style
    worksheet.write_merge(1, 2, 0, 3, 'Second Merge', style)  # Merges row 1 through 2's columns 0 through 3.
    workbook.save('Excel_Workbook.xls')


if __name__ == '__main__':
    main()
