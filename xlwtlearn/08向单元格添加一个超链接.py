#!/usr/bin/env python
# -*- coding:utf-8 -*-
# Author:Smy
import xlwt

def main():
    workbook = xlwt.Workbook()
    worksheet = workbook.add_sheet('My Sheet')
    worksheet.write(0, 0, xlwt.Formula(
        'HYPERLINK("http://www.google.com";"Google")'))  # Outputs the text "Google" linking to http://www.google.com
    workbook.save('Excel_Workbook.xls')


if __name__ == '__main__':
    main()
