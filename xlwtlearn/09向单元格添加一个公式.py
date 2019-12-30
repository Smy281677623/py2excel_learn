#!/usr/bin/env python
# -*- coding:utf-8 -*-
# Author:Smy
import xlwt


def main():
    workbook = xlwt.Workbook()
    worksheet = workbook.add_sheet('My Sheet')
    worksheet.write(0, 0, 5)  # Outputs 5
    worksheet.write(0, 1, 2)  # Outputs 2
    worksheet.write(1, 0, xlwt.Formula('A1*B1'))  # Should output "10" (A1[5] * A2[2])
    worksheet.write(1, 1, xlwt.Formula('SUM(A1,B1)'))  # Should output "7" (A1[5] + A2[2])
    workbook.save('Excel_Workbook.xls')

if __name__ == '__main__':
    main()
