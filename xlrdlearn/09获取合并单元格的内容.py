#!/usr/bin/env python
# -*- coding:utf-8 -*-
# Author:Smy
import xlrd


def main():
    data = xlrd.open_workbook('./联系人.xls')
    sheet_yh2 = data.sheet_by_name('银行2')
    # 获取合并单元格的信息要使用merged_cells属性
    print(sheet_yh2.merged_cells) # [(0, 1, 0, 8), (2, 6, 0, 1)]<br>
    # merged_cells返回的这四个参数的含义是：(row,row_range,col,col_range),其中[row,row_range)包括row,
    # 不包括row_range,col也是一样，下标从0开始。
    #(0, 1, 0, 8) 表示1列-8列合并 (2, 6, 0, 1)表示3行-6行合并<br>
    # 分别获取合并2个单元格的内容：
    print(sheet_yh2.cell(0,0).value)  # 银行2(第一行, 第一列)
    print(sheet_yh2.cell_value(2, 0))  # 银行2(第3行, 第1列)


if __name__ == '__main__':
    main()
