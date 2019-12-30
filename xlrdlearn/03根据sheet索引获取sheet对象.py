# -*- coding:utf-8 -*-
import xlrd


def main():
    """根据sheet索引或者名称获取sheet内容，同时获取sheet名称、行数、列数"""
    # 打开一个xls文件
    data = xlrd.open_workbook('./联系人.xls')
    # 根据sheet索引获取sheet对象
    sheet_one = data.sheet_by_index(0)
    # 根据sheet对象的属性去访问这个sheet中的数据
    print('sheet_one名称:{}\nsheet_one列数: {}\nsheet_one行数: {}'.format(sheet_one.name, sheet_one.ncols, sheet_one.nrows))


if __name__ == '__main__':
    main()