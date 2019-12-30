# -*- coding:utf-8 -*-
import xlrd


def main():
    data = xlrd.open_workbook('./联系人.xls')
    # 根据sheet 的名称获取sheet对象
    sheet1 = data.sheet_by_name('银行2')
    # 获取第2行第1列的信息(默认从0，0开始算的)
    # 获取单元格内容的数据类型
    print(sheet1.cell(1, 0).ctype)  # 第2 行1列内容 ：机构名称为string类型
    print(sheet1.cell(3, 4).ctype)  # 第4行5列内容：999 为number类型
    print(sheet1.cell(3, 6).ctype)  # 第4 行7列内容：2013/7/8 为date类型
    # 说明：ctype : 0 empty,1 string, 2 number, 3 date, 4 boolean, 5 error


if __name__ == '__main__':
    main()