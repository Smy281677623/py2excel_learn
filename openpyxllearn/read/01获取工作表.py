# -*- coding:utf-8 -*-
import openpyxl


def main():
    """获取工作表"""
    # 读取一个xlsx格式文件（格式有严格要求）
    # 通过调用方法load_workbook(filename)进行文件读取，该方法中还有一个read_only参数用于设置文件打开方式，默认为可读可写，该方法最终将返回一个workbook的数据对象
    wb = openpyxl.load_workbook('工作簿.xlsx')
    # 获取所有工作表名（返回值形式为列表）
    sheets = wb.get_sheet_names()
    print(sheets)
    # 根据工作表名获得某一具体的工作表
    sheet_one = wb.get_sheet_by_name('Sheet1')
    print(sheet_one)
    # 获取工作表的表名,可以在任何时候通过Worksheet.title属性修改工作表名
    print(sheet_one.title)
    # 一般来说，表格大多数用到的是打开时显示的工作表，这时可以用active来获取当前工作表
    # 工作表在工作簿创建后,可以通过Workbook.active属性来定位到工作表
    sheet_one = wb.active


if __name__ == '__main__':
    main()