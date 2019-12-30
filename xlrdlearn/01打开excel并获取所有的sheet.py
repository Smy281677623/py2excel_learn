import xlrd


def main():
    # 打开excel文件读取数据
    data = xlrd.open_workbook('./联系人.xls')
    # 获取sheet列表
    sheet_list = data.sheet_names()
    # 打印sheet名称
    print(sheet_list)


if __name__ == '__main__':
    main()