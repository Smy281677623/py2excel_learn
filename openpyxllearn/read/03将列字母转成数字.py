# -*- coding:utf-8 -*-
from openpyxl.utils import get_column_letter, column_index_from_string


def main():
    # 对列进行字母/数字转化
    # 将字母变成数字
    c_num = column_index_from_string('B')  # c_num = 2
    # 将数字变成字母
    c_char = get_column_letter(5)  # c_char = 'E‘
    print(c_num)
    print(c_char)


if __name__ == '__main__':
    main()