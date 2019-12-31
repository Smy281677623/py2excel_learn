# -*- coding:utf-8 -*-
import openpyxl
from openpyxl.styles import Font, Border, Side, PatternFill, colors, Alignment
from datetime import *
# 4个数据集合 传进封装类 实体


class Excel_Template(object):
    def __init__(self, board_voltage, mu_voltage, board_current, mu_current, board_num):
        assert board_voltage and mu_voltage and board_current and mu_current == 0, \
            """board_voltage or mu_voltage or board_current or mu_current can not be None"""
        self.board_voltage_ = board_voltage
        self.mu_voltage_ = mu_voltage
        self.board_current_ = board_current
        self.mu_voltage_ = mu_current
        self.board_num_ = board_num

    def setup_excel(self):
        """创建excel模版"""
        wb = openpyxl.Workbook()
        # 如果不给sheet名称则直接创建默认的sheet1名称
        if self.board_num_ is None:
            ws = wb.create_sheet('sheet1', 0)
        else:
            ws = wb.create_sheet(str(self.board_num_))
        # 设置 1 2 3 4 5 6 列的宽度
        for col in ['A', 'B', 'C', 'D', 'E', 'F']:
            ws.column_dimensions[col].width = 20.0
        # 设置小电流模式
        ws['A1'].value = '小电流模式'
        ws['A1'].font = Font(name='宋体', size=11, bold=True, italic=True, color=colors.BLACK)
        ws['B1'].value = '机器编码'
        ws['B1'].font = Font(name='宋体', size=11, bold=True, italic=True, color=colors.BLACK)
        ws['C1'].font = Font(name='宋体', size=11, bold=True, italic=True, color=colors.BLACK)
        ws.merge_cells('B1:C1')
        ws['D1'].value = '板卡条码'
        ws['D1'].font = Font(name='宋体', size=11, bold=True, italic=True, color=colors.BLACK)
        ws['E1'].font = Font(name='宋体', size=11, bold=True, italic=True, color=colors.BLACK)
        ws['F1'].font = Font(name='宋体', size=11, bold=True, italic=True, color=colors.BLACK)
        # datetime.time()
        # *********************************************************


def main():
    pass


if __name__ == '__main__':
    main()