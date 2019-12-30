#!/usr/bin/env python
# -*- coding:utf-8 -*-
# Author:Smy
"""https://blog.csdn.net/xuezhangjun0121/article/details/91365875"""
import xlwt


def set_style(name, height, bold=False, format_str='', align='center'):
    style = xlwt.XFStyle()  # 初始化样式

    font = xlwt.Font()  # 为样式创建字体
    font.name = name  # 'Times New Roman'
    font.bold = bold
    font.height = height

    borders = xlwt.Borders()  # 为样式创建边框
    borders.left = 2
    borders.right = 2
    borders.top = 0
    borders.bottom = 2

    alignment = xlwt.Alignment()  # 设置排列
    if align == 'center':
        alignment.horz = xlwt.Alignment.HORZ_CENTER
        alignment.vert = xlwt.Alignment.VERT_CENTER
    else:
        alignment.horz = xlwt.Alignment.HORZ_LEFT
        alignment.vert = xlwt.Alignment.VERT_BOTTOM

    style.font = font
    style.borders = borders
    style.num_format_str = format_str
    style.alignment = alignment

    return style



def main():
    wb = xlwt.Workbook()
    sheet_name = str(input('请输入一个机器编码'))
    board_code = str(input('请输入板卡条码'))
    ws = wb.add_sheet(sheet_name, cell_overwrite_ok=True)
    # 设置各列宽为200*30
    for i in range(0, 6):
        ws.col(i).width = 200*30
    # 设置第0行0列的内容(实际为1行1列)
    borders = xlwt.Borders()  # Create Borders
    borders.left = xlwt.Borders.THIN
    # DASHED虚线
    # NO_LINE没有
    # THIN实线
    # May be: NO_LINE, THIN, MEDIUM, DASHED, DOTTED, THICK, DOUBLE, HAIR, MEDIUM_DASHED, THIN_DASH_DOTTED, MEDIUM_DASH_DOTTED, THIN_DASH_DOT_DOTTED, MEDIUM_DASH_DOT_DOTTED, SLANTED_MEDIUM_DASH_DOTTED, or 0x00 through 0x0D.
    borders.left = 2
    borders.right = 2
    borders.top = 0
    borders.bottom = 2
    borders.left_colour = 0x40
    borders.right_colour = 0x40
    borders.top_colour = 0x40
    borders.bottom_colour = 0x40
    style = xlwt.XFStyle()  # Create Style
    style.borders = borders  # Add Borders to Style
    ws.write(0, 0, '小电流模式校准', style)
    # 将第0行1列与2列合并并输入 机器编码：+（我们手动输入的机器编码）,字体居中
    ws.write_merge(0, 0, 1, 2,'机器编码:'+sheet_name,set_style('等线', 200, bold=True,format_str=''))
    # 将第0行的 3 4 5列合并 输入 板卡条码:+（我们手动输入的板卡条码）,字体居中
    ws.write_merge(0, 0, 3, 5, '板卡条码:'+board_code, set_style('等线', 200, bold=True, format_str=''))
    # 对1行0列写入 设置电压， 边框不变
    ws.write(1, 0, '设置电压', style)
    # 小电流模式采样次数
    ws.write(1, 1, '小电流模式采样次数', style)
    # 8001电压读数（mV）
    ws.write(1, 2, '8001电压读数（mV）', style)
    # 万用表电压读数（mV）
    ws.write(1, 3, '万用表电压读数（mV）', style)
    # 8001电流读数(mA)
    ws.write(1, 4, '8001电流读数(mA)', style)
    # 万用表电流读数(mA)
    ws.write(1, 5, '万用表电流读数(mA)', style)
    # 合并 2行的 0 1 2 3 4 5列合并并居中
    ws.write_merge(2, 0, 0, 5, '第一轮老化测试', set_style('等线', 200, bold=True, format_str=''))

    wb.save('my_first_excel.xls')


if __name__ == '__main__':
    main()
