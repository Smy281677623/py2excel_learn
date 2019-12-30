#!/usr/bin/env python
# -*- coding:utf-8 -*-
# Author:Smy
"""https://blog.csdn.net/xuezhangjun0121/article/details/91365875"""
import xlwt


def set_style(name, height, bold=False, format_str='', align='center', color_num=1):
    style = xlwt.XFStyle()  # 初始化样式

    font = xlwt.Font()  # 为样式创建字体
    font.name = name  # 'Times New Roman'
    font.bold = bold
    font.height = height

    borders = xlwt.Borders()  # 为样式创建边框
    borders.left = 2
    borders.right = 2
    borders.top = 2
    borders.bottom = 2

    # 设置颜色
    pattern = xlwt.Pattern()  # Create the Pattern
    pattern.pattern = xlwt.Pattern.SOLID_PATTERN  # May be: NO_PATTERN, SOLID_PATTERN, or 0x00 through 0x12
    # May be: 8 through 63. 0 = Black, 1 = White, 2 = Red, 3 = Green, 4 = Blue, 5 = Yellow, 6 = Magenta, 7 = Cyan, 16 = Maroon, 17 = Dark Green, 18 = Dark Blue, 19 = Dark Yellow , almost brown), 20 = Dark Magenta, 21 = Teal, 22 = Light Gray, 23 = Dark Gray, the list goes on...
    pattern.pattern_fore_colour = color_num

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
    style.pattern = pattern  # Add Pattern to Style

    return style


def set_one_to_ten(x, y, z, ws, temp_color):
    for p, q in zip(x, z * 5):
        ws.write(p, 1, str(q), set_style('等线', 200, bold=True, format_str='', color_num=temp_color))
        if q % 10 == 0:
            temp_color += 1
    for j in x:
        for k in y:
            ws.write(j, k, '1')

def main():
    wb = xlwt.Workbook()
    sheet_name = str(input('请输入一个机器编码'))
    board_code = str(input('请输入板卡条码'))
    ws = wb.add_sheet(sheet_name, cell_overwrite_ok=True)
    # 设置各列宽为200*30
    for i in range(0, 6):
        ws.col(i).width = 200*30
    # 设置第0行0列的内容(实际为1行1列)
    ws.write(0, 0, '小电流模式校准', set_style('等线', 200, bold=True, format_str='', align='center'))
    # 将第0行1列与2列合并并输入 机器编码：+（我们手动输入的机器编码）,字体居中
    ws.write_merge(0, 0, 1, 2,'机器编码:'+sheet_name,set_style('等线', 200, bold=True,format_str=''))
    # 将第0行的 3 4 5列合并 输入 板卡条码:+（我们手动输入的板卡条码）,字体居中
    ws.write_merge(0, 0, 3, 5, '板卡条码:'+board_code, set_style('等线', 200, bold=True, format_str=''))
    # 对1行0列写入 设置电压， 边框不变
    ws.write(1, 0, '设置电压', set_style('等线', 200, bold=True, format_str='', align='center'))
    # 小电流模式采样次数
    ws.write(1, 1, '小电流模式采样次数', set_style('等线', 200, bold=True, format_str='', align='center'))
    # 8001电压读数（mV）
    ws.write(1, 2, '8001电压读数（mV）', set_style('等线', 200, bold=True, format_str='', align='center'))
    # 万用表电压读数（mV）
    ws.write(1, 3, '万用表电压读数（mV）', set_style('等线', 200, bold=True, format_str='', align='center'))
    # 8001电流读数(mA)
    ws.write(1, 4, '8001电流读数(mA)', set_style('等线', 200, bold=True, format_str='', align='center'))
    # 万用表电流读数(mA)
    ws.write(1, 5, '万用表电流读数(mA)', set_style('等线', 200, bold=True, format_str='', align='center'))
    # 合并 2行的 0 1 2 3 4 5列合并并居中
    ws.write_merge(2, 2, 0, 5, '第一轮老化测试', set_style('等线', 200, bold=True, format_str=''))
    # 将 3 到 12 的 0列 合并成 输入3.6 并设置成橙色
    ws.write_merge(3, 12, 0, 0, '3.6V', set_style('等线', 200, bold=True, format_str='', color_num=46))
    ws.write_merge(13, 22, 0, 0, '3.8V', set_style('等线', 200, bold=True, format_str='', color_num=47))
    ws.write_merge(23, 32, 0, 0, '4.0V', set_style('等线', 200, bold=True, format_str='', color_num=48))
    ws.write_merge(33, 42, 0, 0, '4.1V', set_style('等线', 200, bold=True, format_str='', color_num=49))
    ws.write_merge(43, 52, 0, 0, '4.2V', set_style('等线', 200, bold=True, format_str='', color_num=50))

    ws.write_merge(53, 53, 0, 5, '第二轮老化测试', set_style('等线', 200, bold=True, format_str=''))

    ws.write_merge(54, 63, 0, 0, '3.6V', set_style('等线', 200, bold=True, format_str='', color_num=46))
    ws.write_merge(64, 73, 0, 0, '3.8V', set_style('等线', 200, bold=True, format_str='',  color_num=47))
    ws.write_merge(74, 83, 0, 0, '4.0V', set_style('等线', 200, bold=True, format_str='',  color_num=48))
    ws.write_merge(84, 93, 0, 0, '4.1V', set_style('等线', 200, bold=True, format_str='',  color_num=49))
    ws.write_merge(94, 103, 0, 0, '4.2V', set_style('等线', 200, bold=True, format_str='', color_num=50))

    ws.write_merge(104, 104, 0, 5, '第三轮老化测试', set_style('等线', 200, bold=True, format_str=''))

    ws.write_merge(105, 114, 0, 0, '3.6V', set_style('等线', 200, bold=True, format_str='', color_num=46))
    ws.write_merge(115, 124, 0, 0, '3.8V', set_style('等线', 200, bold=True, format_str='',  color_num=47))
    ws.write_merge(125, 134, 0, 0, '4.0V', set_style('等线', 200, bold=True, format_str='',  color_num=48))
    ws.write_merge(135, 144, 0, 0, '4.1V', set_style('等线', 200, bold=True, format_str='',  color_num=49))
    ws.write_merge(145, 154, 0, 0, '4.2V', set_style('等线', 200, bold=True, format_str='',  color_num=50))
    # 从第3行第1列   到   第 52 行 第 5列写入数据并加入边框
    x1 = [x for x in range(3, 53)]
    x2 = [x for x in range(54, 104)]
    x3 = [x for x in range(105, 155)]
    y = [y for y in range(2, 6)]
    z = [z for z in range(1, 11)]
    # 设置采样次数的颜色和数字
    temp_color = 46
    for x in x1,x2,x3:
        set_one_to_ten(x, y, z, ws, temp_color)

    wb.save('my_first_excel.xls')


if __name__ == '__main__':
    main()




