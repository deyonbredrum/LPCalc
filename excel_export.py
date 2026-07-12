# -*- coding: utf-8 -*-
"""
Excel 计算书导出模块 - 使用 openpyxl（稳定版）
"""

import io
from datetime import datetime
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter


def export_excel_report(lp_result, ep_result, building_params,
                        cable_params, c_factors, level_text, ep_level_text,
                        export_mode='combined', has_ep=True,
                        building_attr="", attr_type=""):
    """
    导出 Excel 计算书
    """
    wb = Workbook()

    # 删除默认的空白sheet
    wb.remove(wb.active)

    # ===== Sheet 1: 摘要 =====
    ws1 = wb.create_sheet("摘要", 0)

    # 标题
    ws1.merge_cells('A1:F1')
    cell = ws1['A1']
    cell.value = '建筑物防雷及电子信息防护等级计算书'
    cell.font = Font(name='黑体', size=18, bold=True, color='003366')
    cell.alignment = Alignment(horizontal='center', vertical='center')

    ws1.merge_cells('A2:F2')
    cell = ws1['A2']
    cell.value = f'计算日期：{datetime.now().strftime("%Y年%m月%d日")}'
    cell.font = Font(name='宋体', size=10)
    cell.alignment = Alignment(horizontal='center', vertical='center')

    # 空行
    ws1.row_dimensions[3].height = 10

    # 建筑物参数
    row = 4
    ws1.merge_cells(f'A{row}:F{row}')
    cell = ws1[f'A{row}']
    cell.value = '一、建筑物参数'
    cell.font = Font(name='黑体', size=12, bold=True, color='003366')
    cell.alignment = Alignment(horizontal='left', vertical='center')
    row += 1

    for key, value in building_params.items():
        ws1[f'A{row}'] = key
        ws1[f'B{row}'] = str(value)
        ws1[f'A{row}'].font = Font(name='宋体', size=10)
        ws1[f'B{row}'].font = Font(name='宋体', size=10)
        row += 1

    # 防雷等级结果
    row += 1
    ws1.merge_cells(f'A{row}:F{row}')
    cell = ws1[f'A{row}']
    cell.value = '二、防雷等级计算结果'
    cell.font = Font(name='黑体', size=12, bold=True, color='003366')
    cell.alignment = Alignment(horizontal='left', vertical='center')
    row += 1

    ws1[f'A{row}'] = '年预计雷击次数 N'
    ws1[f'B{row}'] = f"{lp_result.get('N', 0):.6f} 次/a"
    ws1[f'A{row}'].font = Font(name='宋体', size=10)
    ws1[f'B{row}'].font = Font(name='宋体', size=10)
    row += 1

    ws1[f'A{row}'] = '防雷等级'
    ws1[f'B{row}'] = level_text
    ws1[f'A{row}'].font = Font(name='宋体', size=10)
    ws1[f'B{row}'].font = Font(name='宋体', size=12, bold=True)
    ws1[f'B{row}'].alignment = Alignment(horizontal='left', vertical='center')
    row += 1

    # 电子防护等级结果
    if has_ep and ep_result:
        row += 1
        ws1.merge_cells(f'A{row}:F{row}')
        cell = ws1[f'A{row}']
        cell.value = '三、电子信息系统防护等级计算结果'
        cell.font = Font(name='黑体', size=12, bold=True, color='003366')
        cell.alignment = Alignment(horizontal='left', vertical='center')
        row += 1

        ws1[f'A{row}'] = '总年预计雷击次数 N'
        ws1[f'B{row}'] = f"{ep_result.get('N', 0):.6f} 次/a"
        ws1[f'A{row}'].font = Font(name='宋体', size=10)
        ws1[f'B{row}'].font = Font(name='宋体', size=10)
        row += 1

        ws1[f'A{row}'] = '拦截效率 E'
        ws1[f'B{row}'] = f"{ep_result.get('E', 0):.4f}"
        ws1[f'A{row}'].font = Font(name='宋体', size=10)
        ws1[f'B{row}'].font = Font(name='宋体', size=10)
        row += 1

        ws1[f'A{row}'] = '防护等级'
        ws1[f'B{row}'] = ep_level_text
        ws1[f'A{row}'].font = Font(name='宋体', size=10)
        ws1[f'B{row}'].font = Font(name='宋体', size=12, bold=True)
        color_map = {'A级': 'CC0000', 'B级': 'FF8800', 'C级': '0066CC', 'D级': '008800'}
        if ep_level_text in color_map:
            ws1[f'B{row}'].font = Font(name='宋体', size=12, bold=True, color=color_map[ep_level_text])
        ws1[f'B{row}'].alignment = Alignment(horizontal='left', vertical='center')

    # 设置列宽
    for col in ['A', 'B', 'C', 'D', 'E', 'F']:
        ws1.column_dimensions[col].width = 20

    # ===== Sheet 2: 详细计算过程 =====
    ws2 = wb.create_sheet("详细计算过程")

    row = 1
    ws2.merge_cells(f'A{row}:D{row}')
    cell = ws2[f'A{row}']
    cell.value = '防雷等级计算过程'
    cell.font = Font(name='黑体', size=14, bold=True, color='003366')
    cell.alignment = Alignment(horizontal='center', vertical='center')
    row += 1

    # 表头
    headers = ['计算项', '符号', '公式', '结果']
    for col, header in enumerate(headers, 1):
        cell = ws2.cell(row=row, column=col, value=header)
        cell.font = Font(name='黑体', size=10, bold=True)
        cell.fill = PatternFill(start_color='D9E1F2', end_color='D9E1F2', fill_type='solid')
        cell.alignment = Alignment(horizontal='center', vertical='center')
    row += 1

    # 数据行
    lp_data = [
        ('年雷暴日', 'Td', '查表', f"{lp_result.get('Td', 0):.2f} d/a"),
        ('雷击大地密度', 'Ng', '0.1 × Td', f"{lp_result.get('Ng', 0):.4f} 次/(km²·a)"),
        ('扩大宽度', 'D', '√(H(200-H))', f"{lp_result.get('D', 0):.3f} m"),
        ('等效面积', 'Ae', '按周边情况计算', f"{lp_result.get('Ae', 0):.6f} km²"),
        ('校正系数', 'k', '查表', f"{lp_result.get('k', 0):.2f}"),
        ('年预计雷击次数', 'N', 'k × Ng × Ae', f"{lp_result.get('N', 0):.6f} 次/a"),
        ('防雷等级', '-', '根据N值判定', level_text),
    ]

    for data_row in lp_data:
        for col, value in enumerate(data_row, 1):
            cell = ws2.cell(row=row, column=col, value=value)
            cell.font = Font(name='宋体', size=10)
            cell.alignment = Alignment(horizontal='center', vertical='center')
        row += 1

    # 电子防护计算
    if has_ep and ep_result:
        row += 1
        ws2.merge_cells(f'A{row}:D{row}')
        cell = ws2[f'A{row}']
        cell.value = '电子信息系统防护等级计算过程'
        cell.font = Font(name='黑体', size=14, bold=True, color='003366')
        cell.alignment = Alignment(horizontal='center', vertical='center')
        row += 1

        # 表头
        for col, header in enumerate(headers, 1):
            cell = ws2.cell(row=row, column=col, value=header)
            cell.font = Font(name='黑体', size=10, bold=True)
            cell.fill = PatternFill(start_color='D9E1F2', end_color='D9E1F2', fill_type='solid')
            cell.alignment = Alignment(horizontal='center', vertical='center')
        row += 1

        ep_data = [
            ('年预计雷击次数 N1', 'N1', '来自防雷计算', f"{ep_result.get('N1', 0):.6f} 次/a"),
            ('电源线缆截收面积', "Ae1'", '按线缆类型计算', f"{ep_result.get('Ae1', 0):.6f} km²"),
            ('信号线缆截收面积', "Ae2'", '按线缆类型计算', f"{ep_result.get('Ae2', 0):.6f} km²"),
            ('入户设施 N2', 'N2', 'Ng × (Ae1 + Ae2)', f"{ep_result.get('N2', 0):.6f} 次/a"),
            ('总年预计雷击次数', 'N', 'N1 + N2', f"{ep_result.get('N', 0):.6f} 次/a"),
            ('C因子总和', 'C', 'C1+C2+C3+C4+C5+C6', f"{ep_result.get('C', 0):.2f}"),
            ('可接受最大Nc', 'Nc', '0.58/C', f"{ep_result.get('Nc', 0):.6f} 次/a"),
            ('拦截效率', 'E', '1 - Nc/N', f"{ep_result.get('E', 0):.4f}"),
            ('防护等级', '-', '根据E值判定', ep_level_text),
        ]

        for data_row in ep_data:
            for col, value in enumerate(data_row, 1):
                cell = ws2.cell(row=row, column=col, value=value)
                cell.font = Font(name='宋体', size=10)
                cell.alignment = Alignment(horizontal='center', vertical='center')
            row += 1

    # 设置列宽
    ws2.column_dimensions['A'].width = 25
    ws2.column_dimensions['B'].width = 15
    ws2.column_dimensions['C'].width = 30
    ws2.column_dimensions['D'].width = 25

    # ===== Sheet 3: 输入参数 =====
    ws3 = wb.create_sheet("输入参数")

    row = 1
    ws3.merge_cells(f'A{row}:B{row}')
    cell = ws3[f'A{row}']
    cell.value = '输入参数汇总'
    cell.font = Font(name='黑体', size=14, bold=True, color='003366')
    cell.alignment = Alignment(horizontal='center', vertical='center')
    row += 1

    for key, value in building_params.items():
        ws3[f'A{row}'] = key
        ws3[f'B{row}'] = str(value)
        ws3[f'A{row}'].font = Font(name='宋体', size=10)
        ws3[f'B{row}'].font = Font(name='宋体', size=10)
        row += 1

    if c_factors:
        row += 1
        ws3.merge_cells(f'A{row}:B{row}')
        cell = ws3[f'A{row}']
        cell.value = 'C1~C6 因子'
        cell.font = Font(name='黑体', size=12, bold=True, color='003366')
        cell.alignment = Alignment(horizontal='left', vertical='center')
        row += 1

        for key, value in c_factors.items():
            ws3[f'A{row}'] = key
            ws3[f'B{row}'] = f"{value.get('type', '')} = {value.get('val', 0)}"
            ws3[f'A{row}'].font = Font(name='宋体', size=10)
            ws3[f'B{row}'].font = Font(name='宋体', size=10)
            row += 1

    ws3.column_dimensions['A'].width = 30
    ws3.column_dimensions['B'].width = 40

    # ===== 保存 =====
    file_buffer = io.BytesIO()
    wb.save(file_buffer)
    file_buffer.seek(0)
    return file_buffer


def export_excel_separate(lp_result, ep_result, building_params,
                          cable_params, c_factors, level_text, ep_level_text,
                          building_attr="", attr_type=""):
    """
    分开导出两份 Excel
    """
    # ===== 防雷等级计算书 =====
    wb_lp = Workbook()
    wb_lp.remove(wb_lp.active)

    ws1 = wb_lp.create_sheet("防雷等级计算", 0)

    # 标题
    ws1.merge_cells('A1:F1')
    cell = ws1['A1']
    cell.value = '建筑物防雷等级计算书'
    cell.font = Font(name='黑体', size=18, bold=True, color='003366')
    cell.alignment = Alignment(horizontal='center', vertical='center')

    ws1.merge_cells('A2:F2')
    cell = ws1['A2']
    cell.value = f'计算日期：{datetime.now().strftime("%Y年%m月%d日")}'
    cell.font = Font(name='宋体', size=10)
    cell.alignment = Alignment(horizontal='center', vertical='center')

    row = 4
    for key, value in building_params.items():
        ws1[f'A{row}'] = key
        ws1[f'B{row}'] = str(value)
        row += 1

    row += 1
    ws1.merge_cells(f'A{row}:F{row}')
    cell = ws1[f'A{row}']
    cell.value = '防雷等级计算结果'
    cell.font = Font(name='黑体', size=12, bold=True, color='003366')
    cell.alignment = Alignment(horizontal='left', vertical='center')
    row += 1

    ws1[f'A{row}'] = '年预计雷击次数 N'
    ws1[f'B{row}'] = f"{lp_result.get('N', 0):.6f} 次/a"
    row += 1

    ws1[f'A{row}'] = '防雷等级'
    ws1[f'B{row}'] = level_text
    ws1[f'B{row}'].font = Font(name='黑体', size=14, bold=True, color='003366')

    for col in ['A', 'B']:
        ws1.column_dimensions[col].width = 25

    lp_buffer = io.BytesIO()
    wb_lp.save(lp_buffer)
    lp_buffer.seek(0)

    # ===== 电子防护等级计算书 =====
    wb_ep = Workbook()
    wb_ep.remove(wb_ep.active)

    ws1 = wb_ep.create_sheet("电子防护等级计算", 0)

    ws1.merge_cells('A1:F1')
    cell = ws1['A1']
    cell.value = '电子信息系统雷电防护等级计算书'
    cell.font = Font(name='黑体', size=18, bold=True, color='003366')
    cell.alignment = Alignment(horizontal='center', vertical='center')

    ws1.merge_cells('A2:F2')
    cell = ws1['A2']
    cell.value = f'计算日期：{datetime.now().strftime("%Y年%m月%d日")}'
    cell.font = Font(name='宋体', size=10)
    cell.alignment = Alignment(horizontal='center', vertical='center')

    row = 4
    for key, value in building_params.items():
        ws1[f'A{row}'] = key
        ws1[f'B{row}'] = str(value)
        row += 1

    row += 1
    ws1.merge_cells(f'A{row}:F{row}')
    cell = ws1[f'A{row}']
    cell.value = '电子信息系统防护等级计算结果'
    cell.font = Font(name='黑体', size=12, bold=True, color='003366')
    cell.alignment = Alignment(horizontal='left', vertical='center')
    row += 1

    if ep_result:
        ws1[f'A{row}'] = '总年预计雷击次数 N'
        ws1[f'B{row}'] = f"{ep_result.get('N', 0):.6f} 次/a"
        row += 1
        ws1[f'A{row}'] = '拦截效率 E'
        ws1[f'B{row}'] = f"{ep_result.get('E', 0):.4f}"
        row += 1
        ws1[f'A{row}'] = '防护等级'
        ws1[f'B{row}'] = ep_level_text
        ws1[f'B{row}'].font = Font(name='黑体', size=14, bold=True, color='003366')

    for col in ['A', 'B']:
        ws1.column_dimensions[col].width = 25

    ep_buffer = io.BytesIO()
    wb_ep.save(ep_buffer)
    ep_buffer.seek(0)

    return lp_buffer, ep_buffer