# -*- coding: utf-8 -*-
"""
Excel 计算书导出模块
"""

import io
from datetime import datetime
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill


def get_lp_conclusion_text(level_text, N, building_attr="", attr_type=""):
    """
    根据防雷等级和建筑属性生成规范依据文本
    """
    if level_text == "一类":
        clause = "第3.0.2条"
        detail = f"制造、使用或贮存爆炸危险物质，且年平均雷击次数N={N:.6f}次/a大于0.05次/a"
    elif level_text == "二类":
        if attr_type in ["crowded", "important", "heritage", "office", "comm", "station", "airport"]:
            clause = "第3.0.3条"
            detail = f"人员密集的公共建筑物或重要场所，且年平均雷击次数N={N:.6f}次/a大于0.05次/a"
        else:
            clause = "第3.0.3条"
            detail = f"一般民用建筑，年平均雷击次数N={N:.6f}次/a大于0.25次/a"
    elif level_text == "三类":
        if attr_type in ["crowded", "important", "heritage", "office", "comm", "station", "airport"]:
            clause = "第3.0.4条"
            detail = f"人员密集的公共建筑物或重要场所，年平均雷击次数N={N:.6f}次/a在0.01~0.05次/a之间"
        else:
            clause = "第3.0.4条"
            detail = f"一般民用建筑，年平均雷击次数N={N:.6f}次/a在0.05~0.25次/a之间"
    else:
        clause = "第3.0.4条"
        detail = f"年平均雷击次数N={N:.6f}次/a小于0.01次/a，可不设防雷"

    return f"根据GB 50057-2010《建筑物防雷设计规范》{clause}规定，{detail}，应判定为{level_text}防雷建筑物。"


def get_ep_conclusion_text(ep_level_text, E):
    """根据电子防护等级生成规范依据文本"""
    if ep_level_text == "A级":
        detail = f"雷电防护等级为A级，拦截效率E={E:.4f} > 0.98"
    elif ep_level_text == "B级":
        detail = f"雷电防护等级为B级，拦截效率0.95 < E={E:.4f} ≤ 0.98"
    elif ep_level_text == "C级":
        detail = f"雷电防护等级为C级，拦截效率0.90 < E={E:.4f} ≤ 0.95"
    elif ep_level_text == "D级":
        detail = f"雷电防护等级为D级，拦截效率0.80 < E={E:.4f} ≤ 0.90"
    else:
        detail = f"雷电防护等级低于D级，拦截效率E={E:.4f} ≤ 0.80，可不设防护"

    return f"根据GB 50343-2012《建筑物电子信息系统防雷技术规范》第5.4.1条规定，{detail}，应判定为{ep_level_text}。"


def export_excel_report(lp_result, ep_result, building_params,
                        cable_params, c_factors, level_text, ep_level_text,
                        export_mode='combined', has_ep=True,
                        building_attr="", attr_type=""):
    """
    导出 Excel 计算书
    has_ep=True: 完整计算书（防雷 + 电子防护，C1~C6出现在电子防护部分）
    has_ep=False: 仅防雷计算书（不包含C1~C6）
    """
    wb = Workbook()
    wb.remove(wb.active)

    # ===== Sheet 1: 摘要 =====
    ws1 = wb.create_sheet("摘要", 0)

    if has_ep and ep_result:
        title_text = '建筑物防雷及电子信息防护等级计算书'
    else:
        title_text = '建筑物防雷等级计算书'

    ws1.merge_cells('A1:F1')
    cell = ws1['A1']
    cell.value = title_text
    cell.font = Font(name='黑体', size=18, bold=True, color='003366')
    cell.alignment = Alignment(horizontal='center', vertical='center')

    ws1.merge_cells('A2:F2')
    cell = ws1['A2']
    cell.value = f'计算日期：{datetime.now().strftime("%Y年%m月%d日")}'
    cell.font = Font(name='宋体', size=10)
    cell.alignment = Alignment(horizontal='center', vertical='center')

    ws1.row_dimensions[3].height = 10

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
        row += 1

    # ===== 防雷等级结果 =====
    row += 1
    ws1.merge_cells(f'A{row}:F{row}')
    cell = ws1[f'A{row}']
    cell.value = '二、防雷等级计算结果'
    cell.font = Font(name='黑体', size=12, bold=True, color='003366')
    cell.alignment = Alignment(horizontal='left', vertical='center')
    row += 1

    N = lp_result.get('N', 0)
    ws1[f'A{row}'] = '年预计雷击次数 N'
    ws1[f'B{row}'] = f"{N:.6f} 次/a"
    row += 1

    ws1[f'A{row}'] = '防雷等级'
    ws1[f'B{row}'] = level_text
    ws1[f'B{row}'].font = Font(name='宋体', size=12, bold=True, color='003366')
    row += 1

    # 防雷判断依据
    ws1[f'A{row}'] = '判断依据'
    conclusion = get_lp_conclusion_text(level_text, N, building_attr, attr_type)
    ws1.merge_cells(f'B{row}:F{row}')
    ws1[f'B{row}'] = conclusion
    ws1[f'B{row}'].font = Font(name='宋体', size=10)
    ws1[f'B{row}'].alignment = Alignment(horizontal='left', vertical='center', wrap_text=True)
    ws1.row_dimensions[row].height = 30
    row += 1

    # ===== 电子防护等级结果（仅当 has_ep=True 时显示） =====
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
        row += 1

        E = ep_result.get('E', 0)
        ws1[f'A{row}'] = '拦截效率 E'
        ws1[f'B{row}'] = f"{E:.4f}"
        row += 1

        ws1[f'A{row}'] = '防护等级'
        ws1[f'B{row}'] = ep_level_text
        ws1[f'B{row}'].font = Font(name='宋体', size=12, bold=True)
        color_map = {'A级': 'CC0000', 'B级': 'FF8800', 'C级': '0066CC', 'D级': '008800'}
        if ep_level_text in color_map:
            ws1[f'B{row}'].font = Font(name='宋体', size=12, bold=True, color=color_map[ep_level_text])
        row += 1

        # 电子防护判断依据
        ws1[f'A{row}'] = '判断依据'
        ep_conclusion = get_ep_conclusion_text(ep_level_text, E)
        ws1.merge_cells(f'B{row}:F{row}')
        ws1[f'B{row}'] = ep_conclusion
        ws1[f'B{row}'].font = Font(name='宋体', size=10)
        ws1[f'B{row}'].alignment = Alignment(horizontal='left', vertical='center', wrap_text=True)
        ws1.row_dimensions[row].height = 30

    for col in ['A', 'B', 'C', 'D', 'E', 'F']:
        ws1.column_dimensions[col].width = 20

    # ===== Sheet 2: 详细计算过程 =====
    ws2 = wb.create_sheet("详细计算过程")

    # ---- 防雷等级计算 ----
    row = 1
    ws2.merge_cells(f'A{row}:D{row}')
    cell = ws2[f'A{row}']
    cell.value = '一、防雷等级计算过程'
    cell.font = Font(name='黑体', size=14, bold=True, color='003366')
    cell.alignment = Alignment(horizontal='center', vertical='center')
    row += 1

    headers = ['计算项', '符号', '公式', '结果']
    for col, header in enumerate(headers, 1):
        cell = ws2.cell(row=row, column=col, value=header)
        cell.font = Font(name='黑体', size=10, bold=True)
        cell.fill = PatternFill(start_color='D9E1F2', end_color='D9E1F2', fill_type='solid')
        cell.alignment = Alignment(horizontal='center', vertical='center')
    row += 1

    lp_data = [
        ('年雷暴日', 'Td', '查表', f"{lp_result.get('Td', 0):.2f} d/a"),
        ('雷击大地密度', 'Ng', '0.1 × Td', f"{lp_result.get('Ng', 0):.4f} 次/(km²·a)"),
        ('扩大宽度', 'D', '√(H(200-H))', f"{lp_result.get('D', 0):.3f} m"),
        ('等效面积', 'Ae', '按周边情况计算', f"{lp_result.get('Ae', 0):.6f} km²"),
        ('校正系数', 'k', '查表', f"{lp_result.get('k', 0):.2f}"),
        ('年预计雷击次数', 'N', 'k × Ng × Ae', f"{lp_result.get('N', 0):.6f} 次/a"),
    ]

    for data_row in lp_data:
        for col, value in enumerate(data_row, 1):
            cell = ws2.cell(row=row, column=col, value=value)
            cell.font = Font(name='宋体', size=10)
            cell.alignment = Alignment(horizontal='center', vertical='center')
        row += 1

    # ---- 电子防护等级计算（仅当 has_ep=True 时显示） ----
    if has_ep and ep_result:
        row += 2
        ws2.merge_cells(f'A{row}:D{row}')
        cell = ws2[f'A{row}']
        cell.value = '二、电子信息系统防护等级计算过程'
        cell.font = Font(name='黑体', size=14, bold=True, color='003366')
        cell.alignment = Alignment(horizontal='center', vertical='center')
        row += 1

        for col, header in enumerate(headers, 1):
            cell = ws2.cell(row=row, column=col, value=header)
            cell.font = Font(name='黑体', size=10, bold=True)
            cell.fill = PatternFill(start_color='D9E1F2', end_color='D9E1F2', fill_type='solid')
            cell.alignment = Alignment(horizontal='center', vertical='center')
        row += 1

        # C1~C6 因子显示在电子防护计算过程中
        ep_data = [
            ('年预计雷击次数 N1', 'N1', '来自防雷计算', f"{ep_result.get('N1', 0):.6f} 次/a"),
            ('电源线缆截收面积', "Ae1'", '按线缆类型计算', f"{ep_result.get('Ae1', 0):.6f} km²"),
            ('信号线缆截收面积', "Ae2'", '按线缆类型计算', f"{ep_result.get('Ae2', 0):.6f} km²"),
            ('入户设施 N2', 'N2', 'Ng × (Ae1 + Ae2)', f"{ep_result.get('N2', 0):.6f} 次/a"),
            ('总年预计雷击次数', 'N', 'N1 + N2', f"{ep_result.get('N', 0):.6f} 次/a"),
        ]
        for data_row in ep_data:
            for col, value in enumerate(data_row, 1):
                cell = ws2.cell(row=row, column=col, value=value)
                cell.font = Font(name='宋体', size=10)
                cell.alignment = Alignment(horizontal='center', vertical='center')
            row += 1

        # C因子单独显示（带具体选项）
        row += 1
        ws2.merge_cells(f'A{row}:D{row}')
        cell = ws2[f'A{row}']
        cell.value = 'C因子明细'
        cell.font = Font(name='黑体', size=11, bold=True, color='003366')
        cell.alignment = Alignment(horizontal='center', vertical='center')
        row += 1

        factor_names = {
            'C1': '建筑物材料结构',
            'C2': '信息系统重要程度',
            'C3': '设备耐冲击能力',
            'C4': '雷电防护区',
            'C5': '雷击事故后果',
            'C6': '区域雷暴等级',
        }
        for key, value in c_factors.items():
            ws2[f'A{row}'] = key
            ws2[f'B{row}'] = value.get('type', '')
            ws2[f'C{row}'] = str(value.get('val', 0))
            ws2[f'D{row}'] = ''  # 留空
            row += 1

        # C因子总和及后续计算
        ws2[f'A{row}'] = 'C = C1+C2+C3+C4+C5+C6'
        ws2[f'B{row}'] = f"{ep_result.get('C', 0):.2f}"
        row += 1

        ep_result_data = [
            ('可接受最大Nc', 'Nc', '0.58/C', f"{ep_result.get('Nc', 0):.6f} 次/a"),
            ('拦截效率', 'E', '1 - Nc/N', f"{ep_result.get('E', 0):.4f}"),
        ]
        for data_row in ep_result_data:
            for col, value in enumerate(data_row, 1):
                cell = ws2.cell(row=row, column=col, value=value)
                cell.font = Font(name='宋体', size=10)
                cell.alignment = Alignment(horizontal='center', vertical='center')
            row += 1

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
        row += 1

    # 输入参数中不显示 C1~C6，它们只出现在电子防护计算过程中
    ws3.column_dimensions['A'].width = 30
    ws3.column_dimensions['B'].width = 40

    file_buffer = io.BytesIO()
    wb.save(file_buffer)
    file_buffer.seek(0)
    return file_buffer


def export_excel_separate(lp_result, ep_result, building_params,
                          cable_params, c_factors, level_text, ep_level_text,
                          building_attr="", attr_type=""):
    """分开导出两份 Excel"""
    # ===== 防雷等级计算书（不包含任何电子防护内容） =====
    wb_lp = Workbook()
    wb_lp.remove(wb_lp.active)

    ws1 = wb_lp.create_sheet("防雷等级计算", 0)

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

    N = lp_result.get('N', 0)
    ws1[f'A{row}'] = '年预计雷击次数 N'
    ws1[f'B{row}'] = f"{N:.6f} 次/a"
    row += 1

    ws1[f'A{row}'] = '防雷等级'
    ws1[f'B{row}'] = level_text
    ws1[f'B{row}'].font = Font(name='黑体', size=14, bold=True, color='003366')
    row += 1

    # 判断依据
    ws1[f'A{row}'] = '判断依据'
    conclusion = get_lp_conclusion_text(level_text, N, building_attr, attr_type)
    ws1.merge_cells(f'B{row}:F{row}')
    ws1[f'B{row}'] = conclusion
    ws1[f'B{row}'].font = Font(name='宋体', size=10)
    ws1[f'B{row}'].alignment = Alignment(horizontal='left', vertical='center', wrap_text=True)
    ws1.row_dimensions[row].height = 30

    for col in ['A', 'B']:
        ws1.column_dimensions[col].width = 25

    lp_buffer = io.BytesIO()
    wb_lp.save(lp_buffer)
    lp_buffer.seek(0)

    # ===== 电子防护等级计算书（包含防雷计算 + 电子防护计算） =====
    ep_buffer = export_excel_report(
        lp_result=lp_result,
        ep_result=ep_result,
        building_params=building_params,
        cable_params=cable_params,
        c_factors=c_factors,
        level_text=level_text,
        ep_level_text=ep_level_text,
        export_mode='combined',
        has_ep=True,
        building_attr=building_attr,
        attr_type=attr_type
    )

    return lp_buffer, ep_buffer