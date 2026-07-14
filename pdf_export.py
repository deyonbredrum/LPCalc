# -*- coding: utf-8 -*-
"""
PDF 计算书导出模块 - 完整计算过程版
"""

import io
import math
from datetime import datetime
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import cm
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.cidfonts import UnicodeCIDFont
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, PageBreak
from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER, TA_LEFT

try:
    pdfmetrics.registerFont(UnicodeCIDFont('STSong-Light'))
    FONT_NAME = 'STSong-Light'
except:
    FONT_NAME = 'Helvetica'

from excel_export import get_lp_conclusion_text, get_ep_conclusion_text


def create_pdf_report(lp_result, ep_result, building_params,
                       cable_params=None, c_factors=None, level_text="", ep_level_text="",
                       has_ep=True, building_attr="", attr_type=""):
    """
    生成 PDF 计算书
    """
    # 安全处理：确保参数不为 None
    if cable_params is None:
        cable_params = {}
    if c_factors is None:
        c_factors = {}
    if building_params is None:
        building_params = {}
    if lp_result is None:
        lp_result = {}
    if ep_result is None:
        ep_result = {}

    buffer = io.BytesIO()
    # ... 后续代码保持不变 ...

    doc = SimpleDocTemplate(
        buffer,
        pagesize=A4,
        rightMargin=72,
        leftMargin=72,
        topMargin=72,
        bottomMargin=72,
        title="建筑物防雷计算书"
    )

    styles = getSampleStyleSheet()

    title_style = ParagraphStyle(
        'TitleStyle',
        parent=styles['Heading1'],
        fontName=FONT_NAME,
        fontSize=20,
        alignment=TA_CENTER,
        spaceAfter=20,
        textColor=colors.HexColor('#003366')
    )

    heading1_style = ParagraphStyle(
        'Heading1Style',
        parent=styles['Heading2'],
        fontName=FONT_NAME,
        fontSize=14,
        spaceAfter=10,
        textColor=colors.HexColor('#003366')
    )

    heading2_style = ParagraphStyle(
        'Heading2Style',
        parent=styles['Heading3'],
        fontName=FONT_NAME,
        fontSize=12,
        spaceAfter=8
    )

    body_style = ParagraphStyle(
        'BodyStyle',
        parent=styles['Normal'],
        fontName=FONT_NAME,
        fontSize=10,
        leading=16,
        spaceAfter=6
    )

    story = []

    # ===== 主标题 =====
    if has_ep and ep_result:
        title_text = "建筑物防雷及电子信息防护等级计算书"
    else:
        title_text = "建筑物防雷等级计算书"

    story.append(Paragraph(title_text, title_style))
    story.append(Spacer(1, 10))
    story.append(Paragraph(f"计算日期：{datetime.now().strftime('%Y年%m月%d日')}", body_style))
    story.append(Spacer(1, 20))

    # ===== 一、基本信息 =====
    story.append(Paragraph("一、基本信息", heading1_style))
    story.append(Paragraph(f"计算日期：{datetime.now().strftime('%Y年%m月%d日')}", body_style))
    story.append(Paragraph("依据规范：GB 50057-2010《建筑物防雷设计规范》", body_style))
    if has_ep and ep_result:
        story.append(Paragraph("          GB 50343-2012《建筑物电子信息系统防雷技术规范》", body_style))
    story.append(Spacer(1, 10))

    # ===== 二、输入参数 =====
    story.append(Paragraph("二、输入参数", heading1_style))
    story.append(Paragraph("2.1 建筑物参数", heading2_style))

    param_data = [["参数名称", "数值"]]
    for key, value in building_params.items():
        param_data.append([key, str(value)])

    param_table = Table(param_data, colWidths=[4*cm, 10*cm])
    param_table.setStyle(TableStyle([
        ('FONTNAME', (0, 0), (-1, -1), FONT_NAME),
        ('FONTSIZE', (0, 0), (-1, -1), 9),
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#D9E1F2')),
        ('BACKGROUND', (0, 1), (-1, -1), colors.HexColor('#F5F5F5')),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('GRID', (0, 0), (-1, -1), 0.5, colors.grey),
        ('TOPPADDING', (0, 0), (-1, -1), 4),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 4),
    ]))
    story.append(param_table)
    story.append(Spacer(1, 10))

    # ===== 三、防雷等级详细计算过程 =====
    story.append(PageBreak())
    story.append(Paragraph("三、防雷等级详细计算过程", heading1_style))

    # 获取数据
    Td = lp_result.get('Td', 0)
    Ng = lp_result.get('Ng', 0)
    D = lp_result.get('D', 0)
    Ae = lp_result.get('Ae', 0)
    k = lp_result.get('k', 0)
    N = lp_result.get('N', 0)
    city = lp_result.get('city', '')
    L = lp_result.get('L', 0)
    W = lp_result.get('W', 0)
    H = lp_result.get('H', 0)
    surrounding_type = lp_result.get('surrounding_type', '')
    Lz = lp_result.get('Lz', 0)

    # 步骤1
    story.append(Paragraph("步骤1：确定年雷暴日", heading2_style))
    story.append(Paragraph(f"依据：GB 50057-2010 附录A，查表得 {city} 年雷暴日", body_style))
    story.append(Paragraph(f"计算公式：Td = 查表值", body_style))
    story.append(Paragraph(f"代入数据：Td = {Td:.2f} d/a", body_style))
    story.append(Paragraph(f"<b>计算结果：Td = {Td:.2f} d/a</b>", body_style))
    story.append(Spacer(1, 8))

    # 步骤2
    story.append(Paragraph("步骤2：计算雷击大地密度", heading2_style))
    story.append(Paragraph("依据：GB 50057-2010 公式（附A.4）", body_style))
    story.append(Paragraph("计算公式：Ng = 0.1 × Td", body_style))
    story.append(Paragraph(f"代入数据：Ng = 0.1 × {Td:.2f}", body_style))
    story.append(Paragraph(f"<b>计算结果：Ng = {Ng:.4f} 次/(km²·a)</b>", body_style))
    story.append(Spacer(1, 8))

    # 步骤3
    story.append(Paragraph("步骤3：计算扩大宽度", heading2_style))
    story.append(Paragraph("依据：GB 50057-2010 附录A", body_style))
    story.append(Paragraph("计算公式：D = √(H × (200 - H))", body_style))
    story.append(Paragraph(f"代入数据：D = √({H:.1f} × (200 - {H:.1f}))", body_style))
    story.append(Paragraph(f"<b>计算结果：D = {D:.3f} m</b>", body_style))
    story.append(Spacer(1, 8))

    # 步骤4
    story.append(Paragraph("步骤4：计算等效面积", heading2_style))
    story.append(Paragraph("依据：GB 50057-2010 附录A", body_style))
    story.append(Paragraph(f"周边情况：{surrounding_type}", body_style))

    if surrounding_type == "周边2D范围内无其他建筑":
        formula = "Ae = LW + 2(L+W)D + πD²"
        formula_with_values = f"Ae = {L}×{W} + 2×({L}+{W})×{D:.3f} + π×{D:.3f}²"
    elif surrounding_type == "周边2D范围内有等高或比它低的其他建筑物，但不在所考虑建筑物以外100m的保护范围内":
        formula = "Ae = LW + 2(L+W)D + πD² - Lz×D/2"
        formula_with_values = f"Ae = {L}×{W} + 2×({L}+{W})×{D:.3f} + π×{D:.3f}² - {Lz}×{D:.3f}/2"
    elif surrounding_type == "周边2D范围内有等高或比它低的其他建筑物":
        formula = "Ae = LW + (L+W)D + πD²/4"
        formula_with_values = f"Ae = {L}×{W} + ({L}+{W})×{D:.3f} + π×{D:.3f}²/4"
    elif surrounding_type == "周边2D范围内有比它高的其他建筑物":
        formula = "Ae = LW + 2(L+W)D + πD² - Lz×D"
        formula_with_values = f"Ae = {L}×{W} + 2×({L}+{W})×{D:.3f} + π×{D:.3f}² - {Lz}×{D:.3f}"
    elif surrounding_type == "周边2D范围内都有比它高的其他建筑物":
        formula = "Ae = LW"
        formula_with_values = f"Ae = {L}×{W}"
    else:
        formula = "Ae = LW + 2(L+W)D + πD²"
        formula_with_values = f"Ae = {L}×{W} + 2×({L}+{W})×{D:.3f} + π×{D:.3f}²"

    story.append(Paragraph(f"计算公式：{formula}", body_style))
    story.append(Paragraph(f"代入数据：{formula_with_values}", body_style))
    story.append(Paragraph(f"<b>计算结果：Ae = {Ae:.6f} km²</b>", body_style))
    story.append(Spacer(1, 8))

    # 步骤5
    story.append(Paragraph("步骤5：确定校正系数", heading2_style))
    story.append(Paragraph(f"依据：GB 50057-2010 第4.2.1条，{building_params.get('建筑物状况', '一般情况')}", body_style))
    story.append(Paragraph("计算公式：k = 查表值", body_style))
    story.append(Paragraph(f"代入数据：k = {k:.2f}", body_style))
    story.append(Paragraph(f"<b>计算结果：k = {k:.2f}</b>", body_style))
    story.append(Spacer(1, 8))

    # 步骤6
    story.append(Paragraph("步骤6：计算年预计雷击次数", heading2_style))
    story.append(Paragraph("依据：GB 50057-2010 公式（附A.2）", body_style))
    story.append(Paragraph("计算公式：N = k × Ng × Ae", body_style))
    story.append(Paragraph(f"代入数据：N = {k:.2f} × {Ng:.4f} × {Ae:.6f}", body_style))
    story.append(Paragraph(f"<b>计算结果：N = {N:.6f} 次/a</b>", body_style))
    story.append(Spacer(1, 8))

    # 步骤7
    story.append(Paragraph("步骤7：防雷等级判定", heading2_style))
    story.append(Paragraph(f"依据：GB 50057-2010 第3.0.2~3.0.4条，建筑物属性：{building_params.get('建筑物属性', '一般民用建筑')}", body_style))
    if level_text == "一类":
        condition = "N > 0.05 次/a 且 爆炸危险场所"
    elif level_text == "二类":
        condition = "N > 0.05 次/a（人员密集/重要场所）或 N > 0.25 次/a（一般民用）"
    elif level_text == "三类":
        condition = "N > 0.01 次/a（人员密集/重要场所）或 N > 0.05 次/a（一般民用）"
    else:
        condition = "N ≤ 0.01 次/a（人员密集/重要场所）或 N ≤ 0.05 次/a（一般民用）"
    story.append(Paragraph(f"判定条件：{condition}", body_style))
    story.append(Paragraph(f"<b>判定结果：防雷等级 {level_text}</b>", body_style))
    story.append(Spacer(1, 8))

    # 防雷等级结论 + 判断依据
    story.append(Paragraph("▶ 防雷等级判定结论", heading2_style))
    conclusion = get_lp_conclusion_text(level_text, N, building_attr, attr_type)
    story.append(Paragraph(f"📖 {conclusion}", body_style))
    story.append(Spacer(1, 10))

    # ===== 四、电子防护详细计算过程 =====
    if has_ep and ep_result:
        story.append(PageBreak())
        story.append(Paragraph("四、电子信息系统防护等级详细计算过程", heading1_style))

        # 获取数据
        N1 = ep_result.get('N1', 0)
        N2 = ep_result.get('N2', 0)
        N_total = ep_result.get('N', 0)
        C_total = ep_result.get('C', 0)
        Nc = ep_result.get('Nc', 0)
        E = ep_result.get('E', 0)
        Ae1 = ep_result.get('Ae1', 0)
        Ae2 = ep_result.get('Ae2', 0)

        # 步骤1
        story.append(Paragraph("步骤1：确定入户设施截收面积", heading2_style))
        story.append(Paragraph("依据：GB 50343-2012 附录A", body_style))
        story.append(Paragraph(f"电源线缆入户方式：{cable_params.get('电源线缆入户方式', '未选择')}", body_style))
        story.append(Paragraph(f"信号线缆入户方式：{cable_params.get('信号线缆入户方式', '未选择')}", body_style))
        story.append(Paragraph(f"<b>截收面积结果：Ae1' = {Ae1:.6f} km²，Ae2' = {Ae2:.6f} km²</b>", body_style))
        story.append(Spacer(1, 8))

        # 步骤2
        story.append(Paragraph("步骤2：计算入户设施年预计雷击次数 N2", heading2_style))
        story.append(Paragraph("依据：GB 50343-2012 附录A", body_style))
        story.append(Paragraph("计算公式：N2 = Ng × (Ae1' + Ae2')", body_style))
        story.append(Paragraph(f"代入数据：N2 = {Ng:.4f} × ({Ae1:.6f} + {Ae2:.6f})", body_style))
        story.append(Paragraph(f"<b>计算结果：N2 = {N2:.6f} 次/a</b>", body_style))
        story.append(Spacer(1, 8))

        # 步骤3
        story.append(Paragraph("步骤3：计算总年预计雷击次数 N", heading2_style))
        story.append(Paragraph("计算公式：N = N1 + N2", body_style))
        story.append(Paragraph(f"代入数据：N = {N1:.6f} + {N2:.6f}", body_style))
        story.append(Paragraph(f"<b>计算结果：N = {N_total:.6f} 次/a</b>", body_style))
        story.append(Spacer(1, 8))

        # 步骤4: C因子
        story.append(Paragraph("步骤4：C因子确定", heading2_style))
        story.append(Paragraph("依据：GB 50343-2012 表5.4.3-2 ~ 表5.4.3-3", body_style))

        factor_data = [["因子", "选择项", "取值"]]
        factor_names = {
            'C1': '建筑物材料结构',
            'C2': '信息系统重要程度',
            'C3': '设备耐冲击能力',
            'C4': '雷电防护区',
            'C5': '雷击事故后果',
            'C6': '区域雷暴等级',
        }
        for key, value in c_factors.items():
            factor_data.append([
                f"{key}（{factor_names.get(key, '')}）",
                value.get('type', ''),
                str(value.get('val', ''))
            ])

        factor_table = Table(factor_data, colWidths=[3.5*cm, 7*cm, 3.5*cm])
        factor_table.setStyle(TableStyle([
            ('FONTNAME', (0, 0), (-1, -1), FONT_NAME),
            ('FONTSIZE', (0, 0), (-1, -1), 9),
            ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#D9E1F2')),
            ('BACKGROUND', (0, 1), (-1, -1), colors.HexColor('#F5F5F5')),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
            ('GRID', (0, 0), (-1, -1), 0.5, colors.grey),
            ('TOPPADDING', (0, 0), (-1, -1), 4),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 4),
        ]))
        story.append(factor_table)
        story.append(Spacer(1, 5))
        story.append(Paragraph(f"<b>C因子总和：C = {C_total:.2f}</b>", body_style))
        story.append(Spacer(1, 8))

        # 步骤5
        story.append(Paragraph("步骤5：计算可接受最大年雷击次数 Nc", heading2_style))
        story.append(Paragraph("依据：GB 50343-2012 附录A", body_style))
        story.append(Paragraph("计算公式：Nc = 0.58 / C", body_style))
        story.append(Paragraph(f"代入数据：Nc = 0.58 / {C_total:.2f}", body_style))
        story.append(Paragraph(f"<b>计算结果：Nc = {Nc:.6f} 次/a</b>", body_style))
        story.append(Spacer(1, 8))

        # 步骤6
        story.append(Paragraph("步骤6：计算拦截效率 E", heading2_style))
        story.append(Paragraph("依据：GB 50343-2012 附录A", body_style))
        story.append(Paragraph("计算公式：E = 1 - Nc / N", body_style))
        story.append(Paragraph(f"代入数据：E = 1 - {Nc:.6f} / {N_total:.6f}", body_style))
        story.append(Paragraph(f"<b>计算结果：E = {E:.4f}</b>", body_style))
        story.append(Spacer(1, 8))

        # 步骤7
        story.append(Paragraph("步骤7：防护等级判定", heading2_style))
        story.append(Paragraph("依据：GB 50343-2012 第5.4.1条", body_style))
        if ep_level_text == "A级":
            condition = "E > 0.98"
        elif ep_level_text == "B级":
            condition = "0.95 < E ≤ 0.98"
        elif ep_level_text == "C级":
            condition = "0.90 < E ≤ 0.95"
        elif ep_level_text == "D级":
            condition = "0.80 < E ≤ 0.90"
        else:
            condition = "E ≤ 0.80"
        story.append(Paragraph(f"判定条件：{condition}", body_style))

        level = ep_level_text if ep_level_text else "未计算"
        if 'A' in level:
            color = '#CC0000'
        elif 'B' in level:
            color = '#FF8800'
        elif 'C' in level:
            color = '#0066CC'
        else:
            color = '#008800'
        story.append(Paragraph(f'<b>判定结果：防护等级 <font color="{color}">{level}</font></b>', body_style))
        story.append(Spacer(1, 8))

        # 电子防护结论 + 判断依据
        story.append(Paragraph("▶ 电子信息系统防护等级判定结论", heading2_style))
        ep_conclusion = get_ep_conclusion_text(level, E)
        story.append(Paragraph(f"📖 {ep_conclusion}", body_style))
        story.append(Spacer(1, 10))

    # ===== 五、结论 =====
    story.append(PageBreak())
    story.append(Paragraph("五、结论", heading1_style))

    N = lp_result.get('N', 0)
    conclusion = get_lp_conclusion_text(level_text, N, building_attr, attr_type)
    story.append(Paragraph(
        f'1. {conclusion}',
        body_style
    ))
    story.append(Spacer(1, 5))

    if has_ep and ep_result:
        level = ep_level_text if ep_level_text else "未计算"
        E = ep_result.get('E', 0)
        ep_conclusion = get_ep_conclusion_text(level, E)
        story.append(Paragraph(
            f'2. {ep_conclusion}',
            body_style
        ))

    story.append(Spacer(1, 20))

    # 页脚
    story.append(Paragraph("_" * 60, body_style))
    story.append(Paragraph(
        "注：本计算书由 LPCalc 工具自动生成，仅供参考。正式设计需经专业审核。",
        ParagraphStyle(
            'FooterStyle',
            parent=styles['Normal'],
            fontName=FONT_NAME,
            fontSize=8,
            alignment=TA_CENTER,
            textColor=colors.grey
        )
    ))

    doc.build(story)
    buffer.seek(0)
    return buffer


def create_pdf_separate(lp_result, building_params,
                         c_factors, level_text,
                         building_attr="", attr_type=""):
    """仅导出防雷等级 PDF"""
    return create_pdf_report(
        lp_result=lp_result,
        ep_result=None,
        building_params=building_params,
        cable_params={},
        c_factors={},
        level_text=level_text,
        ep_level_text=None,
        has_ep=False,
        building_attr=building_attr,
        attr_type=attr_type
    )