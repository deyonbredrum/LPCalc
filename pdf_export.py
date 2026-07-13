# -*- coding: utf-8 -*-
"""
PDF 计算书导出模块
使用 reportlab 生成专业 PDF 报告
"""

import io
from datetime import datetime
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.units import mm, cm
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.fonts import addMapping
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, PageBreak
from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_RIGHT

# 注册中文字体（使用 reportlab 内置的字体映射）
try:
    # 尝试注册中文字体，如果失败则使用默认字体
    from reportlab.pdfbase import pdfmetrics
    from reportlab.pdfbase.cidfonts import UnicodeCIDFont

    pdfmetrics.registerFont(UnicodeCIDFont('STSong-Light'))
    FONT_NAME = 'STSong-Light'
except:
    FONT_NAME = 'Helvetica'


def create_pdf_report(lp_result, ep_result, building_params,
                      c_factors, level_text, ep_level_text,
                      has_ep=True, building_attr="", attr_type=""):
    """
    生成 PDF 计算书
    """
    buffer = io.BytesIO()

    # 创建 PDF 文档
    doc = SimpleDocTemplate(
        buffer,
        pagesize=A4,
        rightMargin=72,
        leftMargin=72,
        topMargin=72,
        bottomMargin=72,
        title="建筑物防雷计算书"
    )

    # 样式设置
    styles = getSampleStyleSheet()

    # 标题样式
    title_style = ParagraphStyle(
        'TitleStyle',
        parent=styles['Heading1'],
        fontName=FONT_NAME,
        fontSize=20,
        alignment=TA_CENTER,
        spaceAfter=20,
        textColor=colors.HexColor('#003366')
    )

    # 一级标题样式
    heading1_style = ParagraphStyle(
        'Heading1Style',
        parent=styles['Heading2'],
        fontName=FONT_NAME,
        fontSize=14,
        spaceAfter=10,
        textColor=colors.HexColor('#003366')
    )

    # 二级标题样式
    heading2_style = ParagraphStyle(
        'Heading2Style',
        parent=styles['Heading3'],
        fontName=FONT_NAME,
        fontSize=12,
        spaceAfter=8
    )

    # 正文样式
    body_style = ParagraphStyle(
        'BodyStyle',
        parent=styles['Normal'],
        fontName=FONT_NAME,
        fontSize=10,
        leading=16,
        spaceAfter=6
    )

    # 表头样式
    header_style = ParagraphStyle(
        'HeaderStyle',
        parent=styles['Normal'],
        fontName=FONT_NAME,
        fontSize=10,
        alignment=TA_CENTER,
        textColor=colors.white,
        backColor=colors.HexColor('#D9E1F2')
    )

    # 构建内容
    story = []

    # ===== 主标题 =====
    if has_ep:
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
    if has_ep:
        story.append(Paragraph("          GB 50343-2012《建筑物电子信息系统防雷技术规范》", body_style))
    story.append(Spacer(1, 10))

    # ===== 二、输入参数 =====
    story.append(Paragraph("二、输入参数", heading1_style))
    story.append(Paragraph("2.1 建筑物参数", heading2_style))

    # 建筑物参数表格
    param_data = [["参数名称", "数值"]]
    for key, value in building_params.items():
        param_data.append([key, str(value)])

    param_table = Table(param_data, colWidths=[4 * cm, 10 * cm])
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

    # C1~C6 因子
    if c_factors:
        story.append(Paragraph("2.2 电子信息系统防护因子", heading2_style))
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

        factor_table = Table(factor_data, colWidths=[3.5 * cm, 7 * cm, 3.5 * cm])
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
        story.append(Spacer(1, 10))

    # ===== 三、防雷等级计算 =====
    story.append(PageBreak())
    story.append(Paragraph("三、防雷等级计算", heading1_style))

    calc_data = [
        ["计算项", "结果"],
        ["年雷暴日 Td", f"{lp_result.get('Td', 0):.2f} d/a"],
        ["雷击大地密度 Ng", f"{lp_result.get('Ng', 0):.4f} 次/(km²·a)"],
        ["扩大宽度 D", f"{lp_result.get('D', 0):.3f} m"],
        ["等效面积 Ae", f"{lp_result.get('Ae', 0):.6f} km²"],
        ["校正系数 k", f"{lp_result.get('k', 0):.2f}"],
        ["年预计雷击次数 N", f"{lp_result.get('N', 0):.6f} 次/a"],
    ]

    calc_table = Table(calc_data, colWidths=[6 * cm, 8 * cm])
    calc_table.setStyle(TableStyle([
        ('FONTNAME', (0, 0), (-1, -1), FONT_NAME),
        ('FONTSIZE', (0, 0), (-1, -1), 10),
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#D9E1F2')),
        ('BACKGROUND', (0, 1), (-1, -1), colors.HexColor('#F5F5F5')),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('GRID', (0, 0), (-1, -1), 0.5, colors.grey),
        ('TOPPADDING', (0, 0), (-1, -1), 6),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 6),
    ]))
    story.append(calc_table)
    story.append(Spacer(1, 15))

    # 防雷等级结论
    story.append(Paragraph("▶ 防雷等级判定结果", heading2_style))
    story.append(Paragraph(f"该建筑物的防雷等级为：<b>{level_text}</b>", body_style))
    story.append(Spacer(1, 10))

    # ===== 四、电子防护等级计算 =====
    if has_ep and ep_result:
        story.append(PageBreak())
        story.append(Paragraph("四、电子信息系统雷电防护等级计算", heading1_style))

        ep_data = [
            ["计算项", "结果"],
            ["建筑物年预计雷击次数 N1", f"{ep_result.get('N1', 0):.6f} 次/a"],
            ["电源线缆截收面积 Ae1'", f"{ep_result.get('Ae1', 0):.6f} km²"],
            ["信号线缆截收面积 Ae2'", f"{ep_result.get('Ae2', 0):.6f} km²"],
            ["入户设施年预计雷击次数 N2", f"{ep_result.get('N2', 0):.6f} 次/a"],
            ["总年预计雷击次数 N", f"{ep_result.get('N', 0):.6f} 次/a"],
            ["C = C1+C2+C3+C4+C5+C6", f"{ep_result.get('C', 0):.2f}"],
            ["可接受最大年雷击次数 Nc", f"{ep_result.get('Nc', 0):.6f} 次/a"],
            ["拦截效率 E", f"{ep_result.get('E', 0):.4f}"],
        ]

        ep_table = Table(ep_data, colWidths=[6 * cm, 8 * cm])
        ep_table.setStyle(TableStyle([
            ('FONTNAME', (0, 0), (-1, -1), FONT_NAME),
            ('FONTSIZE', (0, 0), (-1, -1), 10),
            ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#D9E1F2')),
            ('BACKGROUND', (0, 1), (-1, -1), colors.HexColor('#F5F5F5')),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
            ('GRID', (0, 0), (-1, -1), 0.5, colors.grey),
            ('TOPPADDING', (0, 0), (-1, -1), 6),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 6),
        ]))
        story.append(ep_table)
        story.append(Spacer(1, 15))

        story.append(Paragraph("▶ 电子信息系统防护等级判定结果", heading2_style))
        level = ep_level_text if ep_level_text else "未计算"
        if 'A' in level:
            color = '#CC0000'
        elif 'B' in level:
            color = '#FF8800'
        elif 'C' in level:
            color = '#0066CC'
        else:
            color = '#008800'
        story.append(
            Paragraph(f'该建筑物的电子信息系统防护等级为：<b><font color="{color}">{level}</font></b>', body_style))
        story.append(Spacer(1, 10))

    # ===== 五、结论 =====
    story.append(PageBreak())
    story.append(Paragraph("五、结论", heading1_style))

    story.append(Paragraph(
        f'1. 根据 GB 50057-2010 计算，该建筑物的防雷等级为：<b><font color="#003366">{level_text}</font></b>',
        body_style
    ))

    if has_ep and ep_result:
        level = ep_level_text if ep_level_text else "未计算"
        if 'A' in level:
            color = '#CC0000'
        elif 'B' in level:
            color = '#FF8800'
        elif 'C' in level:
            color = '#0066CC'
        else:
            color = '#008800'
        story.append(Paragraph(
            f'2. 根据 GB 50343-2012 计算，该建筑物的电子信息系统防护等级为：<b><font color="{color}">{level}</font></b>',
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

    # 生成 PDF
    doc.build(story)
    buffer.seek(0)
    return buffer