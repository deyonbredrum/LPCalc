# -*- coding: utf-8 -*-
"""
Word 计算书导出模块 - 使用 docxtpl（更稳定）
"""

import io
import os
from datetime import datetime
from docxtpl import DocxTemplate
from docx.shared import Pt, RGBColor, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement


def set_cell_font(cell, text, font_name='宋体', font_size=10.5, bold=False,
                  color=None, alignment='left'):
    """设置单元格内容和字体"""
    cell.text = text
    paragraph = cell.paragraphs[0]
    run = paragraph.runs[0]
    run.font.name = font_name
    run.font.size = Pt(font_size)
    run.font.bold = bold
    if color:
        run.font.color.rgb = RGBColor(*color)
    run._element.rPr.rFonts.set(qn('w:eastAsia'), font_name)

    if alignment == 'center':
        paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    elif alignment == 'right':
        paragraph.alignment = WD_ALIGN_PARAGRAPH.RIGHT


def set_heading_font(paragraph, font_name='黑体', font_size=16, bold=True, color=(0, 0, 0)):
    run = paragraph.runs[0] if paragraph.runs else paragraph.add_run()
    run.font.name = font_name
    run.font.size = Pt(font_size)
    run.font.bold = bold
    run._element.rPr.rFonts.set(qn('w:eastAsia'), font_name)
    if color:
        run.font.color.rgb = RGBColor(*color)


def set_cell_background(cell, color_hex='D9E1F2'):
    tc = cell._element
    tcPr = tc.get_or_add_tcPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), color_hex)
    tcPr.append(shd)


def build_basic_info_section(doc):
    """构建基本信息部分（不含电子防护）"""
    heading = doc.add_heading('一、基本信息', level=1)
    set_heading_font(heading, '黑体', 16, True, (0, 51, 102))

    now = datetime.now()
    p = doc.add_paragraph()
    p.add_run('计算日期：').bold = True
    p.add_run(now.strftime("%Y年%m月%d日"))

    p = doc.add_paragraph()
    p.add_run('依据规范：').bold = True
    p.add_run('GB 50057-2010《建筑物防雷设计规范》')

    doc.add_paragraph()


def build_input_params_section(doc, building_params, include_ep=True):
    """构建输入参数部分"""
    heading = doc.add_heading('二、输入参数', level=1)
    set_heading_font(heading, '黑体', 16, True, (0, 51, 102))

    heading3 = doc.add_heading('2.1 建筑物参数', level=2)
    set_heading_font(heading3, '黑体', 14, True, (0, 0, 0))

    table1 = doc.add_table(rows=len(building_params) + 1, cols=2)
    table1.style = 'Light Grid Accent 1'
    table1.alignment = WD_TABLE_ALIGNMENT.CENTER
    table1.columns[0].width = Cm(4.5)
    table1.columns[1].width = Cm(11)

    hdr_cells = table1.rows[0].cells
    hdr_cells[0].text = '参数名称'
    hdr_cells[1].text = '数值'
    for cell in hdr_cells:
        cell.paragraphs[0].runs[0].font.name = '黑体'
        cell.paragraphs[0].runs[0].font.size = Pt(10.5)
        cell.paragraphs[0].runs[0].font.bold = True
        cell.paragraphs[0].runs[0]._element.rPr.rFonts.set(qn('w:eastAsia'), '黑体')
        set_cell_background(cell, 'D9E1F2')

    for i, (key, value) in enumerate(building_params.items(), 1):
        row = table1.rows[i]
        set_cell_font(row.cells[0], key, '宋体', 10.5, False)
        set_cell_font(row.cells[1], str(value), '宋体', 10.5, False)

    doc.add_paragraph()
    return doc


def build_lp_section(doc, lp_result, level_text):
    heading = doc.add_heading('三、防雷等级计算', level=1)
    set_heading_font(heading, '黑体', 16, True, (0, 51, 102))

    calc_steps = [
        ('年雷暴日 Td', f"{lp_result.get('Td', 0):.2f} d/a"),
        ('雷击大地密度 Ng', f"{lp_result.get('Ng', 0):.4f} 次/(km²·a)"),
        ('扩大宽度 D', f"{lp_result.get('D', 0):.3f} m"),
        ('等效面积 Ae', f"{lp_result.get('Ae', 0):.6f} km²"),
        ('校正系数 k', f"{lp_result.get('k', 0):.2f}"),
        ('年预计雷击次数 N', f"{lp_result.get('N', 0):.6f} 次/a"),
    ]

    table = doc.add_table(rows=len(calc_steps) + 1, cols=2)
    table.style = 'Light Grid Accent 1'
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.columns[0].width = Cm(5)
    table.columns[1].width = Cm(10.5)

    hdr_cells = table.rows[0].cells
    hdr_cells[0].text = '计算项'
    hdr_cells[1].text = '结果'
    for cell in hdr_cells:
        cell.paragraphs[0].runs[0].font.name = '黑体'
        cell.paragraphs[0].runs[0].font.size = Pt(10.5)
        cell.paragraphs[0].runs[0].font.bold = True
        cell.paragraphs[0].runs[0]._element.rPr.rFonts.set(qn('w:eastAsia'), '黑体')
        set_cell_background(cell, 'D9E1F2')

    for i, (name, value) in enumerate(calc_steps, 1):
        row = table.rows[i]
        set_cell_font(row.cells[0], name, '宋体', 10.5, False)
        set_cell_font(row.cells[1], str(value), 'Arial', 10.5, False)

    doc.add_paragraph()
    p = doc.add_paragraph()
    p.add_run('▶ 防雷等级判定结果：').bold = True
    run = p.add_run(level_text)
    run.bold = True
    run.font.color.rgb = RGBColor(0, 112, 192)
    run.font.size = Pt(12)
    doc.add_paragraph()
    return doc


def build_conclusion_section(doc, lp_result, level_text, ep_result=None,
                             ep_level_text=None, has_ep=False, building_attr="", attr_type=""):
    N = lp_result.get('N', 0)

    heading = doc.add_heading('四、结论', level=1)
    set_heading_font(heading, '黑体', 16, True, (0, 51, 102))

    p = doc.add_paragraph()
    run = p.add_run('1. ')
    run.bold = True
    p.add_run(f'根据 GB 50057-2010 计算，该建筑物的防雷等级为：')
    run = p.add_run(f' {level_text}')
    run.bold = True
    run.font.color.rgb = RGBColor(0, 112, 192)
    run.font.size = Pt(12)

    if has_ep and ep_result and ep_level_text:
        p = doc.add_paragraph()
        run = p.add_run('2. ')
        run.bold = True
        p.add_run(f'根据 GB 50343-2012 计算，该建筑物的电子信息系统雷电防护等级为：')
        run = p.add_run(f' {ep_level_text}')
        run.bold = True
        if 'A' in ep_level_text:
            run.font.color.rgb = RGBColor(192, 0, 0)
        elif 'B' in ep_level_text:
            run.font.color.rgb = RGBColor(255, 128, 0)
        elif 'C' in ep_level_text:
            run.font.color.rgb = RGBColor(0, 112, 192)
        else:
            run.font.color.rgb = RGBColor(0, 128, 0)
        run.font.size = Pt(12)

    doc.add_paragraph()
    doc.add_paragraph('_' * 50).alignment = WD_ALIGN_PARAGRAPH.CENTER
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run('注：本计算书由 LPCalc 工具自动生成，仅供参考。正式设计需经专业审核。')
    run.font.name = '宋体'
    run.font.size = Pt(9)
    run.font.color.rgb = RGBColor(128, 128, 128)
    run._element.rPr.rFonts.set(qn('w:eastAsia'), '宋体')

    return doc


def export_calculation_report(lp_result, ep_result, building_params,
                              cable_params, c_factors, level_text, ep_level_text,
                              export_mode='combined', has_ep=True,
                              building_attr="", attr_type=""):
    """使用 python-docx-template 导出 Word（稳定版）"""
    from docx import Document as DocxDocument

    doc = DocxDocument()

    style = doc.styles['Normal']
    style.font.name = '宋体'
    style.font.size = Pt(10.5)
    style._element.rPr.rFonts.set(qn('w:eastAsia'), '宋体')

    # 主标题
    title_text = '建筑物防雷等级计算书' if not has_ep else '建筑物防雷及电子信息防护等级计算书'
    title = doc.add_heading(title_text, level=0)
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    set_heading_font(title, '黑体', 22, True, (0, 51, 102))
    doc.add_paragraph('_' * 50).alignment = WD_ALIGN_PARAGRAPH.CENTER
    doc.add_paragraph()

    # 基本信息
    build_basic_info_section(doc)

    # 输入参数
    build_input_params_section(doc, building_params)

    # 防雷计算
    build_lp_section(doc, lp_result, level_text)

    # 结论
    build_conclusion_section(doc, lp_result, level_text, ep_result,
                             ep_level_text, has_ep, building_attr, attr_type)

    file_buffer = io.BytesIO()
    doc.save(file_buffer)
    file_buffer.seek(0)
    return file_buffer


def export_separate_reports(lp_result, ep_result, building_params,
                            cable_params, c_factors, level_text, ep_level_text,
                            building_attr="", attr_type=""):
    """分开导出（防雷和电子防护各一个文件）"""
    from docx import Document as DocxDocument

    # ===== 防雷等级计算书 =====
    doc_lp = DocxDocument()
    style = doc_lp.styles['Normal']
    style.font.name = '宋体'
    style.font.size = Pt(10.5)
    style._element.rPr.rFonts.set(qn('w:eastAsia'), '宋体')

    title = doc_lp.add_heading('建筑物防雷等级计算书', level=0)
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    set_heading_font(title, '黑体', 22, True, (0, 51, 102))
    doc_lp.add_paragraph('_' * 50).alignment = WD_ALIGN_PARAGRAPH.CENTER
    doc_lp.add_paragraph()

    build_basic_info_section(doc_lp)
    build_input_params_section(doc_lp, building_params)
    build_lp_section(doc_lp, lp_result, level_text)

    heading = doc_lp.add_heading('四、结论', level=1)
    set_heading_font(heading, '黑体', 16, True, (0, 51, 102))
    p = doc_lp.add_paragraph()
    p.add_run('根据 GB 50057-2010 计算，该建筑物的防雷等级为：').bold = True
    run = p.add_run(f' {level_text}')
    run.bold = True
    run.font.color.rgb = RGBColor(0, 112, 192)
    run.font.size = Pt(12)

    doc_lp.add_paragraph()
    doc_lp.add_paragraph('_' * 50).alignment = WD_ALIGN_PARAGRAPH.CENTER
    p = doc_lp.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run('注：本计算书由 LPCalc 工具自动生成，仅供参考。正式设计需经专业审核。')
    run.font.name = '宋体'
    run.font.size = Pt(9)
    run.font.color.rgb = RGBColor(128, 128, 128)
    run._element.rPr.rFonts.set(qn('w:eastAsia'), '宋体')

    lp_buffer = io.BytesIO()
    doc_lp.save(lp_buffer)
    lp_buffer.seek(0)

    # ===== 电子防护等级计算书 =====
    doc_ep = DocxDocument()
    style = doc_ep.styles['Normal']
    style.font.name = '宋体'
    style.font.size = Pt(10.5)
    style._element.rPr.rFonts.set(qn('w:eastAsia'), '宋体')

    title = doc_ep.add_heading('电子信息系统雷电防护等级计算书', level=0)
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    set_heading_font(title, '黑体', 22, True, (0, 51, 102))
    doc_ep.add_paragraph('_' * 50).alignment = WD_ALIGN_PARAGRAPH.CENTER
    doc_ep.add_paragraph()

    # 基本信息（含电子防护规范）
    heading = doc_ep.add_heading('一、基本信息', level=1)
    set_heading_font(heading, '黑体', 16, True, (0, 51, 102))
    now = datetime.now()
    p = doc_ep.add_paragraph()
    p.add_run('计算日期：').bold = True
    p.add_run(now.strftime("%Y年%m月%d日"))
    p = doc_ep.add_paragraph()
    p.add_run('依据规范：').bold = True
    p.add_run('GB 50057-2010《建筑物防雷设计规范》')
    p = doc_ep.add_paragraph()
    p.add_run('          ').bold = True
    p.add_run('GB 50343-2012《建筑物电子信息系统防雷技术规范》')
    doc_ep.add_paragraph()

    build_input_params_section(doc_ep, building_params)

    # 电子防护等级计算
    heading = doc_ep.add_heading('三、电子信息系统雷电防护等级计算', level=1)
    set_heading_font(heading, '黑体', 16, True, (0, 51, 102))

    if ep_result:
        ep_steps = [
            ('建筑物年预计雷击次数 N1', f"{ep_result.get('N1', 0):.6f} 次/a"),
            ("电源线缆截收面积 Ae1'", f"{ep_result.get('Ae1', 0):.6f} km²"),
            ("信号线缆截收面积 Ae2'", f"{ep_result.get('Ae2', 0):.6f} km²"),
            ('入户设施年预计雷击次数 N2', f"{ep_result.get('N2', 0):.6f} 次/a"),
            ('总年预计雷击次数 N', f"{ep_result.get('N', 0):.6f} 次/a"),
            ('C = C1+C2+C3+C4+C5+C6', f"{ep_result.get('C', 0):.2f}"),
            ('可接受最大年雷击次数 Nc', f"{ep_result.get('Nc', 0):.6f} 次/a"),
            ('拦截效率 E', f"{ep_result.get('E', 0):.4f}"),
        ]

        table = doc_ep.add_table(rows=len(ep_steps) + 1, cols=2)
        table.style = 'Light Grid Accent 1'
        table.alignment = WD_TABLE_ALIGNMENT.CENTER
        table.columns[0].width = Cm(5)
        table.columns[1].width = Cm(10.5)

        hdr_cells = table.rows[0].cells
        hdr_cells[0].text = '计算项'
        hdr_cells[1].text = '结果'
        for cell in hdr_cells:
            cell.paragraphs[0].runs[0].font.name = '黑体'
            cell.paragraphs[0].runs[0].font.size = Pt(10.5)
            cell.paragraphs[0].runs[0].font.bold = True
            cell.paragraphs[0].runs[0]._element.rPr.rFonts.set(qn('w:eastAsia'), '黑体')
            set_cell_background(cell, 'D9E1F2')

        for i, (name, value) in enumerate(ep_steps, 1):
            row = table.rows[i]
            set_cell_font(row.cells[0], name, '宋体', 10.5, False)
            set_cell_font(row.cells[1], str(value), 'Arial', 10.5, False)

        doc_ep.add_paragraph()
        p = doc_ep.add_paragraph()
        p.add_run('▶ 电子信息系统防护等级判定结果：').bold = True
        run = p.add_run(ep_level_text)
        run.bold = True
        if 'A' in ep_level_text:
            run.font.color.rgb = RGBColor(192, 0, 0)
        elif 'B' in ep_level_text:
            run.font.color.rgb = RGBColor(255, 128, 0)
        elif 'C' in ep_level_text:
            run.font.color.rgb = RGBColor(0, 112, 192)
        else:
            run.font.color.rgb = RGBColor(0, 128, 0)
        run.font.size = Pt(12)
    else:
        doc_ep.add_paragraph('电子信息系统防护等级：未计算')

    heading = doc_ep.add_heading('四、结论', level=1)
    set_heading_font(heading, '黑体', 16, True, (0, 51, 102))
    if ep_result:
        p = doc_ep.add_paragraph()
        p.add_run('根据 GB 50343-2012 计算，该建筑物的电子信息系统雷电防护等级为：').bold = True
        run = p.add_run(f' {ep_level_text}')
        run.bold = True
        if 'A' in ep_level_text:
            run.font.color.rgb = RGBColor(192, 0, 0)
        elif 'B' in ep_level_text:
            run.font.color.rgb = RGBColor(255, 128, 0)
        elif 'C' in ep_level_text:
            run.font.color.rgb = RGBColor(0, 112, 192)
        else:
            run.font.color.rgb = RGBColor(0, 128, 0)
        run.font.size = Pt(12)
    else:
        doc_ep.add_paragraph('电子信息系统防护等级：未计算')

    doc_ep.add_paragraph()
    doc_ep.add_paragraph('_' * 50).alignment = WD_ALIGN_PARAGRAPH.CENTER
    p = doc_ep.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run('注：本计算书由 LPCalc 工具自动生成，仅供参考。正式设计需经专业审核。')
    run.font.name = '宋体'
    run.font.size = Pt(9)
    run.font.color.rgb = RGBColor(128, 128, 128)
    run._element.rPr.rFonts.set(qn('w:eastAsia'), '宋体')

    ep_buffer = io.BytesIO()
    doc_ep.save(ep_buffer)
    ep_buffer.seek(0)

    return lp_buffer, ep_buffer