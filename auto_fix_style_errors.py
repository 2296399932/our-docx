#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
样式错误自动修正工具
自动检测并修正文档中的样式错误，根据样式定义应用正确的格式
"""

import os
import json
import copy
import time

from docx_namespace import DocxElementParser
from compare_styles import compare_styles
from clean_style_statistics import clean_style_statistics


# def auto_fix_style_errors(doc_path, classification_path, style_mapping_path, api_params_path,
#                           output_path=None, save_statistics=True, fix_paragraph=True,
#                           fix_run=True, fix_table=True, interactive=False, clean_statistics=False):
def auto_fix_style_errors(doc_path, classification, style_mapping_path, api_params,
                              output_path=None, save_statistics=True, fix_paragraph=True,
                              fix_run=True, fix_table=True, interactive=False, clean_statistics=False):
    """
    自动检测并修正文档中的样式错误

    参数:
        doc_path: Word文档路径
        classification_path: 分类结果JSON文件路径
        style_mapping_path: 样式映射JSON文件路径
        api_params_path: API参数格式JSON文件路径
        output_path: 输出文档路径，默认为原文件名_fixed.docx
        save_statistics: 是否保存样式统计信息
        fix_paragraph: 是否修复段落样式
        fix_run: 是否修复run样式
        fix_table: 是否修复表格样式
        interactive: 是否交互式修复（每个错误询问是否修复）
        clean_statistics: 是否清理统计结果，减少重复处理，默认为False

    返回:
        str: 输出文件路径
    """
    # 如果未指定输出路径，则生成默认输出路径
    if output_path is None:
        file_name, file_ext = os.path.splitext(doc_path)
        output_path = f"{file_name}_fixed{file_ext}"

    # 执行样式比较并获取结果
    print(f"正在分析文档样式错误: {doc_path}")
    statistics = compare_styles(doc_path, classification, style_mapping_path, api_params)

    # 判断是否需要清理统计结果
    if clean_statistics:
        print("清理样式错误统计结果，优化修复流程...")
        cleaned_statistics = clean_style_statistics(statistics)
        print(f"清理前错误数: {len(statistics['elements'])}, 清理后错误数: {len(cleaned_statistics['elements'])}")
    else:
        print(f"直接处理所有样式错误 (共 {len(statistics['elements'])} 个)")
        cleaned_statistics = statistics



    # # 加载API参数文件，用于获取目标样式值
    # with open(api_params_path, 'r', encoding='utf-8') as f:
    #     api_params = json.load(f)

    # 创建DocxElementParser实例用于修改文档
    doc = DocxElementParser(doc_path)



    # 记录修复的错误数量
    fixed_errors = {
        "paragraph": 0,
        "run": 0,
        "table": 0,
        "total": 0,
        "skipped": 0
    }

    # 自动修复样式错误
    print("\n开始自动修复样式错误...")
    if 'elements' in cleaned_statistics:
        for element_info in cleaned_statistics['elements']:
            element_type = element_info.get('type')
            element = element_info.get('element')
            result = element_info.get('result', [])

            if not result:  # 跳过没有错误的元素
                continue

            # 根据元素类型修复样式
            if element_type == 'paragraph' and fix_paragraph:
                # 修复段落样式
                try:
                    para_index = element_info.get('index')

                    para_index_ = doc.get_paragraph_index_from_element_index(para_index)
                    if para_index is not None:
                        # 收集要应用的样式属性
                        style_props = extract_style_properties(result, 'paragraph')
                        print(f"提取的样式属性: {style_props}")

                        # 应用样式属性
                        if style_props:
                            doc.update_paragraph_style(para_index_, **style_props)

                            fixed_errors['paragraph'] += 1
                            fixed_errors['total'] += 1
                            print(doc.get_paragraph_style_from_element(doc.paragraphs[para_index_]['element']))
                            print(f"已修复段落 {para_index} 的样式错误")

                except Exception as e:
                    print(f"修复段落错误时出错: {e}")

            elif element_type == 'run' and fix_run:
                # 修复run样式
                try:
                    run_indices = element_info.get('index')
                    if run_indices and len(run_indices) == 2:
                        para_index, run_index = run_indices
                        para_index = doc.get_paragraph_index_from_element_index(para_index)

                        # 收集要应用的样式属性
                        style_props = extract_style_properties(result, 'run')

                        # 应用样式属性
                        if style_props:
                            print(f"准备修复段落 {para_index} 中的Run {run_index} 的样式")
                            print(f"修复前的样式: {doc.get_run_style(para_index, run_index)}")
                            print(f"应用的样式属性: {style_props}")

                            # 关键修改：确保字体属性名称正确 (从'font'改为'fonts')
                            if 'font' in style_props:
                                style_props['fonts'] = style_props.pop('font')

                            # 直接调用update_run_style，无需分别处理各属性
                            doc.update_run_style(para_index, run_index, **style_props)

                            print(f"修复后的样式: {doc.get_run_style(para_index, run_index)}")
                            fixed_errors['run'] += 1
                            fixed_errors['total'] += 1
                            print(f"已修复段落 {para_index} 中的Run {run_index} 的样式错误")
                except Exception as e:
                    print(f"修复Run错误时出错: {e}")

            elif element_type == 'table' and fix_table:
                # 修复表格样式
                try:
                    table_index = element_info.get('index')
                    if table_index is not None:
                        # 收集要应用的表格样式属性
                        table_style_props = extract_style_properties(result, 'table')

                        # 收集要应用的文本样式属性
                        table_text_style_props = extract_table_text_style_properties(result)

                        print(f"表格 {table_index} 样式属性: {table_style_props}")
                        print(f"表格 {table_index} 文本样式属性: {table_text_style_props}")

                        # 获取表格尺寸，以便分别处理表头和数据行
                        rows, cols = doc.get_table_dimensions(table_index)
                        print(f"表格尺寸: {rows}行 x {cols}列")

                        # 判断表格类型（是否为三线表）
                        if 'is_three_line_table' in table_style_props:
                            if table_style_props['is_three_line_table']:
                                doc.create_three_line_table(table_index)
                                print(f"已将表格 {table_index} 设置为三线表")
                            del table_style_props['is_three_line_table']

                        # 处理表格宽度
                        if 'table_width' in table_style_props:
                            width_info = table_style_props.pop('table_width')
                            doc.set_table_width(table_index, width=width_info['width'], width_type=width_info['width_type'])
                            print(f"已设置表格 {table_index} 的宽度为 {width_info['width']} ({width_info['width_type']})")

                        # 处理表格对齐方式
                        if 'table_alignment' in table_style_props:
                            alignment = table_style_props.pop('table_alignment')
                            # 直接使用set_table_alignment函数
                            doc.set_table_alignment(table_index, alignment)
                            print(f"已设置表格 {table_index} 的对齐方式为 {alignment}")

                        # 应用其他表格样式属性
                        if table_style_props:
                            # 移除可能导致问题的属性
                            # safe_style_props = {k: v for k, v in table_style_props.items()
                            #                   if k not in ['width', 'alignment']}
                            if table_style_props:
                                try:
                                    doc.set_table_style(table_index, **table_style_props)
                                    print(f"已应用表格 {table_index} 的基本样式: {table_style_props}")
                                except Exception as e:
                                    print(f"应用表格基本样式时出错: {e}")

                        # 应用表格文本样式 - 根据header_style和data_style分别处理
                        if table_text_style_props and rows >= 2:
                            try:
                                # 分别处理表头样式和数据行样式
                                header_style = table_text_style_props.get('column_header', {})
                                data_style = table_text_style_props.get('data_row', {})

                                # 处理font/fonts命名问题
                                if 'font' in header_style:
                                    header_style['fonts'] = header_style.pop('font')
                                if 'font' in data_style:
                                    data_style['fonts'] = data_style.pop('font')
                                if 'font' in table_text_style_props:
                                    table_text_style_props['fonts'] = table_text_style_props.pop('font')

                                # 设置表头样式（第一行）
                                if header_style:
                                    for col in range(cols):
                                        cell_paragraphs = doc.get_table_cell_paragraphs(table_index, 0, col)
                                        for para_element in cell_paragraphs:
                                            doc.set_paragraph_alignment_from_xml(para_element, 'center')
                                            # 直接使用update_runs_style_from_xml设置段落和文本样式
                                            doc.update_runs_style_from_xml(para_element, **header_style)

                                    print(f"已设置表格 {table_index} 的表头样式")

                                # 设置数据行样式（第二行及以后）
                                if data_style:
                                    for row in range(1, rows):
                                        for col in range(cols):
                                            cell_paragraphs = doc.get_table_cell_paragraphs(table_index, row, col)

                                            for para_element in cell_paragraphs:
                                                doc.set_paragraph_alignment_from_xml(para_element,'center')
                                                # 直接使用update_runs_style_from_xml设置段落和文本样式
                                                print(f"应用数据行样式到段落: {data_style}")
                                                doc.update_runs_style_from_xml(para_element, **data_style)

                                    # 更新文档XML以应用所有更改
                                    doc.update_document_xml()

                                    print(f"已设置表格 {table_index} 的数据行样式")
                            except Exception as e:
                                print(f"分别设置表头和数据行样式时出错: {e}")
                                # 如果分别设置失败，尝试使用update_table_text_style进行整体设置
                                try:
                                    # 移除特殊键
                                    common_style = {k: v for k, v in table_text_style_props.items()
                                                if k not in ['column_header', 'data_row']}
                                    if common_style:
                                        doc.update_table_text_style(table_index, **common_style)
                                        print(f"已通过update_table_text_style应用表格 {table_index} 的通用文本样式")
                                except Exception as e2:
                                    print(f"通过update_table_text_style设置样式时出错: {e2}")
                        elif table_text_style_props:
                            # 移除特殊键，应用整体样式
                            common_style = {k: v for k, v in table_text_style_props.items()
                                         if k not in ['column_header', 'data_row']}
                            if common_style:
                                try:
                                    doc.update_table_text_style(table_index, **common_style)
                                    print(f"已应用表格 {table_index} 的通用文本样式")
                                except Exception as e:
                                    print(f"应用表格文本样式时出错: {e}")

                        fixed_errors['table'] += 1
                        fixed_errors['total'] += 1
                        print(f"已修复表格 {table_index} 的样式错误")
                except Exception as e:
                    print(f"修复表格错误时出错: {e}")
                    import traceback
                    traceback.print_exc()

    # 保存统计信息
    if save_statistics:
        output_dir = os.path.dirname(output_path) if os.path.dirname(output_path) else "."
        output_filename = os.path.basename(doc_path).split('.')[0]
        statistics_path = os.path.join(output_dir, f"{output_filename}_fix_statistics.json")

        # 创建修复统计信息
        fix_statistics = {
            "original_errors": len(statistics['elements']),
            "cleaned_errors": len(cleaned_statistics['elements']),
            "fixed_errors": fixed_errors,
            "remaining_errors": len(cleaned_statistics['elements']) - fixed_errors['total'] - fixed_errors['skipped']
        }

        with open(statistics_path, 'w', encoding='utf-8') as f:
            json.dump(fix_statistics, f, ensure_ascii=False, indent=2)
        print(f"修复统计信息已保存到: {statistics_path}")

    # 保存前处理正文：删除空行并将文字改为黑色
    print("正在处理字体颜色...")
    para_count = doc.get_paragraphs_length()
    removed_count = 0
    for i in reversed(range(para_count)):

            # 设置字体颜色为黑色
            run_count = doc.get_run_count_from_xml(doc.paragraphs[i]['element'])
            for run_idx in range(run_count):
                doc.set_run_color(i, run_idx, color="000000")
    print(f"已删除空行 {removed_count} 行，并将正文字体颜色设置为黑色")
    # 插入分页符
    body_heading_indices = []
    if isinstance(classification, dict) and "body_heading_level_1" in classification:
        body_heading_indices = classification["body_heading_level_1"]

    if body_heading_indices:
        print(f"为 {len(body_heading_indices)} 个一级标题插入分页符...")
        for idx in body_heading_indices:
            try:
                para_index = doc.get_paragraph_index_from_element_index(idx)
                doc.insert_page_break_before_paragraph(para_index)
                print(f"已在段落 {para_index}（element_index={idx}）前插入分页符")
            except Exception as e:
                print(f"插入分页符时出错（element_index={idx}）: {e}")

        # 删除空行段落（从第一个一级标题开始，向后遍历）
        first_heading_para_index = doc.get_paragraph_index_from_element_index(body_heading_indices[0])
        para_count = doc.get_paragraphs_length()
        removed_count = 0
        # 倒序遍历，防止删除时索引错乱
        for i in reversed(range(first_heading_para_index, para_count)):
            para_text = doc.get_paragraph_text(doc.paragraphs[i]['element'])
            if not para_text.strip():
                doc.remove_paragraph(i)
                removed_count += 1
        print(f"已删除从第一个一级标题开始后的空行段落 {removed_count} 行")
    # 保存修改后的文档
    doc.save(output_path)
    print(f"\n文档样式错误修复完成并已保存至: {output_path}")
    print(f"共修复了 {fixed_errors['total']} 处样式错误:")
    print(f"  - 段落错误: {fixed_errors['paragraph']}")
    print(f"  - Run错误: {fixed_errors['run']}")
    print(f"  - 表格错误: {fixed_errors['table']}")
    if interactive:
        print(f"  - 跳过错误: {fixed_errors['skipped']}")
    print(f"  - 剩余错误: {len(cleaned_statistics['elements']) - fixed_errors['total'] - fixed_errors['skipped']}")

    return output_path


def extract_style_properties(result, element_type):
    """
    从样式比较结果中提取样式属性和值

    参数:
        result: 样式比较结果
        element_type: 元素类型 ('paragraph', 'run', 'table')

    返回:
        dict: 样式属性和值的字典
    """
    style_props = {}

    for error_dict in result:
        for attr, values in error_dict.items():
            # 获取正确的样式值 (检查success或scuccess字段)
            correct_value = None
            if 'success' in values:
                correct_value = values['success']
            elif 'scuccess' in values:
                correct_value = values['scuccess']

            if correct_value is not None:
                # 根据属性类型设置样式属性
                if element_type == 'paragraph':
                    if attr == 'alignment':
                        style_props['alignment'] = correct_value
                    elif attr == 'first_line':
                        if 'indentation' not in style_props:
                            style_props['indentation'] = {}
                        # 确保值为整数
                        try:
                            style_props['indentation']['firstLine'] = int(correct_value)
                        except (ValueError, TypeError):
                            style_props['indentation']['firstLine'] = correct_value
                    elif attr == 'hanging':
                        if 'indentation' not in style_props:
                            style_props['indentation'] = {}
                        try:
                            style_props['indentation']['hanging'] = int(correct_value)
                        except (ValueError, TypeError):
                            style_props['indentation']['hanging'] = correct_value
                    elif attr == 'before':
                        if 'spacing' not in style_props:
                            style_props['spacing'] = {}
                        try:
                            style_props['spacing']['before'] = int(correct_value)
                        except (ValueError, TypeError):
                            style_props['spacing']['before'] = correct_value
                    elif attr == 'after':
                        if 'spacing' not in style_props:
                            style_props['spacing'] = {}
                        try:
                            style_props['spacing']['after'] = int(correct_value)
                        except (ValueError, TypeError):
                            style_props['spacing']['after'] = correct_value
                    elif attr == 'beforeLines':
                        if 'spacing' not in style_props:
                            style_props['spacing'] = {}
                        try:
                            style_props['spacing']['beforeLines'] = int(correct_value)
                        except (ValueError, TypeError):
                            style_props['spacing']['beforeLines'] = correct_value
                    elif attr == 'afterLines':
                        if 'spacing' not in style_props:
                            style_props['spacing'] = {}
                        try:
                            style_props['spacing']['afterLines'] = int(correct_value)
                        except (ValueError, TypeError):
                            style_props['spacing']['afterLines'] = correct_value
                    elif attr == 'line':
                        if 'spacing' not in style_props:
                            style_props['spacing'] = {}
                        try:
                            style_props['spacing']['line'] = int(correct_value)
                        except (ValueError, TypeError):
                            style_props['spacing']['line'] = correct_value
                    elif attr == 'line_rule':
                        if 'spacing' not in style_props:
                            style_props['spacing'] = {}
                        style_props['spacing']['lineRule'] = correct_value
                    elif attr == 'font_chinese':
                        if 'font' not in style_props:
                            style_props['font'] = {}
                        style_props['font']['eastAsia'] = correct_value
                    elif attr == 'font_ascii':
                        if 'font' not in style_props:
                            style_props['font'] = {}
                        style_props['font']['ascii'] = correct_value
                        style_props['font']['hAnsi'] = correct_value
                    elif attr == 'size':
                        # 确保size为整数
                        try:
                            style_props['size'] = int(correct_value)
                        except (ValueError, TypeError):
                            style_props['size'] = correct_value
                    elif attr == 'bold':
                        # 确保bold为布尔值
                        if isinstance(correct_value, str):
                            style_props['bold'] = correct_value.lower() == 'true'
                        else:
                            style_props['bold'] = bool(correct_value)

                elif element_type == 'run':
                    if attr == 'font_chinese':
                        if 'font' not in style_props:
                            style_props['font'] = {}
                        style_props['font']['eastAsia'] = correct_value
                    elif attr == 'font_ascii':
                        if 'font' not in style_props:
                            style_props['font'] = {}
                        style_props['font']['ascii'] = correct_value
                        style_props['font']['hAnsi'] = correct_value
                    elif attr == 'size':
                        # 确保size为整数
                        try:
                            style_props['size'] = int(correct_value)
                        except (ValueError, TypeError):
                            style_props['size'] = correct_value
                    elif attr == 'bold':
                        # 确保bold为布尔值
                        if isinstance(correct_value, str):
                            style_props['bold'] = correct_value.lower() == 'true'
                        else:
                            style_props['bold'] = bool(correct_value)
                    # 可以根据需要添加更多run属性

                elif element_type == 'table':
                    if attr == 'is_three_line_table':
                        # 确保is_three_line_table为布尔值
                        if isinstance(correct_value, str):
                            style_props['is_three_line_table'] = correct_value.lower() == 'true'
                        else:
                            style_props['is_three_line_table'] = bool(correct_value)
                    elif attr == 'width':
                        # 表格宽度单独处理
                        try:
                            style_props['table_width'] = {
                                'width': int(correct_value),
                                'width_type': 'dxa'  # 默认使用dxa单位
                            }
                        except (ValueError, TypeError):
                            style_props['table_width'] = {
                                'width': correct_value,
                                'width_type': 'dxa'
                            }
                    elif attr == 'alignment' or attr == 'table_alignment':
                        # 表格对齐方式单独处理
                        style_props['table_alignment'] = correct_value
                    # 其他表格基本属性可以在这里添加

    # 调试输出
    print(f"提取的{element_type}样式属性: {style_props}")

    return style_props


def extract_table_text_style_properties(result):
    """
    从样式比较结果中提取表格文本样式属性

    参数:
        result: 样式比较结果，格式为列表，每个元素是包含attribute、scuccess、error等键的字典

    返回:
        dict: 表格文本样式属性字典，包括header_style和data_style
    """
    text_style_props = {}
    column_header_props = {}
    data_row_props = {}
    
    # 打印输入数据结构，帮助调试
    # print(f"输入数据结构: {result}")
    
    try:
        # 遍历每个样式比较结果项
        for item in result:
            # 检查是否是表头(column_header)相关的样式
            if 'column_header' in item:
                attr_info = item['column_header']
                attr_name = attr_info.get('attribute')
                correct_value = attr_info.get('scuccess')
                
                if attr_name and correct_value is not None:
                    if attr_name == 'font_chinese':
                        if 'fonts' not in column_header_props:
                            column_header_props['fonts'] = {}
                        column_header_props['fonts']['eastAsia'] = correct_value
                    elif attr_name == 'font_ascii':
                        if 'fonts' not in column_header_props:
                            column_header_props['fonts'] = {}
                        column_header_props['fonts']['ascii'] = correct_value
                        column_header_props['fonts']['hAnsi'] = correct_value
                    elif attr_name == 'size':
                        try:
                            column_header_props['size'] = int(correct_value)
                        except (ValueError, TypeError):
                            column_header_props['size'] = correct_value
                    elif attr_name == 'bold':
                        if isinstance(correct_value, str):
                            column_header_props['bold'] = correct_value.lower() == 'true'
                        else:
                            column_header_props['bold'] = bool(correct_value)
                    elif attr_name == 'alignment':
                        column_header_props['alignment'] = correct_value
                    # 可以根据需要添加更多属性
            
            # 检查是否是数据行(data_row)相关的样式
            elif 'data_row' in item:
                attr_info = item['data_row']
                attr_name = attr_info.get('attribute')
                correct_value = attr_info.get('scuccess')
                
                if attr_name and correct_value is not None:
                    if attr_name == 'font_chinese':
                        if 'fonts' not in data_row_props:
                            data_row_props['fonts'] = {}
                        data_row_props['fonts']['eastAsia'] = correct_value
                    elif attr_name == 'font_ascii':
                        if 'fonts' not in data_row_props:
                            data_row_props['fonts'] = {}
                        data_row_props['fonts']['ascii'] = correct_value
                        data_row_props['fonts']['hAnsi'] = correct_value
                    elif attr_name == 'size':
                        try:
                            data_row_props['size'] = int(correct_value)
                        except (ValueError, TypeError):
                            data_row_props['size'] = correct_value
                    elif attr_name == 'bold':
                        if isinstance(correct_value, str):
                            data_row_props['bold'] = correct_value.lower() == 'true'
                        else:
                            data_row_props['bold'] = bool(correct_value)
                    elif attr_name == 'alignment':
                        data_row_props['alignment'] = correct_value
                    # 可以根据需要添加更多属性
    except Exception as e:
        print(f"提取表格文本样式属性时出错: {e}")
        import traceback
        traceback.print_exc()
    
    # 将表头和数据行属性添加到结果中
    if column_header_props:
        text_style_props['column_header'] = column_header_props
    if data_row_props:
        text_style_props['data_row'] = data_row_props
    
    # 打印提取的结果，帮助调试
    print(f"提取的表格文本样式属性: {text_style_props}")
    
    return text_style_props


def get_readable_attribute_name(attr):
    """转换属性名称为易读形式"""
    attribute_names = {
        'style_id': '样式ID',
        'font_name': '字体名称',
        'font_size': '字号大小',
        'bold': '粗体',
        'italic': '斜体',
        'underline': '下划线',
        'color': '颜色',
        'highlight': '高亮',
        'alignment': '对齐方式',
        'indentation': '缩进',
        'spacing': '间距',
        'line_spacing': '行间距',
        'border': '边框',
        'shading': '底纹',
        'width': '宽度',
        'column_header': '表头样式',
        'data_row': '数据行样式',
        'table_style': '表格样式',
        'font_chinese': '中文字体',
        'font_ascii': '英文字体',
        'size': '字号大小',
        'before': '段前间距',
        'after': '段后间距',
        'line': '行间距',
        'line_rule': '行距规则',
        'first_line': '首行缩进',
        'hanging': '悬挂缩进',
        'is_three_line_table': '三线表设置',
    }

    # 处理复合属性名（如 column_header.font_chinese）
    if '.' in attr:
        parts = attr.split('.')
        if len(parts) == 2 and parts[0] in attribute_names and parts[1] in attribute_names:
            return f"{attribute_names[parts[0]]}的{attribute_names[parts[1]]}"

    return attribute_names.get(attr, attr)


def fix_specific_elements(doc_path, elements_to_fix, style_mapping_path, api_params_path, output_path=None):
    """
    修复指定元素的样式错误

    参数:
        doc_path: Word文档路径
        elements_to_fix: 要修复的元素列表，格式为 [{'type': 'paragraph', 'index': 10, 'properties': {...}}, ...]
        style_mapping_path: 样式映射JSON文件路径
        api_params_path: API参数格式JSON文件路径
        output_path: 输出文档路径，默认为原文件名_fixed.docx

    返回:
        str: 输出文件路径
    """
    # 如果未指定输出路径，则生成默认输出路径
    if output_path is None:
        file_name, file_ext = os.path.splitext(doc_path)
        output_path = f"{file_name}_fixed{file_ext}"

    # 创建DocxElementParser实例用于修改文档
    doc = DocxElementParser(doc_path)

    # 记录修复的错误数量
    fixed_count = 0

    # 修复每个指定的元素
    for element_info in elements_to_fix:
        element_type = element_info.get('type')
        element_index = element_info.get('index')
        properties = element_info.get('properties', {})

        if element_type == 'paragraph':
            try:
                para_index=doc.get_paragraph_index_from_element_index(element_index)
                doc.update_paragraph_style(para_index, **properties)
                fixed_count += 1
                print(f"已修复段落 {element_index} 的样式")
            except Exception as e:
                print(f"修复段落 {element_index} 时出错: {e}")

        elif element_type == 'run':
            try:
                para_index, run_index = element_index
                para_index_=doc.get_paragraph_index_from_element_index(para_index)

                doc.update_run_style(para_index_, run_index, **properties)
                fixed_count += 1
                print(f"已修复段落 {para_index} 中的Run {run_index} 的样式")
            except Exception as e:
                print(f"修复Run {element_index} 时出错: {e}")

        elif element_type == 'table':
            try:
                # 提取表格样式属性
                table_props = {k: v for k, v in properties.items()
                             if k not in ['text_style', 'is_three_line_table']}
                print(table_props)
                # 提取表格文本样式属性
                text_style = properties.get('text_style', {})

                # 应用表格样式
                if table_props:
                    doc.set_table_style(element_index, **table_props)

                # 应用表格文本样式
                if text_style:

                    doc.update_table_text_style(element_index, **text_style)

                # 设置三线表
                if properties.get('is_three_line_table', False):
                    doc.create_three_line_table(element_index)

                fixed_count += 1
                print(f"已修复表格 {element_index} 的样式")
            except Exception as e:
                print(f"修复表格 {element_index} 时出错: {e}")

    # 保存修改后的文档
    doc.save(output_path)
    print(f"\n文档样式错误修复完成并已保存至: {output_path}")
    print(f"共修复了 {fixed_count} 个元素的样式")

    return output_path


def test_table_style_fix(doc_path, output_path=None):
    """
    测试表格样式修复功能

    Args:
        doc_path (str): 文档路径
        output_path (str, optional): 输出路径。默认为None，将生成默认输出路径。

    Returns:
        str: 处理后的文档路径
    """
    if output_path is None:
        base_name = os.path.basename(doc_path)
        file_name, ext = os.path.splitext(base_name)
        output_path = os.path.join(os.path.dirname(doc_path), f"{file_name}_table_test{ext}")

    doc = DocxElementParser(doc_path)

    # 获取表格数量
    table_count = len(doc.tables)
    print(f"文档中包含 {table_count} 个表格")

    if table_count == 0:
        print("没有表格可以处理")
        return doc_path

    # 表格样式设置
    table_width = "16.0cm"  # 设置表格宽度
    table_alignment = "center"  # 设置表格对齐方式

    # 文本样式设置
    column_header_style = {
        "fonts": {"ascii": "Times New Roman", "eastAsia": "宋体"},  # 设置表头字体
        "size": 21,  # 设置表头字号
        "bold": True,  # 设置表头加粗
        "alignment": "center"  # 设置表头对齐方式
    }

    data_row_style = {
        "fonts": {"ascii": "Times New Roman", "eastAsia": "宋体"},  # 设置数据行字体
        "size": 21,  # 设置数据行字号
        "bold": False,  # 设置数据行不加粗
        "alignment": "center"  # 设置数据行对齐方式
    }

    # 获取第一个表格进行处理
    table_index = 0
    table = doc.tables[table_index]

    print(f"\n处理表格 {table_index + 1}:")

    # 获取表格尺寸
    rows, cols = doc.get_table_dimensions(table_index)
    print(f"表格尺寸: {rows}行 x {cols}列")

    # 设置表格宽度和对齐方式
    try:
        doc.set_table_width(table_index, table_width, width_type='auto')
        print(f"设置表格宽度为 {table_width}")
    except Exception as e:
        print(f"设置表格宽度失败: {e}")

    try:
        doc.set_table_alignment(table_index, table_alignment)
        print(f"设置表格对齐方式为 {table_alignment}")
    except Exception as e:
        print(f"设置表格对齐方式失败: {e}")

    # 根据行列关系获取单元格段落并设置样式
    for row_idx in range(rows):
        for col_idx in range(cols):
            try:
                # 获取单元格中的段落元素
                cell_paragraphs = doc.get_table_cell_paragraphs(table_index, row_idx, col_idx)

                if cell_paragraphs:
                    for para_element in cell_paragraphs:
                        # 根据是否是表头行应用不同的样式
                        if row_idx == 0:  # 表头行
                            print(f"应用表头样式到单元格 ({row_idx}, {col_idx})")
                            # 直接使用update_runs_style_from_xml设置段落和文本样式
                            doc.update_runs_style_from_xml(para_element, **column_header_style)
                        else:  # 数据行
                            print(f"应用数据行样式到单元格 ({row_idx}, {col_idx})")
                            # 直接使用update_runs_style_from_xml设置段落和文本样式
                            doc.update_runs_style_from_xml(para_element, **data_row_style)
            except Exception as e:
                print(f"设置单元格 ({row_idx}, {col_idx}) 样式失败: {e}")

    # 创建三线表
    try:
        doc.create_three_line_table(table_index)
        print("创建三线表成功")
    except Exception as e:
        print(f"创建三线表失败: {e}")

    # 更新文档XML
    doc.update_document_xml()

    # 保存文档
    doc.save(output_path)
    print(f"已将修改后的文档保存到: {output_path}")

    return output_path


if __name__ == "__main__":
    # 设置文件路径 - 根据实际情况修改这些路径
    doc_path = "sdj-毕业论文(1).docx"  # 要处理的Word文档路径
    classification_path = "document_classification_results.json"  # 文档分类结果路径
    style_mapping_path = "document_style_mapping.json"  # 样式映射文件路径
    api_params_path = "智算工程学院毕业设计（论文）模板2025届(1)_api_params.json"  # API参数文件路径
    output_path = "1_fixed.docx"  # 输出文件路径，可以留空使用默认路径
    with open(classification_path, 'r', encoding='utf-8') as f:
        classification = json.load(f)

    # 加载样式映射
    with open(style_mapping_path, 'r', encoding='utf-8') as f:
        style_mapping = json.load(f)

    # 加载API参数
    with open(api_params_path, 'r', encoding='utf-8') as f:
        api_params = json.load(f)
    fixed_file = auto_fix_style_errors(
                doc_path,
                classification,
                style_mapping_path,
                api_params,
                output_path,
                interactive=False,
                clean_statistics=True  # 默认进行清理以提高效率
            )