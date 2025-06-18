#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
样式错误标记工具 - 直接使用文件路径版本
根据样式比较结果在文档中标记出样式错误并添加批注
"""

import os
import json
from docx_namespace import DocxElementParser
from compare_styles import compare_styles
from collections import defaultdict
# 导入清理统计结果的函数
from clean_style_statistics import clean_style_statistics


def mark_style_errors(doc_path, classification_path, style_mapping_path, api_params_path, output_path=None, save_statistics=True, skip_cleaning=False):
    """
    根据样式比较结果，在文档中标记出样式错误并添加批注

    参数:
        doc_path: Word文档路径
        classification_path: 分类结果JSON文件路径
        style_mapping_path: 样式映射JSON文件路径
        api_params_path: API参数格式JSON文件路径
        output_path: 输出文档路径，默认为原文件名_marked.docx
        save_statistics: 是否保存样式统计信息
        skip_cleaning: 跳过统计结果清理步骤，保留所有错误

    返回:
        str: 输出文件路径
    """
    # 如果未指定输出路径，则生成默认输出路径
    if output_path is None:
        file_name, file_ext = os.path.splitext(doc_path)
        output_path = f"{file_name}_marked{file_ext}"

    # 执行样式比较并获取结果
    print(f"正在分析文档样式错误: {doc_path}")
    statistics = compare_styles(doc_path, classification_path, style_mapping_path, api_params_path)

    # 创建DocxElementParser实例用于修改文档和查询原始信息
    doc = DocxElementParser(doc_path)

    # 进行一次预处理，移除空结果
    if 'elements' in statistics:
        filtered_elements = []
        
        for elem in statistics['elements']:
            element_type = elem.get('type')
            if not elem.get('result'):
                continue  # 跳过没有结果的元素
                
            # 对于run元素进行特殊处理
            if element_type == 'run' and 'index' in elem and isinstance(elem['index'], tuple) and len(elem['index']) == 2:
                para_idx, run_idx = elem['index']
                
                # 获取run文本
                run_text = doc._get_run_text(para_idx, run_idx)
                run_text_stripped = run_text.strip()
                
                # 空run不需要处理
                if not run_text_stripped:
                    continue
                
                # 检查是否是标点符号
                is_punctuation = len(run_text_stripped) <= 1 and all(c in ',.!?;:"\'()[]{}""''：；，。！？、（）【】《》' for c in run_text_stripped)
                
                # 检查结果中是否有真正的错误(带error字段)
                has_real_error = False
                error_list = []
                
                for result_item in elem.get('result', []):
                    new_result = {}
                    for attr, values in result_item.items():
                        # 仅当有error字段时视为真正的错误
                        if 'error' in values:
                            has_real_error = True
                            new_result[attr] = values
                    
                    if new_result:
                        error_list.append(new_result)
                
                # 如果是标点符号并且没有真正的错误，跳过
                if is_punctuation and not has_real_error:
                    continue
                
                # 对于非标点的run，即使没有error，也检查是否有缺失的关键属性
                if not has_real_error and not is_punctuation:
                    # 检查是否有size或其他重要属性缺失
                    critical_missing = False
                    for result_item in elem.get('result', []):
                        for attr, values in result_item.items():
                            if attr == 'size' and 'error' not in values and 'success' in values:
                                critical_missing = True
                                break
                    
                    # 如果没有关键缺失属性，跳过
                    if not critical_missing:
                        continue
                
                # 更新结果为已过滤的错误列表
                if has_real_error:
                    elem['result'] = error_list
                
                # 添加到过滤后的元素列表
                filtered_elements.append(elem)
            else:
                # 非run元素直接添加
                filtered_elements.append(elem)
        
        # 更新元素列表
        statistics['elements'] = filtered_elements

    # 在标记错误前清理统计结果，减少重复批注
    if not skip_cleaning:
        print("清理样式错误统计结果，减少重复批注...")
        cleaned_statistics = clean_style_statistics(statistics)
        print(f"清理前错误数: {len(statistics.get('elements', []))}, 清理后错误数: {len(cleaned_statistics.get('elements', []))}")
    else:
        print("跳过统计结果清理步骤，保留所有错误...")
        cleaned_statistics = statistics

    # 加载样式映射文件，用于提取正确的样式值
    with open(style_mapping_path, 'r', encoding='utf-8') as f:
        style_mapping = json.load(f)

    # 记录标记的错误数量
    marked_errors = {
        "paragraph": 0,
        "run": 0,
        "table": 0,
        "total": 0
    }

    # 标记段落和Run的样式错误
    print("\n开始标记样式错误并添加批注...")
    if 'elements' in cleaned_statistics:
        for element_info in cleaned_statistics['elements']:
            element_type = element_info.get('type')
            element = element_info.get('element')
            result = element_info.get('result', [])

            # 跳过没有错误的元素
            if not result:
                continue

            # 双重验证是否需要标记
            if element_type == 'run' and 'index' in element_info:
                para_idx, run_idx = element_info['index']
                
                # 再次获取run文本
                run_text = doc._get_run_text(para_idx, run_idx)
                
                # 对于空run或只有空格的run，跳过
                if not run_text.strip():
                    print(f"跳过空run: 段落 {para_idx} 中的Run {run_idx}")
                    continue
                
                # 对于标点符号，验证是否真的需要标记
                if len(run_text.strip()) <= 1 and all(c in ',.!?;:"\'()[]{}""''：；，。！？、（）【】《》' for c in run_text.strip()):
                    # 检查是否有真正的错误（非缺失属性）
                    has_critical_error = False
                    for error_dict in result:
                        for attr, values in error_dict.items():
                            if 'error' in values and attr not in ['size']:
                                has_critical_error = True
                                break
                    
                    if not has_critical_error:
                        print(f"跳过标点符号run: 段落 {para_idx} 中的Run {run_idx} ({run_text})")
                        continue
            
            # 确定是错误还是缺失（决定标记颜色）
            has_error = any('error' in values for error_dict in result for attr, values in error_dict.items())
            highlight_color = 'red' if has_error else 'yellow'

            # 构建批注内容
            comment_text = generate_comment_text(result, style_mapping, element_type)
            
            # 如果批注内容只包含标题，没有实际错误内容，跳过
            if comment_text.count('\n') <= 1:
                continue

            # 处理段落元素
            if element_type == 'paragraph':
                # 标记段落错误
                try:
                    # 直接使用索引定位段落
                    para_index = element_info.get('index')
                    if para_index is not None:
                        # 获取段落中的所有run数量
                        run_count = doc.get_run_count(para_index)
                        print(f"段落 {para_index} 中的Run数量: {run_count}")
                        para_index_=doc.get_paragraph_index_from_element_index(para_index)
                        # 为段落中的所有run添加高亮
                        for run_idx in range(run_count):
                            doc.set_run_highlight(para_index_, run_idx, highlight_color)
                            print(1111111)

                        # 为段落添加批注
                        doc.add_comment(
                            element_index=para_index,
                            author="样式检查器",
                            comment_text=comment_text,
                            element_type="paragraph"
                        )
                        print(2222222)
                        marked_errors['paragraph'] += 1
                        marked_errors['total'] += 1
                        print(f"已标记段落 {para_index} 的样式错误并添加批注 (颜色: {highlight_color})")
                except Exception as e:
                    print(f"标记段落错误时出错: {e}")

            # 处理run元素
            elif element_type == 'run':
                try:
                    # 直接使用索引定位run
                    run_indices = element_info.get('index')
                    if run_indices and len(run_indices) == 2:
                        para_index, run_index = run_indices
                        para_index_=doc.get_paragraph_index_from_element_index(para_index)

                        # 再次验证run文本，确保不是空run
                        run_text = doc._get_run_text(para_index, run_index)

                        run_text_stripped = run_text.strip()
                        
                        if not run_text_stripped:
                            print(f"跳过空run: 段落 {para_index} 中的Run {run_index}")
                            continue
                        
                        # 对于标点符号，确认是否真的有错误需要标记
                        is_punctuation = len(run_text_stripped) <= 1 and all(c in ',.!?;:"\'()[]{}""''：；，。！？、（）【】《》' for c in run_text_stripped)
                        if is_punctuation:
                            # 验证是否有真正重要的错误
                            has_important_error = False
                            for error_dict in result:
                                for attr, values in error_dict.items():
                                    # 对于标点符号，字体和字号通常不重要
                                    if 'error' in values and attr not in ['size', 'font_ascii'] and is_punctuation:
                                        has_important_error = True
                                        break
                            
                            if not has_important_error:
                                print(f"跳过标点符号run (不需要修复): 段落 {para_index} 中的Run {run_index} ({run_text})")
                                continue
                        
                        # 为run添加高亮
                        doc.set_run_highlight(para_index_, run_index, highlight_color)

                        # 为run添加批注
                        doc.add_comment(
                            element_index=para_index,
                            run_index=run_index,
                            author="样式检查器",
                            comment_text=comment_text,
                            element_type="run"
                        )

                        marked_errors['run'] += 1
                        marked_errors['total'] += 1
                        print(f"已标记段落 {para_index} 中的Run {run_index} 的样式错误并添加批注 (颜色: {highlight_color})")
                        print(f"处理元素: {result}")
                except Exception as e:
                    print(f"标记Run错误时出错: {e}")

            # 处理表格元素
            elif element_type == 'table':
                try:
                    # 直接使用索引定位表格
                    table_index = element_info.get('index')


                    if table_index is not None:
                        # 为表格添加批注
                        doc.add_comment(
                            element_index=table_index,
                            author="样式检查器",
                            comment_text=comment_text,
                            element_type="table"
                        )
                        print(f"表格索引: {table_index}")
                        # 分析错误类型，确定需要标记的单元格
                        column_header_errors = any('column_header' in str(error) for error in result)
                        data_row_errors = any('data_row' in str(error) for error in result)

                        # 获取表格尺寸

                        dimensions = doc.get_table_dimensions(table_index)
                        if not dimensions:
                            continue

                        rows, cols = dimensions

                        # 根据错误类型标记相应单元格
                        if column_header_errors and rows > 0:
                            # 标记表头行
                            for col_idx in range(cols):
                                # 获取单元格中的段落
                                cell_paras = doc.get_table_cell_paragraphs(table_index, 0, col_idx)
                                if cell_paras:
                                    # 查找段落在文档中的索引
                                    for para in cell_paras:

                                        if para is not None:
                                            # 标记段落
                                            run_count = doc.get_run_count_from_xml(para)
                                            for run_idx in range(run_count):
                                                doc.set_run_highlight_from_xml(para, run_idx, highlight_color)

                        if data_row_errors and rows > 1:
                            # 标记数据行（第二行）
                            for col_idx in range(cols):
                                # 获取单元格中的段落
                                cell_paras = doc.get_table_cell_paragraphs(table_index, 1, col_idx)
                                if cell_paras:
                                    # 查找段落在文档中的索引
                                    for para in cell_paras:

                                        if para is not None:
                                            # 标记段落
                                            run_count = doc.get_run_count_from_xml(para)
                                            for run_idx in range(run_count):
                                                doc.set_run_highlight_from_xml(para, run_idx, highlight_color)

                        marked_errors['table'] += 1
                        marked_errors['total'] += 1
                        print(f"已标记表格 {table_index} 的样式错误并添加批注 (颜色: {highlight_color})")
                except Exception as e:
                    print(f"标记表格错误时出错: {e}")

    # 保存统计信息
    if save_statistics:
        output_dir = os.path.dirname(output_path) if os.path.dirname(output_path) else "."
        output_filename = os.path.basename(doc_path).split('.')[0]
        statistics_path = os.path.join(output_dir, f"{output_filename}_style_statistics.json")
        
        # 确保统计信息可以序列化为JSON
        serialized_statistics = serialize_statistics(cleaned_statistics)
        
        with open(statistics_path, 'w', encoding='utf-8') as f:
            json.dump(serialized_statistics, f, ensure_ascii=False, indent=2)
        print(f"样式统计信息已保存到: {statistics_path}")

    # 保存修改后的文档
    doc.save(output_path)
    print(f"\n文档标记和批注完成并已保存至: {output_path}")
    print(f"共标记了 {marked_errors['total']} 处样式错误:")
    print(f"  - 段落错误: {marked_errors['paragraph']}")
    print(f"  - Run错误: {marked_errors['run']}")
    print(f"  - 表格错误: {marked_errors['table']}")

    return output_path


def generate_comment_text(result, style_mapping, element_type):
    """
    生成批注文本，包含错误信息和正确的样式值

    参数:
        result: 样式比较结果
        style_mapping: 样式映射数据
        element_type: 元素类型

    返回:
        str: 批注文本
    """
    comment_lines = []
    comment_lines.append(f"{element_type.capitalize()}样式错误：")

    # 获取对应元素类型的样式映射
    element_mapping = {}
    if element_type == 'paragraph':
        element_mapping = style_mapping.get('paragraph_styles', {})
    elif element_type == 'run':
        element_mapping = style_mapping.get('character_styles', {})
    elif element_type == 'table':
        element_mapping = style_mapping.get('table_styles', {})

    # 判断是黄色标记还是红色标记的情况
    has_error = False
    has_only_success = False
    
    for error_dict in result:
        for attr, values in error_dict.items():
            # 检查新的success键和旧的scuccess键（向后兼容）
            if 'error' in values and ('success' in values or 'scuccess' in values):
                has_error = True
            elif ('success' in values or 'scuccess' in values) and 'error' not in values:
                has_only_success = True
    
    # 处理每个错误结果
    for error_dict in result:
        for attr, values in error_dict.items():
            attribute_name = get_readable_attribute_name(attr)
            
            # 获取正确值（兼容新旧键名）
            expected_value = values.get('success', values.get('scuccess'))
            
            # 红色标记：同时有error和success/scuccess
            if 'error' in values and expected_value is not None:
                current_value = values.get('error')
                
                comment_lines.append(f"• {attribute_name}:")
                comment_lines.append(f"  - 当前值: {current_value}")
                comment_lines.append(f"  - 正确值: {expected_value}")
                
                # 查找更多关于此属性的信息
                if attr in element_mapping:
                    mapping_info = element_mapping[attr]
                    if isinstance(mapping_info, dict) and 'description' in mapping_info:
                        comment_lines.append(f"  - 说明: {mapping_info['description']}")
            
            # 黄色标记：只有success/scuccess没有error
            elif expected_value is not None and 'error' not in values:
                comment_lines.append(f"• {attribute_name}:")
                comment_lines.append(f"  - 缺少设置")
                comment_lines.append(f"  - 正确值: {expected_value}")
                
                # 查找更多关于此属性的信息
                if attr in element_mapping:
                    mapping_info = element_mapping[attr]
                    if isinstance(mapping_info, dict) and 'description' in mapping_info:
                        comment_lines.append(f"  - 说明: {mapping_info['description']}")
    
    # 添加总结信息
    if has_error:
        comment_lines.append("\n红色标记表示样式设置错误，需要修正。")
    if has_only_success:
        comment_lines.append("\n黄色标记表示缺少样式设置，需要添加。")
    
    return "\n".join(comment_lines)


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
    }

    return attribute_names.get(attr, attr)








def serialize_statistics(statistics):
    """
    将统计信息序列化为JSON可表示的格式
    
    参数:
        statistics: 原始统计信息字典
        
    返回:
        dict: 可序列化的统计信息字典
    """
    serialized = {}
    
    # 深拷贝字典，避免修改原始数据
    import copy
    serialized = copy.deepcopy(statistics)
    
    # 处理elements列表
    if 'elements' in serialized:
        new_elements = []
        for elem in serialized['elements']:
            # 如果元素包含XML Element对象，将它们转换为描述字符串
            elem_copy = copy.deepcopy(elem)
            if 'element' in elem_copy:
                elem_copy['element'] = str(elem_copy['element'])
            new_elements.append(elem_copy)
        serialized['elements'] = new_elements
    
    return serialized


# 在这里直接设置文件路径
if __name__ == "__main__":
    # 设置文件路径 - 根据实际情况修改这些路径
    doc_path = "1.docx"  # 要处理的Word文档路径
    classification_path = "document_classification_results.json"  # 文档分类结果路径
    style_mapping_path = "document_style_mapping.json"  # 样式映射文件路径
    api_params_path = "智算工程学院毕业设计（论文）模板2025届(1)_api_params.json"  # API参数文件路径
    output_path = "1_marked.docx"  # 输出文件路径，可以留空使用默认路径

    # 检查文件是否存在
    missing_files = []
    for path, desc in [
        (doc_path, "文档"),
        (classification_path, "分类结果文件"),
        (style_mapping_path, "样式映射文件"),
        (api_params_path, "API参数文件")
    ]:
        if not os.path.exists(path):
            missing_files.append(f"{desc}: {path}")

    if missing_files:
        print("错误: 以下文件不存在:")
        for missing in missing_files:
            print(f"  - {missing}")
    else:
        # 执行标记过程
        marked_file = mark_style_errors(doc_path, classification_path, style_mapping_path, api_params_path, output_path)
        print(f"标记和批注完成: {marked_file}")