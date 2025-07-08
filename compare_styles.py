#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
样式比较工具 - 比较文档实际样式与目标样式
"""

import json
import os
from docx_namespace import DocxElementParser
from style_analyzer import StyleAnalyzer
from collections import defaultdict


# def compare_styles(doc_path, classification_path, style_mapping_path, api_params_path):
def compare_styles(doc_path, classification, style_mapping_path, api_params):

    """
    比较文档中实际的段落样式与目标样式参数
    
    参数:
        doc_path: Word文档路径
        classification_path: 分类结果JSON文件路径
        style_mapping_path: 样式映射JSON文件路径
        api_params_path: API参数格式JSON文件路径
    """
    # 加载文档
    doc = DocxElementParser(doc_path)
    # 创建样式分析器实例获取完整样式信息
    style_analyzer = StyleAnalyzer(doc_path)
    
    # 加载分类结果
    # with open(classification_path, 'r', encoding='utf-8') as f:
    #     classification = json.load(f)
    
    # 加载样式映射
    with open(style_mapping_path, 'r', encoding='utf-8') as f:
        style_mapping = json.load(f)
    
    # 加载API参数
    # with open(api_params_path, 'r', encoding='utf-8') as f:
    #     api_params = json.load(f)
    
    # 结果统计
    statistics = {
        "element_types": 0,  # 比较的元素类型数量
        'elements': [],  # 记录run和段落位置和错误信息的列表
        "paragraphs": 0,  # 比较的段落总数
        "runs": 0,  # 比较的run总数（只计算非空的）
        "paragraphs_matching": 0,  # 样式完全匹配的段落数
        "runs_matching": 0,  # 样式完全匹配的run数（只计算非空的）
        "attribute_matches": defaultdict(int),  # 各属性的匹配次数
        "attribute_totals": defaultdict(int)  # 各属性的比较总次数
    }
    
    print(f"=== 开始比较文档样式: {doc_path} ===\n")
    
    # 遍历映射中的每个部分
    for section, section_mapping in style_mapping.items():
        print(f"\n### 部分: {section} ###")
        
        # 检查该部分在API参数中是否存在
        if section not in api_params:
            print(f"警告: API参数中没有找到'{section}'部分")
            continue
        
        # 获取该部分的API参数
        section_params = api_params[section]
        
        # 特殊处理表格部分
        if section == "正文" and "表格" in section_params:
            print("\n>>> 开始分析表格样式 <<<")
            compare_tables_style(doc, section_params["表格"], statistics)
        
        # 遍历该部分中的每个元素类型
        for element_class, style_class in section_mapping.items():
            # 跳过映射为null的元素
            if style_class is None:
                continue
            
            # 检查元素类型在分类结果中是否存在
            if element_class not in classification:
                print(f"警告: 分类结果中没有找到'{element_class}'元素类型")
                continue
            
            # 检查样式类型在API参数中是否存在
            if style_class not in section_params:
                print(f"警告: API参数中没有找到'{section}.{style_class}'样式")
                continue
            
            # 获取目标样式参数
            target_style = section_params[style_class]
            
            # 获取该元素类型的所有段落索引
            para_indices = classification[element_class]
            
            # 如果没有段落，跳过
            if not para_indices:
                print(f"信息: '{element_class}'没有包含任何段落")
                continue
            
            statistics["element_types"] += 1
            
            # 打印元素类型信息
            print(f"\n>>> 元素类型: {element_class} -> 样式类型: {style_class} (包含 {len(para_indices)} 个段落)")
            
            # 逐个分析每个段落
            paragraphs_matching = 0
            runs_matching = 0
            total_runs = 0  # 只计算非空的run
            
            # 用于收集每个属性的匹配情况
            attribute_matches = defaultdict(int)
            attribute_totals = defaultdict(int)
            
            # 分析每个段落
            for para_idx in para_indices:
                try:
                    # 获取段落的实际样式 - 使用StyleAnalyzer的get_paragraph_complete_style_info方法
                    para_element = doc.elements[para_idx]['element'] if isinstance(doc.elements[para_idx],
                                                                                   dict) and 'element' in doc.elements[
                                                                            para_idx] else doc.elements[para_idx]
                    
                    # 使用新方法获取完整样式信息
                    complete_style_info = style_analyzer.get_paragraph_complete_style_info(para_element)
                    actual_para_style = complete_style_info['effective_style']
                    
                    # 打印段落的样式ID和使用的模板样式(如果有)
                    style_id = complete_style_info.get('style_id')
                    print(f"\n  段落 {para_idx}:")
                    if style_id:
                        print(f"    应用的样式ID: {style_id}")
                    else:
                        print(f"    无样式ID (使用默认样式)")
                    
                    # 比较段落样式
                    para_results = compare_paragraph_style(actual_para_style, target_style, prefix="    段落样式: ")
                    statistics['elements'].append({
                        "type": "paragraph", 
                        'element': doc.elements[para_idx]['element'],
                        'result': para_results['results'],
                        'index': para_idx  # 添加段落索引
                    })
                    # 更新属性统计
                    for attr, matched in para_results.items():
                        if attr != 'results':  # 跳过results键
                            attribute_matches[attr] += 1 if matched else 0
                            attribute_totals[attr] += 1
                    
                    # 判断段落样式是否完全匹配
                    para_match = all(matched for attr, matched in para_results.items() if
                                     attr != 'results') if para_results else False
                    if para_match:
                        paragraphs_matching += 1
                    
                    # 获取段落中所有run的样式
                    run_count = doc.get_run_count(para_idx)
                    non_empty_runs = 0  # 计数非空run
                    
                    matching_runs_in_para = 0
                    if run_count > 0:
                        # 先计算非空run的数量
                        non_empty_run_count = 0
                        for run_idx in range(run_count):
                            run_text = doc._get_run_text(para_idx, run_idx)
                            if run_text.strip():
                                non_empty_run_count += 1
                        
                        print(f"    包含 {run_count} 个run (其中 {non_empty_run_count} 个非空):")

                        # 分析每个run
                        for run_idx in range(run_count):
                            # 获取run文本
                            run_text = doc._get_run_text(para_idx, run_idx)
                            
                            # 跳过空白run
                            if not run_text.strip():
                                continue
                            
                            non_empty_runs += 1  # 计数非空run
                            
                            # 获取run的完整样式信息，使用style_analyzer的get_run_complete_style_info方法
                            # 首先获取run元素
                            run_element = doc._get_run_element(para_idx, run_idx)

                            # 使用style_analyzer获取run的完整样式信息
                            run_complete_style_info = style_analyzer.get_run_complete_style_info(para_element,
                                                                                                 run_element, run_idx)
                            run_effective_style = run_complete_style_info['effective_style']
                            
                            print(f"    Run {run_idx} (\"{run_text[:20]}{'...' if len(run_text) > 20 else ''}\")")
                            
                            # 如果样式ID存在，打印出来
                            if 'style_id' in run_complete_style_info:
                                print(f"      应用的样式ID: {run_complete_style_info['style_id']}")

                            # 比较run样式
                            run_results = compare_run_style(run_effective_style, target_style, prefix="      ")
                            statistics['elements'].append({
                                "type": "run", 
                                'element': run_element,
                                'result': run_results['results'],
                                'index': (para_idx, run_idx)  # 添加(段落索引,run索引)元组
                            })
                            # 更新属性统计
                            for attr, matched in run_results.items():
                                if attr != 'results':  # 跳过results键
                                    attribute_matches[attr] += 1 if matched else 0
                                    attribute_totals[attr] += 1

                            # 判断run样式是否完全匹配
                            run_match = all(matched for attr, matched in run_results.items() if
                                            attr != 'results') if run_results else False
                            if run_match:
                                matching_runs_in_para += 1

                    runs_matching += matching_runs_in_para
                    total_runs += non_empty_runs  # 只计算非空的run

                    statistics["paragraphs"] += 1

                except Exception as e:
                    print(f"  错误: 提取段落 {para_idx} 的样式时出错: {e}")

            # 更新全局统计
            statistics["paragraphs_matching"] += paragraphs_matching
            statistics["runs_matching"] += runs_matching
            statistics["runs"] += total_runs  # 更新非空run的总数

            for attr in attribute_matches:
                statistics["attribute_matches"][attr] += attribute_matches[attr]
                statistics["attribute_totals"][attr] += attribute_totals[attr]

    # 打印全局统计信息
    print(f"\n=== 样式比较总结 ===")
    print(f"比较了 {statistics['element_types']} 种元素类型")
    print(
           f"段落: {statistics['paragraphs_matching']}/{statistics['paragraphs']} 匹配 ({statistics['paragraphs_matching'] / statistics['paragraphs'] * 100:.1f}% 如果有段落)")

    if statistics['runs'] > 0:
        print(
            f"非空Runs: {statistics['runs_matching']}/{statistics['runs']} 匹配 ({statistics['runs_matching'] / statistics['runs'] * 100:.1f}%)")

    return statistics


def compare_run_style(run_effective_style, target_style, prefix=""):
    """
    比较run的实际样式与目标样式

    参数:
        run_effective_style: run的有效样式 (从style_analyzer的get_run_complete_style_info获取)
        target_style: 目标样式字典
        prefix: 输出前缀

    返回:
        dict: 各属性的匹配结果 {attr: matched}
    """
    # 匹配结果
    match_results = {
        'results': []
    }

    # 确保我们有run_properties
    if 'run_properties' not in run_effective_style:
        print(f"{prefix}! 警告: run的有效样式中缺少run_properties")
        run_properties = {}
    else:
        run_properties = run_effective_style['run_properties']

    # 比较字体
    fonts = run_properties.get('fonts', {})
    if 'font_chinese' in target_style:
        if 'eastAsia' in fonts:
            target_font = target_style['font_chinese']
            actual_font = fonts['eastAsia']
            matches = actual_font == target_font
            match_results['font_chinese'] = matches
            match_indicator = "✓" if matches else "✗"
            if match_indicator == "✗":
                match_results['results'].append({
                    'font_chinese': {'success': target_font, 'error': actual_font}})
            print(f"{prefix}{match_indicator} font_chinese: 目标值={target_font}, 实际值={actual_font}")
        else:
            print(f"{prefix}! font_chinese: 目标值={target_style['font_chinese']}, 实际值=未定义")
            match_results['font_chinese'] = False
            match_results['results'].append({
                'font_chinese': {'success': target_style['font_chinese']}})

    if 'font_ascii' in target_style:
        if 'ascii' in fonts:
            target_font = target_style['font_ascii']
            actual_font = fonts['ascii']
            matches = actual_font == target_font
            match_results['font_ascii'] = matches
            match_indicator = "✓" if matches else "✗"
            print(f"{prefix}{match_indicator} font_ascii: 目标值={target_font}, 实际值={actual_font}")
            if match_indicator == "✗":
                match_results['results'].append({
                    'font_ascii': {'success': target_font, 'error': actual_font}})
        else:
            print(f"{prefix}! font_ascii: 目标值={target_style['font_ascii']}, 实际值=未定义")
            match_results['font_ascii'] = False
            match_results['results'].append({
                'font_ascii': {'success': target_style['font_ascii']}})

    # 比较字号
    if 'size' in target_style:
        if 'size' in run_properties:
            target_size = target_style['size']
            actual_size = run_properties['size']
            try:
                actual_size = int(actual_size)

                # 特殊处理中文字号对应关系

                if target_size == actual_size:
                    matches = True
                    print(f"{prefix}✓ size: 目标值={target_size}, 实际值={actual_size} ")

                else:
                    matches = False
                    match_indicator = "✗"
                    print(f"{prefix}{match_indicator} size: 目标值={target_size}, 实际值={actual_size}")
                    # 只有不匹配时才添加到结果
                    match_results['results'].append({
                        'size': {'success': target_size, 'error': actual_size}})

                match_results['size'] = matches
            except:
                match_results['size'] = False
                print(f"{prefix}? size: 目标值={target_size}, 实际值={actual_size} (无法比较)")
                match_results['results'].append({
                    'size': {'success': target_size}
                })
        else:
            print(f"{prefix}! size: 目标值={target_style['size']}, 实际值=未定义")
            match_results['size'] = False
            # 仅当size缺失时添加到结果
            match_results['results'].append({
                'size': {'success': target_style['size']}
            })

    # 比较加粗
    if 'bold' in target_style:
        if 'bold' in run_properties:
            target_bold = target_style['bold']
            actual_bold = run_properties['bold']
            if isinstance(actual_bold, str):
                actual_bold = actual_bold.lower() in ["true", "yes", "1"]
            matches = actual_bold == target_bold
            match_results['bold'] = matches
            match_indicator = "✓" if matches else "✗"
            print(f"{prefix}{match_indicator} bold: 目标值={target_bold}, 实际值={actual_bold}")
            if match_indicator == "✗":
                match_results['results'].append({
                    'bold': {'success': target_bold, 'error': actual_bold}})
        else:
            # 如果在实际样式中没有bold属性，则默认为False
            actual_bold = False
            target_bold = target_style['bold']

            # 如果目标值也是False，则视为匹配
            if target_bold == False:
                matches = True
                print(f"{prefix}✓ bold: 目标值={target_bold}, 实际值={actual_bold} (默认为False)")
            else:
                matches = False
                print(f"{prefix}✗ bold: 目标值={target_bold}, 实际值={actual_bold} (默认为False)")
                match_results['results'].append({
                    'bold': {'success': target_bold, 'error': actual_bold}
                })

            match_results['bold'] = matches

    return match_results


def merge_run_styles(paragraph_style, run_direct_style):
    """
    合并段落样式和run直接样式，模拟Word中run样式继承段落样式的行为

    注意: 此函数现在可能不再需要，因为StyleAnalyzer的get_run_complete_style_info已经处理了样式继承
    但保留此函数以维护向后兼容性，或者在无StyleAnalyzer情况下使用

    参数:
        paragraph_style: 段落有效样式
        run_direct_style: run直接样式

    返回:
        dict: 合并后的run有效样式
    """
    # 创建有效样式的副本
    run_effective_style = {}

    # 首先，从段落样式中提取run相关属性
    if 'run_properties' in paragraph_style:
        para_run_props = paragraph_style['run_properties']

        # 复制所有run属性
        for key, value in para_run_props.items():
            if isinstance(value, dict):
                run_effective_style[key] = value.copy()
            else:
                run_effective_style[key] = value

    # 然后，应用run的直接样式，但只覆盖run中实际设置的样式
    for key, value in run_direct_style.items():
        # 跳过未设置的属性（None或空值）
        if value is None or value == '':
            continue

        if key == 'font_name_eastAsia' and value:
            if 'fonts' not in run_effective_style:
                run_effective_style['fonts'] = {}
            run_effective_style['fonts']['eastAsia'] = value
        elif key == 'font_name_ascii' and value:
            if 'fonts' not in run_effective_style:
                run_effective_style['fonts'] = {}
            run_effective_style['fonts']['ascii'] = value
        elif key == 'font_size' and value:
            run_effective_style['size'] = value
        elif key == 'is_bold' and value is not None:  # 显式设置了加粗属性才覆盖
            run_effective_style['bold'] = value
        elif value:  # 其他有效属性也仅在有值时覆盖
            # 其他属性直接复制
            run_effective_style[key] = value

    # 对于目标比较样式中期望的关键属性，检查是否存在
    # 如果run_properties中有bold属性，但直接样式没有设置is_bold，保留原有bold属性
    if 'bold' in run_effective_style and 'is_bold' not in run_direct_style:
        # 保留段落中的加粗设置
        pass  # 不做任何覆盖，保留原有值

    return run_effective_style


def compare_paragraph_style(actual_style, target_style, prefix=""):
    """
    比较段落的实际样式与目标样式

    参数:
        actual_style: 实际段落样式字典 (get_paragraph_complete_style_info的effective_style)
        target_style: 目标样式字典
        prefix: 输出前缀

    返回:
        dict: 各属性的匹配结果 {attr: matched}
    """
    # 匹配结果
    match_results = {
        'results': []
    }

    # 11.py. 比较对齐方式 (alignment)
    if 'alignment' in target_style:
        # 获取实际的alignment值，如果不存在则默认为"both"
        actual_alignment = None
        if 'paragraph_properties' in actual_style and 'alignment' in actual_style['paragraph_properties']:
            actual_alignment = actual_style['paragraph_properties']['alignment']
        else:
            # 当alignment未设置时，Word默认值为"both"
            actual_alignment = "both"

        target_alignment = target_style['alignment']

            # 特殊处理: left和both视为相同的对齐方式
        if (target_alignment == 'left' and actual_alignment == 'both') or (
                target_alignment == 'both' and actual_alignment == 'left'):
                matches = True
                print(f"{prefix}✓ alignment: 目标值={target_alignment}, 实际值={actual_alignment} (left和both视为相同)")
        else:
                matches = actual_alignment == target_alignment
                match_indicator = "✓" if matches else "✗"
                print(f"{prefix}{match_indicator} alignment: 目标值={target_alignment}, 实际值={actual_alignment}")
                if match_indicator == "✗":
                    match_results['results'].append(
                    {'alignment': {'success': target_alignment, 'error': actual_alignment}})

        match_results['alignment'] = matches

    # 2. 比较缩进 (first_line, hanging)
    if 'first_line' in target_style:
        if 'paragraph_properties' in actual_style and 'indentation' in actual_style[
            'paragraph_properties'] and 'firstLine' in actual_style['paragraph_properties']['indentation']:
            target_first_line = target_style['first_line']
            actual_first_line = actual_style['paragraph_properties']['indentation'].get('firstLine', None)
            if actual_first_line is not None:
                # 允许一定的误差
                try:
                    actual_first_line = int(actual_first_line)
                    # 增加误差容忍度，考虑中文字符宽度 (440 twip) 的影响

                    if abs(actual_first_line - target_first_line) <= 10:
                        matches = True
                        print(
                            f"{prefix}✓ first_line: 目标值={target_first_line}, 实际值={actual_first_line} (允许较大误差)")
                    else:
                        matches = False
                        match_indicator = "✗"
                        print(
                            f"{prefix}{match_indicator} first_line: 目标值={target_first_line}, 实际值={actual_first_line}")
                        match_results['results'].append({
                            'first_line': {'success': target_first_line, 'error': actual_first_line}})
                    match_results['first_line'] = matches
                except:
                    match_results['first_line'] = False
                    print(f"{prefix}? first_line: 目标值={target_first_line}, 实际值={actual_first_line} (无法比较)")
                    match_results['results'].append({
                        'first_line': {'success': target_first_line}})
        else:
            print(f"{prefix}! first_line: 目标值={target_style['first_line']}, 实际值=未定义")
            match_results['first_line'] = False
            match_results['results'].append({
                'first_line': {'success': target_style['first_line']}})

    if 'hanging' in target_style:
        if 'paragraph_properties' in actual_style and 'indentation' in actual_style[
            'paragraph_properties'] and 'hanging' in actual_style['paragraph_properties']['indentation']:
            target_hanging = target_style['hanging']
            actual_hanging = actual_style['paragraph_properties']['indentation'].get('hanging', None)
            if actual_hanging is not None:
                try:
                    actual_hanging = int(actual_hanging)
                    matches = abs(actual_hanging - target_hanging) < 20
                    match_results['hanging'] = matches
                    match_indicator = "✓" if matches else "✗"
                    print(f"{prefix}{match_indicator} hanging: 目标值={target_hanging}, 实际值={actual_hanging}")
                    if match_indicator == "✗":
                        match_results['results'].append({
                            'hanging': {'success': target_hanging, 'error': actual_hanging}})
                except:
                    match_results['hanging'] = False
                    print(f"{prefix}? hanging: 目标值={target_hanging}, 实际值={actual_hanging} (无法比较)")
                    match_results['results'].append({
                        'hanging': {'success': target_hanging}})
        else:
            print(f"{prefix}! hanging: 目标值={target_style['hanging']}, 实际值=未定义")
            match_results['hanging'] = False
            match_results['results'].append({
                'hanging': {'success': target_style['hanging']}})

    # 3. 比较行距和段落间距
    if 'line' in target_style:
        if 'paragraph_properties' in actual_style and 'spacing' in actual_style['paragraph_properties'] and 'line' in \
                actual_style['paragraph_properties']['spacing']:
            target_line = target_style['line']
            actual_line = actual_style['paragraph_properties']['spacing'].get('line', None)

            # 获取行距规则
            target_line_rule = target_style.get('line_rule', None)
            actual_line_rule = actual_style['paragraph_properties']['spacing'].get('lineRule', None)

            if actual_line is not None:
                try:
                    actual_line = int(actual_line)

                    # 特殊处理行距比较 - 考虑不同行距规则下的等效性
                    if target_line_rule and actual_line_rule and target_line_rule != actual_line_rule:
                        # 当规则不同时，进行特殊比较
                        if (target_line_rule == 'exact' and actual_line_rule == 'auto' and
                            abs(actual_line - target_line) < 120):  # 120 twips (6 points) 的误差容忍度
                            matches = True
                            print(
                                f"{prefix}✓ line: 目标值={target_line}({target_line_rule}), 实际值={actual_line}({actual_line_rule}) (不同规则但等效)")
                        else:
                            matches = False
                            match_indicator = "✗"
                            print(
                                f"{prefix}{match_indicator} line: 目标值={target_line}({target_line_rule}), 实际值={actual_line}({actual_line_rule})")
                            match_results['results'].append({
                                'line': {'success': target_line, 'error': actual_line}})
                    else:
                        # 相同规则下的常规比较
                        matches = abs(actual_line - target_line) < 20
                        match_indicator = "✓" if matches else "✗"
                        print(f"{prefix}{match_indicator} line: 目标值={target_line}, 实际值={actual_line}")
                        if match_indicator == "✗":
                            match_results['results'].append({
                                'line': {'success': target_line, 'error': actual_line}})

                    match_results['line'] = matches
                except:
                    match_results['line'] = False
                    print(f"{prefix}? line: 目标值={target_line}, 实际值={actual_line} (无法比较)")
                    match_results['results'].append({
                        'line': {'success': target_line}})
        else:
            print(f"{prefix}! line: 目标值={target_style['line']}, 实际值=未定义")
            match_results['line'] = False
            match_results['results'].append({
                'line': {'success': target_style['line']}})

    if 'line_rule' in target_style:
        if 'paragraph_properties' in actual_style and 'spacing' in actual_style[
            'paragraph_properties'] and 'lineRule' in actual_style['paragraph_properties']['spacing']:
            target_line_rule = target_style['line_rule']
            actual_line_rule = actual_style['paragraph_properties']['spacing'].get('lineRule', None)

            if actual_line_rule is not None:
                # 特殊处理行距规则 - 某些情况下auto和exact可以视为等效
                if ((target_line_rule == 'exact' and actual_line_rule == 'auto') or
                    (target_line_rule == 'auto' and actual_line_rule == 'exact')):
                    # 获取行距值，检查是否接近
                    target_line = target_style.get('line', 0)
                    actual_line = actual_style['paragraph_properties']['spacing'].get('line', 0)
                    if isinstance(actual_line, str):
                        try:
                            actual_line = int(actual_line)
                        except:
                            actual_line = 0

                    if abs(int(target_line) - int(actual_line)) < 120:  # 120 twips (6 points) 的误差容忍度
                        matches = True
                        print(
                            f"{prefix}✓ line_rule: 目标值={target_line_rule}, 实际值={actual_line_rule} (考虑行距值后视为等效)")
                    else:
                        matches = False
                        match_indicator = "✗"
                        print(
                            f"{prefix}{match_indicator} line_rule: 目标值={target_line_rule}, 实际值={actual_line_rule}")
                        match_results['results'].append({
                            'line_rule': {'success': target_line_rule, 'error': actual_line_rule}})
                else:
                    # 常规比较
                    matches = actual_line_rule == target_line_rule
                    match_indicator = "✓" if matches else "✗"
                    print(f"{prefix}{match_indicator} line_rule: 目标值={target_line_rule}, 实际值={actual_line_rule}")
                    if match_indicator == "✗":
                        match_results['results'].append({
                            'line_rule': {'success': target_line_rule, 'error': actual_line_rule}})

                match_results['line_rule'] = matches
        else:
            print(f"{prefix}! line_rule: 目标值={target_style['line_rule']}, 实际值=未定义")
            match_results['results'].append({
                'line_rule': {'success': target_style['line_rule']}})
            match_results['line_rule'] = False

    if 'before' in target_style:
        if 'paragraph_properties' in actual_style and 'spacing' in actual_style['paragraph_properties'] and 'before' in \
                actual_style['paragraph_properties']['spacing']:
            target_before = target_style['before']
            actual_before = actual_style['paragraph_properties']['spacing'].get('before', None)
            if actual_before is not None:
                try:
                    actual_before = int(actual_before)
                    # 特殊处理：如果actual_before很小(<=50)，则视为0
                    if actual_before <= 20 and target_before == 0:
                        matches = True
                        print(f"{prefix}✓ before: 目标值={target_before}, 实际值={actual_before} (小值视为0)")
                    else:
                        matches = abs(actual_before - target_before) < 20
                    match_results['before'] = matches
                    match_indicator = "✓" if matches else "✗"
                    print(f"{prefix}{match_indicator} before: 目标值={target_before}, 实际值={actual_before}")
                    if match_indicator == "✗":
                        match_results['results'].append({
                            'before': {'success': target_before, 'error': actual_before}})
                except:
                    match_results['before'] = False
                    print(f"{prefix}? before: 目标值={target_before}, 实际值={actual_before} (无法比较)")
                    match_results['results'].append({
                        'before': {'success': target_before}})
        else:
            # 如果在实际样式中没有before属性，则默认为0
            actual_before = 0
            target_before = target_style['before']

            # 如果目标值也是0，则视为匹配
            if target_before == 0:
                matches = True
                print(f"{prefix}✓ before: 目标值={target_before}, 实际值={actual_before} (默认为0)")
            else:
                matches = False
                print(f"{prefix}✗ before: 目标值={target_before}, 实际值={actual_before} (默认为0)")
                match_results['results'].append({
                    'before': {'success': target_before, 'error': actual_before}})
            match_results['before'] = matches

    if 'after' in target_style:
        if 'paragraph_properties' in actual_style and 'spacing' in actual_style['paragraph_properties'] and 'after' in \
                actual_style['paragraph_properties']['spacing']:
            target_after = target_style['after']
            actual_after = actual_style['paragraph_properties']['spacing'].get('after', None)
            if actual_after is not None:
                try:
                    actual_after = int(actual_after)
                    # 特殊处理：如果actual_after很小(<=50)，则视为0
                    if actual_after <= 20 and target_after == 0:
                        matches = True
                        print(f"{prefix}✓ after: 目标值={target_after}, 实际值={actual_after} (小值视为0)")
                    else:
                        matches = abs(actual_after - target_after) < 20
                    match_results['after'] = matches
                    match_indicator = "✓" if matches else "✗"
                    print(f"{prefix}{match_indicator} after: 目标值={target_after}, 实际值={actual_after}")
                    if match_indicator == "✗":
                        match_results['results'].append({
                            'after': {'success': target_after, 'error': actual_after}})
                except:
                    match_results['after'] = False
                    print(f"{prefix}? after: 目标值={target_after}, 实际值={actual_after} (无法比较)")
                    match_results['results'].append({
                        'after': {'success': target_after}})

        else:
            # 如果在实际样式中没有after属性，则默认为0
            actual_after = 0
            target_after = target_style['after']

            # 如果目标值也是0，则视为匹配
            if target_after == 0:
                matches = True
                print(f"{prefix}✓ after: 目标值={target_after}, 实际值={actual_after} (默认为0)")
            else:
                matches = False
                print(f"{prefix}✗ after: 目标值={target_after}, 实际值={actual_after} (默认为0)")
                match_results['results'].append({
                    'after': {'success': target_after, 'error': actual_after}})

            match_results['after'] = matches
    # 添加对beforeLines的处理
    if 'beforeLines' in target_style:
        if 'paragraph_properties' in actual_style and 'spacing' in actual_style['paragraph_properties'] and 'beforeLines' in \
                actual_style['paragraph_properties']['spacing']:
            target_beforeLines = target_style['beforeLines']
            actual_beforeLines = actual_style['paragraph_properties']['spacing'].get('beforeLines', None)
            if actual_beforeLines is not None:
                try:
                    actual_beforeLines = int(actual_beforeLines)
                    # 基于行的间距通常不需要像点值间距那样有误差容忍
                    matches = actual_beforeLines == target_beforeLines
                    match_results['beforeLines'] = matches
                    match_indicator = "✓" if matches else "✗"
                    print(f"{prefix}{match_indicator} beforeLines: 目标值={target_beforeLines}, 实际值={actual_beforeLines}")
                    if match_indicator == "✗":
                        match_results['results'].append({
                            'beforeLines': {'success': target_beforeLines, 'error': actual_beforeLines}})
                except:
                    match_results['beforeLines'] = False
                    print(f"{prefix}? beforeLines: 目标值={target_beforeLines}, 实际值={actual_beforeLines} (无法比较)")
                    match_results['results'].append({
                        'beforeLines': {'success': target_beforeLines}})
        else:
            # 如果在实际样式中没有beforeLines属性，则默认为0
            actual_beforeLines = 0
            target_beforeLines = target_style['beforeLines']

            # 如果目标值也是0，则视为匹配
            if target_beforeLines == 0:
                matches = True
                print(f"{prefix}✓ beforeLines: 目标值={target_beforeLines}, 实际值={actual_beforeLines} (默认为0)")
            else:
                matches = False
                print(f"{prefix}✗ beforeLines: 目标值={target_beforeLines}, 实际值={actual_beforeLines} (默认为0)")
                match_results['results'].append({
                    'beforeLines': {'success': target_beforeLines, 'error': actual_beforeLines}})
            match_results['beforeLines'] = matches

    # 添加对afterLines的处理
    if 'afterLines' in target_style:
        if 'paragraph_properties' in actual_style and 'spacing' in actual_style['paragraph_properties'] and 'afterLines' in \
                actual_style['paragraph_properties']['spacing']:
            target_afterLines = target_style['afterLines']
            actual_afterLines = actual_style['paragraph_properties']['spacing'].get('afterLines', None)
            if actual_afterLines is not None:
                try:
                    actual_afterLines = int(actual_afterLines)
                    # 基于行的间距通常不需要像点值间距那样有误差容忍
                    matches = actual_afterLines == target_afterLines
                    match_results['afterLines'] = matches
                    match_indicator = "✓" if matches else "✗"
                    print(f"{prefix}{match_indicator} afterLines: 目标值={target_afterLines}, 实际值={actual_afterLines}")
                    if match_indicator == "✗":
                        match_results['results'].append({
                            'afterLines': {'success': target_afterLines, 'error': actual_afterLines}})
                except:
                    match_results['afterLines'] = False
                    print(f"{prefix}? afterLines: 目标值={target_afterLines}, 实际值={actual_afterLines} (无法比较)")
                    match_results['results'].append({
                        'afterLines': {'success': target_afterLines}})
        else:
            # 如果在实际样式中没有afterLines属性，则默认为0
            actual_afterLines = 0
            target_afterLines = target_style['afterLines']

            # 如果目标值也是0，则视为匹配
            if target_afterLines == 0:
                matches = True
                print(f"{prefix}✓ afterLines: 目标值={target_afterLines}, 实际值={actual_afterLines} (默认为0)")
            else:
                matches = False
                print(f"{prefix}✗ afterLines: 目标值={target_afterLines}, 实际值={actual_afterLines} (默认为0)")
                match_results['results'].append({
                    'afterLines': {'success': target_afterLines, 'error': actual_afterLines}})
            match_results['afterLines'] = matches
    # 检查run_properties中的字体和字号
    if 'run_properties' in actual_style:
        run_props = actual_style.get('run_properties', {})

        # 比较字体
        fonts = run_props.get('fonts', {})
        if 'font_chinese' in target_style:
            if 'eastAsia' in fonts:
                target_font = target_style['font_chinese']
                actual_font = fonts['eastAsia']
                matches = actual_font == target_font
                match_results['font_chinese'] = matches
                match_indicator = "✓" if matches else "✗"
                print(f"{prefix}{match_indicator} font_chinese: 目标值={target_font}, 实际值={actual_font}")
                if match_indicator == "✗":
                    match_results['results'].append({
                        'font_chinese': {'success': target_font, 'error': actual_font}})
            else:
                print(f"{prefix}! font_chinese: 目标值={target_style['font_chinese']}, 实际值=未定义")
                match_results['font_chinese'] = False
                match_results['results'].append({
                    'font_chinese': {'success': target_style['font_chinese']}})

        if 'font_ascii' in target_style:
            if 'ascii' in fonts:
                target_font = target_style['font_ascii']
                actual_font = fonts['ascii']
                matches = actual_font == target_font
                match_results['font_ascii'] = matches
                match_indicator = "✓" if matches else "✗"
                print(f"{prefix}{match_indicator} font_ascii: 目标值={target_font}, 实际值={actual_font}")
                if match_indicator == "✗":
                    match_results['results'].append({
                        'font_ascii': {'success': target_font, 'error': actual_font}})
            else:
                print(f"{prefix}! font_ascii: 目标值={target_style['font_ascii']}, 实际值=未定义")
                match_results['font_ascii'] = False
                match_results['results'].append({
                    'font_ascii': {'success': target_style['font_ascii']}
                })

        # 比较字号
        if 'size' in target_style:
            if 'size' in run_props:
                target_size = target_style['size']
                actual_size = run_props['size']
                try:
                    actual_size = int(actual_size)

                    # 特殊处理中文字号对应关系
                    # 小四号(24) 和 五号(21) 的特殊处理
                    if target_size ==  actual_size:
                        matches = True
                        print(f"{prefix}✓ size: 目标值={target_size}, 实际值={actual_size}")

                    else:
                        matches = False
                        match_indicator = "✗"
                        print(f"{prefix}{match_indicator} size: 目标值={target_size}, 实际值={actual_size}")
                        # 只有不匹配时才添加到结果
                        match_results['results'].append({
                            'size': {'success': target_size, 'error': actual_size}})

                    match_results['size'] = matches
                except:
                    match_results['size'] = False
                    print(f"{prefix}? size: 目标值={target_size}, 实际值={actual_size} (无法比较)")
                    match_results['results'].append({
                        'size': {'success': target_size}
                    })
            else:
                print(f"{prefix}! size: 目标值={target_style['size']}, 实际值=未定义")
                match_results['size'] = False
                # 仅当size缺失时添加到结果
                match_results['results'].append({
                    'size': {'success': target_style['size']}
                })

        # 比较加粗
        if 'bold' in target_style:
            if 'bold' in run_props:
                target_bold = target_style['bold']
                actual_bold = run_props['bold']
                if isinstance(actual_bold, str):
                    actual_bold = actual_bold.lower() in ["true", "yes", "11.py"]
                matches = actual_bold == target_bold
                match_results['bold'] = matches
                match_indicator = "✓" if matches else "✗"
                print(f"{prefix}{match_indicator} bold: 目标值={target_bold}, 实际值={actual_bold}")
                if match_indicator == "✗":
                    match_results['results'].append({
                        'bold': {'success': target_bold, 'error': actual_bold}})
            else:
                # 如果在实际样式中没有bold属性，则默认为False
                actual_bold = False
                target_bold = target_style['bold']

                # 如果目标值也是False，则视为匹配
                if target_bold == False:
                    matches = True
                    print(f"{prefix}✓ bold: 目标值={target_bold}, 实际值={actual_bold} (默认为False)")
                else:
                    matches = False
                    print(f"{prefix}✗ bold: 目标值={target_bold}, 实际值={actual_bold} (默认为False)")
                    match_results['results'].append({
                        'bold': {'success': target_bold, 'error': actual_bold}
                    })

                match_results['bold'] = matches
    else:
        # 如果实际样式中没有run_properties，但目标样式中有相关属性
        for attr in ['font_chinese', 'font_ascii', 'size', 'bold']:
            if attr in target_style:
                if attr == 'bold':
                    # 特殊处理bold属性，默认为False
                    actual_bold = False
                    target_bold = target_style['bold']

                    # 如果目标值也是False，则视为匹配
                    if target_bold == False:
                        matches = True
                        print(
                            f"{prefix}✓ bold: 目标值={target_bold}, 实际值={actual_bold} (默认为False，无run_properties)")
                    else:
                        matches = False
                        print(
                            f"{prefix}✗ bold: 目标值={target_bold}, 实际值={actual_bold} (默认为False，无run_properties)")
                        match_results['results'].append({
                            'bold': {'success': target_bold, 'error': actual_bold}
                        })

                    match_results['bold'] = matches
                else:
                    # 其他属性标记为未定义
                    print(f"{prefix}! {attr}: 目标值={target_style[attr]}, 实际值=未定义 (无run_properties)")
                    match_results[attr] = False
                    match_results['results'].append({
                        attr: {'success': target_style[attr]}
                    })

    return match_results


def compare_tables_style(doc, table_styles, statistics):
    """
    分析文档中所有表格，并与目标样式进行比较

    参数:
        doc: DocxElementParser实例
        table_styles: 目标表格样式
        statistics: 统计信息字典，用于累计结果
    """
    print("\n=== 开始分析文档中的表格 ===")

    # 获取文档中的所有表格
    tables = doc.get_all_tables()
    table_count = len(tables)
    print(f"文档中共有 {table_count} 个表格")

    if table_count == 0:
        print("没有找到表格，跳过表格样式分析")
        return

    # 用于收集表格样式匹配情况的统计信息
    table_matches = {
        "total": table_count,
        "matching": 0,
        "header_row_matching": 0,
        "data_row_matching": 0,
        "three_line_matching": 0
    }

    # 分析每个表格
    for i in range(table_count):
        print(f"\n>>> 表格 {i} <<<")

        # 提取表格的行样式
        extracted_styles = extract_table_row_styles(doc, i)
        print("\n提取的表格样式:")
        print(json.dumps(extracted_styles, indent=2, ensure_ascii=False))

        print("\n目标表格样式:")
        print(json.dumps(table_styles, indent=2, ensure_ascii=False))

        # 比较表格样式
        match_results = compare_table_styles(extracted_styles, table_styles, prefix="  ")

        # 将表格元素和错误信息添加到statistics['elements']中
        if 'results' in match_results:
            current_table = tables[i]
            table_element = current_table['element'] if isinstance(current_table,
                                                                   dict) and 'element' in current_table else current_table

            # 添加表格元素和错误信息到statistics
            if statistics is not None and 'elements' in statistics and len(match_results['results']) > 0:
                statistics['elements'].append({
                    "type": "table",
                    'element': table_element,
                    'result': match_results['results'],
                    'index': i  # 添加表格索引
                })

        # 更新统计
        if match_results["overall_match"]:
            table_matches["matching"] += 1

        if match_results["列名行文本样式"]["match"]:
            table_matches["header_row_matching"] += 1

        if match_results["数据行文本样式"]["match"]:
            table_matches["data_row_matching"] += 1

        if match_results["is_three_line_table"]["match"]:
            table_matches["three_line_matching"] += 1

        # 打印比较总结
        print("\n  表格样式比较总结:")
        print(f"  整体匹配: {'是' if match_results['overall_match'] else '否'}")
        print(f"  列名行匹配: {'是' if match_results['列名行文本样式']['match'] else '否'}")
        print(f"  数据行匹配: {'是' if match_results['数据行文本样式']['match'] else '否'}")
        print(f"  三线表匹配: {'是' if match_results['is_three_line_table']['match'] else '否'}")

        # 如果有差异，打印差异详情
        if not match_results["overall_match"]:
            print("\n  差异详情:")

            if not match_results["列名行文本样式"]["match"]:
                print("  列名行差异:")
                for key, diff in match_results["列名行文本样式"]["differences"].items():
                    print(f"    {key}: 目标值={diff['target']}, 实际值={diff['extracted']}")

            if not match_results["数据行文本样式"]["match"]:
                print("  数据行差异:")
                for key, diff in match_results["数据行文本样式"]["differences"].items():
                    print(f"    {key}: 目标值={diff['target']}, 实际值={diff['extracted']}")

    # 打印表格分析总结
    print("\n=== 表格样式分析总结 ===")
    print(f"总表格数: {table_count}")
    print(
        f"完全匹配的表格: {table_matches['matching']}/{table_count} ({table_matches['matching'] / table_count * 100:.1f}%)")
    print(
        f"列名行匹配的表格: {table_matches['header_row_matching']}/{table_count} ({table_matches['header_row_matching'] / table_count * 100:.1f}%)")
    print(
        f"数据行匹配的表格: {table_matches['data_row_matching']}/{table_count} ({table_matches['data_row_matching'] / table_count * 100:.1f}%)")
    print(
        f"三线表状态匹配的表格: {table_matches['three_line_matching']}/{table_count} ({table_matches['three_line_matching'] / table_count * 100:.1f}%)")

    # 如果统计信息不为None，更新全局统计
    if statistics is not None:
        statistics["table_matches"] = table_matches
        # 更新element_types计数
        statistics["element_types"] += 1


def extract_table_row_styles(doc, table_index):
    """
    提取表格的列名行（第一行）和数据行（第二行）的样式信息

    参数:
        doc: DocxElementParser实例
        table_index: 表格索引

    返回:
        dict: 包含列名行和数据行样式信息的字典
    """
    result = {
        "列名行文本样式": {},
        "数据行文本样式": {}
    }

    try:
        # 获取表格维度
        dims = doc.get_table_dimensions(table_index)
        if not dims or dims[0] < 2:  # 确保表格至少有两行
            print(f"表格 {table_index} 不存在或行数不足")
            return result

        rows, cols = dims

        # 创建样式分析器来获取完整样式
        style_analyzer = None
        if hasattr(doc, "docx_path"):
            try:
                from style_analyzer import StyleAnalyzer
                style_analyzer = StyleAnalyzer(doc.docx_path)
            except (ImportError, Exception) as e:
                print(f"无法创建StyleAnalyzer: {e}")

        # 处理列名行（第一行）
        row = 0
        col = 0  # 使用第一列作为样式参考

        # 获取单元格中的段落
        header_paragraphs = doc.get_table_cell_paragraphs(table_index, row, col)
        if header_paragraphs and style_analyzer:
            para = header_paragraphs[0]

            # 使用样式分析器获取完整样式
            try:
                # 获取段落的完整样式信息
                complete_style = style_analyzer.get_paragraph_complete_style_info(para)
                para_style = complete_style['effective_style']

                # 提取段落属性
                if 'paragraph_properties' in para_style:
                    pp = para_style['paragraph_properties']

                    # 提取对齐方式
                    if 'alignment' in pp:
                        result["列名行文本样式"]["alignment"] = pp['alignment']

                    # 提取行距
                    if 'spacing' in pp:
                        spacing = pp['spacing']
                        if 'line' in spacing:
                            result["列名行文本样式"]["line"] = int(spacing['line'])
                        if 'lineRule' in spacing:
                            result["列名行文本样式"]["line_rule"] = spacing['lineRule']

                # 提取run样式信息
                runs = doc.get_runs_from_paragraph(para)
                if runs:
                    run = runs[0]
                    # 使用style_analyzer获取run的完整样式信息
                    run_complete_style = style_analyzer.get_run_complete_style_info(para, run)
                    run_effective_style = run_complete_style['effective_style']

                    # 提取run_properties
                    if 'run_properties' in run_effective_style:
                        rp = run_effective_style['run_properties']

                        # 提取字体
                        if 'fonts' in rp:
                            fonts = rp['fonts']
                            if 'eastAsia' in fonts:
                                result["列名行文本样式"]["font_chinese"] = fonts['eastAsia']
                            if 'ascii' in fonts:
                                result["列名行文本样式"]["font_ascii"] = fonts['ascii']

                        # 提取字号
                        if 'size' in rp:
                            result["列名行文本样式"]["size"] = int(rp['size'])

                        # 提取加粗
                        if 'bold' in rp:
                            bold_value = rp['bold']
                            # 确保布尔值
                            if isinstance(bold_value, str):
                                bold_value = bold_value.lower() in ["true", "yes", "11.py"]
                            result["列名行文本样式"]["bold"] = bold_value

            except Exception as e:
                print(f"使用StyleAnalyzer提取列名行样式时出错: {e}")

        # 如果StyleAnalyzer不可用或提取失败，使用原始方法
        if not result["列名行文本样式"] and header_paragraphs:
            para = header_paragraphs[0]
            para_style = doc.get_paragraph_style_from_element(para)

            # 提取段落样式中的对齐方式
            if para_style and 'alignment' in para_style:
                result["列名行文本样式"]["alignment"] = para_style['alignment']

            # 提取段落样式中的行距
            if para_style and 'line' in para_style:
                result["列名行文本样式"]["line"] = int(para_style['line'])
                if 'line_rule' in para_style:
                    result["列名行文本样式"]["line_rule"] = para_style['line_rule']

            # 获取run元素的样式
            runs = doc.get_runs_from_paragraph(para)
            if runs:
                run = runs[0]
                run_style = doc.get_run_style_from_element(run)

                # 提取字体信息
                if 'fonts' in run_style:
                    font = run_style['fonts']
                    if 'eastAsia' in font:
                        result["列名行文本样式"]["font_chinese"] = font['eastAsia']
                    if 'ascii' in font:
                        result["列名行文本样式"]["font_ascii"] = font['ascii']

                # 提取字号
                if 'size' in run_style:
                    result["列名行文本样式"]["size"] = int(run_style['size'])

                # 提取加粗信息
                if 'bold' in run_style:
                    result["列名行文本样式"]["bold"] = run_style['bold']
                elif 'is_bold' in run_style:  # 兼容不同版本的属性名
                    result["列名行文本样式"]["bold"] = run_style['is_bold']

        # 处理数据行（第二行）
        if rows > 1:
            row = 1

            # 获取单元格中的段落
            data_paragraphs = doc.get_table_cell_paragraphs(table_index, row, col)
            if data_paragraphs and style_analyzer:
                para = data_paragraphs[0]

                # 使用样式分析器获取完整样式
                try:
                    # 获取段落的完整样式信息
                    complete_style = style_analyzer.get_paragraph_complete_style_info(para)
                    para_style = complete_style['effective_style']

                    # 提取段落属性
                    if 'paragraph_properties' in para_style:
                        pp = para_style['paragraph_properties']

                        # 提取对齐方式
                        if 'alignment' in pp:
                            result["数据行文本样式"]["alignment"] = pp['alignment']

                        # 提取行距
                        if 'spacing' in pp:
                            spacing = pp['spacing']
                            if 'line' in spacing:
                                result["数据行文本样式"]["line"] = int(spacing['line'])
                            if 'lineRule' in spacing:
                                result["数据行文本样式"]["line_rule"] = spacing['lineRule']

                    # 提取run样式信息
                    runs = doc.get_runs_from_paragraph(para)
                    if runs:
                        run = runs[0]
                        # 使用style_analyzer获取run的完整样式信息
                        run_complete_style = style_analyzer.get_run_complete_style_info(para, run)
                        run_effective_style = run_complete_style['effective_style']

                        # 提取run_properties
                        if 'run_properties' in run_effective_style:
                            rp = run_effective_style['run_properties']

                            # 提取字体
                            if 'fonts' in rp:
                                fonts = rp['fonts']
                                if 'eastAsia' in fonts:
                                    result["数据行文本样式"]["font_chinese"] = fonts['eastAsia']
                                if 'ascii' in fonts:
                                    result["数据行文本样式"]["font_ascii"] = fonts['ascii']

                            # 提取字号
                            if 'size' in rp:
                                result["数据行文本样式"]["size"] = int(rp['size'])

                            # 提取加粗
                            if 'bold' in rp:
                                bold_value = rp['bold']
                                # 确保布尔值
                                if isinstance(bold_value, str):
                                    bold_value = bold_value.lower() in ["true", "yes", "1"]
                                result["数据行文本样式"]["bold"] = bold_value

                except Exception as e:
                    print(f"使用StyleAnalyzer提取数据行样式时出错: {e}")

            # 如果StyleAnalyzer不可用或提取失败，使用原始方法
            if not result["数据行文本样式"] and data_paragraphs:
                para = data_paragraphs[0]
                para_style = doc.get_paragraph_style_from_element(para)

                # 提取段落样式中的对齐方式
                if para_style and 'alignment' in para_style:
                    result["数据行文本样式"]["alignment"] = para_style['alignment']

                # 提取段落样式中的行距
                if para_style and 'line' in para_style:
                    result["数据行文本样式"]["line"] = int(para_style['line'])
                    if 'line_rule' in para_style:
                        result["数据行文本样式"]["line_rule"] = para_style['line_rule']

                # 获取run元素的样式
                runs = doc.get_runs_from_paragraph(para)
                if runs:
                    run = runs[0]
                    run_style = doc.get_run_style_from_element(run)

                    # 提取字体信息
                    if 'fonts' in run_style:
                        font = run_style['fonts']
                        if 'eastAsia' in font:
                            result["数据行文本样式"]["font_chinese"] = font['eastAsia']
                        if 'ascii' in font:
                            result["数据行文本样式"]["font_ascii"] = font['ascii']

                    # 提取字号
                    if 'size' in run_style:
                        result["数据行文本样式"]["size"] = int(run_style['size'])

                    # 提取加粗信息
                    if 'bold' in run_style:
                        result["数据行文本样式"]["bold"] = run_style['bold']
                    elif 'is_bold' in run_style:  # 兼容不同版本的属性名
                        result["数据行文本样式"]["bold"] = run_style['is_bold']

        # 添加三线表判断
        three_line_result = is_three_line_table(doc, table_index)
        result["is_three_line_table"] = three_line_result["is_three_line_table"]

        # 设置默认值，确保关键属性不缺失
        for style_key in ["列名行文本样式", "数据行文本样式"]:
            # 如果没有提取到加粗信息，默认为False
            if "bold" not in result[style_key]:
                result[style_key]["bold"] = False
        print(f"提取的表格行样式1: {result}")
        return result

    except Exception as e:
        print(f"提取表格行样式时出错: {e}")
        return result

def is_three_line_table(doc, table_index):
        """
        判断指定索引的表格是否是标准三线表

        标准三线表具有以下特征:
        1. 表头上方有一条横线（通常较粗）
        2. 表头下方有一条横线（区分表头和数据区域）
        3. 表格底部有一条横线（通常较粗）
        4. 无内部水平分隔线（表头和底部之间）
        5. 无垂直分隔线

        参数:
            doc: DocxElementParser实例
            table_index: 表格的索引

        返回:
            dict: 包含判断结果和详细信息的字典
        """

        result = {
            "is_three_line_table": False,
            "reasons": [],
            "borders": {
                "header_top": None,
                "header_bottom": None,
                "table_bottom": None,
                "inner_h": None,
                "vertical": None
            }
        }

        try:
            # 获取表格
            tables = doc.get_all_tables()
            if table_index >= len(tables):
                result["reasons"].append(f"表格索引 {table_index} 超出范围")
                return result

            table = tables[table_index]
            if isinstance(table, dict) and 'element' in table:
                table = table['element']
            NAMESPACES = doc.NAMESPACES

            # 11.py. 检查表格全局边框设置
            tblBorders = table.find(f".//{{{NAMESPACES['w']}}}tblPr/{{{NAMESPACES['w']}}}tblBorders", NAMESPACES)
            if tblBorders is None:
                result["reasons"].append("表格没有定义边框")
                return result

            # 记录表格级别的顶部和底部边框
            top_border = tblBorders.find(f".//{{{NAMESPACES['w']}}}top", NAMESPACES)
            bottom_border = tblBorders.find(f".//{{{NAMESPACES['w']}}}bottom", NAMESPACES)
            table_top = None
            table_bottom = None

            if top_border is not None:
                table_top = {
                    "val": top_border.get(f"{{{NAMESPACES['w']}}}val"),
                    "size": top_border.get(f"{{{NAMESPACES['w']}}}sz")
                }

            if bottom_border is not None:
                table_bottom = {
                    "val": bottom_border.get(f"{{{NAMESPACES['w']}}}val"),
                    "size": bottom_border.get(f"{{{NAMESPACES['w']}}}sz")
                }
                # 先记录表格级别的底部边框
                if table_bottom["val"] != "none":
                    result["borders"]["table_bottom"] = table_bottom

            # 检查全局内部水平线和垂直线设置
            inside_h = tblBorders.find(f".//{{{NAMESPACES['w']}}}insideH", NAMESPACES)
            inside_v = tblBorders.find(f".//{{{NAMESPACES['w']}}}insideV", NAMESPACES)

            # 记录全局内部线设置
            if inside_h is not None:
                result["borders"]["inner_h"] = {
                    "val": inside_h.get(f"{{{NAMESPACES['w']}}}val"),
                    "size": inside_h.get(f"{{{NAMESPACES['w']}}}sz")
                }

            if inside_v is not None:
                result["borders"]["vertical"] = {
                    "val": inside_v.get(f"{{{NAMESPACES['w']}}}val"),
                    "size": inside_v.get(f"{{{NAMESPACES['w']}}}sz")
                }

            # 2. 检查表格行
            rows = table.findall(f".//{{{NAMESPACES['w']}}}tr", NAMESPACES)
            if not rows:
                result["reasons"].append("表格没有行")
                return result

            # 3. 检查第一行（表头）的单元格边框
            if len(rows) > 0:
                first_row_cells = rows[0].findall(f".//{{{NAMESPACES['w']}}}tc", NAMESPACES)
                if first_row_cells:
                    first_cell = first_row_cells[0]

                    # 检查表头顶部边框
                    top_border = first_cell.find(
                        f".//{{{NAMESPACES['w']}}}tcPr/{{{NAMESPACES['w']}}}tcBorders/{{{NAMESPACES['w']}}}top", NAMESPACES)
                    if top_border is not None:
                        result["borders"]["header_top"] = {
                            "val": top_border.get(f"{{{NAMESPACES['w']}}}val"),
                            "size": top_border.get(f"{{{NAMESPACES['w']}}}sz")
                        }
                    elif table_top is not None and table_top["val"] != "none":
                        # 如果单元格没有定义，但表格有定义，则使用表格定义
                        result["borders"]["header_top"] = table_top

                    # 检查表头底部边框
                    bottom_border = first_cell.find(
                        f".//{{{NAMESPACES['w']}}}tcPr/{{{NAMESPACES['w']}}}tcBorders/{{{NAMESPACES['w']}}}bottom",
                        NAMESPACES)
                    if bottom_border is not None:
                        result["borders"]["header_bottom"] = {
                            "val": bottom_border.get(f"{{{NAMESPACES['w']}}}val"),
                            "size": bottom_border.get(f"{{{NAMESPACES['w']}}}sz")
                        }

            # 4. 检查最后一行的边框设置
            if len(rows) > 0:
                last_row = rows[-1]
                last_row_cells = last_row.findall(f".//{{{NAMESPACES['w']}}}tc", NAMESPACES)

                # 检查行级别的边框继承
                row_borders = last_row.find(f".//{{{NAMESPACES['w']}}}tblPrEx/{{{NAMESPACES['w']}}}tblBorders", NAMESPACES)
                if row_borders is not None:
                    row_bottom = row_borders.find(f".//{{{NAMESPACES['w']}}}bottom", NAMESPACES)
                    if row_bottom is not None:
                        bottom_val = row_bottom.get(f"{{{NAMESPACES['w']}}}val")
                        bottom_size = row_bottom.get(f"{{{NAMESPACES['w']}}}sz")
                        if bottom_val != "none":
                            result["borders"]["table_bottom"] = {
                                "val": bottom_val,
                                "size": bottom_size
                            }

                # 如果行级别没有定义，检查单元格级别
                if result["borders"]["table_bottom"] is None and last_row_cells:
                    last_cell = last_row_cells[0]
                    cell_bottom = last_cell.find(
                        f".//{{{NAMESPACES['w']}}}tcPr/{{{NAMESPACES['w']}}}tcBorders/{{{NAMESPACES['w']}}}bottom",
                        NAMESPACES)
                    if cell_bottom is not None:
                        bottom_val = cell_bottom.get(f"{{{NAMESPACES['w']}}}val")
                        bottom_size = cell_bottom.get(f"{{{NAMESPACES['w']}}}sz")
                        if bottom_val != "none":
                            result["borders"]["table_bottom"] = {
                                "val": bottom_val,
                                "size": bottom_size
                            }

            # 5. 判断是否符合三线表特征
            is_three_line = True
            reasons = []

            # 检查表头顶部是否有边框
            if result["borders"]["header_top"] is None or result["borders"]["header_top"]["val"] == "none":
                is_three_line = False
                reasons.append("表头顶部没有边框")

            # 检查表头底部是否有边框
            if result["borders"]["header_bottom"] is None or result["borders"]["header_bottom"]["val"] == "none":
                is_three_line = False
                reasons.append("表头底部没有边框（表头与数据区域之间）")

            # 检查表格底部是否有边框
            if result["borders"]["table_bottom"] is None or result["borders"]["table_bottom"]["val"] == "none":
                is_three_line = False
                reasons.append("表格底部没有边框")

            # 检查内部水平线是否为空
            if result["borders"]["inner_h"] is not None and result["borders"]["inner_h"]["val"] != "none":
                is_three_line = False
                reasons.append("表格有内部水平分隔线（表头和底部之间）")

            # 检查垂直线是否为空
            if result["borders"]["vertical"] is not None and result["borders"]["vertical"]["val"] != "none":
                is_three_line = False
                reasons.append("表格有垂直分隔线")

            result["is_three_line_table"] = is_three_line
            result["reasons"] = reasons if reasons else ["符合三线表标准"]

            # 添加表格的行列数据
            dims = doc.get_table_dimensions(table_index)
            if dims:
                result["rows"] = dims[0]
                result["columns"] = dims[1]

            return result

        except Exception as e:
            result["reasons"].append(f"分析时出错: {str(e)}")
            return result

def compare_table_styles(extracted_styles, target_styles, prefix=""):
    match_results = {
        'results': [],
        "overall_match": True,
        "列名行文本样式": {"match": True, "differences": {}},
        "数据行文本样式": {"match": True, "differences": {}},
        "is_three_line_table": {"match": True}
    }

    print(f"{prefix}比较表格样式:")

    # 检查三线表匹配
    if "is_three_line_table" in extracted_styles and "is_three_line_table" in target_styles:
        is_match = extracted_styles["is_three_line_table"] == target_styles["is_three_line_table"]
        match_results["is_three_line_table"]["match"] = is_match
        match_indicator = "✓" if is_match else "✗"
        print(f"{prefix}{match_indicator} is_three_line_table: 目标值={target_styles['is_three_line_table']}, 实际值={extracted_styles['is_three_line_table']}")
        if match_indicator == "✗":
            match_results['results'].append({'is_three_line_table': {'scuccess': target_styles['is_three_line_table'],
                                                      'error': extracted_styles['is_three_line_table']}})

        if not is_match:
            match_results["overall_match"] = False

    # 字体相似性检查函数
    def are_fonts_similar(font1, font2):
        """检查两个字体名称是否相似或兼容"""
        if font1 == font2:
            return True

        # 中文字体相似性检查
        cn_font_families = {
            "宋体": ["宋体", "SimSun", "宋"],
            "黑体": ["黑体", "SimHei", "黑"],
            "楷体": ["楷体", "KaiTi", "楷"],
            "仿宋": ["仿宋", "FangSong", "仿"],
            "微软雅黑": ["微软雅黑", "Microsoft YaHei", "雅黑"]
        }

        # 英文字体相似性检查
        en_font_families = {
            "Times New Roman": ["Times New Roman", "Times", "TNR"],
            "Arial": ["Arial", "Helvetica"],
            "Calibri": ["Calibri"],
            "Courier New": ["Courier New", "Courier"]
        }

        # 检查中文字体家族
        for family, variants in cn_font_families.items():
            if font1 in variants and font2 in variants:
                return True

        # 检查英文字体家族
        for family, variants in en_font_families.items():
            if font1 in variants and font2 in variants:
                return True

        return False

    # 比较列名行样式
    print(f"{prefix}列名行文本样式:")
    for key in ["font_chinese", "font_ascii", "size", "bold", "alignment", "line", "line_rule"]:
        if key in target_styles["列名行文本样式"]:
            target_value = target_styles["列名行文本样式"][key]

            # 如果目标值是"unknown"，则任何值或缺失值都匹配
            if target_value == "unknown":
                if key in extracted_styles["列名行文本样式"]:
                    extracted_value = extracted_styles["列名行文本样式"][key]
                    print(f"{prefix}  ✓ {key}: 目标值=未知, 实际值={extracted_value} (任意值都匹配)")
                else:
                    print(f"{prefix}  ✓ {key}: 目标值=未知, 实际值=未定义 (任意值都匹配)")
                continue

            # 处理提取值缺失的情况
            if key not in extracted_styles["列名行文本样式"]:
                # 对于加粗属性，缺失通常意味着"false"
                if key == "bold" and target_value is False:
                    print(f"{prefix}  ✓ {key}: 目标值={target_value}, 实际值=未定义 (默认为False)")
                    continue

                print(f"{prefix}  ! {key}: 目标值={target_value}, 实际值=未定义")
                match_results["列名行文本样式"]["differences"][key] = {"target": target_value, "extracted": "未定义"}
                match_results["列名行文本样式"]["match"] = False
                match_results["overall_match"] = False
                match_results['results'].append({'column_header': {'attribute': key, 'scuccess': target_value}})
                continue

            extracted_value = extracted_styles["列名行文本样式"][key]
            is_match = False

            # 特殊处理各属性的比较
            if key in ["font_chinese", "font_ascii"]:
                # 使用字体相似性检查
                is_match = are_fonts_similar(extracted_value, target_value)
                match_indicator = "✓" if is_match else "✗"
                print(f"{prefix}  {match_indicator} {key}: 目标值={target_value}, 实际值={extracted_value}")
                if match_indicator == "✗":
                    match_results['results'].append(
                        {'column_header': {'attribute': key, 'scuccess': target_value, 'error': extracted_value}})

            elif key == "size":
                try:
                    # 特殊处理字号对应关系
                    if abs(extracted_value - target_value) <= 1:
                        is_match = True
                        print(f"{prefix}  ✓ {key}: 目标值={target_value}, 实际值={extracted_value} ")
                    else:
                        print(f"{prefix}  x {key}: 目标值={target_value}, 实际值={extracted_value}")
                        match_results['results'].append(
                            {'column_header': {'attribute': key, 'scuccess': target_value, 'error': extracted_value}})
                except (ValueError, TypeError):
                    is_match = str(extracted_value) == str(target_value)
                    match_indicator = "✓" if is_match else "✗"
                    print(f"{prefix}  {match_indicator} {key}: 目标值={target_value}, 实际值={extracted_value}")
                    if match_indicator == "✗":
                        match_results['results'].append(
                            {'column_header': {'attribute': key, 'scuccess': target_value, 'error': extracted_value}})
            elif key == "line":
                try:
                    # 允许±30个twip的误差
                    target_line = target_value if isinstance(target_value, int) else int(target_value)
                    extracted_line = extracted_value if isinstance(extracted_value, int) else int(extracted_value)
                    is_match = abs(extracted_line - target_line) <= 30
                    match_indicator = "✓" if is_match else "✗"
                    print(f"{prefix}  {match_indicator} {key}: 目标值={target_value}, 实际值={extracted_value}")
                    if match_indicator == "✗":
                        match_results['results'].append(
                            {'column_header': {'attribute': key, 'scuccess': target_value, 'error': extracted_value}})
                except (ValueError, TypeError):
                    is_match = str(extracted_value) == str(target_value)
                    match_indicator = "✓" if is_match else "✗"
                    print(f"{prefix}  {match_indicator} {key}: 目标值={target_value}, 实际值={extracted_value}")
                    if match_indicator == "✗":
                        match_results['results'].append(
                            {'column_header': {'attribute': key, 'scuccess': target_value, 'error': extracted_value}})
            elif key == "alignment":
                # 现在要求对齐方式必须完全一致
                is_match = extracted_value == target_value
                match_indicator = "✓" if is_match else "✗"
                print(f"{prefix}  {match_indicator} {key}: 目标值={target_value}, 实际值={extracted_value}")
                if match_indicator == "✗":
                    match_results['results'].append(
                        {'column_header': {'attribute': key, 'scuccess': target_value, 'error': extracted_value}})
            elif key == "line_rule":
                # 特殊处理行距规则，auto和exact在某些情况下可互换
                if (target_value == "auto" and extracted_value == "exact") or \
                   (target_value == "exact" and extracted_value == "auto"):
                    # 检查行距值是否接近
                    if "line" in target_styles["列名行文本样式"] and "line" in extracted_styles["列名行文本样式"]:
                        try:
                            target_line = int(target_styles["列名行文本样式"]["line"])
                            extracted_line = int(extracted_styles["列名行文本样式"]["line"])
                            if abs(target_line - extracted_line) <= 30:
                                is_match = True
                                print(f"{prefix}  ✓ {key}: 目标值={target_value}, 实际值={extracted_value} (考虑行距值后视为等效)")
                            else:
                                is_match = False
                                match_indicator = "✗"
                                print(f"{prefix}  {match_indicator} {key}: 目标值={target_value}, 实际值={extracted_value}")
                                match_results['results'].append({'column_header': {'attribute': key,
                                                                                   'scuccess': target_value,
                                                                                   'error': extracted_value}})
                        except (ValueError, TypeError):
                            is_match = False
                            match_indicator = "✗"
                            print(f"{prefix}  {match_indicator} {key}: 目标值={target_value}, 实际值={extracted_value}")
                            match_results['results'].append({'column_header': {'attribute': key,
                                                                               'scuccess': target_value,
                                                                               'error': extracted_value}})
                    else:
                        is_match = True
                        print(f"{prefix}  ✓ {key}: 目标值={target_value}, 实际值={extracted_value} (auto和exact视为等效)")
                else:
                    is_match = extracted_value == target_value
                    match_indicator = "✓" if is_match else "✗"
                    print(f"{prefix}  {match_indicator} {key}: 目标值={target_value}, 实际值={extracted_value}")
                    if match_indicator == "✗":
                        match_results['results'].append(
                            {'column_header': {'attribute': key, 'scuccess': target_value, 'error': extracted_value}})
            else:
                # 其他属性直接比较
                is_match = extracted_value == target_value
                match_indicator = "✓" if is_match else "✗"
                print(f"{prefix}  {match_indicator} {key}: 目标值={target_value}, 实际值={extracted_value}")
                if match_indicator == "✗":
                    match_results['results'].append(
                        {'column_header': {'attribute': key, 'scuccess': target_value, 'error': extracted_value}})

            if not is_match:
                match_results["列名行文本样式"]["differences"][key] = {
                    "target": target_value,
                    "extracted": extracted_value
                }
                match_results["列名行文本样式"]["match"] = False
                match_results["overall_match"] = False

    # 比较数据行样式 - 使用与列名行相同的逻辑
    print(f"{prefix}数据行文本样式:")
    for key in ["font_chinese", "font_ascii", "size", "bold", "alignment", "line", "line_rule"]:
        if key in target_styles["数据行文本样式"]:
            target_value = target_styles["数据行文本样式"][key]

            # 如果目标值是"unknown"，则任何值或缺失值都匹配
            if target_value == "unknown":
                if key in extracted_styles["数据行文本样式"]:
                    extracted_value = extracted_styles["数据行文本样式"][key]
                    print(f"{prefix}  ✓ {key}: 目标值=未知, 实际值={extracted_value} (任意值都匹配)")
                else:
                    print(f"{prefix}  ✓ {key}: 目标值=未知, 实际值=未定义 (任意值都匹配)")
                continue

            # 处理提取值缺失的情况
            if key not in extracted_styles["数据行文本样式"]:
                # 对于加粗属性，缺失通常意味着"false"
                if key == "bold" and target_value is False:
                    print(f"{prefix}  ✓ {key}: 目标值={target_value}, 实际值=未定义 (默认为False)")
                    continue

                print(f"{prefix}  ! {key}: 目标值={target_value}, 实际值=未定义")
                match_results["数据行文本样式"]["differences"][key] = {"target": target_value, "extracted": "未定义"}
                match_results["数据行文本样式"]["match"] = False
                match_results["overall_match"] = False
                match_results['results'].append({'data_row': {'attribute': key, 'scuccess': target_value}})
                continue

            extracted_value = extracted_styles["数据行文本样式"][key]
            is_match = False

            # 特殊处理各属性的比较
            if key in ["font_chinese", "font_ascii"]:
                # 使用字体相似性检查
                is_match = are_fonts_similar(extracted_value, target_value)
                match_indicator = "✓" if is_match else "✗"
                print(f"{prefix}  {match_indicator} {key}: 目标值={target_value}, 实际值={extracted_value}")
                if match_indicator == "✗":
                    match_results['results'].append(
                        {'data_row': {'attribute': key, 'scuccess': target_value, 'error': extracted_value}})
            elif key == "size":
                try:
                    # 特殊处理字号对应关系
                    if abs(extracted_value - target_value) <= 1:
                        is_match = True
                        print(f"{prefix}  ✓ {key}: 目标值={target_value}, 实际值={extracted_value} ")
                    else:
                        print(f"{prefix}  x {key}: 目标值={target_value}, 实际值={extracted_value}")
                        match_results['results'].append(
                            {'data_row': {'attribute': key, 'scuccess': target_value, 'error': extracted_value}})
                except (ValueError, TypeError):
                    is_match = str(extracted_value) == str(target_value)
                    match_indicator = "✓" if is_match else "✗"
                    print(f"{prefix}  {match_indicator} {key}: 目标值={target_value}, 实际值={extracted_value}")
                    if match_indicator == "✗":
                        match_results['results'].append(
                            {'data_row': {'attribute': key, 'scuccess': target_value, 'error': extracted_value}})
            elif key == "line":
                try:
                    # 允许±30个twip的误差
                    target_line = target_value if isinstance(target_value, int) else int(target_value)
                    extracted_line = extracted_value if isinstance(extracted_value, int) else int(extracted_value)
                    is_match = abs(extracted_line - target_line) <= 30
                    match_indicator = "✓" if is_match else "✗"
                    print(f"{prefix}  {match_indicator} {key}: 目标值={target_value}, 实际值={extracted_value}")
                    if match_indicator == "✗":
                        match_results['results'].append(
                            {'data_row': {'attribute': key, 'scuccess': target_value, 'error': extracted_value}})
                except (ValueError, TypeError):
                    is_match = str(extracted_value) == str(target_value)
                    match_indicator = "✓" if is_match else "✗"
                    print(f"{prefix}  {match_indicator} {key}: 目标值={target_value}, 实际值={extracted_value}")
                    if match_indicator == "✗":
                        match_results['results'].append(
                            {'data_row': {'attribute': key, 'scuccess': target_value, 'error': extracted_value}})
            elif key == "alignment":
                # 特殊处理对齐方式，left和both视为等效
                if target_value in ["left", "both"]:
                    is_match = extracted_value in ["left", "both"]
                else:
                    is_match = extracted_value == target_value
                match_indicator = "✓" if is_match else "✗"
                print(f"{prefix}  {match_indicator} {key}: 目标值={target_value}, 实际值={extracted_value}")
                if match_indicator == "✗":
                    match_results['results'].append(
                        {'data_row': {'attribute': key, 'scuccess': target_value, 'error': extracted_value}})
            elif key == "line_rule":
                # 特殊处理行距规则，auto和exact在某些情况下可互换
                if (target_value == "auto" and extracted_value == "exact") or \
                   (target_value == "exact" and extracted_value == "auto"):
                    # 检查行距值是否接近
                    if "line" in target_styles["数据行文本样式"] and "line" in extracted_styles["数据行文本样式"]:
                        try:
                            target_line = int(target_styles["数据行文本样式"]["line"])
                            extracted_line = int(extracted_styles["数据行文本样式"]["line"])
                            if abs(target_line - extracted_line) <= 30:
                                is_match = True
                                print(f"{prefix}  ✓ {key}: 目标值={target_value}, 实际值={extracted_value} (考虑行距值后视为等效)")
                            else:
                                is_match = False
                                match_indicator = "✗"
                                print(f"{prefix}  {match_indicator} {key}: 目标值={target_value}, 实际值={extracted_value}")
                                match_results['results'].append({'data_row': {'attribute': key,
                                                                              'scuccess': target_value,
                                                                              'error': extracted_value}})
                        except (ValueError, TypeError):
                            is_match = False
                            match_indicator = "✗"
                            print(f"{prefix}  {match_indicator} {key}: 目标值={target_value}, 实际值={extracted_value}")
                            match_results['results'].append(
                                {'data_row': {'attribute': key, 'scuccess': target_value, 'error': extracted_value}})
                    else:
                        is_match = True
                        print(f"{prefix}  ✓ {key}: 目标值={target_value}, 实际值={extracted_value} (auto和exact视为等效)")
                else:
                    is_match = extracted_value == target_value
                    match_indicator = "✓" if is_match else "✗"
                    print(f"{prefix}  {match_indicator} {key}: 目标值={target_value}, 实际值={extracted_value}")
                    if match_indicator == "✗":
                        match_results['results'].append(
                            {'data_row': {'attribute': key, 'scuccess': target_value, 'error': extracted_value}})
            else:
                # 其他属性直接比较
                is_match = extracted_value == target_value
                match_indicator = "✓" if is_match else "✗"
                print(f"{prefix}  {match_indicator} {key}: 目标值={target_value}, 实际值={extracted_value}")
                if match_indicator == "✗":
                    match_results['results'].append(
                        {'data_row': {'attribute': key, 'scuccess': target_value, 'error': extracted_value}})

            if not is_match:
                match_results["数据行文本样式"]["differences"][key] = {
                    "target": target_value,
                    "extracted": extracted_value
                }
                match_results["数据行文本样式"]["match"] = False
                match_results["overall_match"] = False

    return match_results


if __name__ == "__main__":
    # 文件路径
    doc_path = "1_fixed.docx"
    classification_path = "document_classification_results.json"
    style_mapping_path = "document_style_mapping.json"
    api_params_path = "智算工程学院毕业设计（论文）模板2025届(1)_api_params.json"

    # 加载分类结果
    with open(classification_path, 'r', encoding='utf-8') as f:
        classification = json.load(f)

    # 加载样式映射
    with open(style_mapping_path, 'r', encoding='utf-8') as f:
        style_mapping = json.load(f)

    # 加载API参数
    with open(api_params_path, 'r', encoding='utf-8') as f:
        api_params = json.load(f)
    # 检查文件是否存在
    missing_files = []
    for path in [doc_path, classification_path, style_mapping_path, api_params_path]:
        if not os.path.exists(path):
            missing_files.append(path)

    if missing_files:
        print("错误: 以下文件不存在:")
        for path in missing_files:
            print(f"  - {path}")
    else:
        # 执行样式比较
      compare_styles(doc_path, classification, style_mapping_path, api_params)