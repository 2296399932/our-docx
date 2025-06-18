#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
检查指定段落的有效样式和目标样式
"""

import json
import os
from docx_namespace import DocxElementParser
from style_analyzer import StyleAnalyzer
from compare_styles import compare_paragraph_style, merge_run_styles

def check_paragraph_style(doc_path, classification_path, style_mapping_path, api_params_path, para_index=68):
    """
    检查特定段落的有效样式和目标样式
    
    参数:
        doc_path: Word文档路径
        classification_path: 分类结果JSON文件路径
        style_mapping_path: 样式映射JSON文件路径
        api_params_path: API参数格式JSON文件路径
        para_index: 要检查的段落索引
    """
    # 加载文档
    doc = DocxElementParser(doc_path)
    # 创建样式分析器实例获取完整样式信息
    style_analyzer = StyleAnalyzer(doc_path)
    
    # 加载分类结果
    with open(classification_path, 'r', encoding='utf-8') as f:
        classification = json.load(f)
    
    # 加载样式映射
    with open(style_mapping_path, 'r', encoding='utf-8') as f:
        style_mapping = json.load(f)
    
    # 加载API参数
    with open(api_params_path, 'r', encoding='utf-8') as f:
        api_params = json.load(f)
    
    # 先从md.py的方式获取有效样式
    print("=== 使用md.py的方式获取有效样式 ===")
    md_complete_style_info = style_analyzer.get_paragraph_complete_style_info(style_analyzer.elements[para_index]['element'])
    md_actual_para_style = md_complete_style_info['effective_style']
    print("有效样式 (md.py方式):")
    print(json.dumps(md_actual_para_style, indent=2, ensure_ascii=False))
    
    print("\n=== 使用compare_styles.py的逻辑获取有效样式和目标样式 ===")
    
    # 首先确定段落属于哪个元素类型
    element_type = None
    style_class = None
    element_class = None
    section = None
    
    # 搜索分类结果，找到包含该段落索引的元素类型
    for class_name, para_indices in classification.items():
        if para_index in para_indices:
            element_class = class_name
            break
    
    if element_class:
        print(f"索引 {para_index} 的段落属于元素类型: {element_class}")
        
        # 从样式映射中查找该元素类型的样式类
        for sec, section_mapping in style_mapping.items():
            if element_class in section_mapping:
                section = sec
                style_class = section_mapping[element_class]
                break
        
        if style_class and section:
            print(f"对应的样式类型: {section}.{style_class}")
            
            # 从API参数中获取目标样式
            if section in api_params and style_class in api_params[section]:
                target_style = api_params[section][style_class]
                print("目标样式:")
                print(json.dumps(target_style, indent=2, ensure_ascii=False))
                
                # 获取段落的有效样式
                para_element = doc.elements[para_index]['element'] if isinstance(doc.elements[para_index], dict) and 'element' in doc.elements[para_index] else doc.elements[para_index]
                
                # 使用StyleAnalyzer获取完整样式信息
                complete_style_info = style_analyzer.get_paragraph_complete_style_info(para_element)
                actual_para_style = complete_style_info['effective_style']
                
                print("\n有效样式 (compare_styles.py方式):")
                print(json.dumps(actual_para_style, indent=2, ensure_ascii=False))
                
                # 比较段落样式
                print("\n样式比较结果:")
                para_results = compare_paragraph_style(actual_para_style, target_style, prefix="  ")
                
                # 检查是否匹配
                is_matching = all(matched for attr, matched in para_results.items() if attr != 'results')
                print(f"\n段落样式是否完全匹配: {'是' if is_matching else '否'}")
                
                # 如果有results字段，输出详细的错误信息
                if 'results' in para_results and para_results['results']:
                    print("\n详细错误信息:")
                    for error_dict in para_results['results']:
                        for attr, values in error_dict.items():
                            print(f"  属性: {attr}")
                            if 'scuccess' in values:
                                print(f"    正确值: {values['scuccess']}")
                            if 'error' in values:
                                print(f"    当前值: {values['error']}")
            else:
                print(f"API参数中没有找到 {section}.{style_class} 的样式定义")
        else:
            print(f"样式映射中没有找到元素类型 {element_class} 的映射")
    else:
        print(f"分类结果中没有找到包含段落索引 {para_index} 的元素类型")
    
    # 检查相同点和不同点
    print("\n=== md.py与compare_styles.py获取的有效样式比较 ===")
    print("相同点:")
    for key in md_actual_para_style.keys() & actual_para_style.keys():
        if key == 'run_properties' or key == 'paragraph_properties':
            continue
        if md_actual_para_style[key] == actual_para_style[key]:
            print(f"  {key}: {md_actual_para_style[key]}")
    
    print("\n不同点:")
    for key in md_actual_para_style.keys() | actual_para_style.keys():
        if key not in md_actual_para_style:
            print(f"  {key}: 仅在compare_styles.py中存在, 值为 {actual_para_style[key]}")
        elif key not in actual_para_style:
            print(f"  {key}: 仅在md.py中存在, 值为 {md_actual_para_style[key]}")
        elif md_actual_para_style[key] != actual_para_style[key] and key not in ('run_properties', 'paragraph_properties'):
            print(f"  {key}: md.py中为 {md_actual_para_style[key]}, compare_styles.py中为 {actual_para_style[key]}")
    
    # 比较run_properties
    if 'run_properties' in md_actual_para_style and 'run_properties' in actual_para_style:
        print("\nrun_properties比较:")
        md_run_props = md_actual_para_style['run_properties']
        actual_run_props = actual_para_style['run_properties']
        
        print("相同点:")
        for key in md_run_props.keys() & actual_run_props.keys():
            if md_run_props[key] == actual_run_props[key]:
                print(f"  {key}: {md_run_props[key]}")
        
        print("\n不同点:")
        for key in md_run_props.keys() | actual_run_props.keys():
            if key not in md_run_props:
                print(f"  {key}: 仅在compare_styles.py中存在, 值为 {actual_run_props[key]}")
            elif key not in actual_run_props:
                print(f"  {key}: 仅在md.py中存在, 值为 {md_run_props[key]}")
            elif md_run_props[key] != actual_run_props[key]:
                print(f"  {key}: md.py中为 {md_run_props[key]}, compare_styles.py中为 {actual_run_props[key]}")
    
    # 比较paragraph_properties
    if 'paragraph_properties' in md_actual_para_style and 'paragraph_properties' in actual_para_style:
        print("\nparagraph_properties比较:")
        md_para_props = md_actual_para_style['paragraph_properties']
        actual_para_props = actual_para_style['paragraph_properties']
        
        print("相同点:")
        for key in md_para_props.keys() & actual_para_props.keys():
            if isinstance(md_para_props[key], (dict, list)):
                # 对于复杂结构，简单判断是否相同
                is_same = md_para_props[key] == actual_para_props[key]
                print(f"  {key}: {'相同' if is_same else '不同'}")
            elif md_para_props[key] == actual_para_props[key]:
                print(f"  {key}: {md_para_props[key]}")
        
        print("\n不同点:")
        for key in md_para_props.keys() | actual_para_props.keys():
            if key not in md_para_props:
                print(f"  {key}: 仅在compare_styles.py中存在")
            elif key not in actual_para_props:
                print(f"  {key}: 仅在md.py中存在")
            elif md_para_props[key] != actual_para_props[key]:
                if isinstance(md_para_props[key], (dict, list)) or isinstance(actual_para_props[key], (dict, list)):
                    print(f"  {key}: 值不同 (复杂结构)")
                else:
                    print(f"  {key}: md.py中为 {md_para_props[key]}, compare_styles.py中为 {actual_para_props[key]}")

if __name__ == "__main__":
    # 文件路径
    doc_path = "1.docx"
    classification_path = "document_classification_results.json"
    style_mapping_path = "document_style_mapping.json"
    api_params_path = "智算工程学院毕业设计（论文）模板2025届(1)_api_params.json"
    para_index = 68
    
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
        # 执行段落样式检查
        check_paragraph_style(doc_path, classification_path, style_mapping_path, api_params_path, para_index) 