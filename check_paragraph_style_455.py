#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
检测索引为455的段落的所有样式信息
"""

import os
import json
from docx_namespace import DocxElementParser
from style_analyzer import StyleAnalyzer  # 如果有这个模块

# 设置要分析的文档路径
DOC_PATH = "1.docx"  # 替换为实际文档路径
PARA_INDEX = 455  # 要检查的段落索引

def print_header(title):
    """打印带格式的标题"""
    print("\n" + "=" * 50)
    print(f" {title} ".center(50, "="))
    print("=" * 50)

def print_dict(data, indent=0):
    """美观地打印嵌套字典"""
    for key, value in data.items():
        if isinstance(value, dict):
            print(" " * indent + f"{key}:")
            print_dict(value, indent + 2)
        else:
            print(" " * indent + f"{key}: {value}")

def main():
    """主函数，检测指定段落的样式"""
    if not os.path.exists(DOC_PATH):
        print(f"错误: 文档 '{DOC_PATH}' 不存在")
        return

    # 创建DocxElementParser实例
    doc = DocxElementParser(DOC_PATH)
    
    # 检查段落索引是否有效
    if PARA_INDEX >= len(doc.elements):
        print(f"错误: 段落索引 {PARA_INDEX} 超出文档范围 (0-{len(doc.elements)-1})")
        return
    
    # 获取段落文本，验证是否为所需段落
    element_text = doc.get_element_text(PARA_INDEX)
    print(f"正在分析段落 {PARA_INDEX}:")
    print(f"段落文本: {element_text[:100]}{'...' if len(element_text) > 100 else ''}")
    
    print_header("段落基本信息")
    
    # 尝试获取元素类型
    element = doc.elements[PARA_INDEX]
    element_type = "paragraph"
    if isinstance(element, dict) and "type" in element:
        element_type = element["type"]
    print(f"元素类型: {element_type}")
    
    # 获取段落ID
    para_element = element
    if isinstance(element, dict) and "element" in element:
        para_element = element["element"]
    
    para_id = para_element.get(f'{{{doc.NAMESPACES["w14"]}}}paraId')
    print(f"段落ID: {para_id}")
    
    # 获取样式ID
    style_id = None
    pPr = para_element.find(f'.//{{{doc.NAMESPACES["w"]}}}pPr', doc.NAMESPACES)
    if pPr is not None:
        pStyle = pPr.find(f'.//{{{doc.NAMESPACES["w"]}}}pStyle', doc.NAMESPACES)
        if pStyle is not None:
            style_id = pStyle.get(f'{{{doc.NAMESPACES["w"]}}}val')
    print(f"样式ID: {style_id}")
    
    print_header("段落样式属性")
    
    # 获取段落对齐方式
    alignment = doc.get_paragraph_alignment(PARA_INDEX)
    print(f"对齐方式: {alignment}")
    
    # 获取段落缩进
    indentation = doc.get_paragraph_indentation(PARA_INDEX)
    print("缩进设置:")
    print_dict(indentation, 2)
    
    # 获取段落间距
    spacing = doc.get_paragraph_spacing(PARA_INDEX)
    print("间距设置:")
    print_dict(spacing, 2)
    
    # 获取段落边框
    borders = doc.get_paragraph_borders(PARA_INDEX)
    print("边框设置:")
    print_dict(borders, 2)
    
    # 获取段落底纹
    shading = doc.get_paragraph_shading(PARA_INDEX)
    print("底纹设置:")
    print_dict(shading, 2)
    
    # 获取段落编号
    numbering = doc.get_paragraph_numbering(PARA_INDEX)
    print("编号设置:")
    print_dict(numbering, 2)
    
    # 获取段落字体
    font = doc.get_paragraph_font(PARA_INDEX)
    print("字体设置:")
    print_dict(font, 2)
    
    # 获取更完整的段落样式
    all_styles = doc.get_all_paragraph_styles(PARA_INDEX)
    print("完整段落样式:")
    print_dict(all_styles, 2)
    
    # 使用StyleAnalyzer获取更深入的样式信息（如果可用）
    try:
        style_analyzer = StyleAnalyzer(DOC_PATH)
        complete_style_info = style_analyzer.get_paragraph_complete_style_info(para_element)
        print_header("StyleAnalyzer样式分析")
        
        # 打印应用的样式ID和样式名称
        if "style_id" in complete_style_info:
            print(f"应用的样式ID: {complete_style_info['style_id']}")
        if "style_name" in complete_style_info:
            print(f"样式名称: {complete_style_info['style_name']}")
        
        # 打印有效样式（考虑了继承和直接格式化）
        if "effective_style" in complete_style_info:
            print("\n有效样式 (考虑继承和直接格式化):")
            # 以JSON格式打印，更易读
            print(json.dumps(complete_style_info["effective_style"], indent=2, ensure_ascii=False))
    except (ImportError, NameError, Exception) as e:
        print(f"\n无法使用StyleAnalyzer: {e}")
    
    print_header("Run样式分析")
    
    # 获取段落中的run数量
    run_count = doc.get_run_count(PARA_INDEX)
    print(f"段落中共有 {run_count} 个run")
    
    # 分析每个run
    for run_idx in range(run_count):
        run_text = doc.get_run_text(PARA_INDEX, run_idx)
        
        # 跳过空白run
        if not run_text.strip():
            continue
            
        print(f"\n--- Run {run_idx}: \"{run_text[:30]}{'...' if len(run_text) > 30 else ''}\" ---")
        
        # 获取run字体
        font = doc.get_run_font(PARA_INDEX, run_idx)
        print("字体:")
        print_dict(font, 2)
        
        # 获取run字号
        size = doc.get_run_size(PARA_INDEX, run_idx)
        print(f"字号: {size} (约 {size/2 if size else 'N/A'} 磅)")
        
        # 获取run格式化
        formatting = doc.get_run_formatting(PARA_INDEX, run_idx)
        print("格式化:")
        print_dict(formatting, 2)
        
        # 获取run颜色
        color = doc.get_run_color(PARA_INDEX, run_idx)
        print("颜色:")
        print_dict(color, 2)
        
        # 获取完整的run样式
        style = doc.get_run_style(PARA_INDEX, run_idx)
        print("完整样式:")
        print_dict(style, 2)
    
    print_header("段落中的批注")
    
    # 检查是否有批注
    comments = doc.get_comment_at_paragraph(PARA_INDEX)
    if comments:
        print(f"找到 {len(comments)} 条批注:")
        for i, comment in enumerate(comments):
            print(f"\n批注 {i+1}:")
            print(f"  作者: {comment.get('author', 'Unknown')}")
            print(f"  日期: {comment.get('date', 'Unknown')}")
            print(f"  内容: {comment.get('text', '')}")
    else:
        print("没有找到批注")

if __name__ == "__main__":
    main() 