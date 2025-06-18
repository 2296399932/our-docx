#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
验证段落索引和元素索引转换的正确性
"""

from docx_namespace import DocxElementParser

# 创建解析器
doc = DocxElementParser("1.docx")

# 设置段落索引
para_index = 171  # 您的实际段落索引

# 获取对应的元素索引
element_index = doc.get_element_index_from_paragraph_index(para_index)

print(f"段落索引: {para_index}")
print(f"对应的元素索引: {element_index}")

# 使用段落索引获取文本
if para_index < len(doc.paragraphs):
    para_text = doc.get_paragraph_text(doc.paragraphs[para_index])
    print(f"\n使用段落索引({para_index})获取的文本:")
    print(f"文本长度: {len(para_text)}")
    print(f"文本内容: {para_text[:100]}{'...' if len(para_text) > 100 else ''}")
else:
    print(f"\n段落索引({para_index})超出范围，最大索引为: {len(doc.paragraphs)-1}")

# 使用元素索引获取文本
if element_index >= 0 and element_index < len(doc.elements):
    element_text = doc.get_element_text(element_index)
    print(f"\n使用元素索引({element_index})获取的文本:")
    print(f"文本长度: {len(element_text)}")
    print(f"文本内容: {element_text[:100]}{'...' if len(element_text) > 100 else ''}")
else:
    print(f"\n元素索引({element_index})无效或超出范围，最大索引为: {len(doc.elements)-1}")

# 验证反向转换
if element_index >= 0:
    reverse_para_index = doc.get_paragraph_index_from_element_index(element_index)
    print(f"\n反向转换检查:")
    print(f"元素索引 {element_index} 对应的段落索引: {reverse_para_index}")
    print(f"转换正确: {para_index == reverse_para_index}")

# 获取段落样式信息
if para_index < len(doc.paragraphs):
    print(f"\n段落 {para_index} 的样式信息:")
    para_style = doc.get_all_paragraph_styles(para_index)
    print(f"样式ID: {para_style.get('style_id')}")
    print(f"对齐方式: {para_style.get('alignment')}")
    print(f"缩进: {para_style.get('indentation')}")
    print(f"间距: {para_style.get('spacing')}")

# 使用元素索引获取段落样式
if element_index >= 0 and element_index < len(doc.elements):
    element = doc.elements[element_index].get('element')
    if element is not None:
        print(f"\n使用元素索引({element_index})获取的段落样式:")
        element_style = doc.extract_paragraph_style(element)
        print(f"样式ID: {element_style.get('style_id')}")
        print(f"对齐方式: {element_style.get('alignment')}")
        print(f"缩进: {element_style.get('indentation')}")
        print(f"间距: {element_style.get('spacing')}")

# 尝试更新段落间距并检查更新结果
try:
    print(f"\n尝试更新段落 {para_index} 的间距:")
    spacing_values = {'before': 50, 'after': 50}
    print(f"更新值: {spacing_values}")
    
    # 使用段落索引更新
    result = doc.set_paragraph_spacing(para_index, **spacing_values)
    print(f"更新结果: {'成功' if result else '失败'}")
    
    # 检查更新后的值
    updated_style = doc.get_all_paragraph_styles(para_index)
    print(f"更新后的间距: {updated_style.get('spacing')}")
    
    # 保存修改到新文件
    output_path = "1_spacing_updated.docx"
    doc.save(output_path)
    print(f"已保存修改到: {output_path}")
    
except Exception as e:
    print(f"更新段落间距时出错: {e}") 