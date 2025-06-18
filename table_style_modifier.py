#!/usr/bin/env python
# -*- coding: utf-8 -*-
import re

from docx_namespace import DocxElementParser
import os
import json
from openai import OpenAI
from google import genai
from google.genai import types


client = genai.Client(api_key="AIzaSyBqmGlZHCAFViUE0Jw-ox_Hi171dCK-XQw")
def convert_style_values(style_dict):
    """
    将描述性的样式值转换为实际的API参数值
    
    参数:
        style_dict (dict): 包含样式描述的字典
        
    返回:
        dict: 转换后的样式字典，可直接用于API调用
    """
    # 创建结果字典
    result = {}
    
    # 遍历输入字典进行转换
    for section_name, section_data in style_dict.items():
        if not isinstance(section_data, dict):
            result[section_name] = section_data
            continue
            
        result[section_name] = {}
        
        for element_name, element_data in section_data.items():
            if not isinstance(element_data, dict):
                result[section_name][element_name] = element_data
                continue
                
            result[section_name][element_name] = {}
            
            for attr_name, attr_value in element_data.items():
                # 跳过空值
                if attr_value is None or attr_value == "" or attr_value == "未知":
                    continue
                    
                # 1. 中文字体/英文字体转换
                if attr_name == "中文字体":
                    result[section_name][element_name]["font_chinese"] = attr_value
                elif attr_name == "英文字体":
                    result[section_name][element_name]["font_ascii"] = attr_value
                
                # 2. 字号转换
                elif attr_name == "字号":
                    # 处理中文字号
                    if "小四" in attr_value:
                        result[section_name][element_name]["size"] = 24
                    elif "四号" in attr_value:
                        result[section_name][element_name]["size"] = 28
                    elif "小三" in attr_value:
                        result[section_name][element_name]["size"] = 30
                    elif "三号" in attr_value:
                        result[section_name][element_name]["size"] = 32
                    elif "小二" in attr_value:
                        result[section_name][element_name]["size"] = 36
                    elif "二号" in attr_value:
                        result[section_name][element_name]["size"] = 44
                    elif "小一" in attr_value:
                        result[section_name][element_name]["size"] = 48
                    elif "一号" in attr_value:
                        result[section_name][element_name]["size"] = 52
                    elif "小初" in attr_value:
                        result[section_name][element_name]["size"] = 56
                    elif "初号" in attr_value:
                        result[section_name][element_name]["size"] = 84
                    elif "五号" in attr_value:
                        result[section_name][element_name]["size"] = 21
                    elif "小五" in attr_value:
                        result[section_name][element_name]["size"] = 18
                    elif "六号" in attr_value:
                        result[section_name][element_name]["size"] = 15
                    elif "小六" in attr_value:
                        result[section_name][element_name]["size"] = 13
                    elif "七号" in attr_value:
                        result[section_name][element_name]["size"] = 11
                    else:
                        # 尝试提取数字
                        num_match = re.search(r'(\d+(\.\d+)?)', attr_value)
                        if num_match:
                            size_value = float(num_match.group(1))
                            # 检查是否有磅或pt单位
                            if "pt" in attr_value or "磅" in attr_value:
                                # Word中字号是磅值的2倍
                                result[section_name][element_name]["size"] = int(size_value * 2)
                            else:
                                # 假定是半磅值
                                result[section_name][element_name]["size"] = int(size_value)
                
                # 3. 加粗转换
                elif attr_name == "加粗":
                    result[section_name][element_name]["bold"] = attr_value.lower() in ["是", "true", "yes", "1"]
                
                # 4. 对齐方式转换
                elif attr_name == "对齐方式":
                    alignment_map = {
                        "左对齐": "left", 
                        "居中": "center",
                        "右对齐": "right",
                        "两端对齐": "both"
                    }
                    
                    for key, value in alignment_map.items():
                        if key in attr_value:
                            result[section_name][element_name]["alignment"] = value
                            break
                    else:
                        # 默认值
                        if "对齐" in attr_value:
                            result[section_name][element_name]["alignment"] = "left"
                
                # 5. 缩进转换
                elif attr_name == "缩进":
                    # 首行缩进
                    first_line_match = re.search(r'首行[缩进]*[^\d]*(\d+(\.\d+)?)[字符]*', attr_value)
                    if first_line_match:
                        chars = float(first_line_match.group(1))
                        # 一个汉字约为440 twip
                        result[section_name][element_name]["first_line"] = int(chars * 440)
                    
                    # 悬挂缩进
                    hanging_match = re.search(r'悬挂[缩进]*[^\d]*(\d+(\.\d+)?)[字符]*', attr_value)
                    if hanging_match:
                        chars = float(hanging_match.group(1))
                        result[section_name][element_name]["hanging"] = int(chars * 440)
                    
                    # 左侧缩进
                    left_match = re.search(r'左[侧边]?[缩进]*[^\d]*(\d+(\.\d+)?)[字符厘米]*', attr_value)
                    if left_match:
                        value = float(left_match.group(1))
                        # 检查单位
                        if "厘米" in attr_value or "cm" in attr_value:
                            # 1厘米约为567 twip
                            result[section_name][element_name]["left"] = int(value * 567)
                        else:
                            # 假定是字符
                            result[section_name][element_name]["left"] = int(value * 440)
                    
                    # 右侧缩进
                    right_match = re.search(r'右[侧边]?[缩进]*[^\d]*(\d+(\.\d+)?)[字符厘米]*', attr_value)
                    if right_match:
                        value = float(right_match.group(1))
                        if "厘米" in attr_value or "cm" in attr_value:
                            result[section_name][element_name]["right"] = int(value * 567)
                        else:
                            result[section_name][element_name]["right"] = int(value * 440)
                
                # 6. 行距转换
                elif attr_name == "行距":
                    # 固定值行距
                    fixed_match = re.search(r'固定[值]?[^\d]*(\d+(\.\d+)?)[磅pt]*', attr_value)
                    if fixed_match:
                        pts = float(fixed_match.group(1))
                        # 1磅 = 20 twip
                        result[section_name][element_name]["line"] = int(pts * 20)
                        result[section_name][element_name]["line_rule"] = "exact"
                    
                    # 多倍行距
                    multi_match = re.search(r'[多倍行距]*(\d+(\.\d+)?)[倍行距]*', attr_value)
                    if multi_match and ("倍" in attr_value or "多倍" in attr_value):
                        times = float(multi_match.group(1))
                        # 多倍行距基数是240 twip
                        result[section_name][element_name]["line"] = int(times * 240)
                        result[section_name][element_name]["line_rule"] = "auto"
                    
                    # 最小值行距
                    if "最小值" in attr_value:
                        min_match = re.search(r'最小值[^\d]*(\d+(\.\d+)?)[磅pt]*', attr_value)
                        if min_match:
                            pts = float(min_match.group(1))
                            result[section_name][element_name]["line"] = int(pts * 20)
                            result[section_name][element_name]["line_rule"] = "atLeast"
                
                # 7. 段前段后距离转换
                elif attr_name == "段前距离":
                    # 行为单位
                    line_match = re.search(r'(\d+(\.\d+)?)[行]*', attr_value)
                    if line_match and ("行" in attr_value):
                        lines = float(line_match.group(1))
                        # 使用beforeLines参数，保留原始行数值
                        result[section_name][element_name]["beforeLines"] = lines
                    
                    # 磅为单位
                    pt_match = re.search(r'(\d+(\.\d+)?)[磅pt]*', attr_value)
                    if pt_match and ("磅" in attr_value or "pt" in attr_value):
                        pts = float(pt_match.group(1))
                        result[section_name][element_name]["before"] = int(pts * 20)
                    
                    # 如果只有数字，默认为磅
                    elif pt_match and not ("行" in attr_value):
                        pts = float(pt_match.group(1))
                        result[section_name][element_name]["before"] = int(pts * 20)
                
                elif attr_name == "段后距离":
                    # 行为单位
                    line_match = re.search(r'(\d+(\.\d+)?)[行]*', attr_value)
                    if line_match and ("行" in attr_value):
                        lines = float(line_match.group(1))
                        # 使用afterLines参数，保留原始行数值
                        result[section_name][element_name]["afterLines"] = lines
                    
                    # 磅为单位
                    pt_match = re.search(r'(\d+(\.\d+)?)[磅pt]*', attr_value)
                    if pt_match and ("磅" in attr_value or "pt" in attr_value):
                        pts = float(pt_match.group(1))
                        result[section_name][element_name]["after"] = int(pts * 20)
                    
                    # 如果只有数字，默认为磅
                    elif pt_match and not ("行" in attr_value):
                        pts = float(pt_match.group(1))
                        result[section_name][element_name]["after"] = int(pts * 20)
                
                # 保留其他未处理的属性
                else:
                    if isinstance(attr_value, dict):
                        result[section_name][element_name][attr_name] = convert_style_values({attr_name: attr_value})[attr_name]
                    else:
                        result[section_name][element_name][attr_name] = attr_value
    
    return result

def extract_and_print_all_content(docx_path):
    """
    提取并打印Word文档的所有内容，包括段落、表格、批注、页眉页脚、脚注等
    
    参数:
        docx_path (str): 输入文档路径
    """
    print(f"\n==================== 开始分析文档：{docx_path} ====================\n")
    
    # 创建DocxElementParser实例
    doc = DocxElementParser(docx_path)
    
    # 11.py. 打印文档基本信息
    para_count = doc.get_paragraphs_length()
    table_count = doc.get_table_length()
    
    print(f"文档基本信息:")
    print(f"  - 总段落数: {para_count}")
    print(f"  - 总表格数: {table_count}")
    
    # 2. 打印所有段落内容
    print("\n---------------------- 段落内容 -----------------------")
    paragraphs_text = doc.get_all_paragraphs_text()
    for i, text in enumerate(paragraphs_text):
        if text.strip():  # 只打印非空段落
            # 获取段落样式信息
            try:
                para_element = doc.paragraphs[i]
                style_info = doc.extract_paragraph_style(para_element['element'] if isinstance(para_element, dict) and 'element' in para_element else para_element)
                style_summary = ""
                if 'style_id' in style_info and style_info['style_id']:
                    style_summary = f"[样式ID: {style_info['style_id']}]"
                elif 'name' in style_info and style_info['name']:
                    style_summary = f"[样式名: {style_info['name']}]"
            except:
                style_summary = "[样式未知]"
            
            # 限制文本长度，防止输出过长
            display_text = text[:100] + "..." if len(text) > 100 else text
            print(f"段落 {i}: {style_summary} {display_text}")
    
    # 3. 打印所有表格内容（精简版）
    print("\n---------------------- 表格内容 -----------------------")
    for table_idx in range(table_count):
        print(f"\n表格 {table_idx}:")
        try:
            # 获取表格元素
            table_element = doc.tables[table_idx]
            
            # 获取表格样式信息
            style_name = "默认"
            try:
                table_style = doc.get_table_style(table_idx)
                if table_style and 'style_id' in table_style:
                    style_name = table_style.get('style_id')
            except:
                pass
            
            # 提取表格内容（仅用于分析，不完全打印）
            table_content = []
            try:
                table_content = doc.extract_table_content(table_element['element'])
            except:
                pass
            
            # 计算表格尺寸
            row_count = len(table_content) if table_content else 0
            col_count = max([len(row) for row in table_content]) if table_content and row_count > 0 else 0
            
            # 打印表格基本信息
            print(f"  - 样式: {style_name}")
            print(f"  - 尺寸: {row_count}行 x {col_count}列")
            
            # 仅打印第一行内容作为预览（如果有内容）
            if table_content and row_count > 0 and col_count > 0:
                first_row = table_content[0]
                preview = " | ".join([str(cell)[:15] + "..." if len(str(cell)) > 15 else str(cell) for cell in first_row])
                print(f"  - 内容预览: {preview}")
                
                # 如果行数超过1，显示更多行数量信息
                if row_count > 1:
                    print(f"  - 更多行: {row_count-1}行未显示")
            else:
                print("  - 内容: 表格为空或无法解析")
                
        except Exception as e:
            print(f"  - 错误: 解析表格时出错: {str(e)}")
    
    # 4. 提取并打印批注
    print("\n---------------------- 批注内容 -----------------------")
    try:
        comments = doc.extract_comments()
        print(f"文档中共有 {len(comments)} 条批注")
        
        for idx, comment in enumerate(comments):
            print(f"\n批注 {idx+1}:")
            print(f"  ID: {comment['id']}")
            print(f"  作者: {comment['author']}")
            print(f"  日期: {comment['date']}")
            print(f"  内容: {comment['text']}")
            print(f"  引用文本: {comment['referenced_text']}")
            print(f"  所在段落索引: {comment['paragraph_index']}")
    except Exception as e:
        print(f"提取批注时出错: {e}")
    
    # 5. 提取并打印页眉页脚
    print("\n---------------------- 页眉页脚 -----------------------")
    try:
        # 尝试获取页眉
        if hasattr(doc, 'parts') and 'headers' in doc.parts:
            headers = doc.parts['headers']
            print(f"找到 {len(headers)} 个页眉")
            
            for header_name, header_tree in headers.items():
                print(f"\n页眉 {header_name}:")
                try:
                    # 提取页眉中的段落文本
                    header_paragraphs = header_tree.findall('.//{%s}p' % doc.NAMESPACES['w'])
                    for p_idx, p in enumerate(header_paragraphs):
                        text = doc.get_paragraph_text(p)
                        if text.strip():
                            print(f"  段落 {p_idx}: {text}")
                except Exception as e:
                    print(f"  解析页眉文本时出错: {e}")
        else:
            print("未找到页眉信息")
            
        # 尝试获取页脚
        if hasattr(doc, 'parts') and 'footers' in doc.parts:
            footers = doc.parts['footers']
            print(f"\n找到 {len(footers)} 个页脚")
            
            for footer_name, footer_tree in footers.items():
                print(f"\n页脚 {footer_name}:")
                try:
                    # 提取页脚中的段落文本
                    footer_paragraphs = footer_tree.findall('.//{%s}p' % doc.NAMESPACES['w'])
                    for p_idx, p in enumerate(footer_paragraphs):
                        text = doc.get_paragraph_text(p)
                        if text.strip():
                            print(f"  段落 {p_idx}: {text}")
                except Exception as e:
                    print(f"  解析页脚文本时出错: {e}")
        else:
            print("未找到页脚信息")
    except Exception as e:
        print(f"提取页眉页脚时出错: {e}")
    
    # 6. 提取并打印脚注和尾注（精简版）
    print("\n---------------------- 脚注和尾注 -----------------------")
    try:
        # 尝试获取脚注
        if hasattr(doc, 'parts') and 'footnotes' in doc.parts and doc.parts['footnotes'] is not None:
            footnotes_root = doc.parts['footnotes'].getroot()
            footnote_elements = footnotes_root.findall('.//{%s}footnote' % doc.NAMESPACES['w'])
            
            # 过滤掉ID为0或1的特殊脚注（通常是分隔符和连续符）
            valid_footnotes = [fn for fn in footnote_elements if fn.get('{%s}id' % doc.NAMESPACES['w']) not in ['0', '1']]
            
            print(f"找到 {len(valid_footnotes)} 条脚注")
            if valid_footnotes:
                for fn_idx, footnote in enumerate(valid_footnotes[:3]):  # 只显示前3个
                    fn_id = footnote.get('{%s}id' % doc.NAMESPACES['w'])
                    # 提取脚注中的文本
                    fn_text = ""
                    for p in footnote.findall('.//{%s}p' % doc.NAMESPACES['w']):
                        fn_text += doc.get_paragraph_text(p) + " "
                    
                    # 限制长度
                    display_text = fn_text[:50] + "..." if len(fn_text) > 50 else fn_text
                    print(f"  脚注 {fn_idx+1} (ID={fn_id}): {display_text}")
                
                # 如果有更多脚注，显示数量信息
                if len(valid_footnotes) > 3:
                    print(f"  ...还有 {len(valid_footnotes)-3} 条脚注未显示")
        else:
            print("未找到脚注信息")
            
        # 尝试获取尾注
        if hasattr(doc, 'parts') and 'endnotes' in doc.parts and doc.parts['endnotes'] is not None:
            endnotes_root = doc.parts['endnotes'].getroot()
            endnote_elements = endnotes_root.findall('.//{%s}endnote' % doc.NAMESPACES['w'])
            
            # 过滤掉ID为0或1的特殊尾注
            valid_endnotes = [en for en in endnote_elements if en.get('{%s}id' % doc.NAMESPACES['w']) not in ['0', '1']]
            
            print(f"\n找到 {len(valid_endnotes)} 条尾注")
            if valid_endnotes:
                for en_idx, endnote in enumerate(valid_endnotes[:3]):  # 只显示前3个
                    en_id = endnote.get('{%s}id' % doc.NAMESPACES['w'])
                    # 提取尾注中的文本
                    en_text = ""
                    for p in endnote.findall('.//{%s}p' % doc.NAMESPACES['w']):
                        en_text += doc.get_paragraph_text(p) + " "
                    
                    # 限制长度
                    display_text = en_text[:50] + "..." if len(en_text) > 50 else en_text
                    print(f"  尾注 {en_idx+1} (ID={en_id}): {display_text}")
                
                # 如果有更多尾注，显示数量信息
                if len(valid_endnotes) > 3:
                    print(f"  ...还有 {len(valid_endnotes)-3} 条尾注未显示")
        else:
            print("未找到尾注信息")
    except Exception as e:
        print(f"提取脚注和尾注时出错: {e}")
    
    # 7. 提取图片信息（精简版）
    print("\n---------------------- 图片信息 -----------------------")
    try:
        # 使用简单方法统计图片
        image_count = doc.count_images_simple()
        print(f"文档中共有 {image_count} 张图片")
        
        # 尝试获取图片所在段落
        if hasattr(doc, 'get_image_paragraphs_indices'):
            image_paragraphs = doc.get_image_paragraphs_indices()
            if image_paragraphs:
                # 只显示图片段落数量和前3个图片位置
                print(f"图片分布在 {len(image_paragraphs)} 个段落中")
                
                for idx, (para_idx, rel_ids) in enumerate(image_paragraphs[:3]):
                    # 尝试获取段落文本简介
                    para_text = ""
                    if para_idx < len(doc.paragraphs):
                        para_text = doc.get_paragraph_text(doc.paragraphs[para_idx])
                        para_text = para_text[:30] + "..." if len(para_text) > 30 else para_text
                    
                    print(f"  图片 {idx+1}: 段落{para_idx} {para_text}")
                
                # 如果有更多图片，显示数量信息
                if len(image_paragraphs) > 3:
                    print(f"  ...还有 {len(image_paragraphs)-3} 个段落中的图片未显示")
        else:
                    print("未找到图片相关段落")
    except Exception as e:
        print(f"提取图片信息时出错: {e}")
    
    print(f"\n==================== 文档分析完成 ====================\n")
    
    # 返回提取的文档内容，以便后续AI分析
    return {
        "paragraphs": [{ "text": text} for i, text in enumerate(paragraphs_text) if text.strip()],
        "tables":[ doc.extract_table_content(table['element']) for table in doc.tables],


    }

def analyze_document_styles_with_ai(docx_path, api_key=None, model="qwen-max-latest"):
    """
    提取文档内容并发送给AI进行分析，返回各个模块的样式说明

    参数:
        docx_path (str): 输入文档路径
        api_key (str, optional): API密钥，默认为None时使用环境变量
        model (str, optional): 使用的AI模型，默认为"qwen-max-latest"
        
    返回:
        dict: 包含样式分析和API参数的结果
    """
    print(f"开始提取文档内容: {docx_path}")
    doc_content = extract_and_print_all_content(docx_path)
    print(doc_content)
    
    # 创建OpenAI客户端

    # 先使用AI整理文档内容和识别样式规范
    print("正在使用AI整理文档内容和识别样式规范...")
    
    # 发送给AI的系统提示
    preprocessing_system_prompt = """
    你是一个专业的学术论文模板分析专家。你需要分析提供的Word论文模板的内容和格式规范，清晰识别各部分的结构、内容和格式要求。
    
    请执行以下详细任务：
    1. 识别论文模板的主要组成部分，如封面、诚信声明、中英文摘要、目录、正文章节、参考文献、致谢、附录等
    2. 分析每个部分的格式规范要求，包括字体、字号、段落间距、缩进、对齐方式、加粗（默认为否）等
    3. 识别模板中的示例内容与格式说明文字，区分哪些是样例、哪些是说明 
    4. 注意模板中可能包含的特殊格式要求和注意事项
    5. 整理各部分的内容和格式规范，形成清晰的文档结构说明
    6.目录的标题一般与正文的标题格式不同请你注意区分
    你的回答应该详细描述论文各部分的精确格式规范，以便后续进行样式分析和应用。
    """
    
    # 准备用户提示
    preprocessing_user_prompt = f"""
    请分析以下从论文模板中提取的内容，识别并详细整理各部分的结构、内容和格式规范：
    
    模板内容：
    {json.dumps(doc_content, ensure_ascii=False, indent=2)}
    
    请将分析结果整理为以下详细结构：
    

    
    1. 详细结构分析：
       按照论文顺序，对每个部分进行详细分析，包括：
 
       - 中英文摘要与关键词
       - 目录
       - 正文（各级标题与正文段落）
       - 表格
       - 图片（包括图片标题）
       - 图表（包括图表标题,还有列名行和数据行的内容文字要求）
       - 公式（包括公式编号）
    
       - 其他格式要求（如页边距、页码格式等）
       - 参考文献
       - 致谢
       - 附录等
    
  
    
    请尽可能详细地提取模板中隐含的各种格式规范，这对后续准确分析论文样式至关重要。
    """
    
    try:


        preprocessing_completion = f"{preprocessing_system_prompt}\n\n{preprocessing_user_prompt}"
        response = client.models.generate_content(
            model="gemini-2.5-flash-preview-05-20",
            contents=[preprocessing_completion],
            config=types.GenerateContentConfig(
                max_output_tokens=40000,
                temperature=0.3
            )
        )

        # 提取响应内容
        organized_content = response.text
        # 获取整理后的内容

        print("文档内容整理完成")
        
        # 保存整理后的内容到文件
        organized_content_file = f"{os.path.splitext(docx_path)[0]}_organized_content.txt"
        with open(organized_content_file, 'w', encoding='utf-8') as f:
            f.write(organized_content)
        print(f"整理后的内容已保存到: {organized_content_file}")
        
        # 定义分类标准（原有代码移到这里）
        classification_standards = {
            "目录": """
            目录部分通常包含以下元素：
            1. 目录标题(目录标题)：如"目录"或"CONTENTS"
            2. 一级目录项(一级目录项)：如"第一章 绪论"、"致谢"、"参考文献"等
            3. 二级目录项(二级目录项)：如"1.1 研究背景"
            4. 三级目录项(三级目录项)：如"1.1.1 问题提出"
            5. 目录页码(页码)：目录项后面的页码数字
            """,

            "中文摘要": """
            中文摘要部分通常包含以下元素：
            1. 摘要标题(摘要标题)：如"摘要"
            2. 摘要内容(摘要内容)
            3. 关键词标签(关键词标签)：如"关键词："
            4. 关键词内容(关键词内容)
            """,

            "英文摘要": """
            英文摘要部分通常包含以下元素：
            1. 摘要标题(摘要标题)：如"Abstract"
            2. 摘要内容(摘要内容)
            3. 关键词标签(关键词标签)：如"Keywords:"
            4. 关键词内容(关键词内容)
            """,

            "正文": """
            正文部分通常包含以下元素：
            1. 一级标题(一级标题)：如"第一章 绪论"、"第二章 相关技术"等。一级标题通常以"第X章"开头。
            2. 二级标题(二级标题)：如"1.1 研究背景"、"2.3 系统架构"等。二级标题通常采用"X.Y 内容"的格式。
            3. 三级标题(三级标题)：如"1.1.1 问题提出"等。三级标题通常采用"X.Y.Z 内容"的格式。
            4. 正文段落(正文段落)：普通的正文内容。
            5. 图片(图片)：任何图片段落。
            6. 图表标题(图表标题)：图片的标题或说明文字。
            7. 表格(表格)：表格元素。
            8. 表格标题(表格标题)：表格的标题或说明文字。
            9. 公式(公式)：数学公式。
            10. 代码(代码)：代码块或程序代码。

            标题级别分类的详细规则：
            - 一级标题(一级标题)：
              * 以"第X章"开头的标题（如"第一章"、"第二章"等）
              * "致谢"、"参考文献"、"附录"等独立大章节的标题
              * 通常字号最大、格式最突出

            - 二级标题(二级标题)：
              * 采用"X.Y"编号格式的标题（如"1.1"、"2.3"等）
              * 作为一级标题的下一级标题
              * 通常字号比一级标题小，但比正文大

            - 三级标题(三级标题)：
              * 采用"X.Y.Z"编号格式的标题（如"1.1.1"、"2.3.2"等）
              * 作为二级标题的下一级标题
              * 通常字号可能比二级标题小，但仍比正文大
            """,

            "参考文献": """
            参考文献部分通常包含以下元素：
            1. 参考文献标题(标题)：如"参考文献"或"References"
            2. 参考文献条目(条目)
            """,

            "致谢": """
            致谢部分通常包含以下元素：
            1. 致谢标题(标题)：如"致谢"
            2. 致谢内容(内容)
            """,

            "附录": """
            附录部分通常包含以下元素：
            1. 附录标题(标题)：如"附录"或"Appendix"、"外文原文及译文"等
            2. 附录子标题(子标题)
            3. 附录内容(内容)
            4. 附录图表(图表)
            5. 附录表格(表格)
            """
        }

        # 创建预定义的JSON模板，包含所有需要填充的样式属性
        style_template = {
            "目录": {
                "目录标题": {
                    "中文字体": "",
                    "英文字体": "",
                    "字号": "",
                    "加粗": "",
                    "对齐方式": "",
                    "段前距离": "",
                    "段后距离": "",
                    "缩进": "",
                    "行距": ""
                },
                "一级目录项": {
                    "中文字体": "",
                    "英文字体": "",
                    "字号": "",
                    "加粗": "",
                    "对齐方式": "",
                    "段前距离": "",
                    "段后距离": "",
                    "缩进": "",
                    "行距": ""
                },
                "二级目录项": {
                    "中文字体": "",
                    "英文字体": "",
                    "字号": "",
                    "加粗": "",
                    "对齐方式": "",
                    "段前距离": "",
                    "段后距离": "",
                    "缩进": "",
                    "行距": "",
                },
                "三级目录项": {
                    "中文字体": "",
                    "英文字体": "",
                    "字号": "",
                    "加粗": "",
                    "对齐方式": "",
                    "段前距离": "",
                    "段后距离": "",
                    "缩进": "",
                    "行距": ""
                },
                "页码": {
                    "中文字体": "",
                    "英文字体": "",
                    "字号": "",
                    "加粗": "",
                    "对齐方式": ""
                }
            },
            "中文摘要": {
                "摘要标题": {
                    "中文字体": "",
                    "英文字体": "",
                    "字号": "",
                    "加粗": "",
                    "对齐方式": "",
                    "段前距离": "",
                    "段后距离": "",
                    "缩进": "",
                    "行距": ""
                },
                "摘要内容": {
                    "中文字体": "",
                    "英文字体": "",
                    "字号": "",
                    "加粗": "",
                    "对齐方式": "",
                    "段前距离": "",
                    "段后距离": "",
                    "缩进": "",
                    "行距": ""
                },
                "关键词标签": {
                    "中文字体": "",
                    "英文字体": "",
                    "字号": "",
                    "加粗": "",
                    "对齐方式": "",
                    "段前距离": "",
                    "段后距离": "",
                    "缩进": "",
                    "行距": ""
                },
                "关键词内容": {
                    "中文字体": "",
                    "英文字体": "",
                    "字号": "",
                    "加粗": "",
                    "对齐方式": "",
                    "段前距离": "",
                    "段后距离": "",
                    "缩进": "",
                    "行距": ""
                }
            },
            "英文摘要": {
                "摘要标题": {
                    "中文字体": "",
                    "英文字体": "",
                    "字号": "",
                    "加粗": "",
                    "对齐方式": "",
                    "段前距离": "",
                    "段后距离": "",
                    "缩进": "",
                    "行距": ""
                },
                "摘要内容": {
                    "中文字体": "",
                    "英文字体": "",
                    "字号": "",
                    "加粗": "",
                    "对齐方式": "",
                    "段前距离": "",
                    "段后距离": "",
                    "缩进": "",
                    "行距": ""
                },
                "关键词标签": {
                    "中文字体": "",
                    "英文字体": "",
                    "字号": "",
                    "加粗": "",
                    "对齐方式": "",
                    "段前距离": "",
                    "段后距离": ""
                },
                "关键词内容": {
                    "中文字体": "",
                    "英文字体": "",
                    "字号": "",
                    "加粗": "",
                    "对齐方式": "",
                    "段前距离": "",
                    "段后距离": ""
                }
            },
            "正文": {
                "一级标题": {
                    "中文字体": "",
                    "英文字体": "",
                    "字号": "",
                    "加粗": "",
                    "对齐方式": "",
                    "段前距离": "",
                    "段后距离": "",
                    "缩进": "",
                    "行距": ""
                },
                "二级标题": {
                    "中文字体": "",
                    "英文字体": "",
                    "字号": "",
                    "加粗": "",
                    "对齐方式": "",
                    "段前距离": "",
                    "段后距离": "",
                    "缩进": "",
                    "行距": ""
                },
                "三级标题": {
                    "中文字体": "",
                    "英文字体": "",
                    "字号": "",
                    "加粗": "",
                    "对齐方式": "",
                    "段前距离": "",
                    "段后距离": "",
                    "缩进": "",
                    "行距": ""
                },
                "正文段落": {
                    "中文字体": "",
                    "英文字体": "",
                    "字号": "",
                    "加粗": "",
                    "对齐方式": "",
                    "段前距离": "",
                    "段后距离": "",
                    "缩进": "",
                    "行距": ""
                },
                "图表标题": {
                    "中文字体": "",
                    "英文字体": "",
                    "字号": "",
                    "加粗": "",
                    "对齐方式": "",
                    "段前距离": "",
                    "段后距离": "",
                    "缩进": "",
                    "行距": ""
                },
                "表格标题": {
                    "中文字体": "",
                    "英文字体": "",
                    "字号": "",
                    "加粗": "",
                    "对齐方式": "",
                    "段前距离": "",
                    "段后距离": "",
                    "缩进": "",
                    "行距": ""
                },
                "表格": {
                    "列名行文本样式": {
                        "中文字体": "",
                        "英文字体": "",
                        "字号": "",
                        "加粗": "",
                        "对齐方式": "",
                        "行距": ""
                    },

                    "是否为三线表": "",
                    "数据行文本样式": {
                        "中文字体": "",
                        "英文字体": "",
                        "字号": "",
                        "加粗": "",
                        "对齐方式": "",
                        "行距": ""
                    }
                },
                "公式": {
                    "中文字体": "",
                    "英文字体": "",
                    "字号": "",
                    "对齐方式": "",
                    "段前距离": "",
                    "段后距离": ""
                },
                "代码": {
                    "中文字体": "",
                    "英文字体": "",
                    "字号": "",
                    "加粗": "",
                    "对齐方式": "",
                    "段前距离": "",
                    "段后距离": "",
                    "缩进": "",
                    "行距": ""
                }
                },


            "参考文献": {
                "标题": {
                    "中文字体": "",
                    "英文字体": "",
                    "字号": "",
                    "加粗": "",
                    "对齐方式": "",
                    "段前距离": "",
                    "段后距离": "",
                    "缩进": "",
                    "行距": ""
                },
                "条目": {
                    "中文字体": "",
                    "英文字体": "",
                    "字号": "",
                    "加粗": "",
                    "对齐方式": "",
                    "段前距离": "",
                    "段后距离": "",
                    "缩进": "",
                    "行距": ""
                }
            },
            "致谢": {
                "标题": {
                    "中文字体": "",
                    "英文字体": "",
                    "字号": "",
                    "加粗": "",
                    "对齐方式": "",
                    "段前距离": "",
                    "段后距离": "",
                    "缩进": "",
                    "行距": ""
                },
                "内容": {
                    "中文字体": "",
                    "英文字体": "",
                    "字号": "",
                    "加粗": "",
                    "对齐方式": "",
                    "段前距离": "",
                    "段后距离": "",
                    "缩进": "",
                    "行距": ""
                }
            },
            "附录": {
                "标题": {
                    "中文字体": "",
                    "英文字体": "",
                    "字号": "",
                    "加粗": "",
                    "对齐方式": "",
                    "段前距离": "",
                    "段后距离": "",
                    "缩进": "",
                    "行距": ""
                },
                "子标题": {
                    "中文字体": "",
                    "英文字体": "",
                    "字号": "",
                    "加粗": "",
                    "对齐方式": "",
                    "段前距离": "",
                    "段后距离": "",
                    "缩进": "",
                    "行距": ""
                },
                "内容": {
                    "中文字体": "",
                    "英文字体": "",
                    "字号": "",
                    "加粗": "",
                    "对齐方式": "",
                    "段前距离": "",
                    "段后距离": "",
                    "缩进": "",
                    "行距": ""
                }
            }
        }
        
        # 将模板转换为JSON字符串
        template_json = json.dumps(style_template, ensure_ascii=False, indent=2)
        
        print("开始使用AI分析文档样式...")
        
        # 准备发送给AI的系统提示
        system_prompt = """
        你是一个专业的学术论文格式分析专家。你需要分析提供的Word文档内容和整理后的文档结构，识别各部分的样式特征，并在给定的JSON模板中填充正确的样式信息。
        
        请仔细分析段落的样式信息，包括字体、大小、缩进、间距等，为每种类型的元素提供详细的样式描述。
        模板中的每个空字段都需要填充，如果无法确定某个字段的值，请填写"未知"或最可能的默认值。
        
        你的回答应该只包含填充完成的JSON对象，不要包含其他解释或额外内容。
        """
        
        # 修改用户提示，加入整理后的内容
        user_prompt = f"""
        请分析以下Word文档内容和整理后的文档结构，根据给定的分类标准，在提供的JSON模板中填充各部分元素的样式特征。
        注意：  必须严格按照模板格式填充，不能有任何多余的字段。如果某个字段没有对应内容，请填写"未知"或最可能的默认值 ，但是 加粗默认为否。
        
        文档原始内容：
        {json.dumps(doc_content, ensure_ascii=False, indent=2)}
        
        文档整理后的结构和内容：
        {organized_content}

        分类标准：
        {json.dumps(classification_standards, ensure_ascii=False, indent=2)}

        样式模板（请在此模板中填充样式信息）：
        {template_json}
        
        说明：
        1. 中文字体：填写中文字体名称，如"宋体"、"黑体"、"楷体"等
        2. 英文字体：填写英文字体名称，如"Times New Roman"、"Arial"、"Calibri"等
        3. 字号：填写数值，如"小四"对应"12pt"或"24半磅" "五号"对应"10.5pt"或"21半磅"
        4. 加粗：填写"是"或"否" 注意没有特别说明需要加粗默认否
        5. 对齐方式：填写"左对齐"、"居中"、"右对齐"或"两端对齐"
        6. 段前距离/段后距离：明确指定单位，如"0.5行"或"6磅"。请注意区分以行为单位和以磅为单位的情况
        7. 缩进：填写数值，如"首行缩进2字符"或"悬挂缩进1厘米"或者"首行缩进2英文半角空格"等
        8. 行距：填写数值，如"1.5倍行距"或"固定值20磅" 
        
        注意 " 段前后。。。指代段前段后距离都为一个值 ，中文字体没有明确说明就默认为宋体" 
        请确保填充模板中的所有字段，返回完整的JSON对象。只需返回JSON数据，不需要其他解释。
        10. 只返回JSON对象，不要有任何解释、注释或多余内容。
        11. "列名行文本样式"和"数据行文本样式"这两个字段的字段名必须保留中文，其内部所有属性必须全部转换为英文API参数名和标准值。
        请严格按照这些规则转换样式值，确保所有转换后的样式对象都可以直接用于函数调用。
        """
        
        # 使用非流式处理模式调用模型
        print(f"正在调用AI模型({model})分析文档样式...")


        completion = f"{system_prompt}\n\n{user_prompt}"
        response = client.models.generate_content(
            model="gemini-2.5-flash-preview-05-20",
            contents=[completion],
            config=types.GenerateContentConfig(
                max_output_tokens=40000,
                temperature=0.5
            )
        )

        # 提取响应内容
        response_content = response.text


        
        # 解析JSON结果
        try:
            # 查找JSON内容
            import re
            json_match = re.search(r'```json\s*([\s\S]*?)\s*```', response_content)
            
            if json_match:
                # 从代码块中提取JSON
                json_content = json_match.group(1)
                result = json.loads(json_content)
            else:
                # 尝试直接解析整个响应
                result = json.loads(response_content)
                
            print("AI分析完成，成功获取样式说明")
            
            # 保存分析结果到文件
            output_file = f"{os.path.splitext(docx_path)[0]}_style_analysis.json"
            with open(output_file, 'w', encoding='utf-8') as f:
                json.dump(result, f, ensure_ascii=False, indent=2)
            print(f"分析结果已保存到: {output_file}")
            
            # 将样式分析结果转换为API参数格式
            print("正在将样式转换为API参数格式...")
            api_params = convert_style_to_api_params(result, docx_path)
            
            return {
                "style_analysis": result,
                "api_params": api_params
            }
            
        except json.JSONDecodeError as e:
            print(f"解析AI响应时出错: {e}")
            print("原始响应内容:")
            print(response_content)
            return {"error": "无法解析AI响应", "raw_response": response_content}
            
    except Exception as e:
        print(f"调用AI服务时出错: {e}")
        return {"error": str(e)}


def convert_style_to_api_params(style_dict, docx_path, api_key=None, model="qwen-max-latest"):
    """
    将描述性样式转换为可直接用于docx_namespace.py中函数的参数格式
    
    参数:
        style_dict: 样式分析结果字典
        docx_path: 文档路径，用于生成输出文件名
        api_key: API密钥
        model: 使用的AI模型
        
    返回:
        dict: 可直接用于函数调用的参数格式
    """
    # 设置输出文件路径
    api_params_file = f"{os.path.splitext(docx_path)[0]}_api_params.json"
    
    # 创建OpenAI客户端

    
    # 准备系统提示
    system_prompt = """
    你是一个文档处理API专家。你需要将描述性的样式定义转换为可直接用于Python API函数的参数格式。

    目标函数是docx_namespace.py中的update_runs_style(para_index, **style_properties)和update_paragraph_style(para_index, **style_properties)。
    
    请严格按照以下规则转换样式值:
    
    1. 中文字体/英文字体 → "font_chinese"/"font_ascii": 
       - 直接使用字体名称字符串，如 "宋体"→"font_chinese": "宋体"
    
    2. 字号 → "size": 整数值，将描述转换为半磅值(不是磅值)
       - 数字直接转为整数，如"24" → "size": 24
       - 中文字号必须准确转换:
         * "小四" → "size": 24
         * "四号" → "size": 28
         * "小三" → "size": 30
         * "三号" → "size": 32
         * "小二" → "size": 36
         * "二号" → "size": 44
         * "小一" → "size": 48
         * "一号" → "size": 52
         * "小初" → "size": 56
         * "初号" → "size": 84
         * "五号" → "size": 21
         * "小五" → "size": 18
         * "六号" → "size": 15
         * "小六" → "size": 13
         * "七号" → "size": 11
    
    3. 加粗 → "bold": 布尔值，"是"→true, "否"→false
       - "是", "true", "yes", "1" → "bold": true
       - "否", "false", "no", "0" → "bold": false
       如果没有明确说明需要加粗一律不加粗
    
    4. 对齐方式 → "alignment": 字符串
       - "左对齐" → "alignment": "left"
       - "居中" → "alignment": "center" 
       - "右对齐" → "alignment": "right"
       - "两端对齐" → "alignment": "both"
    
    5. 缩进 → "first_line"/"hanging"/"left"/"right": 整数值(twip单位)
       - "首行缩进2字符" → "first_line": 420 (一个汉字约210 twip)
       - "悬挂缩进4字符" → "hanging": 1760
       - “英文2字符”或“英文2半角空格” → "first_line": 210(一个英文约105 twip)。
       - "左侧缩进1厘米" → "left": 567 (1厘米约567 twip)
       - "右侧缩进2字符" → "right": 880
    
    6. 行距 → "line"/"line_rule": 整数和字符串
       - "固定值21磅" → "line": 420, "line_rule": "exact" (1磅=20 twip)
       - "多倍行距1.25" → "line": 300, "line_rule": "auto" (1.25*240=300)
       - "最小值24磅" → "line": 480, "line_rule": "atLeast"
    
    7. 段落间距 → 根据单位类型选择不同参数:
       - 以"行"为单位时:
         * "0.5行" → "beforeLines": 50 或 "afterLines": 50 (保留原始行数值，不转换为twip)
         * "1.5行" → "beforeLines": 150 或 "afterLines": 150
       - 以"磅"为单位时:
         * "6磅" → "before": 120 或 "after": 120 (1磅=20 twip)
         * "12pt" → "before": 240 或 "after": 240
       - 如果值为0或接近0，一定要设置为0，不能省略
       - 必须根据原始描述中的单位选择正确的参数名称，不要混用
       
    8. 表格样式处理：
       - 对于"表格"对象中的所有嵌套结构(如"列名行文本样式"、"数据行文本样式")，必须递归应用上述所有转换规则
       - "是否为三线表": "是" → "is_three_line_table": true
       - "是否为三线表": "否" → "is_three_line_table": false
       - 表格中的特殊行距描述如"固定值21磅"、"多倍行距1.25倍"必须按第6条规则转换为标准参数
       - 示例:
         * "中文字体": "宋体" → "font_chinese": "宋体"
         * "字号": "21" → "size": 21 (必须为数字类型，不是字符串)
         * "加粗": "否" → "bold": false (必须为布尔类型，不是字符串)
         * "行距": "固定值21磅" → "line": 420, "line_rule": "exact"
         * "行距": "多倍行距1.25倍" → "line": 300, "line_rule": "auto"
    
    必须严格遵循这些规则，所有数值必须是实际的数字(不是字符串)，布尔值必须是true/false(不是字符串)。
    千万不要省略任何转换，即使原始样式中没有某个属性，也应该在需要时转换为默认值。
    遍历并转换JSON中的每一个嵌套层级，无论多深，确保所有描述性文本都被转换为API参数格式。
    """
    
    # 准备用户提示
    user_prompt = f"""
    请将以下描述性样式定义转换为可直接用于docx_namespace.py中update_runs_style和update_paragraph_style函数的参数格式：
    
    ```json
    {json.dumps(style_dict, ensure_ascii=False, indent=2)}
    ```
    
    请确保转换后的每个样式对象都可以直接作为**kwargs参数传递给函数。所有数值必须是实际的数字(不是字符串)，如"字号":"24"应转换为"size":24。
    段落间距等属性如果原值为0或很小的值(如50 twips以下)，请确保设置为0而不是省略。
    
    特别注意：
    1. 段落间距必须根据单位类型选择不同参数：
       - 以"行"为单位时使用"beforeLines"/"afterLines"参数，如"0.5行"→"beforeLines":50
       - 以"磅"为单位时使用"before"/"after"参数，如"6磅"→"before":120
    
    2. 表格样式("表格"对象)中的所有嵌套属性也必须完全转换，包括"列名行文本样式"和"数据行文本样式"中的所有属性，
    将它们的中文键名替换为英文API参数名，将描述性值转换为实际数值、布尔值或标准字符串。
    例如，将"是否为三线表":"是"转换为"is_three_line_table":true，将"行距":"多倍行距1.25倍"转换为"line":300和"line_rule":"auto"。
    3.未特别说明需要加粗就默认为False
    4. 只允许英文key，不允许出现任何中文key。
    5. 所有布尔值必须为 true/false，不能为字符串"是"或"否"。
    6. 所有数值必须为数字类型，不能为字符串。
    7. 不允许出现未定义的字段，必须严格按照API参数格式输出。
    8. 表格相关的所有嵌套字段（如"列名行文本样式"、"数据行文本样式"）也必须递归转换为英文key和标准值。
    9. 如果无法确定某个值，必须用最常见的默认值，不允许留空或用"未知"。
    10. 只返回JSON对象，不要有任何解释、注释或多余内容。
    11. "列名行文本样式"和"数据行文本样式"这两个字段的字段名必须保留中文，其内部所有属性必须全部转换为英文API参数名和标准值。
    请严格按照这些规则转换样式值，确保所有转换后的样式对象都可以直接用于函数调用。
    """

    
    try:
        # 调用AI进行转换
        print(f"使用AI({model})将样式转换为API参数格式...")

        conversion = f"{system_prompt}\n\n{user_prompt}"
        response = client.models.generate_content(
            model="gemini-2.5-flash-preview-05-20",
            contents=[conversion],
            config=types.GenerateContentConfig(
                max_output_tokens=40000,
                temperature=0.2
            ))

        # 提取响应内容
        conversion_result = response.text

        
        # 解析JSON结果
        try:
            # 查找JSON内容
            json_match = re.search(r'```json\s*([\s\S]*?)\s*```', conversion_result)
            
            if json_match:
                # 从代码块中提取JSON
                json_content = json_match.group(1)
                api_params = json.loads(json_content)
            else:
                # 尝试直接解析整个响应
                api_params = json.loads(conversion_result)
                
            print("样式转换完成，成功获取API参数格式")
            
            # 保存API参数到文件
            with open(api_params_file, 'w', encoding='utf-8') as f:
                json.dump(api_params, f, ensure_ascii=False, indent=2)
            print(f"API参数已保存到: {api_params_file}")
            
            return api_params
            
        except json.JSONDecodeError as e:
            print(f"解析API参数时出错: {e}")
            print("原始转换结果:")
            print(conversion_result)
            
            # 备用方案：使用本地转换函数
            print("尝试使用本地转换函数...")
            api_params = convert_style_values(style_dict)
            
            # 保存API参数到文件
            with open(api_params_file, 'w', encoding='utf-8') as f:
                json.dump(api_params, f, ensure_ascii=False, indent=2)
            print(f"API参数已保存到: {api_params_file} (使用本地转换)")
            
            return api_params
            
    except Exception as e:
        print(f"调用AI进行转换时出错: {e}")
        print("使用本地转换函数...")
        
        # 使用本地转换函数作为备用方案
        api_params = convert_style_values(style_dict)
        
        # 保存API参数到文件
        with open(api_params_file, 'w', encoding='utf-8') as f:
            json.dump(api_params, f, ensure_ascii=False, indent=2)
        print(f"API参数已保存到: {api_params_file} (使用本地转换)")
        
        return api_params

if __name__ == "__main__":
    # 使用示例 - 使用AI分析文档样式
    doc_path = "智算工程学院毕业设计（论文）模板2025届(1).docx"
    result = analyze_document_styles_with_ai(doc_path)

    print("\n===================== AI分析结果摘要 =====================")
    # 打印主要部分的样式摘要
    for section, styles in result.items():
        if isinstance(styles, dict):
            print(f"\n{section}部分样式特征:")
            for element_type, style in list(styles.items())[:3]:  # 只显示每部分的前3个元素类型
                print(f"  - {element_type}: {style}")
            if len(styles) > 3:
                print(f"  ... 还有{len(styles)-3}种元素类型未显示")
    print("\n详细分析结果已保存到JSON文件中")