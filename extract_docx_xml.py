#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
Word文档XML结构提取工具

此脚本将Word文档(docx)解压缩并提取其内部XML文件，
便于查看和分析文档的内部结构。
同时提供将XML目录重新打包为docx文件的功能。
"""

import os
import zipfile
import shutil
import xml.dom.minidom
import argparse

def extract_docx(docx_path, output_dir=None):
    """
    提取docx文件中的XML文件并美化输出
    
    参数:
        docx_path: docx文件路径
        output_dir: 输出目录，默认为docx文件名_xml
    """
    # 检查文件是否存在
    if not os.path.exists(docx_path):
        print(f"错误: 文件 '{docx_path}' 不存在")
        return False
    
    # 检查文件是否为docx文件
    if not docx_path.lower().endswith('.docx'):
        print(f"警告: 文件 '{docx_path}' 可能不是Word文档(docx)文件")
    
    # 设置输出目录
    if output_dir is None:
        base_name = os.path.splitext(os.path.basename(docx_path))[0]
        output_dir = f"{base_name}_xml"
    
    # 创建输出目录
    if os.path.exists(output_dir):
        print(f"警告: 输出目录 '{output_dir}' 已存在，将被覆盖")
        shutil.rmtree(output_dir)
    
    os.makedirs(output_dir)
    print(f"创建输出目录: {output_dir}")
    
    # 解压docx文件
    print(f"正在解压 {docx_path}...")
    with zipfile.ZipFile(docx_path, 'r') as zip_ref:
        zip_ref.extractall(output_dir)
    
    # 遍历提取的文件，美化XML文件
    xml_files = []
    for root, dirs, files in os.walk(output_dir):
        for file in files:
            if file.endswith('.xml') or file.endswith('.rels'):
                file_path = os.path.join(root, file)
                try:
                    # 解析XML
                    dom = xml.dom.minidom.parse(file_path)
                    # 美化XML并写回文件
                    pretty_xml = dom.toprettyxml(indent="  ")
                    with open(file_path, 'w', encoding='utf-8') as f:
                        f.write(pretty_xml)
                    
                    # 收集XML文件信息
                    rel_path = os.path.relpath(file_path, output_dir)
                    xml_files.append(rel_path)
                    print(f"美化XML: {rel_path}")
                except Exception as e:
                    print(f"警告: 无法解析或美化 {file_path}: {e}")
    
    # 创建索引文件
    index_path = os.path.join(output_dir, "索引.html")
    with open(index_path, 'w', encoding='utf-8') as f:
        f.write(f"""<!DOCTYPE html>
<html>
<head>
    <meta charset="utf-8">
    <title>Word文档结构 - {os.path.basename(docx_path)}</title>
    <style>
        body {{ font-family: Arial, sans-serif; margin: 20px; }}
        h1 {{ color: #2c3e50; }}
        ul {{ list-style-type: none; padding: 0; }}
        li {{ margin: 8px 0; }}
        a {{ color: #3498db; text-decoration: none; }}
        a:hover {{ text-decoration: underline; }}
        .file-path {{ color: #7f8c8d; font-size: 0.8em; }}
        .important {{ font-weight: bold; color: #e74c3c; }}
    </style>
</head>
<body>
    <h1>Word文档内部结构 - {os.path.basename(docx_path)}</h1>
    <p>以下是文档中的XML文件列表:</p>
    <ul>
""")
        
        # 添加重要文件到顶部
        important_files = [
            "word/document.xml",
            "word/styles.xml",
            "word/numbering.xml",
            "word/settings.xml",
            "[Content_Types].xml",
            "_rels/.rels"
        ]
        
        # 首先添加重要文件
        for file_path in important_files:
            if file_path in xml_files or any(file_path.lower() == p.lower() for p in xml_files):
                if os.path.exists(os.path.join(output_dir, file_path)):
                    f.write(f"""        <li><a href="{file_path}" class="important">{file_path}</a> <span class="file-path">(重要文件)</span></li>\n""")
        
        # 然后添加其他文件
        for file_path in sorted(xml_files):
            if not any(file_path.lower() == p.lower() for p in important_files):
                f.write(f"""        <li><a href="{file_path}">{file_path}</a></li>\n""")
        
        f.write("""    </ul>
    <h2>主要文件说明:</h2>
    <ul>
        <li><strong>word/document.xml</strong> - 包含文档的主要内容，段落、表格等</li>
        <li><strong>word/styles.xml</strong> - 定义文档中使用的样式</li>
        <li><strong>word/numbering.xml</strong> - 定义编号列表和大纲级别</li>
        <li><strong>word/settings.xml</strong> - 文档设置，如页面大小、边距等</li>
        <li><strong>[Content_Types].xml</strong> - 描述包中各部分的内容类型</li>
        <li><strong>_rels/.rels</strong> - 描述包内各部分之间的关系</li>
    </ul>
    <p><em>提示: 点击链接可查看具体XML文件内容</em></p>
</body>
</html>""")
    
    print(f"\n提取完成! 可以在 {output_dir} 目录中查看文件")
    print(f"请打开 {index_path} 以浏览文档结构")
    
    return True

def repack_xml_to_docx(xml_dir, output_docx=None):
    """
    将提取的XML目录重新打包为docx文件
    
    参数:
        xml_dir: 包含XML文件的目录路径
        output_docx: 输出的docx文件路径，默认为XML目录名去掉_xml后缀 + .docx
    
    返回:
        成功返回True，失败返回False
    """
    # 检查目录是否存在
    if not os.path.exists(xml_dir) or not os.path.isdir(xml_dir):
        print(f"错误: 目录 '{xml_dir}' 不存在或不是一个目录")
        return False
    
    # 设置输出文件路径
    if output_docx is None:
        # 假设XML目录是通过extract_docx函数创建的，尝试移除_xml后缀
        if xml_dir.endswith('_xml'):
            base_name = xml_dir[:-4]  # 移除_xml后缀
        else:
            base_name = xml_dir
        output_docx = f"{base_name}.docx"
    
    # 检查输出文件是否已存在
    if os.path.exists(output_docx):
        print(f"警告: 输出文件 '{output_docx}' 已存在，将被覆盖")
    
    print(f"正在将 {xml_dir} 打包为 {output_docx}...")
    
    try:
        # 创建临时ZIP文件
        with zipfile.ZipFile(output_docx, 'w') as zip_file:
            # 遍历XML目录中的所有文件
            for root, dirs, files in os.walk(xml_dir):
                for file in files:
                    # 跳过生成的索引文件
                    if file == "索引.html":
                        continue
                    
                    file_path = os.path.join(root, file)
                    
                    # 计算在zip中的相对路径
                    arcname = os.path.relpath(file_path, xml_dir)
                    
                    # 将文件添加到zip
                    print(f"添加文件: {arcname}")
                    zip_file.write(file_path, arcname)
        
        print(f"\n打包完成! 生成的Word文档: {output_docx}")
        return True
    except Exception as e:
        print(f"错误: 打包失败 - {str(e)}")
        # 如果有错误发生，尝试删除已创建的不完整文件
        if os.path.exists(output_docx):
            try:
                os.remove(output_docx)
            except:
                pass
        return False

def main():
    parser = argparse.ArgumentParser(description="Word文档(docx)XML结构处理工具")
    subparsers = parser.add_subparsers(dest="command", help="命令")
    
    # 解压命令
    extract_parser = subparsers.add_parser("extract", help="提取docx文件为XML")
    extract_parser.add_argument("docx_path", help="Word文档(.docx)文件路径")
    extract_parser.add_argument("--output", "-o", help="输出目录路径")
    
    # 打包命令
    repack_parser = subparsers.add_parser("repack", help="将XML目录重新打包为docx")
    repack_parser.add_argument("xml_dir", help="包含XML文件的目录路径")
    repack_parser.add_argument("--output", "-o", help="输出的docx文件路径")
    
    args = parser.parse_args()
    
    if args.command == "extract":
        extract_docx(args.docx_path, args.output)
    elif args.command == "repack":
        repack_xml_to_docx(args.xml_dir, args.output)
    else:
        # 默认行为，提取示例文件
        extract_docx("1_fixed.docx")

if __name__ == "__main__":
    # 如果没有传入命令行参数，使用默认的文件路径
    repack_xml_to_docx("2_xml", "新文件.docx")