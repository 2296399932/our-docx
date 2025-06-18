#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
提取用户表结构信息
这个脚本专门用于提取和分析文档中的数据库表结构描述表
"""

from docx_namespace import DocxElementParser
import json
import re

def print_table_content(table_content):
    """
    打印表格内容（美化格式）
    
    参数:
        table_content: 表格内容数组
    """
    if not table_content:
        print("表格内容为空")
        return
    
    # 获取每列的最大宽度
    col_widths = []
    for row in table_content:
        for i, cell in enumerate(row):
            # 确保列表长度足够
            while len(col_widths) <= i:
                col_widths.append(0)
            # 更新最大宽度
            cell_text = str(cell) if cell is not None else ""
            col_widths[i] = max(col_widths[i], len(cell_text))
    
    # 打印表格头部分隔线
    header_sep = "+"
    for width in col_widths:
        header_sep += "-" * (width + 2) + "+"
    print(header_sep)
    
    # 打印表格内容
    for row_idx, row in enumerate(table_content):
        row_str = "|"
        for i, cell in enumerate(row):
            cell_text = str(cell) if cell is not None else ""
            # 计算填充
            padding = col_widths[i] - len(cell_text)
            row_str += " " + cell_text + " " * padding + " |"
        print(row_str)
        
        # 在第一行后（表头）打印分隔符
        if row_idx == 0:
            header_sep = "+"
            for width in col_widths:
                header_sep += "-" * (width + 2) + "+"
            print(header_sep)
    
    # 打印表格底部分隔线
    footer_sep = "+"
    for width in col_widths:
        footer_sep += "-" * (width + 2) + "+"
    print(footer_sep)

def extract_table_structure(table_content):
    """
    提取表格结构信息，用于数据库表结构描述表
    
    参数:
        table_content: 表格内容数组
        
    返回:
        dict: 表结构信息，包含字段定义
    """
    # 初始化表结构
    table_structure = {
        "table_name": "",
        "fields": []
    }
    
    if not table_content or len(table_content) <= 1:
        return table_structure
    
    # 获取表头
    headers = [str(cell).lower() if cell is not None else "" for cell in table_content[0]]
    
    # 查找关键列的索引
    column_idx = next((i for i, h in enumerate(headers) if re.search(r'(列名|column|字段|field)', h)), None)
    type_idx = next((i for i, h in enumerate(headers) if re.search(r'(类型|type)', h)), None)
    desc_idx = next((i for i, h in enumerate(headers) if re.search(r'(描述|desc)', h)), None)
    
    # 如果没有找到必要的列，返回空结构
    if column_idx is None or type_idx is None:
        return table_structure
    
    # 遍历数据行提取字段信息
    for row_idx in range(1, len(table_content)):
        row = table_content[row_idx]
        if len(row) > max(column_idx, type_idx):
            field_name = str(row[column_idx]) if row[column_idx] is not None else ""
            field_type = str(row[type_idx]) if row[type_idx] is not None else ""
            
            # 如果字段名为空，跳过
            if not field_name.strip():
                continue
            
            # 创建字段定义
            field_def = {
                "name": field_name,
                "type": field_type
            }
            
            # 添加描述（如果有）
            if desc_idx is not None and len(row) > desc_idx:
                field_def["description"] = str(row[desc_idx]) if row[desc_idx] is not None else ""
            
            # 提取额外信息（如主键、自增、外键等）
            if field_def.get("description"):
                # 检查是否为主键
                if re.search(r'主键', field_def["description"], re.IGNORECASE):
                    field_def["is_primary_key"] = True
                
                # 检查是否自增
                if re.search(r'自增', field_def["description"], re.IGNORECASE):
                    field_def["is_auto_increment"] = True
                
                # 检查是否为外键
                foreign_key_match = re.search(r'外键.*?关联\s+(\w+)', field_def["description"], re.IGNORECASE)
                if foreign_key_match:
                    field_def["is_foreign_key"] = True
                    field_def["references_table"] = foreign_key_match.group(1)
            
            # 添加到字段列表
            table_structure["fields"].append(field_def)
    
    return table_structure

def generate_sql_create_table(table_structure, table_name="users"):
    """
    根据表结构生成SQL CREATE TABLE语句
    
    参数:
        table_structure: 表结构信息
        table_name: 表名
        
    返回:
        str: SQL CREATE TABLE语句
    """
    sql = f"CREATE TABLE {table_name} (\n"
    
    # 生成字段定义
    for i, field in enumerate(table_structure["fields"]):
        # 字段名和类型
        sql += f"    {field['name']} {field['type']}"
        
        # 主键
        if field.get("is_primary_key"):
            sql += " PRIMARY KEY"
        
        # 自增
        if field.get("is_auto_increment"):
            sql += " AUTO_INCREMENT"
        
        # 如果不是最后一个字段，加逗号
        if i < len(table_structure["fields"]) - 1:
            sql += ","
        
        # 添加注释
        if field.get("description"):
            sql += f" COMMENT '{field['description']}'"
        
        sql += "\n"
    
    # 外键约束（如果有）
    foreign_keys = []
    for field in table_structure["fields"]:
        if field.get("is_foreign_key") and field.get("references_table"):
            foreign_key = f"    FOREIGN KEY ({field['name']}) REFERENCES {field['references_table']}(id)"
            foreign_keys.append(foreign_key)
    
    # 添加外键约束
    if foreign_keys:
        sql += ",\n" + ",\n".join(foreign_keys) + "\n"
    
    # 结束CREATE TABLE语句
    sql += ");"
    
    return sql

def main():
    # 加载文档
    doc_path = "1.docx"
    doc = DocxElementParser(doc_path)
    
    # 在用户示例中提到的表格索引
    table_index = 214
    
    # 读取文档中的表格内容
    try:
        # 获取表格元素
        table_element = doc.elements[table_index]['element'] if isinstance(doc.elements[table_index], dict) and 'element' in doc.elements[table_index] else doc.elements[table_index]
        
        # 获取表格内容
        table_content = doc.extract_table_content(table_element)
        
        # 打印原始表格内容
        print("\n原始表格内容:")
        print_table_content(table_content)
        
        # 提取表结构
        table_structure = extract_table_structure(table_content)
        
        # 打印表结构
        print("\n提取的表结构:")
        print(json.dumps(table_structure, indent=2, ensure_ascii=False))
        
        # 生成建表SQL
        table_name = "users"  # 可以从表格内容或文档上下文推断
        sql = generate_sql_create_table(table_structure, table_name)
        
        # 打印SQL
        print(f"\n{table_name}表的创建SQL:")
        print(sql)
    
    except Exception as e:
        print(f"错误: 处理表格索引 {table_index} 时出错: {e}")

if __name__ == "__main__":
    main() 