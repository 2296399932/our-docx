#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
Word样式修改器 - 修改Word文档中的样式定义

此脚本扩展了StyleAnalyzer，添加功能用于:
11.py. 修改现有样式定义
2. 创建新样式
3. 批量应用样式修改
4. 导出修改后的文档
"""

import os
import json
import copy
import xml.etree.ElementTree as ET
from style_analyzer import StyleAnalyzer

class StyleModifier(StyleAnalyzer):
    """修改Word文档样式的类"""
    
    def __init__(self, path):
        """初始化样式修改器"""
        super().__init__(path)
        self.modified = False  # 标记是否进行了修改
        self.backup_styles = None  # 存储原始样式以便恢复
    
    def backup_current_styles(self):
        """备份当前样式"""
        if self.backup_styles is None:
            self.backup_styles = copy.deepcopy(self.style_map)
            print("已备份原始样式")
    
    def restore_styles(self):
        """恢复到原始样式"""
        if self.backup_styles is not None:
            self.style_map = copy.deepcopy(self.backup_styles)
            self._update_styles_xml()
            self.modified = True
            print("已恢复原始样式")
        else:
            print("错误: 未找到样式备份")
    
    def modify_style(self, style_id, properties):
        """
        修改指定样式的属性
        
        Args:
            style_id: 要修改的样式ID
            properties: 要修改的属性字典，格式如:
                {
                    'name': '新样式名',
                    'paragraph_properties': {
                        'alignment': 'center',
                        'indentation': {'firstLine': '420'},
                        'spacing': {'line': '360', 'lineRule': 'auto'}
                    },
                    'run_properties': {
                        'fonts': {'eastAsia': '宋体', 'ascii': 'Times New Roman'},
                        'size': '24',
                        'bold': 'true'
                    }
                }
        
        Returns:
            bool: 是否成功修改
        """
        # 检查样式是否存在
        if style_id not in self.style_map:
            print(f"错误: 样式 '{style_id}' 不存在")
            return False
        
        # 备份当前样式
        self.backup_current_styles()
        
        # 获取现有样式
        style = self.style_map[style_id]
        
        # 更新基本属性
        if 'name' in properties:
            style['name'] = properties['name']
        
        if 'basedOn' in properties:
            style['basedOn'] = properties['basedOn']
        
        if 'next' in properties:
            style['next'] = properties['next']
        
        # 更新段落属性
        if 'paragraph_properties' in properties:
            para_props = properties['paragraph_properties']
            
            # 确保style中有paragraph_properties
            if 'paragraph_properties' not in style:
                style['paragraph_properties'] = {}
            
            # 更新段落属性
            for prop, value in para_props.items():
                if isinstance(value, dict):
                    # 处理嵌套属性如indentation, spacing
                    if prop not in style['paragraph_properties']:
                        style['paragraph_properties'][prop] = {}
                    style['paragraph_properties'][prop].update(value)
                else:
                    style['paragraph_properties'][prop] = value
        
        # 更新文本属性
        if 'run_properties' in properties:
            run_props = properties['run_properties']
            
            # 确保style中有run_properties
            if 'run_properties' not in style:
                style['run_properties'] = {}
            
            # 更新文本属性
            for prop, value in run_props.items():
                if isinstance(value, dict):
                    # 处理嵌套属性如fonts
                    if prop not in style['run_properties']:
                        style['run_properties'][prop] = {}
                    style['run_properties'][prop].update(value)
                else:
                    style['run_properties'][prop] = value
        
        # 标记已修改
        self.modified = True
        
        # 更新XML
        self._update_styles_xml()
        
        # 重新计算有效样式
        self._recalculate_effective_styles()
        
        print(f"已修改样式: {style_id}")
        return True
    
    def create_style(self, style_id, style_properties):
        """
        创建新样式
        
        Args:
            style_id: 新样式ID
            style_properties: 样式属性，必须包含'type'和'name'
        
        Returns:
            bool: 是否成功创建
        """
        # 检查必要属性
        if 'type' not in style_properties or 'name' not in style_properties:
            print("错误: 新样式必须包含'type'和'name'属性")
            return False
        
        # 检查是否已存在
        if style_id in self.style_map:
            print(f"错误: 样式 '{style_id}' 已存在")
            return False
        
        # 备份当前样式
        self.backup_current_styles()
        
        # 创建新样式
        new_style = {
            'style_id': style_id,
            'type': style_properties['type'],
            'name': style_properties['name']
        }
        
        # 添加其他属性
        for key, value in style_properties.items():
            if key not in ['style_id', 'type', 'name']:
                new_style[key] = value
        
        # 添加到样式映射
        self.style_map[style_id] = new_style
        
        # 如果有基础样式，更新继承关系
        if 'basedOn' in new_style:
            parent_id = new_style['basedOn']
            if parent_id not in self.style_hierarchy:
                self.style_hierarchy[parent_id] = []
            self.style_hierarchy[parent_id].append(style_id)
        
        # 标记已修改
        self.modified = True
        
        # 更新XML
        self._update_styles_xml()
        
        # 重新计算有效样式
        self._recalculate_effective_styles()
        
        print(f"已创建样式: {style_id}")
        return True
    
    def delete_style(self, style_id):
        """
        删除样式
        
        Args:
            style_id: 要删除的样式ID
        
        Returns:
            bool: 是否成功删除
        """
        # 检查样式是否存在
        if style_id not in self.style_map:
            print(f"错误: 样式 '{style_id}' 不存在")
            return False
        
        # 检查是否有其他样式依赖于此样式
        dependent_styles = []
        for sid, style in self.style_map.items():
            if 'basedOn' in style and style['basedOn'] == style_id:
                dependent_styles.append(sid)
        
        if dependent_styles:
            print(f"警告: 无法删除样式 '{style_id}'，以下样式依赖于它: {', '.join(dependent_styles)}")
            return False
        
        # 备份当前样式
        self.backup_current_styles()
        
        # 从继承关系中移除
        if 'basedOn' in self.style_map[style_id]:
            parent_id = self.style_map[style_id]['basedOn']
            if parent_id in self.style_hierarchy and style_id in self.style_hierarchy[parent_id]:
                self.style_hierarchy[parent_id].remove(style_id)
        
        # 删除样式
        del self.style_map[style_id]
        
        # 从有效样式中移除
        if style_id in self.effective_styles:
            del self.effective_styles[style_id]
        
        # 标记已修改
        self.modified = True
        
        # 更新XML
        self._update_styles_xml()
        
        print(f"已删除样式: {style_id}")
        return True
    
    def apply_style_to_paragraphs(self, style_id, paragraph_indices):
        """
        将样式应用到指定段落
        
        Args:
            style_id: 样式ID
            paragraph_indices: 段落索引列表
        
        Returns:
            int: 成功应用样式的段落数量
        """
        # 检查样式是否存在
        if style_id not in self.style_map:
            print(f"错误: 样式 '{style_id}' 不存在")
            return 0
        
        # 检查样式类型
        style = self.style_map[style_id]
        if style.get('type') != 'paragraph':
            print(f"错误: 样式 '{style_id}' 不是段落样式")
            return 0
        
        # 计数成功应用的段落
        success_count = 0
        
        # 应用样式到每个段落
        for para_idx in paragraph_indices:
            try:
                # 获取段落元素
                para_element = self.elements[para_idx].get('element')
                if para_element is None:
                    print(f"警告: 段落 {para_idx} 不存在")
                    continue
                
                # 获取或创建pPr元素
                pPr = para_element.find(f".//{{{self.NAMESPACES['w']}}}pPr")
                if pPr is None:
                    pPr = ET.SubElement(para_element, f"{{{self.NAMESPACES['w']}}}pPr")
                
                # 获取或创建pStyle元素
                pStyle = pPr.find(f".//{{{self.NAMESPACES['w']}}}pStyle")
                if pStyle is None:
                    pStyle = ET.SubElement(pPr, f"{{{self.NAMESPACES['w']}}}pStyle")
                
                # 设置样式ID
                pStyle.set(f"{{{self.NAMESPACES['w']}}}val", style_id)
                
                success_count += 1
            except Exception as e:
                print(f"错误: 无法将样式应用到段落 {para_idx}: {e}")
        
        # 更新XML
        if success_count > 0:
            self.modified = True
            print(f"已将样式 '{style_id}' 应用到 {success_count} 个段落")
        
        return success_count
    
    def _update_styles_xml(self):
        """更新styles.xml以反映样式修改"""
        # 获取styles.xml的根元素
        styles_root = self.parts['styles'].getroot()
        
        # 清除现有样式
        for style_elem in styles_root.findall(f".//{{{self.NAMESPACES['w']}}}style", self.NAMESPACES):
            styles_root.remove(style_elem)
        
        # 添加修改后的样式
        for style_id, style in self.style_map.items():
            # 跳过默认样式
            if style_id == "默认样式":
                continue
            
            # 创建style元素
            new_style = ET.Element(f"{{{self.NAMESPACES['w']}}}style")
            new_style.set(f"{{{self.NAMESPACES['w']}}}type", style.get('type', 'paragraph'))
            new_style.set(f"{{{self.NAMESPACES['w']}}}styleId", style_id)
            
            # 添加name
            if 'name' in style:
                name_elem = ET.SubElement(new_style, f"{{{self.NAMESPACES['w']}}}name")
                name_elem.set(f"{{{self.NAMESPACES['w']}}}val", style['name'])
            
            # 添加basedOn
            if 'basedOn' in style:
                based_on_elem = ET.SubElement(new_style, f"{{{self.NAMESPACES['w']}}}basedOn")
                based_on_elem.set(f"{{{self.NAMESPACES['w']}}}val", style['basedOn'])
            
            # 添加next
            if 'next' in style:
                next_elem = ET.SubElement(new_style, f"{{{self.NAMESPACES['w']}}}next")
                next_elem.set(f"{{{self.NAMESPACES['w']}}}val", style['next'])
            
            # 添加link
            if 'link' in style:
                link_elem = ET.SubElement(new_style, f"{{{self.NAMESPACES['w']}}}link")
                link_elem.set(f"{{{self.NAMESPACES['w']}}}val", style['link'])
            
            # 添加quickFormat
            if style.get('quickFormat'):
                ET.SubElement(new_style, f"{{{self.NAMESPACES['w']}}}qFormat")
            
            # 添加uiPriority
            if 'uiPriority' in style:
                ui_priority_elem = ET.SubElement(new_style, f"{{{self.NAMESPACES['w']}}}uiPriority")
                ui_priority_elem.set(f"{{{self.NAMESPACES['w']}}}val", style['uiPriority'])
            
            # 添加段落属性
            if 'paragraph_properties' in style and style['paragraph_properties']:
                pPr = ET.SubElement(new_style, f"{{{self.NAMESPACES['w']}}}pPr")
                self._add_paragraph_properties(pPr, style['paragraph_properties'])
            
            # 添加文本属性
            if 'run_properties' in style and style['run_properties']:
                rPr = ET.SubElement(new_style, f"{{{self.NAMESPACES['w']}}}rPr")
                self._add_run_properties(rPr, style['run_properties'])
            
            # 将样式添加到styles.xml
            styles_root.append(new_style)
    
    def _add_paragraph_properties(self, pPr_elem, properties):
        """向pPr元素添加段落属性"""
        # 添加对齐方式
        if 'alignment' in properties:
            jc = ET.SubElement(pPr_elem, f"{{{self.NAMESPACES['w']}}}jc")
            jc.set(f"{{{self.NAMESPACES['w']}}}val", properties['alignment'])
        
        # 添加缩进
        if 'indentation' in properties and properties['indentation']:
            ind = ET.SubElement(pPr_elem, f"{{{self.NAMESPACES['w']}}}ind")
            for attr, val in properties['indentation'].items():
                ind.set(f"{{{self.NAMESPACES['w']}}}{attr}", str(val))
        
        # 添加间距
        if 'spacing' in properties and properties['spacing']:
            spacing = ET.SubElement(pPr_elem, f"{{{self.NAMESPACES['w']}}}spacing")
            for attr, val in properties['spacing'].items():
                spacing.set(f"{{{self.NAMESPACES['w']}}}{attr}", str(val))
        
        # 添加保持行和下一段
        if properties.get('keepNext') in ['true', True, '11.py']:
            ET.SubElement(pPr_elem, f"{{{self.NAMESPACES['w']}}}keepNext")
        
        if properties.get('keepLines') in ['true', True, '11.py']:
            ET.SubElement(pPr_elem, f"{{{self.NAMESPACES['w']}}}keepLines")
        
        # 添加大纲级别
        if 'outlineLevel' in properties:
            outline_lvl = ET.SubElement(pPr_elem, f"{{{self.NAMESPACES['w']}}}outlineLvl")
            outline_lvl.set(f"{{{self.NAMESPACES['w']}}}val", str(properties['outlineLevel']))
    
    def _add_run_properties(self, rPr_elem, properties):
        """向rPr元素添加文本属性"""
        # 添加字体
        if 'fonts' in properties and properties['fonts']:
            rFonts = ET.SubElement(rPr_elem, f"{{{self.NAMESPACES['w']}}}rFonts")
            for attr, val in properties['fonts'].items():
                rFonts.set(f"{{{self.NAMESPACES['w']}}}{attr}", val)
        
        # 添加大小
        if 'size' in properties:
            sz = ET.SubElement(rPr_elem, f"{{{self.NAMESPACES['w']}}}sz")
            sz.set(f"{{{self.NAMESPACES['w']}}}val", str(properties['size']))
        
        if 'sizeCs' in properties:
            szCs = ET.SubElement(rPr_elem, f"{{{self.NAMESPACES['w']}}}szCs")
            szCs.set(f"{{{self.NAMESPACES['w']}}}val", str(properties['sizeCs']))
        
        # 添加加粗
        if properties.get('bold') in ['true', True, '11.py']:
            ET.SubElement(rPr_elem, f"{{{self.NAMESPACES['w']}}}b")
        
        if properties.get('boldCs') in ['true', True, '11.py']:
            ET.SubElement(rPr_elem, f"{{{self.NAMESPACES['w']}}}bCs")
        
        # 添加斜体
        if properties.get('italic') in ['true', True, '11.py']:
            ET.SubElement(rPr_elem, f"{{{self.NAMESPACES['w']}}}i")
        
        # 添加颜色
        if 'color' in properties:
            color = ET.SubElement(rPr_elem, f"{{{self.NAMESPACES['w']}}}color")
            color.set(f"{{{self.NAMESPACES['w']}}}val", properties['color'])
        
        # 添加下划线
        if 'underline' in properties:
            u = ET.SubElement(rPr_elem, f"{{{self.NAMESPACES['w']}}}u")
            u.set(f"{{{self.NAMESPACES['w']}}}val", properties['underline'])
    
    def _recalculate_effective_styles(self):
        """重新计算所有有效样式"""
        self.effective_styles = {}
        for style_id in self.style_map:
            if style_id != "默认样式":
                self.effective_styles[style_id] = self._calculate_effective_style(style_id)
    
    def save_document(self, output_path=None):
        """
        保存修改后的文档
        
        Args:
            output_path: 输出文件路径，如果为None则使用原文件名+_modified
        
        Returns:
            str: 保存的文件路径
        """
        if not self.modified:
            print("警告: 文档未修改，无需保存")
            return None
        
        # 确定输出路径
        if output_path is None:
            base_name = os.path.splitext(self.path)[0]
            output_path = f"{base_name}_modified.docx"
        
        # 保存文档
        self.save(output_path)
        print(f"已保存修改后的文档: {output_path}")
        return output_path
    
    def load_styles_from_json(self, json_path):
        """
        从JSON文件加载样式定义
        
        Args:
            json_path: JSON文件路径
            
        Returns:
            bool: 是否成功加载
        """
        try:
            with open(json_path, 'r', encoding='utf-8') as f:
                styles_data = json.load(f)
            
            # 备份当前样式
            self.backup_current_styles()
            
            # 检查JSON结构
            if 'style_map' in styles_data:
                # 完整样式数据
                self.style_map = styles_data['style_map']
                if 'style_hierarchy' in styles_data:
                    self.style_hierarchy = styles_data['style_hierarchy']
                if 'effective_styles' in styles_data:
                    self.effective_styles = styles_data['effective_styles']
            else:
                # 假设是简化的样式映射
                for style_id, style_props in styles_data.items():
                    if style_id in self.style_map:
                        # 更新现有样式
                        self.modify_style(style_id, style_props)
                    else:
                        # 创建新样式
                        if 'type' in style_props and 'name' in style_props:
                            self.create_style(style_id, style_props)
                        else:
                            print(f"警告: 无法创建样式 '{style_id}'，缺少必要属性")
            
            # 更新XML
            self._update_styles_xml()
            
            # 重新计算有效样式
            self._recalculate_effective_styles()
            
            self.modified = True
            print(f"已从 {json_path} 加载样式定义")
            return True
            
        except Exception as e:
            print(f"错误: 无法从JSON加载样式: {e}")
            return False
    
    def batch_modify_styles(self, styles_dict):
        """
        批量修改多个样式
        
        Args:
            styles_dict: 样式修改字典 {style_id: properties}
            
        Returns:
            int: 成功修改的样式数量
        """
        success_count = 0
        
        for style_id, properties in styles_dict.items():
            if self.modify_style(style_id, properties):
                success_count += 1
        
        print(f"已批量修改 {success_count}/{len(styles_dict)} 个样式")
        return success_count

# 使用示例
if __name__ == "__main__":


    # 创建样式修改器实例
    modifier = StyleModifier("1.docx")

    # 修改样式1的字体大小为3号（24半磅）
    modifier.modify_style("11.py", {
        "run_properties": {
            "size": "24",  # 24半磅 = 12磅 = 小四号
            "sizeCs": "24"  # 同时设置复杂脚本的字体大小
        }
    })

    # 保存修改后的文档
    output_path = modifier.save_document("1_modified.docx")

    print(f"修改完成，已将样式1的字体改为3号（小四号），保存到：{output_path}")
    # 2. 创建新的自定义样式 "CustomStyle"
    modifier.create_style("CustomStyle", {
        "type": "paragraph",  # 样式类型：段落样式
        "name": "伟大自定义样式",  # 样式显示名称
        "basedOn": "11.py",  # 基于样式1
        # 段落属性
        "paragraph_properties": {
            "alignment": "both",  # 两端对齐
            "indentation": {
                "firstLine": "840"  # 首行缩进2字符 (约420单位/字符)
            },
            "spacing": {
                "line": "360",  # 行距值
                "lineRule": "auto",  # 行距规则：多倍行距
                "before": "120",  # 段前间距 (120 = 0.5行)
                "after": "120"  # 段后间距 (120 = 0.5行)
            }
        },
        # 字符属性
        "run_properties": {
            "fonts": {
                "eastAsia": "宋体",  # 中文字体
                "ascii": "Times New Roman",  # 英文字体
                "hAnsi": "Times New Roman"  # 西欧字体
            },
            "size": "28",  # 字号：14磅 (28半磅)
            "sizeCs": "28"  # 复杂脚本字号
        }
    })

    # 3. 将新样式应用到索引为331的段落
    modifier.apply_style_to_paragraphs("CustomStyle", [356])

    # 4. 保存修改后的文档
    output_path = modifier.save_document("1_styled.docx")

    # 打印完成信息
    print(f"已完成以下操作:")
    print(f"11.py. 修改样式1的字体为3号（小四号）")
    print(f"2. 创建新样式'自定义样式'")
    print(f"3. 将新样式应用于索引为331的段落")
    print(f"4. 文档已保存到：{output_path}")