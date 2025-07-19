#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
Word样式分析器 - 解析和可视化Word文档的样式关系

此脚本扩展了DocxElementParser，添加功能用于:
11.py. 提取所有样式定义
2. 分析样式继承链
3. 计算样式的有效属性 (应用继承)
4. 可视化样式关系
"""

import os
import json

from util.docx_namespace import DocxElementParser


class StyleAnalyzer(DocxElementParser):
    """分析Word文档样式结构的类"""

    def __init__(self, path):
        """初始化样式分析器"""
        super().__init__(path)
        self.style_map = {}  # 存储所有样式
        self.style_hierarchy = {}  # 存储样式继承关系
        self.effective_styles = {}  # 存储计算后的有效样式
        self.default_paragraph_style_id = None  # 存储默认段落样式ID
        self.default_character_style_id = None  # 存储默认字符样式ID
        self.default_table_style_id = None  # 存储默认表格样式ID
        self._analyze_styles()

    def _analyze_styles(self):
        """分析文档中的所有样式定义"""
        if self.parts['styles'] is None:
            print("警告: 找不到styles.xml")
            return

        # 获取根元素和所有样式元素
        styles_root = self.parts['styles'].getroot()
        style_elements = styles_root.findall(".//w:style", self.NAMESPACES)

        # 提取默认样式信息
        default_styles = {}
        default_elements = styles_root.findall(".//w:docDefaults", self.NAMESPACES)
        if default_elements:
            default_styles = self._extract_default_styles(default_elements[0])
            self.style_map["默认样式"] = default_styles

        # 处理每个样式元素
        for style_elem in style_elements:
            style_info = self._extract_style_info(style_elem)
            self.style_map[style_info['style_id']] = style_info

            # 检查是否为默认样式
            is_default = style_elem.get(f"{{{self.NAMESPACES['w']}}}default") == "1"
            style_type = style_elem.get(f"{{{self.NAMESPACES['w']}}}type")

            if is_default:
                if style_type == "paragraph":
                    self.default_paragraph_style_id = style_info['style_id']
                elif style_type == "character":
                    self.default_character_style_id = style_info['style_id']
                elif style_type == "table":
                    self.default_table_style_id = style_info['style_id']

            # 如果样式名称是"Normal"且没找到默认段落样式，则可能是默认样式
            if style_type == "paragraph" and 'name' in style_info:
                if style_info['name'] == "Normal" and not self.default_paragraph_style_id:
                    self.default_paragraph_style_id = style_info['style_id']

            # 添加到继承层级
            if 'basedOn' in style_info:
                parent_id = style_info['basedOn']
                if parent_id not in self.style_hierarchy:
                    self.style_hierarchy[parent_id] = []
                self.style_hierarchy[parent_id].append(style_info['style_id'])

        # 处理特殊情况：如果仍未找到默认段落样式，尝试使用ID为"1"的样式作为备选
        if not self.default_paragraph_style_id and "1" in self.style_map:
            style_info = self.style_map["1"]
            if style_info.get('type') == "paragraph":
                self.default_paragraph_style_id = "1"
                print("注意: 未找到显式标记的默认段落样式，使用ID为'1'的样式作为默认样式")

        # 计算有效样式 (应用继承)
        for style_id in self.style_map:
            if style_id != "默认样式":  # 跳过默认样式
                self.effective_styles[style_id] = self._calculate_effective_style(style_id)

    def _extract_default_styles(self, default_elem):
        """从文档默认值提取样式信息"""
        default_styles = {
            'style_id': '默认样式',
            'name': '文档默认样式',
            'type': 'default',
            'paragraph_properties': {},
            'run_properties': {}
        }

        # 提取段落默认值
        para_defaults = default_elem.find(".//w:pPrDefault", self.NAMESPACES)
        if para_defaults is not None:
            para_props = para_defaults.find(".//w:pPr", self.NAMESPACES)
            if para_props is not None:
                default_styles['paragraph_properties'] = self.extract_paragraph_style(para_defaults)

        # 提取文本运行默认值
        run_defaults = default_elem.find(".//w:rPrDefault", self.NAMESPACES)
        if run_defaults is not None:
            run_props = run_defaults.find(".//w:rPr", self.NAMESPACES)
            if run_props is not None:
                default_styles['run_properties'] = self._extract_run_properties_from_element(run_props)

        return default_styles

    def _extract_style_info(self, style_elem):
        """从样式元素提取样式信息"""
        # 基本信息
        style_id = style_elem.get(f"{{{self.NAMESPACES['w']}}}styleId")
        style_type = style_elem.get(f"{{{self.NAMESPACES['w']}}}type")

        style_info = {
            'style_id': style_id,
            'type': style_type,
            'paragraph_properties': {},
            'run_properties': {}
        }

        # 提取样式名称
        name_elem = style_elem.find(".//w:name", self.NAMESPACES)
        if name_elem is not None:
            style_info['name'] = name_elem.get(f"{{{self.NAMESPACES['w']}}}val")

        # 提取基础样式
        based_on_elem = style_elem.find(".//w:basedOn", self.NAMESPACES)
        if based_on_elem is not None:
            style_info['basedOn'] = based_on_elem.get(f"{{{self.NAMESPACES['w']}}}val")

        # 提取下一个样式
        next_elem = style_elem.find(".//w:next", self.NAMESPACES)
        if next_elem is not None:
            style_info['next'] = next_elem.get(f"{{{self.NAMESPACES['w']}}}val")

        # 提取链接样式
        link_elem = style_elem.find(".//w:link", self.NAMESPACES)
        if link_elem is not None:
            style_info['link'] = link_elem.get(f"{{{self.NAMESPACES['w']}}}val")

        # 提取大纲级别
        if style_type == 'paragraph':
            para_props = style_elem.find(".//w:pPr", self.NAMESPACES)
            if para_props is not None:

                style_info['paragraph_properties'] = self.extract_paragraph_style(style_elem)
                # 检查是否有大纲级别
                outline_elem = para_props.find(".//w:outlineLvl", self.NAMESPACES)
                if outline_elem is not None:
                    style_info['outlineLevel'] = outline_elem.get(f"{{{self.NAMESPACES['w']}}}val")

        # 提取文本属性
        run_props = style_elem.find(".//w:rPr", self.NAMESPACES)
        if run_props is not None:
            style_info['run_properties'] = self._extract_run_properties_from_element(run_props)

        # 提取其他特殊属性
        ui_priority = style_elem.find(".//w:uiPriority", self.NAMESPACES)
        if ui_priority is not None:
            style_info['uiPriority'] = ui_priority.get(f"{{{self.NAMESPACES['w']}}}val")

        # 检查快速格式标志
        q_format = style_elem.find(".//w:qFormat", self.NAMESPACES)
        if q_format is not None:
            style_info['quickFormat'] = True

        # 检查是否隐藏
        semi_hidden = style_elem.find(".//w:semiHidden", self.NAMESPACES)
        if semi_hidden is not None:
            style_info['semiHidden'] = semi_hidden.get(f"{{{self.NAMESPACES['w']}}}val", "true")

        return style_info

    def _calculate_effective_style(self, style_id, processed=None):
        """计算样式的有效属性 (应用继承)"""
        if processed is None:
            processed = set()

        # 防止循环依赖
        if style_id in processed:
            print(f"警告: 检测到样式 {style_id} 的循环依赖")
            return {}

        processed.add(style_id)

        # 获取当前样式
        if style_id not in self.style_map:
            print(f"警告: 引用了未定义的样式 {style_id}")
            return {}

        current_style = self.style_map[style_id]

        # 如果没有基础样式，直接返回当前样式
        if 'basedOn' not in current_style:
            return current_style.copy()

        # 获取基础样式的有效属性
        parent_id = current_style['basedOn']
        parent_style = self._calculate_effective_style(parent_id, processed)

        # 合并样式，当前样式优先
        effective_style = parent_style.copy()

        # 更新除基础样式引用外的所有顶层属性
        for key, value in current_style.items():
            if key != 'basedOn':
                if key in ['paragraph_properties', 'run_properties'] and key in effective_style:
                    # 合并嵌套属性
                    effective_style[key] = {**effective_style[key], **value}
                else:
                    effective_style[key] = value

        return effective_style

    def get_style_hierarchy(self):
        """获取样式的继承层级"""
        return self.style_hierarchy

    def get_style_info(self, style_id):
        """获取指定样式ID的样式信息"""
        if style_id in self.style_map:
            return self.style_map[style_id]
        return None

    def get_effective_style(self, style_id):
        """获取指定样式ID的有效样式 (应用继承后)"""
        if style_id in self.effective_styles:
            return self.effective_styles[style_id]
        return None

    def get_default_style_id(self, style_type='paragraph'):
        """
        获取指定类型的默认样式ID

        Args:
            style_type: 样式类型，可选值: 'paragraph', 'character', 'table'

        Returns:
            str: 默认样式ID，如果未找到则返回None
        """
        if style_type == 'paragraph':
            return self.default_paragraph_style_id
        elif style_type == 'character':
            return self.default_character_style_id
        elif style_type == 'table':
            return self.default_table_style_id
        return None

    def get_paragraph_complete_style_info(self, para_element):
        """
        获取段落的完整样式信息，包括模板样式、直接样式、有效样式和潜在影响

        Args:
            para_element: 段落XML元素

        Returns:
            dict: 包含以下信息的字典:
                - direct_style: 段落直接定义的样式
                - template_style: 段落引用的模板样式 (如果有)
                - default_style: 文档默认样式 (如果没有其他样式时使用)
                - effective_style: 最终有效的样式 (综合所有因素)
                - possible_influences: 可能影响段落显示的其他因素
        """
        # 1. 提取段落的直接样式
        direct_style = self.get_paragraph_style_from_element(para_element)
        runs = self.get_runs_from_paragraph(para_element)

        # 2. 获取段落的样式ID
        style_id = None
        pPr = para_element.find(".//w:pPr", self.NAMESPACES)
        if pPr is not None:
            pStyle = pPr.find(".//w:pStyle", self.NAMESPACES)
            if pStyle is not None:
                style_id = pStyle.get(f"{{{self.NAMESPACES['w']}}}val")

        # 3. 获取模板样式 (如果有)
        template_style = None
        if style_id and style_id in self.style_map:
            template_style = self.style_map[style_id]

        # 4. 获取默认样式
        default_style = None
        if self.default_paragraph_style_id and self.default_paragraph_style_id in self.style_map:
            default_style = self.style_map[self.default_paragraph_style_id]
        elif "默认样式" in self.style_map:  # 如果没有默认段落样式，使用文档默认样式
            default_style = self.style_map["默认样式"]
        elif "1" in self.style_map:  # 最后尝试使用ID为"1"的样式
            default_style = self.style_map["1"]

        # 5. 计算有效样式 (合并所有样式)
        effective_style = {}

        # 5.1 首先应用默认样式 (始终应用，无论是否有样式ID)
        if default_style:
            import copy
            # 深度复制默认样式的所有属性
            for key, value in default_style.items():
                if key not in ['style_id', 'name', 'type']:
                    if isinstance(value, dict):
                        # 对字典类型进行深拷贝
                        effective_style[key] = copy.deepcopy(value)
                    else:
                        effective_style[key] = value

        # 5.2 然后应用模板样式 (如果有)
        if template_style and style_id in self.effective_styles:
            template_effective = self.effective_styles[style_id]
            for key, value in template_effective.items():
                if key not in ['style_id', 'name', 'type']:
                    if isinstance(value, dict) and key in effective_style:
                        # 对于嵌套字典，使用深度合并
                        for subkey, subvalue in value.items():
                            if isinstance(subvalue, dict) and subkey in effective_style[key]:
                                # 再次深度合并
                                effective_style[key][subkey] = {**effective_style[key][subkey], **subvalue}
                            else:
                                effective_style[key][subkey] = subvalue
                    else:
                        effective_style[key] = value

        # 5.3 最后应用直接样式 (覆盖前面的设置)
        for key, value in direct_style.items():
            if key not in ['style_id'] and value:  # 跳过空值
                # 特殊处理某些顶层属性，将它们映射到正确的嵌套位置
                if key == 'alignment' and 'paragraph_properties' in effective_style:
                    effective_style['paragraph_properties']['alignment'] = value
                elif key == 'indentation' and 'paragraph_properties' in effective_style:
                    # 直接替换继承的缩进属性
                    effective_style['paragraph_properties']['indentation'] = value
                    print()
                elif key == 'spacing' and 'paragraph_properties' in effective_style:
                    # 直接替换继承的间距属性
                    effective_style['paragraph_properties']['spacing'] = value
                elif key == 'run_properties':
                    if not effective_style.get('run_properties'):
                        effective_style['run_properties'] = {}
                    # 合并run_properties中的嵌套字典
                    for subkey, subvalue in value.items():
                        if isinstance(subvalue, dict) and subkey in effective_style['run_properties']:
                            effective_style['run_properties'][subkey].update(subvalue)
                        else:
                            effective_style['run_properties'][subkey] = subvalue
                else:
                    # 其他属性直接复制到顶层
                    effective_style[key] = value

        # 6. 检查其他可能的影响因素
        possible_influences = []

        # 6.1 检查是否有上一个段落可能影响样式
        try:
            parent = para_element.getparent()
            if parent is not None:
                siblings = parent.findall(f".//{{{self.NAMESPACES['w']}}}p")
                idx = siblings.index(para_element)
                if idx > 0:
                    prev_para = siblings[idx - 1]
                    prev_style_id = None
                    prev_pPr = prev_para.find(".//w:pPr", self.NAMESPACES)
                    if prev_pPr is not None:
                        prev_pStyle = prev_pPr.find(".//w:pStyle", self.NAMESPACES)
                        if prev_pStyle is not None:
                            prev_style_id = prev_pStyle.get(f"{{{self.NAMESPACES['w']}}}val")

                    if prev_style_id:
                        # 检查前一段落样式是否会影响当前段落
                        prev_style = self.style_map.get(prev_style_id)
                        if prev_style and 'next' in prev_style and prev_style['next'] == style_id:
                            possible_influences.append({
                                'type': 'previous_paragraph',
                                'style_id': prev_style_id,
                                'relationship': 'next attribute'
                            })
        except:
            # 出错时忽略，这只是额外信息
            pass

        # 7. 整合结果
        result = {
            'direct_style': direct_style,
            'effective_style': effective_style,
            'possible_influences': possible_influences
        }

        if style_id:
            result['style_id'] = style_id
            result['template_style'] = template_style
        else:
            result['style_id'] = None
            result['note'] = '未找到样式ID，使用默认样式'

        if default_style:
            result['default_style'] = default_style

        return result

    def print_style_info(self, style_id=None):
        """打印样式信息"""
        if style_id is None:
            # 打印所有样式
            print("\n=== 文档样式信息 ===")
            for sid in self.style_map:
                style = self.style_map[sid]
                print(f"\n样式ID: {sid}")
                if 'name' in style:
                    print(f"名称: {style['name']}")
                print(f"类型: {style.get('type', '未知')}")
                if 'basedOn' in style:
                    print(f"基于: {style['basedOn']}")
                if 'next' in style:
                    print(f"下一样式: {style['next']}")
                if 'outlineLevel' in style:
                    print(f"大纲级别: {style['outlineLevel']}")
                print("---")
        else:
            # 打印特定样式
            if style_id in self.style_map:
                style = self.style_map[style_id]
                effective_style = self.get_effective_style(style_id)

                print(f"\n=== 样式 '{style_id}' 信息 ===")
                if 'name' in style:
                    print(f"名称: {style['name']}")
                print(f"类型: {style.get('type', '未知')}")

                if 'basedOn' in style:
                    print(f"基于: {style['basedOn']}")
                    if style['basedOn'] in self.style_map:
                        base_name = self.style_map[style['basedOn']].get('name', style['basedOn'])
                        print(f"  (基础样式名称: {base_name})")

                if 'next' in style:
                    print(f"下一样式: {style['next']}")
                if 'outlineLevel' in style:
                    print(f"大纲级别: {style['outlineLevel']}")

                # 打印段落属性
                if 'paragraph_properties' in style and style['paragraph_properties']:
                    print("\n段落属性:")
                    for prop, value in style['paragraph_properties'].items():
                        if isinstance(value, dict):
                            print(f"  {prop}:")
                            for subprop, subval in value.items():
                                print(f"    {subprop}: {subval}")
                        else:
                            print(f"  {prop}: {value}")

                # 打印文本属性
                if 'run_properties' in style and style['run_properties']:
                    print("\n文本属性:")
                    for prop, value in style['run_properties'].items():
                        if isinstance(value, dict):
                            print(f"  {prop}:")
                            for subprop, subval in value.items():
                                print(f"    {subprop}: {subval}")
                        else:
                            print(f"  {prop}: {value}")

                # 打印有效样式 (通过继承计算)
                if effective_style and effective_style != style:
                    print("\n有效属性 (应用继承后):")

                    # 打印有效段落属性
                    if 'paragraph_properties' in effective_style and effective_style['paragraph_properties']:
                        print("\n有效段落属性:")
                        for prop, value in effective_style['paragraph_properties'].items():
                            if isinstance(value, dict):
                                print(f"  {prop}:")
                                for subprop, subval in value.items():
                                    print(f"    {subprop}: {subval}")
                            else:
                                print(f"  {prop}: {value}")

                    # 打印有效文本属性
                    if 'run_properties' in effective_style and effective_style['run_properties']:
                        print("\n有效文本属性:")
                        for prop, value in effective_style['run_properties'].items():
                            if isinstance(value, dict):
                                print(f"  {prop}:")
                                for subprop, subval in value.items():
                                    print(f"    {subprop}: {subval}")
                            else:
                                print(f"  {prop}: {value}")
            else:
                print(f"错误: 样式 '{style_id}' 未找到")

    def print_style_hierarchy(self):
        """打印样式继承层级"""
        print("\n=== 样式继承层级 ===")
        # 找到顶级样式 (没有父级的样式)
        top_styles = set(self.style_map.keys()) - set(
            style for styles in self.style_hierarchy.values() for style in styles)
        for style_id in sorted(top_styles):
            self._print_style_tree(style_id)

    def print_default_styles(self):
        """打印默认样式信息"""
        print("\n=== 默认样式信息 ===")

        # 打印文档默认样式
        if "默认样式" in self.style_map:
            print("\n文档默认样式:")
            default_style = self.style_map["默认样式"]
            if 'paragraph_properties' in default_style and default_style['paragraph_properties']:
                print("  段落属性:")
                for key, value in default_style['paragraph_properties'].items():
                    print(f"    {key}: {value}")
            if 'run_properties' in default_style and default_style['run_properties']:
                print("  文本属性:")
                for key, value in default_style['run_properties'].items():
                    print(f"    {key}: {value}")

        # 打印默认段落样式
        if self.default_paragraph_style_id:
            print(f"\n默认段落样式 (ID: {self.default_paragraph_style_id}):")
            if self.default_paragraph_style_id in self.style_map:
                style = self.style_map[self.default_paragraph_style_id]
                if 'name' in style:
                    print(f"  名称: {style['name']}")
                if 'paragraph_properties' in style and style['paragraph_properties']:
                    print("  段落属性:")
                    for prop, value in style['paragraph_properties'].items():
                        if isinstance(value, dict):
                            print(f"    {prop}:")
                            for subprop, subval in value.items():
                                print(f"      {subprop}: {subval}")
                        else:
                            print(f"    {prop}: {value}")
                if 'run_properties' in style and style['run_properties']:
                    print("  文本属性:")
                    for prop, value in style['run_properties'].items():
                        if isinstance(value, dict):
                            print(f"    {prop}:")
                            for subprop, subval in value.items():
                                print(f"      {subprop}: {subval}")
                        else:
                            print(f"    {prop}: {value}")
            else:
                print("  警告: 找不到此样式的详细信息")
        else:
            print("\n未找到默认段落样式")

        # 打印默认字符样式
        if self.default_character_style_id:
            print(f"\n默认字符样式 (ID: {self.default_character_style_id}):")
            if self.default_character_style_id in self.style_map:
                style = self.style_map[self.default_character_style_id]
                if 'name' in style:
                    print(f"  名称: {style['name']}")
                if 'run_properties' in style and style['run_properties']:
                    print("  文本属性:")
                    for prop, value in style['run_properties'].items():
                        if isinstance(value, dict):
                            print(f"    {prop}:")
                            for subprop, subval in value.items():
                                print(f"      {subprop}: {subval}")
                        else:
                            print(f"    {prop}: {value}")
            else:
                print("  警告: 找不到此样式的详细信息")
        else:
            print("\n未找到默认字符样式")

        # 打印默认表格样式
        if self.default_table_style_id:
            print(f"\n默认表格样式 (ID: {self.default_table_style_id}):")
            if self.default_table_style_id in self.style_map:
                style = self.style_map[self.default_table_style_id]
                if 'name' in style:
                    print(f"  名称: {style['name']}")
                if 'tblPr' in style and style['tblPr']:
                    print("  表格属性:")
                    for prop, value in style['tblPr'].items():
                        if isinstance(value, dict):
                            print(f"    {prop}:")
                            for subprop, subval in value.items():
                                print(f"      {subprop}: {subval}")
                        else:
                            print(f"    {prop}: {value}")
            else:
                print("  警告: 找不到此样式的详细信息")
        else:
            print("\n未找到默认表格样式")

    def _print_style_tree(self, style_id, level=0, prefix=''):
        """递归打印样式树"""
        if style_id not in self.style_map:
            return

        style = self.style_map[style_id]
        style_name = style.get('name', style_id)

        # 打印当前样式
        if level == 0:
            print(f"{prefix}└─ {style_id} ({style_name})")
        else:
            print(f"{prefix}└─ {style_id} ({style_name})")

        # 打印子样式
        children = self.style_hierarchy.get(style_id, [])
        for i, child in enumerate(children):
            is_last = i == len(children) - 1
            new_prefix = prefix + ('    ' if level == 0 or is_last else '│   ')
            self._print_style_tree(child, level + 1, new_prefix)

    def export_styles_json(self, output_path):
        """导出样式到JSON文件"""
        styles_data = {
            'style_map': self.style_map,
            'style_hierarchy': self.style_hierarchy,
            'effective_styles': self.effective_styles
        }

        with open(output_path, 'w', encoding='utf-8') as f:
            json.dump(styles_data, f, ensure_ascii=False, indent=2)

        print(f"样式信息已导出到: {output_path}")

    def _merge_font_properties(self, base_fonts, override_fonts):
        """
        智能合并字体属性，保持语言特定字体的正确应用

        Args:
            base_fonts: 基础字体属性字典
            override_fonts: 需要覆盖的字体属性字典

        Returns:
            dict: 合并后的字体属性字典
        """
        if not base_fonts:
            return override_fonts.copy() if override_fonts else {}

        if not override_fonts:
            return base_fonts.copy()

        result = base_fonts.copy()

        # 处理四种主要字体属性: ascii(英文), eastAsia(中文等), hAnsi(欧洲字符), cs(复杂文字系统)
        for attr in ['ascii', 'hAnsi', 'eastAsia', 'cs']:
            if attr in override_fonts:
                result[attr] = override_fonts[attr]

        # 处理字体提示
        if 'hint' in override_fonts:
            result['hint'] = override_fonts['hint']

        return result

    def _get_theme_fonts(self):
        """
        从文档主题中提取字体信息
        
        Returns:
            dict: 包含主题字体映射的字典
        """
        theme_fonts = {
            'majorEastAsia': '宋体',  # 默认主要东亚字体
            'minorEastAsia': '宋体',  # 默认次要东亚字体
            'majorAscii': 'Times New Roman',  # 默认主要ASCII字体
            'minorAscii': 'Times New Roman',  # 默认次要ASCII字体
            'majorHAnsi': 'Times New Roman',  # 默认主要拉丁字体
            'minorHAnsi': 'Times New Roman',  # 默认次要拉丁字体
            'majorBidi': 'Times New Roman',  # 默认主要复杂文字字体
            'minorBidi': 'Times New Roman',  # 默认次要复杂文字字体
        }
        
        # 尝试从主题文件中获取字体信息
        if 'theme1' in self.parts and self.parts['theme1'] is not None:
            theme_root = self.parts['theme1'].getroot()
            
            # 提取主要字体
            major_fonts = theme_root.findall(".//a:majorFont", self.NAMESPACES)
            if major_fonts:
                # 拉丁字体(ASCII和HANSI)
                latin = major_fonts[0].find(".//a:latin", self.NAMESPACES)
                if latin is not None and latin.get('typeface'):
                    theme_fonts['majorAscii'] = latin.get('typeface')
                    theme_fonts['majorHAnsi'] = latin.get('typeface')
                
                # 东亚字体
                ea = major_fonts[0].find(".//a:ea", self.NAMESPACES)
                if ea is not None and ea.get('typeface'):
                    theme_fonts['majorEastAsia'] = ea.get('typeface')
                
                # 复杂文字字体
                cs = major_fonts[0].find(".//a:cs", self.NAMESPACES)
                if cs is not None and cs.get('typeface'):
                    theme_fonts['majorBidi'] = cs.get('typeface')
            
            # 提取次要字体
            minor_fonts = theme_root.findall(".//a:minorFont", self.NAMESPACES)
            if minor_fonts:
                # 拉丁字体(ASCII和HANSI)
                latin = minor_fonts[0].find(".//a:latin", self.NAMESPACES)
                if latin is not None and latin.get('typeface'):
                    theme_fonts['minorAscii'] = latin.get('typeface')
                    theme_fonts['minorHAnsi'] = latin.get('typeface')
                
                # 东亚字体
                ea = minor_fonts[0].find(".//a:ea", self.NAMESPACES)
                if ea is not None and ea.get('typeface'):
                    theme_fonts['minorEastAsia'] = ea.get('typeface')
                
                # 复杂文字字体
                cs = minor_fonts[0].find(".//a:cs", self.NAMESPACES)
                if cs is not None and cs.get('typeface'):
                    theme_fonts['minorBidi'] = cs.get('typeface')
                    
            # 检查针对中文的特殊字体设置
            hans_fonts = theme_root.findall(".//a:font[@script='Hans']", self.NAMESPACES)
            for font in hans_fonts:
                if font.get('typeface'):
                    # 检查该字体属于majorFont还是minorFont
                    parent = font.getparent().getparent()
                    if parent.tag.endswith('majorFont'):
                        theme_fonts['majorEastAsia'] = font.get('typeface')
                    elif parent.tag.endswith('minorFont'):
                        theme_fonts['minorEastAsia'] = font.get('typeface')
        
        return theme_fonts

    def _get_doc_default_fonts(self):
        """
        从styles.xml中的w:docDefaults元素获取文档默认字体
        
        Returns:
            dict: 包含默认字体信息的字典
        """
        default_fonts = {
            'ascii': 'Times New Roman',
            'hAnsi': 'Times New Roman', 
            'eastAsia': '宋体',
            'cs': 'Times New Roman'
        }
        
        # 检查styles.xml是否存在
        if self.parts['styles'] is None:
            return default_fonts
            
        # 获取docDefaults元素
        styles_root = self.parts['styles'].getroot()
        doc_defaults = styles_root.find(".//w:docDefaults", self.NAMESPACES)
        
        if doc_defaults is not None:
            # 获取rPrDefault元素
            run_defaults = doc_defaults.find(".//w:rPrDefault", self.NAMESPACES)
            if run_defaults is not None:
                run_props = run_defaults.find(".//w:rPr", self.NAMESPACES)
                if run_props is not None:
                    # 获取默认字体设置
                    fonts_elem = run_props.find(".//w:rFonts", self.NAMESPACES)
                    if fonts_elem is not None:
                        # 获取各语言的默认字体
                        for attr in ['ascii', 'hAnsi', 'eastAsia', 'cs']:
                            val = fonts_elem.get(f"{{{self.NAMESPACES['w']}}}{attr}")
                            if val is not None:
                                default_fonts[attr] = val
        
        return default_fonts

    def _get_doc_default_properties(self):
        """
        从styles.xml中的w:docDefaults元素获取文档默认属性(字体、字号等)
        
        优先顺序:
        1. 先从w:docDefaults中获取
        2. 如果没有，则从styleId为1的默认样式(Normal)中获取
        3. 最后才使用硬编码默认值
        
        Returns:
            dict: 包含默认属性的字典
        """
        # 初始化默认值，只有在前面两步都没找到时才会使用
        default_props = {
            'fonts': {
                'ascii': 'Times New Roman',
                'hAnsi': 'Times New Roman', 
                'eastAsia': '宋体',
                'cs': 'Times New Roman'
            },
            'size': '24',       # 默认字号为小四(24值/12磅)
            'size_cs': '24'     # 默认复杂文本字号
        }
        
        # 检查styles.xml是否存在
        if self.parts['styles'] is None:
            return default_props
            
        # 获取styles根元素
        styles_root = self.parts['styles'].getroot()
        
        # 步骤1: 从docDefaults获取默认属性
        doc_defaults_props = self._extract_doc_defaults(styles_root)
        
        # 步骤2: 如果docDefaults中没找到字号，从Normal样式(styleId=1)中获取
        if not doc_defaults_props.get('size') or not doc_defaults_props.get('size_cs'):
            normal_props = self._extract_normal_style_props(styles_root)
            
            # 合并属性，Normal样式只在docDefaults没有的情况下使用
            if not doc_defaults_props.get('size') and normal_props.get('size'):
                doc_defaults_props['size'] = normal_props['size']
            
            if not doc_defaults_props.get('size_cs') and normal_props.get('size_cs'):
                doc_defaults_props['size_cs'] = normal_props['size_cs']
            
            # 合并字体属性
            for attr, font in normal_props.get('fonts', {}).items():
                if attr not in doc_defaults_props.get('fonts', {}):
                    doc_defaults_props['fonts'][attr] = font
        
        # 步骤3: 如果还是没有，使用默认值
        if doc_defaults_props:
            # 确保字体字典存在
            if 'fonts' not in doc_defaults_props:
                doc_defaults_props['fonts'] = default_props['fonts']
                
            # 合并任何缺失的值
            for attr, val in default_props['fonts'].items():
                if attr not in doc_defaults_props['fonts']:
                    doc_defaults_props['fonts'][attr] = val
                    
            # 使用默认字号，如果没有在前面找到
            if not doc_defaults_props.get('size'):
                doc_defaults_props['size'] = default_props['size']
                
            if not doc_defaults_props.get('size_cs'):
                doc_defaults_props['size_cs'] = default_props['size_cs']
                
            return doc_defaults_props
            
        # 如果什么都没找到，返回默认值
        return default_props
        
    def _extract_doc_defaults(self, styles_root):
        """
        从docDefaults元素中提取默认属性
        
        Args:
            styles_root: styles.xml的根元素
            
        Returns:
            dict: 包含从docDefaults中提取的默认属性
        """
        props = {
            'fonts': {}
        }
        
        doc_defaults = styles_root.find(".//w:docDefaults", self.NAMESPACES)
        if doc_defaults is None:
            return props
            
        # 获取rPrDefault元素
        run_defaults = doc_defaults.find(".//w:rPrDefault", self.NAMESPACES)
        if run_defaults is not None:
            run_props = run_defaults.find(".//w:rPr", self.NAMESPACES)
            if run_props is not None:
                # 获取默认字体设置
                fonts_elem = run_props.find(".//w:rFonts", self.NAMESPACES)
                if fonts_elem is not None:
                    for attr in ['ascii', 'hAnsi', 'eastAsia', 'cs']:
                        val = fonts_elem.get(f"{{{self.NAMESPACES['w']}}}{attr}")
                        if val is not None:
                            props['fonts'][attr] = val
                
                # 获取默认字号
                sz_elem = run_props.find(".//w:sz", self.NAMESPACES)
                if sz_elem is not None:
                    props['size'] = sz_elem.get(f"{{{self.NAMESPACES['w']}}}val")
                
                # 获取复杂文本字号
                szCs_elem = run_props.find(".//w:szCs", self.NAMESPACES)
                if szCs_elem is not None:
                    props['size_cs'] = szCs_elem.get(f"{{{self.NAMESPACES['w']}}}val")
                    
                # 获取其他可能的属性...
        
        return props
        
    def _extract_normal_style_props(self, styles_root):
        """
        从Normal样式(styleId=1)中提取属性
        
        Args:
            styles_root: styles.xml的根元素
            
        Returns:
            dict: 包含从Normal样式中提取的属性
        """
        props = {
            'fonts': {}
        }
        
        # 尝试获取styleId为1的样式(通常是Normal)
        normal_style = styles_root.find(".//w:style[@w:styleId='1']", self.NAMESPACES)
        
        # 如果没找到，尝试找到默认段落样式
        if normal_style is None:
            normal_style = styles_root.find(".//w:style[@w:default='1'][@w:type='paragraph']", self.NAMESPACES)
        
        # 如果仍然没找到，尝试找名为"Normal"的样式
        if normal_style is None:
            normal_style = styles_root.find(".//w:style[w:name/@w:val='Normal'][@w:type='paragraph']", self.NAMESPACES)
            
        if normal_style is None:
            return props
            
        # 从Normal样式中提取run属性
        run_props = normal_style.find(".//w:rPr", self.NAMESPACES)
        if run_props is not None:
            # 获取字体
            fonts_elem = run_props.find(".//w:rFonts", self.NAMESPACES)
            if fonts_elem is not None:
                for attr in ['ascii', 'hAnsi', 'eastAsia', 'cs']:
                    val = fonts_elem.get(f"{{{self.NAMESPACES['w']}}}{attr}")
                    if val is not None:
                        props['fonts'][attr] = val
            
            # 获取字号
            sz_elem = run_props.find(".//w:sz", self.NAMESPACES)
            if sz_elem is not None:
                props['size'] = sz_elem.get(f"{{{self.NAMESPACES['w']}}}val")
            
            # 获取复杂文本字号
            szCs_elem = run_props.find(".//w:szCs", self.NAMESPACES)
            if szCs_elem is not None:
                props['size_cs'] = szCs_elem.get(f"{{{self.NAMESPACES['w']}}}val")
                
            # 获取其他可能的属性...
        
        return props

    def get_run_complete_style_info(self, para_element, run_element, run_index=None):
        """
        获取run的完整样式信息，包括直接样式、模板样式、段落继承样式和有效样式

        正确处理样式继承和覆盖关系，确保直接样式覆盖继承样式

        Args:
            para_element: 包含run的段落XML元素
            run_element: run的XML元素
            run_index: run在段落中的索引（可选）

        Returns:
            dict: 包含以下信息的字典:
                - direct_style: run直接定义的样式
                - template_style: run引用的模板样式 (如果有)
                - paragraph_style: 从段落继承的样式
                - effective_style: 最终有效的样式 (综合所有因素)
                - possible_influences: 可能影响run显示的其他因素
        """
        # 1. 提取run的直接样式
        direct_style = self.get_run_style_from_element(run_element)

        # 2. 获取run的样式ID
        style_id = None
        rPr = run_element.find(".//w:rPr", self.NAMESPACES)
        if rPr is not None:
            rStyle = rPr.find(".//w:rStyle", self.NAMESPACES)
            if rStyle is not None:
                style_id = rStyle.get(f"{{{self.NAMESPACES['w']}}}val")

        # 3. 获取run的模板样式 (如果有)
        template_style = None
        if style_id and style_id in self.style_map:
            template_style = self.style_map[style_id]

        # 4. 获取默认字符样式
        default_character_style = None
        if self.default_character_style_id and self.default_character_style_id in self.style_map:
            default_character_style = self.style_map[self.default_character_style_id]

        # 5. 获取段落的样式信息，因为run会继承段落样式
        paragraph_style_info = self.get_paragraph_complete_style_info(para_element)
        paragraph_effective_style = paragraph_style_info.get('effective_style', {})
        
        # 获取段落样式ID，用于处理只有hint属性的情况
        para_style_id = paragraph_style_info.get('style_id')

        # 6. 获取文档默认属性(字体、字号等)
        doc_default_props = self._get_doc_default_properties()
        doc_default_fonts = doc_default_props['fonts']

        # 7. 计算run的有效样式 (合并所有样式)
        effective_style = {}
        import copy

        # 7.1 首先应用文档默认样式 (如果有)
        if "默认样式" in self.style_map and 'run_properties' in self.style_map["默认样式"]:
            effective_style['run_properties'] = copy.deepcopy(self.style_map["默认样式"]['run_properties'])
            
            # 确保字体字典存在
            if 'fonts' not in effective_style['run_properties']:
                effective_style['run_properties']['fonts'] = {}
                
            # 将文档默认字体应用到字体字典中
            for attr, font in doc_default_fonts.items():
                effective_style['run_properties']['fonts'][attr] = font
                
            # 应用默认字号
            effective_style['run_properties']['size'] = doc_default_props.get('size')
            effective_style['run_properties']['size_cs'] = doc_default_props.get('size_cs')
        else:
            # 如果没有找到默认样式，创建一个包含默认属性的run_properties
            effective_style['run_properties'] = {
                'fonts': doc_default_fonts.copy(),
                'size': doc_default_props.get('size'),
                'size_cs': doc_default_props.get('size_cs')
            }
        
        # 7.2 然后应用默认字符样式 (如果有)
        if default_character_style and 'run_properties' in default_character_style:
            # 合并字体属性
            if 'fonts' in default_character_style['run_properties'] and 'fonts' in effective_style['run_properties']:
                effective_style['run_properties']['fonts'] = self._merge_font_properties(
                    effective_style['run_properties']['fonts'], 
                    default_character_style['run_properties'].get('fonts', {})
                )

            # 合并其他属性
            for key, value in default_character_style['run_properties'].items():
                if key != 'fonts':  # 字体已经处理过了
                    effective_style['run_properties'][key] = copy.deepcopy(value)

        # 7.3 然后应用段落的有效样式中与run相关的部分
        if 'run_properties' in paragraph_effective_style:
            # 合并run_properties，段落样式覆盖默认字符样式
            for subkey, subvalue in paragraph_effective_style['run_properties'].items():
                if subkey == 'fonts' and 'fonts' in effective_style['run_properties']:
                    # 使用专门的字体合并方法
                    effective_style['run_properties']['fonts'] = self._merge_font_properties(
                        effective_style['run_properties']['fonts'], subvalue)
                elif isinstance(subvalue, dict) and subkey in effective_style['run_properties']:
                    # 对嵌套字典进行深度合并
                    effective_style['run_properties'][subkey] = {**effective_style['run_properties'][subkey],
                                                                 **subvalue}
                else:
                    effective_style['run_properties'][subkey] = subvalue

        # 7.4 然后应用run的模板样式 (如果有)
        if template_style and style_id in self.effective_styles:
            template_effective = self.effective_styles[style_id]
            if 'run_properties' in template_effective:
                # 合并run_properties，模板样式覆盖段落样式
                for subkey, subvalue in template_effective['run_properties'].items():
                    if subkey == 'fonts' and 'fonts' in effective_style['run_properties']:
                        # 使用专门的字体合并方法
                        effective_style['run_properties']['fonts'] = self._merge_font_properties(
                            effective_style['run_properties']['fonts'], copy.deepcopy(subvalue))
                    elif isinstance(subvalue, dict) and subkey in effective_style['run_properties']:
                        # 对于嵌套字典(如fonts)，进行深度合并而不是完全替换
                        effective_style['run_properties'][subkey] = {
                            **effective_style['run_properties'][subkey],
                            **copy.deepcopy(subvalue)
                        }
                    else:
                        effective_style['run_properties'][subkey] = subvalue

        # 7.5 最后应用run的直接样式 (这部分应该完全覆盖之前的设置)
        # 直接从run的XML元素提取样式
        if rPr is not None:
            # 检查是否是只有hint属性的情况
            only_has_hint = True
            for elem in rPr:
                if not elem.tag.endswith('rFonts'):  # 如果有除rFonts外的其他属性
                    only_has_hint = False
                    break
                    
            if only_has_hint:
                fonts_elem = rPr.find(".//w:rFonts", self.NAMESPACES)
                if fonts_elem is not None:
                    # 检查是否确实只有hint属性
                    has_font_attrs = False
                    for attr in ['ascii', 'hAnsi', 'eastAsia', 'cs']:
                        if fonts_elem.get(f"{{{self.NAMESPACES['w']}}}{attr}") is not None:
                            has_font_attrs = True
                            break
                            
                    hint_val = fonts_elem.get(f"{{{self.NAMESPACES['w']}}}hint")
                    if hint_val and not has_font_attrs:
                        # 只有hint属性的情况
                        # 1. 确保字体字典存在
                        if 'fonts' not in effective_style['run_properties']:
                            effective_style['run_properties']['fonts'] = {}
                        
                        # 2. 首先检查段落是否有样式ID
                        if para_style_id and para_style_id in self.style_map:
                            # 如果段落有样式，使用段落样式中定义的字体和字号
                            para_style = self.style_map[para_style_id]
                            if 'run_properties' in para_style:
                                # 应用段落样式中的字体设置
                                if 'fonts' in para_style['run_properties']:
                                    for attr, font in para_style['run_properties']['fonts'].items():
                                        effective_style['run_properties']['fonts'][attr] = font
                                
                                # 应用段落样式中的字号设置
                                if 'size' in para_style['run_properties']:
                                    effective_style['run_properties']['size'] = para_style['run_properties']['size']
                                if 'size_cs' in para_style['run_properties']:
                                    effective_style['run_properties']['size_cs'] = para_style['run_properties']['size_cs']
                        else:
                            # 如果段落没有样式，根据hint值选择合适的默认字体和字号
                            if hint_val == 'eastAsia':
                                # 对于东亚文本，使用文档默认的eastAsia字体
                                effective_style['run_properties']['fonts']['eastAsia'] = doc_default_fonts.get('eastAsia', '宋体')
                                # 使用默认字号(通常为小四/12磅)
                                effective_style['run_properties']['size'] = doc_default_props.get('size')
                                effective_style['run_properties']['size_cs'] = doc_default_props.get('size_cs')
                            elif hint_val == 'default':
                                # 对于default提示，应用所有默认字体和字号
                                for attr, font in doc_default_fonts.items():
                                    effective_style['run_properties']['fonts'][attr] = font
                                effective_style['run_properties']['size'] = doc_default_props.get('size')
                                effective_style['run_properties']['size_cs'] = doc_default_props.get('size_cs')
                        
                        # 添加hint属性
                        effective_style['run_properties']['fonts']['hint'] = hint_val
            else:
                # 处理字体属性
                fonts_elem = rPr.find(".//w:rFonts", self.NAMESPACES)
                if fonts_elem is not None:
                    # 确保fonts字典存在
                    if 'fonts' not in effective_style['run_properties']:
                        effective_style['run_properties']['fonts'] = {}

                    # 提取run中指定的字体属性
                    current_fonts = {}
                    for attr in ['ascii', 'hAnsi', 'eastAsia', 'cs']:
                        val = fonts_elem.get(f"{{{self.NAMESPACES['w']}}}{attr}")
                        if val is not None:
                            current_fonts[attr] = val

                        # 处理hint属性
                        hint_val = fonts_elem.get(f"{{{self.NAMESPACES['w']}}}hint")
                        if hint_val:
                            effective_style['run_properties']['fonts']['hint'] = hint_val
                    
                    # 合并字体属性，run中指定的覆盖继承的
                    if current_fonts:
                        effective_style['run_properties']['fonts'] = self._merge_font_properties(
                            effective_style['run_properties']['fonts'], current_fonts)

                    # 处理字号
                    sz_elem = rPr.find(".//w:sz", self.NAMESPACES)
                    if sz_elem is not None:
                        effective_style['run_properties']['size'] = sz_elem.get(f"{{{self.NAMESPACES['w']}}}val")
                    
                    szCs_elem = rPr.find(".//w:szCs", self.NAMESPACES)
                    if szCs_elem is not None:
                        effective_style['run_properties']['size_cs'] = szCs_elem.get(f"{{{self.NAMESPACES['w']}}}val")

        # 8. 检查其他可能的影响因素
        possible_influences = []

        # 8.1 检查是否有特殊字符格式化影响
        if run_element.find(".//{%s}tab" % self.NAMESPACES['w']) is not None:
            possible_influences.append({
                'type': 'special_character',
                'description': 'Run包含制表符，可能影响格式化'
            })

        if run_element.find(".//{%s}br" % self.NAMESPACES['w']) is not None:
            possible_influences.append({
                'type': 'special_character',
                'description': 'Run包含换行符，可能影响格式化'
            })

        # 9. 整合结果
        result = {
            'direct_style': direct_style,
            'effective_style': effective_style,
            'possible_influences': possible_influences,
            'paragraph_style': paragraph_effective_style.get('run_properties', {}),
            'doc_default_props': doc_default_props  # 添加文档默认属性信息以便调试
        }

        if style_id:
            result['style_id'] = style_id
            result['template_style'] = template_style

        if default_character_style:
            result['default_character_style'] = default_character_style

        return result

# # # 示例用法
# if __name__ == "__main__":
#     # 检查命令行参数
# import sys
# if len(sys.argv) > 1:
#     docx_path = sys.argv[1]
# else:
#     docx_path = "1.docx"  # 默认文件
#
# if not os.path.exists(docx_path):
#     print(f"错误: 文件 '{docx_path}' 不存在")
#     sys.exit(1)
#
# # 创建分析器
# analyzer = StyleAnalyzer(docx_path)
#
#
# # 打印样式1的完整信息
# print("=== 样式1的详细信息 ===")
# # 打印所有样式信息
#
#
# # 打印样式继承层级
# print(   analyzer.get_paragraph_complete_style_info(analyzer.elements[75].get('element')))
# analyzer.get_paragraph_complete_style_info(analyzer.elements[356].get('element'))
#
