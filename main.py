import xml.etree.ElementTree as ET
def get_structured_body_elements(self):
    """
    提取文档中的所有顶层元素(w:p及其同级标签)并返回结构化信息

    Returns:
        包含每个元素信息的列表，每个元素包含：
        - type: 元素类型 (paragraph, table, section等)
        - tag: 原始XML标签名
        - index: 在文档中的序号位置
        - id: 标识符 (如段落的paraId)
        - preview: 内容预览
        - element: 原始XML元素对象
    """
    body = self.root.find(f".//{{{self.NAMESPACES['w']}}}body")
    if body is None:
        return []

    elements = []
    for index, element in enumerate(body):
        # 获取不带命名空间的标签名
        tag_with_ns = element.tag
        tag_name = tag_with_ns.split('}')[-1] if '}' in tag_with_ns else tag_with_ns

        # 准备元素信息
        elem_info = {
            'tag': tag_with_ns,
            'short_tag': tag_name,
            'index': index,
            'element': element
        }

        # 根据标签类型处理
        if tag_name == 'p':
            elem_info['type'] = 'paragraph'
            # 获取段落ID
            elem_info['id'] = element.get(f"{{{self.NAMESPACES['w14']}}}paraId", '')
            # 获取段落内容预览
            text = self.get_paragraph_text(element)
            elem_info['preview'] = text[:50] + '...' if len(text) > 50 else text

        elif tag_name == 'tbl':
            elem_info['type'] = 'table'
            rows = element.findall(f".//{{{self.NAMESPACES['w']}}}tr")
            cols = len(rows[0].findall(f".//{{{self.NAMESPACES['w']}}}tc")) if rows else 0

            # 获取表格第一个单元格的内容作为预览
            first_cell = element.find(f".//{{{self.NAMESPACES['w']}}}tc")
            preview = ""
            if first_cell is not None:
                cell_paragraphs = first_cell.findall(f".//{{{self.NAMESPACES['w']}}}p")
                if cell_paragraphs:
                    preview = self.get_paragraph_text(cell_paragraphs[0])

            elem_info['preview'] = f"表格({len(rows)}行×{cols}列): {preview[:30]}..."
            elem_info['rows'] = len(rows)
            elem_info['cols'] = cols

        elif tag_name == 'sectPr':
            elem_info['type'] = 'section'
            elem_info['preview'] = '文档节属性'

        else:
            elem_info['type'] = 'other'
            elem_info['preview'] = f'其他元素: {tag_name}'

        elements.append(elem_info)

    return elements
def _parse_xml(self, content):
        """将XML内容解析为ElementTree"""
        try:
            return ET.parse(BytesIO(content))
        except ET.ParseError as e:
            print(f"XML解析错误: {e}")
            return None