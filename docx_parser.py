import zipfile
from io import BytesIO
import xml.etree.ElementTree as ET
import os



class DocxFile:
    """表示一个DOCX文件，结构化存储各部分内容"""

    def __init__(self, path):
        self.path = path
        # 结构化存储各部分
        self.parts = {
            'document': None,  # word/document.xml
            'styles': None,  # word/styles.xml
            'relationships': None,  # word/_rels/document.xml.rels
            'numbering': None,  # word/numbering.xml
            'footnotes': None,  # word/footnotes.xml
            'endnotes': None,  # word/endnotes.xml
            'settings': None,  # word/settings.xml
            'fonts': None,  # word/fontTable.xml
            'comments': None,  # word/comments.xml
            'headers': {},  # word/header[11.py-9].xml
            'footers': {},  # word/footer[11.py-9].xml
            'media': {},  # word/media/下的文件
            'embeddings': {},  # word/embeddings/下的文件
            'other': {},  # 其他未分类的文件
            'people': {},  # 批注人的文件
            'commentsExtended': {},  # commentsExtended.xml
        }
        # 为了兼容docx_namespace.py，添加docx_parts作为parts的别名
        self.docx_parts = self.parts
        self._extract_and_parse()

    def _extract_and_parse(self, output_dir=None):
        """
        解压并结构化解析DOCX文件
        
        Args:
            output_dir: 可选，指定保存解压文件的目录，如果为None则不保存到磁盘
        """
        # 如果指定了输出目录，确保它存在
        if output_dir:
            os.makedirs(output_dir, exist_ok=True)
        
        with zipfile.ZipFile(self.path) as zip_file:
            # 解压并分类所有文件
            for item in zip_file.infolist():
                content = zip_file.read(item.filename)
                
                # 如果指定了输出目录，保存文件到磁盘
                if output_dir:
                    # 构建完整的输出路径，保留原始目录结构
                    output_path = os.path.join(output_dir, item.filename)
                    
                    # 检查是否是目录（以斜杠结尾）
                    if item.filename.endswith('/'):
                        # 如果是目录，只创建目录而不尝试写入文件
                        os.makedirs(output_path, exist_ok=True)
                    else:
                        # 确保目标目录存在
                        os.makedirs(os.path.dirname(output_path), exist_ok=True)
                        # 写入文件
                        with open(output_path, 'wb') as f:
                            f.write(content)
                
                # 分类存储
                if item.filename == 'word/document.xml':
                    self.parts['document'] = self._parse_xml(content)
                elif item.filename == 'word/styles.xml':
                    self.parts['styles'] = self._parse_xml(content)
                elif item.filename == 'word/_rels/document.xml.rels':
                    self.parts['relationships'] = self._parse_xml(content)
                elif item.filename == 'word/numbering.xml':
                    self.parts['numbering'] = self._parse_xml(content)
                elif item.filename == 'word/comments.xml':
                    self.parts['comments'] = self._parse_xml(content)
                    # 为了兼容性，也存储原始XML字符串
                    self.docx_parts['word/comments.xml'] = content
                elif item.filename.startswith('word/header'):
                    header_num = item.filename.split('header')[1]
                    self.parts['headers'][f'header{header_num}'] = self._parse_xml(content)
                elif item.filename== 'word/commentsExtended.xml':

                    self.parts['commentsExtended'] = self._parse_xml(content)
                elif item.filename=='word/people.xml':


                    self.parts['people'] = self._parse_xml(content)
                elif item.filename.startswith('word/footer'):
                    footer_num = item.filename.split('footer')[1]
                    self.parts['footers'][f'footer{footer_num}'] = self._parse_xml(content)
                elif item.filename.startswith('word/media/'):
                    media_name = item.filename.split('media/')[1]
                    self.parts['media'][media_name] = content  # 二进制内容，不解析
                elif item.filename.startswith('word/embeddings/'):
                    embed_name = item.filename.split('embeddings/')[1]
                    self.parts['embeddings'][embed_name] = content  # 二进制内容
                elif item.filename.startswith('word/') and item.filename.endswith('.xml'):
                    # 其他word目录下的xml文件
                    name = item.filename
                    self.parts['other'][name] = self._parse_xml(content)
                    # 为了兼容性，保存原始内容
                    self.docx_parts[name] = content
                elif item.filename == '[Content_Types].xml':
                    name='[Content_Types].xml'
                    self.parts['other'][name] = self._parse_xml(content)
                    # 为了兼容性，保存原始内容
                    self.docx_parts[name] = content
                else:
                    # 其他文件
                    name = item.filename
                    self.parts['other'][name] = content
                    # 为了兼容性，保存原始内容
                    self.docx_parts[name] = content

    def _configure_parser(self):
        """配置XML解析器以更好地处理复杂XML"""
        # 创建自定义解析器
        parser = ET.XMLParser(encoding='utf-8')
        # 如果可能，增加递归限度
        try:
            import sys
            sys.setrecursionlimit(10000)  # 增加Python递归限制
        except Exception as e:
            print(f"无法修改递归限制: {e}")
        return parser
            
    def _parse_xml(self, content):
        """使用优化的解析器解析XML内容"""
        try:
            parser = self._configure_parser()
            return ET.parse(BytesIO(content), parser=parser)
        except ET.ParseError as e:
            print(f"XML解析错误: {e}")
            return None

    def get_header(self, num=1):
        """获取指定编号的页眉"""
        return self.parts['headers'].get(f'header{num}')

    def get_footer(self, num=1):
        """获取指定编号的页脚"""
        return self.parts['footers'].get(f'footer{num}')

    def get_media(self, name):
        """获取指定的媒体文件"""
        return self.parts['media'].get(name)
        
    def get_comments(self):
        """获取文档的批注"""
        return self.parts['comments']

    def add_media(self, name, content):
        """添加媒体文件"""
        self.parts['media'][name] = content

    def save(self, output_path):
        """将 self.parts 中的所有内容按照原始结构保存为新的 DOCX 文件，跳过空部分"""
        with zipfile.ZipFile(output_path, 'w', compression=zipfile.ZIP_DEFLATED) as zip_out:
            # 11.py. 保存主文档文件
            if self.parts['document'] is not None:
                self._write_xml_to_zip(zip_out, 'word/document.xml', self.parts['document'])

            # 2. 保存样式文件
            if self.parts['styles'] is not None:
                self._write_xml_to_zip(zip_out, 'word/styles.xml', self.parts['styles'])

            # 3. 保存关系文件
            if self.parts['relationships'] is not None:
                self._write_xml_to_zip(zip_out, 'word/_rels/document.xml.rels', self.parts['relationships'])

            # 4. 保存其他预定义的XML文件
            predefined_files = {
                'numbering': 'word/numbering.xml',
                'footnotes': 'word/footnotes.xml',
                'endnotes': 'word/endnotes.xml',
                'settings': 'word/settings.xml',
                'fonts': 'word/fontTable.xml',
                'comments': 'word/comments.xml',
                'commentsExtended': 'word/commentsExtended.xml',
                'people': 'word/people.xml'
            }

            for part_name, file_path in predefined_files.items():
                if self.parts[part_name] is not None:
                    self._write_xml_to_zip(zip_out, file_path, self.parts[part_name])

            # 5. 保存页眉 - 检查是否为空字典
            if self.parts['headers'] and isinstance(self.parts['headers'], dict):
                for header_name, header_tree in self.parts['headers'].items():
                    self._write_xml_to_zip(zip_out, f'word/{header_name}', header_tree)

            # 6. 保存页脚 - 检查是否为空字典
            if self.parts['footers'] and isinstance(self.parts['footers'], dict):
                for footer_name, footer_tree in self.parts['footers'].items():
                    self._write_xml_to_zip(zip_out, f'word/{footer_name}', footer_tree)

            # 7. 保存媒体文件 - 检查是否为空字典
            if self.parts['media'] and isinstance(self.parts['media'], dict):
                for media_path, media_content in self.parts['media'].items():
                    zip_out.writestr(f'word/media/{media_path}', media_content)

            # 8. 保存嵌入对象 - 检查是否为空字典
            if self.parts['embeddings'] and isinstance(self.parts['embeddings'], dict):
                for embed_path, embed_content in self.parts['embeddings'].items():
                    zip_out.writestr(embed_path, embed_content)

            # 9. 保存其他文件 - 检查是否为空字典
            if self.parts['other'] and isinstance(self.parts['other'], dict):
                for other_path, other_content in self.parts['other'].items():
                    if isinstance(other_content, ET.ElementTree):
                        self._write_xml_to_zip(zip_out, other_path, other_content)
                    else:
                        zip_out.writestr(other_path, other_content)

            # 10. 添加空检查 - 如果people是空字典则跳过
            if self.parts['people'] and isinstance(self.parts['people'], dict) and len(self.parts['people']) > 0:
                self._write_xml_to_zip(zip_out, 'word/people.xml', self.parts['people'])

            # 11.py. 添加空检查 - 如果commentsExtended是空字典则跳过
            if self.parts['commentsExtended'] and isinstance(self.parts['commentsExtended'], dict) and len(
                    self.parts['commentsExtended']) > 0:
                self._write_xml_to_zip(zip_out, 'word/commentsExtended.xml', self.parts['commentsExtended'])

        print(f"文档已保存到: {output_path}")
        return output_path

    def _write_xml_to_zip(self, zip_out, file_path, xml_tree):
        """将ElementTree对象写入ZIP文件，添加类型检查和错误处理"""
        
        # 在函数开头导入所需模块，确保它们在整个函数中可用
        from io import BytesIO
        import xml.etree.ElementTree as ET
        
        # 添加类型检查
        if not hasattr(xml_tree, 'write') or not callable(getattr(xml_tree, 'write')):
            print(f"错误: 文件 '{file_path}' 的内容不是有效的XML树对象，而是 {type(xml_tree)}")
            
            # 如果是字典类型，打印字典内容帮助调试
            if isinstance(xml_tree, dict):
                print(f"字典内容: {xml_tree}")
                
                # 可以尝试修复字典问题（提供一种紧急解决方案）
                try:
                    print(f"尝试将字典转换为XML树...")
                    # 创建一个简单的替代XML
                    root = ET.Element("root")
                    for key, value in xml_tree.items():
                        # 创建子元素，使用字符串表示值
                        sub = ET.SubElement(root, str(key))
                        if isinstance(value, dict):
                            for k, v in value.items():
                                sub_sub = ET.SubElement(sub, str(k))
                                sub_sub.text = str(v)
                        else:
                            sub.text = str(value)
                    
                    # 创建新的ElementTree
                    xml_tree = ET.ElementTree(root)
                    print("转换成功，将继续保存")
                except Exception as e:
                    print(f"转换失败: {e}")
                    # 跳过此文件，避免整个保存操作失败
                    print(f"跳过文件 '{file_path}'")
                    return
            else:
                # 跳过此文件，避免整个保存操作失败
                print(f"跳过文件 '{file_path}'")
                return
            
        try:
            # 使用BytesIO写入XML内容
            buffer = BytesIO()
            xml_tree.write(buffer, encoding='UTF-8', xml_declaration=True)
            xml_bytes = buffer.getvalue()

            # 确保有XML声明
            if not xml_bytes.startswith(b'<?xml'):
                xml_bytes = b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' + xml_bytes

            # 写入ZIP文件
            zip_out.writestr(file_path, xml_bytes)
            print(f"成功保存文件: {file_path}")
        except Exception as e:
            print(f"保存文件 '{file_path}' 时出错: {e}")
            # 记录更多细节以帮助诊断
            import traceback
            traceback.print_exc()

    def print_document_xml(self):
        """打印document.xml的完整内容"""
        if 'document' in self.parts and self.parts['document'] is not None:
            print("=== document.xml 完整内容 ===")

            # 获取根元素
            root = self.parts['document'].getroot()

            # 使用minidom格式化输出
            import xml.dom.minidom as minidom
            import xml.etree.ElementTree as ET

            # 将整个ElementTree转换为字符串
            rough_string = ET.tostring(root, 'utf-8')

            # 使用minidom解析并格式化
            reparsed = minidom.parseString(rough_string)
            pretty_str = reparsed.toprettyxml(indent="  ")

            print(pretty_str[:10000])
            print("=== XML文档结束 ===")
        else:
            print("文档XML不可用")

    def print_comments_xml(self):
        """打印comments.xml的完整内容"""
        if 'comments' in self.parts and self.parts['comments'] is not None:
            print("=== comments.xml 完整内容 ===")

            # 获取根元素
            root = self.parts['comments'].getroot()

            # 使用minidom格式化输出
            import xml.dom.minidom as minidom
            import xml.etree.ElementTree as ET

            # 将整个ElementTree转换为字符串
            rough_string = ET.tostring(root, 'utf-8')

            # 使用minidom解析并格式化
            reparsed = minidom.parseString(rough_string)
            pretty_str = reparsed.toprettyxml(indent="  ")

            print(pretty_str)
            print("=== XML文档结束 ===")
        else:
            print("批注XML不可用")

    def create_comments_file(self, author="Anonymous"):
        """
        创建新的comments.xml文件并将其添加到parts中

        参数:
            author (str): 默认的批注作者名称

        返回:
            ET.Element: 创建的comments根元素
        """
        # 定义XML命名空间
        ns_w = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
        ns_w14 = "{http://schemas.microsoft.com/office/word/2010/wordml}"
        ns_w15 = "{http://schemas.microsoft.com/office/word/2012/wordml}"

        # 创建comments根元素
        comments_root = ET.Element(f"{ns_w}comments")

        # 添加标准命名空间
        namespaces = {
            "wpc": "http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas",
            "mc": "http://schemas.openxmlformats.org/markup-compatibility/2006",
            "o": "urn:schemas-microsoft-com:office:office",
            "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
            "m": "http://schemas.openxmlformats.org/officeDocument/2006/math",
            "v": "urn:schemas-microsoft-com:vml",
            "wp14": "http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing",
            "wp": "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
            "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
            "w14": "http://schemas.microsoft.com/office/word/2010/wordml",
            "w15": "http://schemas.microsoft.com/office/word/2012/wordml",
            "w10": "urn:schemas-microsoft-com:office:word",
            "wpg": "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup",
            "wpi": "http://schemas.microsoft.com/office/word/2010/wordprocessingInk",
            "wne": "http://schemas.microsoft.com/office/word/2006/wordml",
            "wps": "http://schemas.microsoft.com/office/word/2010/wordprocessingShape"
        }

        # 设置命名空间属性
        for prefix, uri in namespaces.items():
            comments_root.attrib[f"xmlns:{prefix}"] = uri

        # 设置mc:Ignorable属性
        comments_root.attrib["{http://schemas.openxmlformats.org/markup-compatibility/2006}Ignorable"] = "w14 w15 wp14"

        # 创建ElementTree对象
        comments_tree = ET.ElementTree(comments_root)

        # 将comments文件添加到parts中
        self.parts['comments'] = comments_tree

        # 创建commentsExtended.xml文件
        self.create_comments_extended_file()

        # 创建people.xml文件
        self.create_people_file(author)

        # 确保document.xml.rels中包含对comments.xml的引用
        self.add_comments_relationship()

        return comments_root

    def create_comments_extended_file(self):
        """创建新的commentsExtended.xml文件并将其添加到parts中"""
        # 定义XML命名空间
        ns_w15 = "{http://schemas.microsoft.com/office/word/2012/wordml}"

        # 创建commentsEx根元素
        comments_ex_root = ET.Element(f"{ns_w15}commentsEx")

        # 添加标准命名空间
        namespaces = {
            "wpc": "http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas",
            "mc": "http://schemas.openxmlformats.org/markup-compatibility/2006",
            "o": "urn:schemas-microsoft-com:office:office",
            "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
            "m": "http://schemas.openxmlformats.org/officeDocument/2006/math",
            "v": "urn:schemas-microsoft-com:vml",
            "wp14": "http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing",
            "wp": "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
            "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
            "w14": "http://schemas.microsoft.com/office/word/2010/wordml",
            "w15": "http://schemas.microsoft.com/office/word/2012/wordml",
            "w10": "urn:schemas-microsoft-com:office:word",
            "wpg": "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup",
            "wpi": "http://schemas.microsoft.com/office/word/2010/wordprocessingInk",
            "wne": "http://schemas.microsoft.com/office/word/2006/wordml",
            "wps": "http://schemas.microsoft.com/office/word/2010/wordprocessingShape"
        }

        # 设置命名空间属性
        for prefix, uri in namespaces.items():
            comments_ex_root.attrib[f"xmlns:{prefix}"] = uri

        # 设置mc:Ignorable属性
        comments_ex_root.attrib[
            "{http://schemas.openxmlformats.org/markup-compatibility/2006}Ignorable"] = "w14 w15 wp14"

        # 创建ElementTree对象
        comments_ex_tree = ET.ElementTree(comments_ex_root)

        # 将commentsExtended文件添加到parts中
        self.parts['commentsExtended'] = comments_ex_tree

        # 确保document.xml.rels中包含对commentsExtended.xml的引用
        self.add_comments_extended_relationship()

        return comments_ex_root

    def create_people_file(self, author="Anonymous"):
        """
        创建新的people.xml文件并将其添加到parts中

        参数:
            author (str): 默认的批注作者名称
        """
        # 定义XML命名空间
        ns_w15 = "{http://schemas.microsoft.com/office/word/2012/wordml}"

        # 创建people根元素
        people_root = ET.Element(f"{ns_w15}people")

        # 添加标准命名空间
        namespaces = {
            "wpc": "http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas",
            "mc": "http://schemas.openxmlformats.org/markup-compatibility/2006",
            "o": "urn:schemas-microsoft-com:office:office",
            "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
            "m": "http://schemas.openxmlformats.org/officeDocument/2006/math",
            "v": "urn:schemas-microsoft-com:vml",
            "wp14": "http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing",
            "wp": "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
            "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
            "w14": "http://schemas.microsoft.com/office/word/2010/wordml",
            "w15": "http://schemas.microsoft.com/office/word/2012/wordml",
            "w10": "urn:schemas-microsoft-com:office:word",
            "wpg": "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup",
            "wpi": "http://schemas.microsoft.com/office/word/2010/wordprocessingInk",
            "wne": "http://schemas.microsoft.com/office/word/2006/wordml",
            "wps": "http://schemas.microsoft.com/office/word/2010/wordprocessingShape"
        }

        # 设置命名空间属性
        for prefix, uri in namespaces.items():
            people_root.attrib[f"xmlns:{prefix}"] = uri

        # 设置mc:Ignorable属性
        people_root.attrib["{http://schemas.openxmlformats.org/markup-compatibility/2006}Ignorable"] = "w14 w15 wp14"

        # 添加默认作者
        if author:
            person = ET.SubElement(people_root, f"{ns_w15}person")
            person.set(f"{ns_w15}author", author)

            presence_info = ET.SubElement(person, f"{ns_w15}presenceInfo")
            presence_info.set(f"{ns_w15}providerId", "None")
            presence_info.set(f"{ns_w15}userId", author)

        # 创建ElementTree对象
        people_tree = ET.ElementTree(people_root)

        # 将people文件添加到parts中
        self.parts['people'] = people_tree

        # 确保document.xml.rels中包含对people.xml的引用
        self.add_people_relationship()

        return people_root

    def add_comments_relationship(self):
        """确保document.xml.rels中包含对comments.xml的引用"""
        self._add_relationship(
            "comments.xml",
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments"
        )

    def add_comments_extended_relationship(self):
        """确保document.xml.rels中包含对commentsExtended.xml的引用"""
        self._add_relationship(
            "commentsExtended.xml",
            "http://schemas.microsoft.com/office/2011/relationships/commentsExtended"
        )

    def add_people_relationship(self):
        """确保document.xml.rels中包含对people.xml的引用"""
        self._add_relationship(
            "people.xml",
            "http://schemas.microsoft.com/office/2011/relationships/people"
        )

    def _add_relationship(self, target, relationship_type):
        """
        添加文档关系

        参数:
            target (str): 目标文件路径
            relationship_type (str): 关系类型URI

        返回:
            str: 新关系的ID
        """
        # 定义命名空间
        ns_rel = "{http://schemas.openxmlformats.org/package/2006/relationships}"

        # 确保relationships存在
        if 'relationships' not in self.parts or self.parts['relationships'] is None:
            # 创建新的relationships文件
            rels_root = ET.Element(f"{ns_rel}Relationships")
            rels_root.attrib["xmlns"] = "http://schemas.openxmlformats.org/package/2006/relationships"
            self.parts['relationships'] = ET.ElementTree(rels_root)

        # 获取relationships根元素
        rels_root = self.parts['relationships'].getroot()

        # 检查是否已存在该关系
        for rel in rels_root.findall(f".//{ns_rel}Relationship"):
            if rel.get("Target") == target:
                # 关系已存在，返回现有ID
                return rel.get("Id")

        # 关系不存在，创建新关系
        # 找到当前最大的rId并加1
        max_rid = 0
        for rel in rels_root.findall(f".//{ns_rel}Relationship"):
            rid = rel.get("Id", "")
            if rid.startswith("rId"):
                try:
                    rid_num = int(rid[3:])
                    max_rid = max(max_rid, rid_num)
                except ValueError:
                    pass

        # 创建新的关系ID
        new_rel_id = f"rId{max_rid + 1}"

        # 创建新的关系元素
        new_rel = ET.SubElement(rels_root, f"{ns_rel}Relationship")
        new_rel.set("Id", new_rel_id)
        new_rel.set("Type", relationship_type)
        new_rel.set("Target", target)

        return new_rel_id
    # 可以添加更多便捷访问方法...
# docx = DocxFile("智算工程学院毕业设计（论文）模板2025届(1)-王俊豪-6021203526(1).docx")
#
# # 创建comments相关文件
#
#
# # 现在你可以向comments_root添加具体的批注内容
# # ...
#
# # 保存文档
# docx.save("with_comments.docx")