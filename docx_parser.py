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
            'headers': {},  # word/header[1-9].xml
            'footers': {},  # word/footer[1-9].xml
            'media': {},  # word/media/下的文件
            'embeddings': {},  # word/embeddings/下的文件
            'other': {}  # 其他未分类的文件
        }
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
                elif item.filename.startswith('word/header'):
                    header_num = item.filename.split('header')[1]
                    self.parts['headers'][f'header{header_num}'] = self._parse_xml(content)
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
                elif item.filename == '[Content_Types].xml':
                    name='[Content_Types].xml'
                    self.parts['other'][name] = self._parse_xml(content)
                else:
                    # 其他文件
                    name = item.filename
                    self.parts['other'][name] = content

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

    def add_media(self, name, content):
        """添加媒体文件"""
        self.parts['media'][name] = content

    def save(self, output_path):
        """将 self.parts 中的所有内容按照原始结构保存为新的 DOCX 文件"""
        with zipfile.ZipFile(output_path, 'w', compression=zipfile.ZIP_DEFLATED) as zip_out:
            # 1. 保存主文档文件
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
                'fonts': 'word/fontTable.xml'
            }

            for part_name, file_path in predefined_files.items():
                if self.parts[part_name] is not None:
                    self._write_xml_to_zip(zip_out, file_path, self.parts[part_name])

            # 5. 保存页眉
            for header_name, header_tree in self.parts['headers'].items():
                self._write_xml_to_zip(zip_out, f'word/{header_name}', header_tree)

            # 6. 保存页脚
            for footer_name, footer_tree in self.parts['footers'].items():
                self._write_xml_to_zip(zip_out, f'word/footers/{footer_name}', footer_tree)

            # 7. 保存媒体文件
            for media_path, media_content in self.parts['media'].items():

                zip_out.writestr(f'word/media/{media_path}', media_content)

            # 8. 保存嵌入对象
            for embed_path, embed_content in self.parts['embeddings'].items():
                zip_out.writestr(embed_path, embed_content)

            # 9. 保存其他文件
            for other_path, other_content in self.parts['other'].items():
                if isinstance(other_content, ET.ElementTree):
                    self._write_xml_to_zip(zip_out, other_path, other_content)
                else:
                    zip_out.writestr(other_path, other_content)

    def _write_xml_to_zip(self, zip_out, file_path, xml_tree):
        """将ElementTree对象写入ZIP文件"""
        with BytesIO() as f:
            # 保留XML声明和正确的命名空间
            xml_tree.write(f, encoding='UTF-8', xml_declaration=True)
            xml_bytes = f.getvalue()

            # 确保有XML声明
            if not xml_bytes.startswith(b'<?xml'):
                xml_bytes = b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' + xml_bytes

            zip_out.writestr(file_path, xml_bytes)

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

    # 可以添加更多便捷访问方法...
# main_docx = DocxFile('智算工程学院毕业设计（论文）模板2025届(1)-王俊豪-6021203526(1).docx')
# main_docx.save("1.docx")
# main_docx._extract_and_parse(output_dir='extracted_docx')