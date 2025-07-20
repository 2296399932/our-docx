import xml.etree.ElementTree as ET
import re
from docx_parser import DocxFile
import traceback
import xml.dom.minidom as minidom
import pandas as pd
import os
import uuid
import time
from PIL import Image


class DocxElementParser(DocxFile):
    """用于解析Word文档XML的类，提供对文档结构和内容的访问，继承自DocxFile"""

    # 定义常见的XML命名空间
    NAMESPACES = {
        'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
        'w14': 'http://schemas.microsoft.com/office/word/2010/wordml',
        'wp': 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing',
        'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
        'pic': 'http://schemas.openxmlformats.org/drawingml/2006/picture',
        'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
        'm': 'http://schemas.openxmlformats.org/officeDocument/2006/math',
        # 添加以下新的命名空间
        'wpc': 'http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas',
        'mc': 'http://schemas.openxmlformats.org/markup-compatibility/2006',
        'o': 'urn:schemas-microsoft-com:office:office',
        'v': 'urn:schemas-microsoft-com:vml',
        'wp14': 'http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing',
        'w10': 'urn:schemas-microsoft-com:office:word',
        'w15': 'http://schemas.microsoft.com/office/word/2012/wordml',
        'wpg': 'http://schemas.microsoft.com/office/word/2010/wordprocessingGroup',
        'wpi': 'http://schemas.microsoft.com/office/word/2010/wordprocessingInk',
        'wne': 'http://schemas.microsoft.com/office/word/2006/wordml',
        'wps': 'http://schemas.microsoft.com/office/word/2010/wordprocessingShape',
        'wpsCustomData': 'http://www.wps.cn/officeDocument/2013/wpsCustomData',
        'xml': 'http://www.w3.org/XML/1998/namespace'  # 添加xml命名空间
    }

    def __init__(self, path):
        """初始化解析器

        Args:
            path: Word文档的文件路径
        """
        # 调用父类构造函数
        super().__init__(path)

        # 获取文档的XML树
        self.tree = self.parts["document"]
        self.root = self.tree.getroot() if self.tree else None

        # 初始化元素列表
        self.elements = []
        self.paragraphs = []
        self.tables = []
        self.sections = []

        # 注册所有命名空间用于XPath查询
        for prefix, uri in self.NAMESPACES.items():
            ET.register_namespace(prefix, uri)

        # 解析文档结构
        self.get_structured_body_elements()

    def get_element(self):
        """通过ID获取特定元素

        Args:
            element_id: 元素的ID，如段落的paraId

        Returns:
            匹配的元素，如果未找到则返回None
        """
        return self.elements

    def find_elements_by_tag(self, tag_name):
        """查找所有指定标签的元素

        Args:
            tag_name: 标签名称，如'w:p'或'w:tbl'

        Returns:
            符合条件的元素列表
        """
        if ':' in tag_name:
            prefix, name = tag_name.split(':')
            namespace = self.NAMESPACES.get(prefix, '')
            xpath = f".//{{{namespace}}}{name}"
        else:
            xpath = f".//{tag_name}"

        return self.root.findall(xpath)

    def get_body_direct_children(self):
        """获取body元素的直接子元素(段落、表格等)"""
        body = self.root.find(f".//{{{self.NAMESPACES['w']}}}body")
        if body is not None:
            return list(body)
        return []

    def get_all_paragraphs(self):
        """获取所有段落元素"""
        return self.paragraphs
    def get_all_paragraphs_text(self):
        """获取所有段落元素的文本内容"""
        return [self.get_paragraph_text(p['element']) for p in self.paragraphs]

    def get_paragraphs_length(self):
        return len(self.paragraphs)
    def get_table_length(self):
        return len(self.tables)
    def get_all_tables(self):
        """获取所有表格元素"""
        return self.tables

    def get_paragraph_by_id(self, para_id):
        """通过paraId获取特定段落"""
        for p in self.get_all_paragraphs():
            if p.get(f"{{{self.NAMESPACES['w14']}}}paraId") == para_id:
                return p
        return None

    def get_paragraph_text(self, paragraph):
        """提取段落中的所有文本内容"""
        text_elements = paragraph.findall(f".//{{{self.NAMESPACES['w']}}}t")
        return ''.join(elem.text or '' for elem in text_elements)

    def get_all_text(self):
        """提取文档中的所有文本内容"""
        text_elements = self.root.findall(f".//{{{self.NAMESPACES['w']}}}t")
        return ''.join(elem.text or '' for elem in text_elements)

    def get_element_attributes(self, element):
        """获取元素的所有属性"""
        return element.attrib

    def get_structured_body_elements(self):
        """
        提取文档中的所有顶层元素(w:p及其同级标签)并返回结构化信息，
        并将不同类型的元素分别存储到相应的列表中

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

        # 清空元素列表，避免重复调用时出现问题
        self.elements = []
        self.paragraphs = []
        self.tables = []
        self.sections = []
        self.images=[]
        for index, element in enumerate(body):
            # 获取不带命名空间的标签名
            tag_with_ns = element.tag
            tag_name = tag_with_ns.split('}')[-1] if '}' in tag_with_ns else tag_with_ns

            # 准备元素信息
            elem_info = {


                'index': index,
                'element': element
            }

            # 根据标签类型处理
            if tag_name == 'p':
                elem_info['type'] = 'paragraph'
                self.paragraphs.append(elem_info)
                # 获取段落ID
                elem_info['id'] = element.get(f"{{{self.NAMESPACES['w14']}}}paraId", '')
                self.paragraphs.append(elem_info)
                # 检查该段落是否包含图片
                has_image = False
                # 查找段落中的w:drawing或w:pict元素（图片容器）
                drawings = element.findall(f".//{{{self.NAMESPACES['w']}}}drawing") or []
                picts = element.findall(f".//{{{self.NAMESPACES['w']}}}pict") or []

                if drawings or picts:

                    elem_info['type'] = 'image'
                    # 可以添加额外的图片信息提取
                    elem_info['image_info'] = self.get_image_from_pra(element)

                    self.images.append(elem_info)
                    elem_info['type'] = 'paragraph'




            elif tag_name == 'tbl':
                elem_info['type'] = 'table'
                self.tables.append(elem_info)
            elif tag_name == 'sectPr':
                elem_info['type'] = 'section'
                self.sections.append(elem_info)
            elif tag_name == 'bookmarkStart':
                elem_info['type'] = 'bookmarkStart'
            elif tag_name == 'bookmarkEnd':
                elem_info['type'] = 'bookmarkEnd'

            else:
                elem_info['type'] = 'other'

            # 所有元素都添加到主元素列表
            self.elements.append(elem_info)



    def get_element_text(self, num):
        """从元素信息字典中提取文本内容

        Args:
            num: 从get_structured_body_elements返回的元素信息字典索引

        Returns:
            str: 如果元素是段落类型，返回其中所有文本内容；
                如果是表格类型，返回格式化的表格内容；
                否则返回空字符串
        """
        if self.elements==[]:
            return ''
        if self.elements[num].get('type') =='paragraph':

            str = self.elements[num].get('element')
            return self.get_paragraph_text(str)
        elif self.elements[num].get('type') =='table':
            table_element = self.elements[num].get('element')
            return self.extract_table_content(table_element)

    def extract_table_content(self, table_element):
        """提取表格中的所有文本内容

        Args:
            table_element: 表格XML元素

        Returns:
            str: 格式化的表格内容
        """
        result = []
        # 找到所有表格行
        rows = table_element.findall(f".//{{{self.NAMESPACES['w']}}}tr")

        for row in rows:
            row_text = []
            # 找到行中的所有单元格
            cells = row.findall(f".//{{{self.NAMESPACES['w']}}}tc")

            for cell in cells:
                # 找到单元格中的所有段落
                paragraphs = cell.findall(f".//{{{self.NAMESPACES['w']}}}p")
                cell_text = []

                for p in paragraphs:
                    p_text = self.get_paragraph_text(p)
                    if p_text.strip():
                        cell_text.append(p_text)

                row_text.append("".join(cell_text))

            result.append(" | ".join(row_text))

        return "\n".join(result)

    def print_full_xml(self):
        """打印整个XML文档的内容"""
        if self.tree is None:
            print("没有可用的XML文档")
            return



        try:
            # 将整个ElementTree转换为字符串
            rough_string = ET.tostring(self.root, 'utf-8')

            # 使用minidom解析并格式化
            reparsed = minidom.parseString(rough_string)
            pretty_str = reparsed.toprettyxml(indent="  ")

            print("=== XML文档的完整内容 ===")
            print(pretty_str[:10000])
            print("=== XML文档结束 ===")

        except Exception as e:
            print(f"打印XML时发生错误: {e}")

            # 尝试备用方法
            print("尝试直接打印XML元素:")
            print(ET.tostring(self.root, encoding='unicode'))

    def export_table_to_file(self, table_idx, file_path, format='xlsx'):
        """将指定索引的表格导出为xlsx或csv文件

        Args:
            table_idx: self.tables中的表格索引
            file_path: 要保存的文件路径
            format: 文件格式，'xlsx'或'csv'

        Returns:
            bool: 是否成功导出
        """


        # 检查索引是否有效
        if table_idx < 0 or table_idx >= len(self.tables):
            print(f"错误：表格索引{table_idx}超出范围(0-{len(self.tables)-1})")
            return False

        # 获取表格元素
        table_element = self.tables[table_idx]['element']

        # 提取表格数据为二维列表
        table_data = []
        rows = table_element.findall(f".//{{{self.NAMESPACES['w']}}}tr")

        for row in rows:
            row_data = []
            cells = row.findall(f".//{{{self.NAMESPACES['w']}}}tc")

            for cell in cells:
                cell_text = ''
                paragraphs = cell.findall(f".//{{{self.NAMESPACES['w']}}}p")

                for p in paragraphs:
                    p_text = self.get_paragraph_text(p)
                    if cell_text and p_text:
                        cell_text += '\n' + p_text
                    else:
                        cell_text += p_text

                row_data.append(cell_text)

            table_data.append(row_data)

        # 创建pandas DataFrame
        df = pd.DataFrame(table_data)

        # 如果第一行看起来像表头，可以使用它作为列名
        if len(table_data) > 1:
            df.columns = df.iloc[0]
            df = df[1:]

        # 根据格式导出文件
        try:
            if format.lower() == 'xlsx':
                df.to_excel(file_path, index=False)
                print(f"表格已成功导出为Excel文件：{file_path}")
            elif format.lower() == 'csv':
                df.to_csv(file_path, index=False)
                print(f"表格已成功导出为CSV文件：{file_path}")
            else:
                print(f"不支持的文件格式：{format}，请使用'xlsx'或'csv'")
                return False

            return True

        except Exception as e:
            print(f"导出表格时发生错误：{e}")
            return False

    def export_all_tables(self, dir_path, format='xlsx'):
        """将文档中的所有表格导出为xlsx或csv文件

        Args:
            dir_path: 要保存表格的目录路径
            format: 文件格式，'xlsx'或'csv'

        Returns:
            int: 成功导出的表格数量
        """


        # 确保目录存在
        if not os.path.exists(dir_path):
            os.makedirs(dir_path)

        count = 0
        for i in range(len(self.tables)):
            file_name = f"table_{i+1}.{format}"
            file_path = os.path.join(dir_path, file_name)

            if self.export_table_to_file(i, file_path, format):
                count += 1

        print(f"已成功导出{count}个表格到{dir_path}目录")
        return count

    def extract_images_simple(self, output_dir):
        """从文档中提取所有图片到指定目录（简化版）

        Args:
            output_dir: 输出图片的目录路径

        Returns:
            int: 成功提取的图片数量
            list: 提取的图片文件路径列表
        """
        # 确保输出目录存在
        if not os.path.exists(output_dir):
            os.makedirs(output_dir)

        extracted_images = []
        count = 0

        # 直接从self.parts['media']字典获取所有图片
        media_files = self.parts['media']

        if not media_files:
            print("文档中没有找到媒体文件")
            return 0, []

        # 遍历所有媒体文件并保存
        for i, (image_name, image_data) in enumerate(media_files.items()):
            # 获取文件扩展名
            _, ext = os.path.splitext(image_name)
            if not ext:
                # 如果没有扩展名，尝试猜测文件类型
                ext = '.jpg'  # 默认扩展名

            # 构建输出文件路径
            output_file = os.path.join(output_dir, f"image_{i+1}{ext}")

            try:
                # 写入图片文件
                with open(output_file, 'wb') as f:
                    f.write(image_data)

                extracted_images.append(output_file)
                count += 1
                print(f"提取图片: {output_file}")
            except Exception as e:
                print(f"提取图片时出错: {e}")

        print(f"成功提取{count}张图片到{output_dir}目录")
        return count, extracted_images

    def count_images_simple(self):
        """统计文档中的图片数量（简化版）

        Returns:
            int: 文档中图片的数量
        """
        media_count = len(self.parts['media'])
        print(f"文档中包含{media_count}个媒体文件")
        return media_count

    def extract_paragraph_style(self, paragraph_element):
        """提取段落中的所有样式信息

        Args:
            paragraph_element: 段落XML元素对象

        Returns:
            dict: 包含段落样式信息的字典
        """
        style_info = {
            'style_id': None,
            'alignment': None,
            'indentation': {},
            'spacing': {},
            'borders': {},
            'shading': None,
            'numbering': {},
            'run_properties': {},
            'other_properties': {},
            'page_break': {
                'has_page_break': False,
                'type': ""
            }
        }

        # 查找段落属性标签
        pPr = paragraph_element.find(f".//{{{self.NAMESPACES['w']}}}pPr")
        if pPr is None:
            return {'has_style': False, 'message': '段落无样式信息'}

        # 1. 提取样式ID
        style = pPr.find(f".//{{{self.NAMESPACES['w']}}}pStyle")
        if style is not None:
            style_info['style_id'] = style.get(f"{{{self.NAMESPACES['w']}}}val")

        # 2. 提取对齐方式
        jc = pPr.find(f".//{{{self.NAMESPACES['w']}}}jc")
        if jc is not None:
            style_info['alignment'] = jc.get(f"{{{self.NAMESPACES['w']}}}val")

        # 3. 提取缩进信息
        ind = pPr.find(f".//{{{self.NAMESPACES['w']}}}ind")
        if ind is not None:
            for key in ['left', 'right', 'firstLine', 'hanging']:
                val = ind.get(f"{{{self.NAMESPACES['w']}}}{key}")
                if val:
                    style_info['indentation'][key] = val

        # 4. 提取段落间距
        spacing = pPr.find(f".//{{{self.NAMESPACES['w']}}}spacing")
        if spacing is not None:
            for key in ['before', 'after', 'line', 'lineRule', 'beforeLines', 'afterLines']:
                val = spacing.get(f"{{{self.NAMESPACES['w']}}}{key}")
                if val:
                    style_info['spacing'][key] = val

        # 5. 提取段落边框
        pBdr = pPr.find(f".//{{{self.NAMESPACES['w']}}}pBdr")
        if pBdr is not None:
            for border_type in ['top', 'bottom', 'left', 'right']:
                border = pBdr.find(f".//{{{self.NAMESPACES['w']}}}{border_type}")
                if border is not None:
                    style_info['borders'][border_type] = {}
                    for attr in ['val', 'sz', 'space', 'color']:
                        val = border.get(f"{{{self.NAMESPACES['w']}}}{attr}")
                        if val:
                            style_info['borders'][border_type][attr] = val

        # 6. 提取背景填充
        shading = pPr.find(f".//{{{self.NAMESPACES['w']}}}shd")
        if shading is not None:
            style_info['shading'] = {
                'val': shading.get(f"{{{self.NAMESPACES['w']}}}val"),
                'color': shading.get(f"{{{self.NAMESPACES['w']}}}color"),
                'fill': shading.get(f"{{{self.NAMESPACES['w']}}}fill")
            }

        # 7. 提取编号信息
        numPr = pPr.find(f".//{{{self.NAMESPACES['w']}}}numPr")
        if numPr is not None:
            ilvl = numPr.find(f".//{{{self.NAMESPACES['w']}}}ilvl")
            if ilvl is not None:
                style_info['numbering']['level'] = ilvl.get(f"{{{self.NAMESPACES['w']}}}val")

            numId = numPr.find(f".//{{{self.NAMESPACES['w']}}}numId")
            if numId is not None:
                style_info['numbering']['id'] = numId.get(f"{{{self.NAMESPACES['w']}}}val")

        # 8. 提取文字样式属性
        rPr = pPr.find(f".//{{{self.NAMESPACES['w']}}}rPr")
        if rPr is not None:
            # 提取字体
            rFonts = rPr.find(f".//{{{self.NAMESPACES['w']}}}rFonts")
            if rFonts is not None:
                style_info['run_properties']['fonts'] = {}
                for font_type in ['ascii', 'hAnsi', 'eastAsia', 'cs']:
                    font = rFonts.get(f"{{{self.NAMESPACES['w']}}}{font_type}")
                    if font:
                        style_info['run_properties']['fonts'][font_type] = font

            # 提取字号
            sz = rPr.find(f".//{{{self.NAMESPACES['w']}}}sz")
            if sz is not None:
                style_info['run_properties']['size'] = sz.get(f"{{{self.NAMESPACES['w']}}}val")

            szCs = rPr.find(f".//{{{self.NAMESPACES['w']}}}szCs")
            if szCs is not None:
                style_info['run_properties']['sizeCs'] = szCs.get(f"{{{self.NAMESPACES['w']}}}val")

            # 提取加粗、倾斜、下划线等格式
            for style_tag in ['b', 'i', 'u', 'strike', 'caps', 'smallCaps']:
                tag = rPr.find(f".//{{{self.NAMESPACES['w']}}}{style_tag}")
                if tag is not None:
                    val = tag.get(f"{{{self.NAMESPACES['w']}}}val", 'true')
                    style_info['run_properties'][style_tag] = val

            # 提取文字颜色
            color = rPr.find(f".//{{{self.NAMESPACES['w']}}}color")
            if color is not None:
                style_info['run_properties']['color'] = color.get(f"{{{self.NAMESPACES['w']}}}val")

            # 新增: 提取上下标
            vertAlign = rPr.find(f".//{{{self.NAMESPACES['w']}}}vertAlign")
            if vertAlign is not None:
                style_info['run_properties']['vertAlign'] = vertAlign.get(f"{{{self.NAMESPACES['w']}}}val")
                print(f"找到上下标设置: {style_info['run_properties']['vertAlign']}")

            # 新增: 提取文本位置偏移
            position = rPr.find(f".//{{{self.NAMESPACES['w']}}}position")
            if position is not None:
                style_info['run_properties']['position'] = position.get(f"{{{self.NAMESPACES['w']}}}val")
                print(f"找到位置偏移: {style_info['run_properties']['position']}")

            # 新增: 提取字符间距
            spacing = rPr.find(f".//{{{self.NAMESPACES['w']}}}spacing")
            if spacing is not None:
                style_info['run_properties']['char_spacing'] = spacing.get(f"{{{self.NAMESPACES['w']}}}val")
                print(f"找到字符间距: {style_info['run_properties']['char_spacing']}")

            # 新增: 提取高亮色
            highlight = rPr.find(f".//{{{self.NAMESPACES['w']}}}highlight")
            if highlight is not None:
                style_info['run_properties']['highlight'] = highlight.get(f"{{{self.NAMESPACES['w']}}}val")
                print(f"找到高亮色: {style_info['run_properties']['highlight']}")

            # 新增: 提取字符宽度比例
            w_scale = rPr.find(f".//{{{self.NAMESPACES['w']}}}w")
            if w_scale is not None:
                style_info['run_properties']['width_scale'] = w_scale.get(f"{{{self.NAMESPACES['w']}}}val")
                print(f"找到字符宽度比例: {style_info['run_properties']['width_scale']}")

            # 新增: 提取双删除线
            dstrike = rPr.find(f".//{{{self.NAMESPACES['w']}}}dstrike")
            if dstrike is not None:
                style_info['run_properties']['dstrike'] = dstrike.get(f"{{{self.NAMESPACES['w']}}}val", 'true')
                print(f"找到双删除线: {style_info['run_properties']['dstrike']}")

            # 检查段落是否有分页前属性
            page_break_before = pPr.find(f".//{{{self.NAMESPACES['w']}}}pageBreakBefore")
            if page_break_before is not None:
                val = page_break_before.get(f"{{{self.NAMESPACES['w']}}}val")
                # 当val不为"0"或"false"时表示有分页
                if val and val not in ["0", "false"]:
                    style_info['page_break']['has_page_break'] = True
                    style_info['page_break']['type'] = 'paragraph_property'
                    print(f"发现段落分页属性: pageBreakBefore={val}")

            # 检查段落中的文本运行是否包含分页符
            for run in paragraph_element.findall(f".//{{{self.NAMESPACES['w']}}}r"):
                br_elements = run.findall(f".//{{{self.NAMESPACES['w']}}}br")
                for br in br_elements:
                    br_type = br.get(f"{{{self.NAMESPACES['w']}}}type")
                    if br_type == "page":
                        style_info['page_break']['has_page_break'] = True
                        style_info['page_break']['type'] = 'manual_page_break'
                        print(f"发现手动分页符: <w:br w:type=\"page\"/>")
                        break
                if style_info['page_break']['has_page_break']:
                    break
            # 新增: 提取文本底纹
            shd = rPr.find(f".//{{{self.NAMESPACES['w']}}}shd")
            if shd is not None:
                style_info['run_properties']['shading'] = {
                    'val': shd.get(f"{{{self.NAMESPACES['w']}}}val"),
                    'color': shd.get(f"{{{self.NAMESPACES['w']}}}color"),
                    'fill': shd.get(f"{{{self.NAMESPACES['w']}}}fill")
                }
                print(f"找到文本底纹: {style_info['run_properties']['shading']}")

        # 9. 提取其他段落属性
        for child in pPr:
            tag_name = child.tag.split('}')[-1] if '}' in child.tag else child.tag
            # 跳过已经处理过的标签
            if tag_name in ['pStyle', 'jc', 'ind', 'spacing', 'pBdr', 'shd', 'numPr', 'rPr']:
                continue

            # 处理其他标签
            attrs = {}
            for key, value in child.attrib.items():
                # 简化命名空间
                attr_name = key.split('}')[-1] if '}' in key else key
                attrs[attr_name] = value

            style_info['other_properties'][tag_name] = attrs

        return style_info

    def format_paragraph_style(self, style_info):
        """将段落样式信息格式化为易读的字符串

        Args:
            style_info: extract_paragraph_style返回的样式信息字典

        Returns:
            str: 格式化后的样式信息字符串
        """
        if not style_info or style_info.get('has_style') is False:
            return "段落无样式信息"

        lines = []
        lines.append("段落样式信息:")

        if style_info['style_id']:
            lines.append(f"- 样式ID: {style_info['style_id']}")

        if style_info['alignment']:
            alignment_map = {
                'left': '左对齐',
                'right': '右对齐',
                'center': '居中',
                'both': '两端对齐',
                'distribute': '分散对齐'
            }
            align_text = alignment_map.get(style_info['alignment'], style_info['alignment'])
            lines.append(f"- 对齐方式: {align_text}")

        if style_info['indentation']:
            lines.append("- 缩进设置:")
            for key, value in style_info['indentation'].items():
                indent_name = {
                    'left': '左缩进',
                    'right': '右缩进',
                    'firstLine': '首行缩进',
                    'hanging': '悬挂缩进'
                }.get(key, key)
                lines.append(f"  • {indent_name}: {value}")

        if style_info['spacing']:
            lines.append("- 间距设置:")
            for key, value in style_info['spacing'].items():
                spacing_name = {
                    'before': '段前距',
                    'after': '段后距',
                    'line': '行距',
                    'lineRule': '行距规则'
                }.get(key, key)
                lines.append(f"  • {spacing_name}: {value}")

        if style_info['run_properties']:
            lines.append("- 文字属性:")
            if 'fonts' in style_info['run_properties']:
                lines.append("  • 字体:")
                for font_type, font in style_info['run_properties']['fonts'].items():
                    font_type_name = {
                        'ascii': '英文字体',
                        'hAnsi': '西文字体',
                        'eastAsia': '中文字体',
                        'cs': '复杂文种字体'
                    }.get(font_type, font_type)
                    lines.append(f"    ◦ {font_type_name}: {font}")

            if 'size' in style_info['run_properties']:
                # Word中的字号是实际点数的两倍
                size_pt = int(style_info['run_properties']['size']) / 2
                lines.append(f"  • 字号: {size_pt}磅 ({style_info['run_properties']['size']})")

            style_names = {
                'b': '加粗',
                'i': '倾斜',
                'u': '下划线',
                'strike': '删除线',
                'caps': '全大写',
                'smallCaps': '小型大写字母'
            }

            for style_key, style_name in style_names.items():
                if style_key in style_info['run_properties']:
                    val = style_info['run_properties'][style_key]
                    is_on = val.lower() != 'false' if isinstance(val, str) else bool(val)
                    lines.append(f"  • {style_name}: {'是' if is_on else '否'}")

            if 'color' in style_info['run_properties']:
                lines.append(f"  • 文字颜色: {style_info['run_properties']['color']}")

        return "\n".join(lines)

    # 以下是单独提取特定样式的函数

    def get_paragraph_alignment(self, num):
        """获取段落对齐方式

        Args:
            num: 段落XML元素对象num

        Returns:
            dict: 包含对齐信息的字典，如 {'alignment': 'left', 'description': '左对齐'}
        """
        result = {'alignment': None, 'description': '未设置对齐方式'}

        # 查找段落属性标签
        pPr = self.paragraphs[num]['element'].find(f".//{{{self.NAMESPACES['w']}}}pPr")
        if pPr is None:
            return result

        # 提取对齐方式
        jc = pPr.find(f".//{{{self.NAMESPACES['w']}}}jc")
        if jc is not None:
            alignment = jc.get(f"{{{self.NAMESPACES['w']}}}val")
            result['alignment'] = alignment

            # 添加中文描述
            alignment_map = {
                'left': '左对齐',
                'right': '右对齐',
                'center': '居中对齐',
                'both': '两端对齐',
                'distribute': '分散对齐',
                'justified': '两端对齐'
            }
            result['description'] = alignment_map.get(alignment, alignment)

        return result

    def get_paragraph_indentation(self, num):
        """获取段落缩进信息

        Args:
            paragraph_element: 段落XML元素对象

        Returns:
            dict: 包含缩进信息的字典
        """
        result = {
            'left': None,
            'right': None,
            'firstLine': None,
            'hanging': None,
            'description': []
        }

        # 查找段落属性标签
        pPr = self.paragraphs[num]['element'].find(f".//{{{self.NAMESPACES['w']}}}pPr")
        if pPr is None:
            return result

        # 提取缩进信息
        ind = pPr.find(f".//{{{self.NAMESPACES['w']}}}ind")
        if ind is not None:
            for key in ['left', 'right', 'firstLine', 'hanging']:
                val = ind.get(f"{{{self.NAMESPACES['w']}}}{key}")
                if val:
                    result[key] = val
                    indent_name = {
                        'left': '左缩进',
                        'right': '右缩进',
                        'firstLine': '首行缩进',
                        'hanging': '悬挂缩进'
                    }.get(key, key)
                    # Word缩进单位是1/20磅
                    indent_pt = float(val) / 20
                    result['description'].append(f"{indent_name}: {val} (约 {indent_pt:.2f}磅)")

        if not result['description']:
            result['description'] = ['未设置缩进']

        return result

    def get_paragraph_spacing(self, num):
        """
        获取段落的间距设置

        Args:
            num: 段落索引

        Returns:
            dict: 包含以下键的字典:
                before: 段前间距（磅值）
                after: 段后间距（磅值）
                beforeLines: 段前间距（行数）
                afterLines: 段后间距（行数）
                line: 行间距值
                lineRule: 行间距规则
                description: 格式化的描述
        """
        # 检查索引是否有效
        if num < 0 or num >= len(self.paragraphs):
            print(f"错误：段落索引{num}超出范围(0-{len(self.paragraphs) - 1})")
            return {}

        try:
            paragraph = self.paragraphs[num]
            para_element = paragraph.get('element')
            spacing_info = {}
            description = []

            # 查找段落属性元素
            pPr = para_element.find(".//w:pPr", self.NAMESPACES)
            if pPr is not None:
                # 查找间距元素
                spacing = pPr.find(".//w:spacing", self.NAMESPACES)
                if spacing is not None:
                    # 提取段前间距（磅值）
                    before = spacing.get(f"{{{self.NAMESPACES['w']}}}before")
                    if before is not None:
                        spacing_info['before'] = before
                        before_pt = float(before) / 20
                        description.append(f"段前间距: {before_pt}磅")

                    # 提取段前间距（行数）
                    beforeLines = spacing.get(f"{{{self.NAMESPACES['w']}}}beforeLines")
                    if beforeLines is not None:
                        spacing_info['beforeLines'] = beforeLines
                        before_lines = float(beforeLines) / 100
                        description.append(f"段前间距: {before_lines}行")

                    # 提取段后间距（磅值）
                    after = spacing.get(f"{{{self.NAMESPACES['w']}}}after")
                    if after is not None:
                        spacing_info['after'] = after
                        after_pt = float(after) / 20
                        description.append(f"段后间距: {after_pt}磅")

                    # 提取段后间距（行数）
                    afterLines = spacing.get(f"{{{self.NAMESPACES['w']}}}afterLines")
                    if afterLines is not None:
                        spacing_info['afterLines'] = afterLines
                        after_lines = float(afterLines) / 100
                        description.append(f"段后间距: {after_lines}行")

                    # 提取行间距
                    line = spacing.get(f"{{{self.NAMESPACES['w']}}}line")
                    if line is not None:
                        spacing_info['line'] = line

                        # 提取行间距规则
                        lineRule = spacing.get(f"{{{self.NAMESPACES['w']}}}lineRule")
                        if lineRule is not None:
                            spacing_info['lineRule'] = lineRule

                            # 根据规则计算实际行间距
                            if lineRule == 'auto':
                                # 多倍行距
                                line_multiple = float(line) / 240.0
                                description.append(f"行距: {line_multiple:.0%} (约 {line_multiple:.2f}倍)")
                                description.append("行距规则: 多倍行距")
                            elif lineRule == 'exact':
                                # 固定值
                                line_pt = float(line) / 20.0
                                description.append(f"行距: {line_pt}磅")
                                description.append("行距规则: 固定值")
                            elif lineRule == 'atLeast':
                                # 最小值
                                line_pt = float(line) / 20.0
                                description.append(f"行距: 最小 {line_pt}磅")
                                description.append("行距规则: 最小值")
                            else:
                                description.append(f"行距: {line} ({lineRule})")
                        else:
                            # 默认为auto
                            spacing_info['lineRule'] = 'auto'
                            line_multiple = float(line) / 240.0
                            description.append(f"行距: {line_multiple:.0%} (约 {line_multiple:.2f}倍)")
                            description.append("行距规则: 多倍行距 (默认)")

            # 添加描述信息
            if description:
                spacing_info['description'] = description
            else:
                spacing_info['description'] = ["默认间距设置"]

            return spacing_info
        except Exception as e:
            print(f"获取段落间距时出错: {e}")
            return {'error': str(e)}

    def get_paragraph_borders(self, num):
        """获取段落边框信息

        Args:
            paragraph_element: 段落XML元素对象

        Returns:
            dict: 包含边框信息的字典
        """
        result = {
            'top': None,
            'bottom': None,
            'left': None,
            'right': None,
            'description': []
        }

        # 查找段落属性标签
        pPr = self.paragraphs[num]['element'].find(f".//{{{self.NAMESPACES['w']}}}pPr")
        if pPr is None:
            return result

        # 提取边框信息
        pBdr = pPr.find(f".//{{{self.NAMESPACES['w']}}}pBdr")
        if pBdr is not None:
            for border_type in ['top', 'bottom', 'left', 'right']:
                border = pBdr.find(f".//{{{self.NAMESPACES['w']}}}{border_type}")
                if border is not None:
                    result[border_type] = {}
                    border_info = []

                    for attr in ['val', 'sz', 'space', 'color']:
                        val = border.get(f"{{{self.NAMESPACES['w']}}}{attr}")
                        if val:
                            result[border_type][attr] = val
                            if attr == 'val':
                                border_info.append(f"样式: {val}")
                            elif attr == 'sz':
                                # 边框大小以1/8磅为单位
                                border_pt = float(val) / 8
                                border_info.append(f"宽度: {border_pt:.2f}磅")
                            elif attr == 'space':
                                # 边框间距以磅为单位
                                border_info.append(f"间距: {val}磅")
                            elif attr == 'color':
                                border_info.append(f"颜色: {val}")

                    border_name = {'top': '上边框', 'bottom': '下边框', 'left': '左边框', 'right': '右边框'}.get(border_type)
                    if border_info:
                        result['description'].append(f"{border_name}: {', '.join(border_info)}")

        if not result['description']:
            result['description'] = ['无边框']

        return result

    def get_paragraph_shading(self, num):
        """获取段落背景填充信息

        Args:
            paragraph_element: 段落XML元素对象

        Returns:
            dict: 包含背景填充信息的字典
        """
        result = {
            'val': None,
            'color': None,
            'fill': None,
            'description': '无背景填充'
        }

        # 查找段落属性标签
        pPr = self.paragraphs[num]['element'].find(f".//{{{self.NAMESPACES['w']}}}pPr")
        if pPr is None:
            return result

        # 提取背景填充信息
        shading = pPr.find(f".//{{{self.NAMESPACES['w']}}}shd")
        if shading is not None:
            result['val'] = shading.get(f"{{{self.NAMESPACES['w']}}}val")
            result['color'] = shading.get(f"{{{self.NAMESPACES['w']}}}color")
            result['fill'] = shading.get(f"{{{self.NAMESPACES['w']}}}fill")

            descriptions = []
            if result['val']:
                shading_map = {
                    'clear': '清除',
                    'solid': '实心'
                }
                val_text = shading_map.get(result['val'], result['val'])
                descriptions.append(f"类型: {val_text}")

            if result['color']:
                descriptions.append(f"前景色: {result['color']}")

            if result['fill']:
                descriptions.append(f"背景色: {result['fill']}")

            if descriptions:
                result['description'] = '背景填充: ' + ', '.join(descriptions)

        return result

    def get_paragraph_numbering(self, num):
        """获取段落编号信息

        Args:
            paragraph_element: 段落XML元素对象

        Returns:
            dict: 包含编号信息的字典
        """
        result = {
            'id': None,
            'level': None,
            'description': '无编号'
        }

        # 查找段落属性标签
        pPr = self.paragraphs[num]['element'].find(f".//{{{self.NAMESPACES['w']}}}pPr")
        if pPr is None:
            return result

        # 提取编号信息
        numPr = pPr.find(f".//{{{self.NAMESPACES['w']}}}numPr")
        if numPr is not None:
            ilvl = numPr.find(f".//{{{self.NAMESPACES['w']}}}ilvl")
            if ilvl is not None:
                result['level'] = ilvl.get(f"{{{self.NAMESPACES['w']}}}val")

            numId = numPr.find(f".//{{{self.NAMESPACES['w']}}}numId")
            if numId is not None:
                result['id'] = numId.get(f"{{{self.NAMESPACES['w']}}}val")

            descriptions = []
            if result['id']:
                descriptions.append(f"编号ID: {result['id']}")
            if result['level']:
                level_num = int(result['level'])
                descriptions.append(f"级别: {level_num + 1} (内部值: {result['level']})")

            if descriptions:
                result['description'] = '编号设置: ' + ', '.join(descriptions)

        return result


    def get_paragraph_font(self, num):
        """
        获取段落的字体属性

        Args:
            num: 段落索引

        Returns:
            dict: 包含字体属性的字典
        """
        # 检查索引是否有效
        if num < 0 or num >= len(self.paragraphs):
            print(f"错误：段落索引{num}超出范围(0-{len(self.paragraphs) - 1})")
            return {}

        try:
            paragraph = self.paragraphs[num]
            para_element = paragraph.get('element')
            result = {
                'fonts': {},  # 字体名称
                'size': None,  # 字体大小
                'attributes': {},  # 各种属性
                'color': None,  # 颜色
                'description': []  # 描述信息
            }

            # 查找段落属性元素
            pPr = para_element.find(".//w:pPr", self.NAMESPACES)
            if pPr is not None:
                # 查找段落级字体设置
                rPr = pPr.find(".//w:rPr", self.NAMESPACES)
                if rPr is not None:
                    # 提取字体信息
                    rFonts = rPr.find(".//w:rFonts", self.NAMESPACES)
                    if rFonts is not None:
                        # 提取各种字体名称
                        for font_type in ['ascii', 'hAnsi', 'eastAsia', 'cs']:
                            font_val = rFonts.get(f"{{{self.NAMESPACES['w']}}}{font_type}")
                            if font_val:
                                result['fonts'][font_type] = font_val
                                result['description'].append(f"{font_type}字体: {font_val}")

                    # 提取字体大小
                    sz = rPr.find(".//w:sz", self.NAMESPACES)
                    if sz is not None:
                        size_val = sz.get(f"{{{self.NAMESPACES['w']}}}val")
                        if size_val:
                            # Word中字体大小是半磅值，需要除以2
                            size_pt = int(size_val) / 2
                            result['size'] = size_pt
                            result['description'].append(f"字体大小: {size_pt}磅")

                    # 提取颜色
                    color = rPr.find(".//w:color", self.NAMESPACES)
                    if color is not None:
                        color_val = color.get(f"{{{self.NAMESPACES['w']}}}val")
                        result['color'] = color_val
                        result['description'].append(f"颜色: {color_val}")

                    # 提取加粗、斜体等属性
                    attrs = {
                        'b': '加粗',
                        'i': '斜体',
                        'u': '下划线',
                        'strike': '删除线',
                        'caps': '全大写',
                        'smallCaps': '小型大写字母'
                    }
                    for attr, desc in attrs.items():
                        attr_elem = rPr.find(f".//w:{attr}", self.NAMESPACES)
                        if attr_elem is not None:
                            val = attr_elem.get(f"{{{self.NAMESPACES['w']}}}val")
                            # 如果没有val属性或val=true/1，则属性生效
                            is_on = val is None or val.lower() in ['true', '1', 'on']
                            if is_on:
                                result['attributes'][attr] = True
                                result['description'].append(desc)

            # 如果段落级字体设置不存在或不完整，检查第一个run的设置
            # 这对应于Word中"整个段落"的字体设置显示
            if not result['fonts'] and not result['size'] and not result['color']:
                # 获取第一个run
                r_elements = para_element.findall(f".//{{{self.NAMESPACES['w']}}}r")
                if r_elements:
                    first_run = r_elements[0]
                    rPr = first_run.find(".//w:rPr", self.NAMESPACES)
                    if rPr is not None:
                        # 提取字体信息
                        rFonts = rPr.find(".//w:rFonts", self.NAMESPACES)
                        if rFonts is not None:
                            # 提取各种字体名称
                            for font_type in ['ascii', 'hAnsi', 'eastAsia', 'cs']:
                                font_val = rFonts.get(f"{{{self.NAMESPACES['w']}}}{font_type}")
                                if font_val:
                                    if 'fonts' not in result:
                                        result['fonts'] = {}
                                    result['fonts'][font_type] = font_val
                                    if f"{font_type}字体: {font_val}" not in result['description']:
                                        result['description'].append(f"{font_type}字体: {font_val}")

                        # 提取字体大小
                        sz = rPr.find(".//w:sz", self.NAMESPACES)
                        if sz is not None and not result['size']:
                            size_val = sz.get(f"{{{self.NAMESPACES['w']}}}val")
                            if size_val:
                                # Word中字体大小是半磅值，需要除以2
                                size_pt = int(size_val) / 2
                                result['size'] = size_pt
                                if f"字体大小: {size_pt}磅" not in result['description']:
                                    result['description'].append(f"字体大小: {size_pt}磅")

                        # 提取颜色
                        color = rPr.find(".//w:color", self.NAMESPACES)
                        if color is not None and not result['color']:
                            color_val = color.get(f"{{{self.NAMESPACES['w']}}}val")
                            result['color'] = color_val
                            if f"颜色: {color_val}" not in result['description']:
                                result['description'].append(f"颜色: {color_val}")

                        # 提取属性
                        for attr, desc in attrs.items():
                            if attr not in result['attributes']:
                                attr_elem = rPr.find(f".//w:{attr}", self.NAMESPACES)
                                if attr_elem is not None:
                                    val = attr_elem.get(f"{{{self.NAMESPACES['w']}}}val")
                                    # 如果没有val属性或val=true/1，则属性生效
                                    is_on = val is None or val.lower() in ['true', '1', 'on']
                                    if is_on:
                                        result['attributes'][attr] = True
                                        if desc not in result['description']:
                                            result['description'].append(desc)

            if not result['description']:
                result['description'] = ['未设置字体属性']

            return result
        except Exception as e:
            print(f"获取段落字体属性时出错: {e}")
            return {'error': str(e)}

    def get_all_paragraph_styles(self, num):
        """获取段落的所有样式信息

        Args:
            num: 段落XML元素对象索引

        Returns:
            dict: 包含所有样式信息的字典
        """

        return {
            'alignment': self.get_paragraph_alignment(num),
            'indentation': self.get_paragraph_indentation(num),
            'spacing': self.get_paragraph_spacing(num),
            'borders': self.get_paragraph_borders(num),
            'shading': self.get_paragraph_shading(num),
            'numbering': self.get_paragraph_numbering(num),
            'fonts': self.get_paragraph_font(num)
        }

    def get_element_run_text(self, index):
        """提取指定索引元素中所有w:r/w:t的文本内容

        Args:
            index: self.elements的索引

        Returns:
            str: 所有w:r/w:t中的文本内容连接成的字符串
        """
        # 检查索引是否有效
        if index < 0 or index >= len(self.elements):
            print(f"错误：元素索引{index}超出范围(0-{len(self.elements)-1})")
            return ""

        # 获取指定索引的元素
        element = self.elements[index]['element']

        # 查找所有w:r元素
        r_elements = element.findall(f".//{{{self.NAMESPACES['w']}}}r")

        # 提取所有w:t的文本内容
        texts = []
        for r in r_elements:
            t_elements = r.findall(f".//{{{self.NAMESPACES['w']}}}t")
            for t in t_elements:
                if t.text:
                    texts.append(t.text)


        return texts

    def get_paragraph_run_text(self, index):
            """提取指定索引元素中所有w:r/w:t的文本内容

            Args:
                index: self.paragraphs的索引

            Returns:
                str: 所有w:r/w:t中的文本内容连接成的字符串
            """
            # 检查索引是否有效
            if index < 0 or index >= len(self.elements):
                print(f"错误：元素索引{index}超出范围(0-{len(self.elements) - 1})")
                return ""

            # 获取指定索引的元素
            element = self.paragraphs[index]['element']

            # 查找所有w:r元素
            r_elements = element.findall(f".//{{{self.NAMESPACES['w']}}}r")

            # 提取所有w:t的文本内容
            texts = []
            for r in r_elements:
                t_elements = r.findall(f".//{{{self.NAMESPACES['w']}}}t")
                for t in t_elements:
                    if t.text:
                        texts.append(t.text)

            return texts
    def get_element_run_content(self, index):
        """提取指定索引元素中所有w:r元素的详细内容

        Args:
            index: self.elements的索引

        Returns:
            list: 包含每个w:r元素内容信息的列表
        """
        # 检查索引是否有效
        if index < 0 or index >= len(self.elements):
            print(f"错误：元素索引{index}超出范围(0-{len(self.elements)-1})")
            return []

        # 获取指定索引的元素
        element = self.elements[index]['element']

        # 查找所有w:r元素
        r_elements = element.findall(f".//{{{self.NAMESPACES['w']}}}r")

        # 提取每个w:r的内容信息
        r_contents = []
        for r in r_elements:
            r_info = {'text': '', 'has_drawing': False, 'has_symbol': False, 'has_tab': False}

            # 提取文本内容
            t_elements = r.findall(f".//{{{self.NAMESPACES['w']}}}t")
            r_info['text'] = "".join([t.text if t.text else '' for t in t_elements])

            # 检查是否包含图片
            drawing = r.find(f".//{{{self.NAMESPACES['w']}}}drawing")
            if drawing is not None:
                r_info['has_drawing'] = True

                # 尝试提取图片描述信息
                docPr = drawing.find(f".//{{{self.NAMESPACES['wp']}}}docPr")
                if docPr is not None:
                    r_info['drawing_name'] = docPr.get('name', '')
                    r_info['drawing_description'] = docPr.get('descr', '')

                # 尝试提取图片关系ID
                blip = drawing.find(f".//{{{self.NAMESPACES['a']}}}blip")
                if blip is not None:
                    r_info['drawing_relationship'] = blip.get(f"{{{self.NAMESPACES['r']}}}embed", '')

            # 检查是否包含符号
            sym = r.find(f".//{{{self.NAMESPACES['w']}}}sym")
            if sym is not None:
                r_info['has_symbol'] = True
                r_info['symbol_font'] = sym.get(f"{{{self.NAMESPACES['w']}}}font", '')
                r_info['symbol_char'] = sym.get(f"{{{self.NAMESPACES['w']}}}char", '')

            # 检查是否包含制表符
            tab = r.find(f".//{{{self.NAMESPACES['w']}}}tab")
            if tab is not None:
                r_info['has_tab'] = True

            # 添加到结果列表
            r_contents.append(r_info)

        return r_contents

    def get_image_by_relation_id(self, relation_id):
        """通过关系ID找到对应的图片

        Args:
            relation_id: 图片的关系ID (例如 'rId38')

        Returns:
            tuple: (图片名称, 图片二进制数据) 或者 (None, None)
        """
        # 获取文档关系数据
        relationships = self.parts['relationships']
        if relationships is None:
            print("无法获取文档关系")
            return None, None

        # 在关系中查找指定ID
        target_path = None
        rels_root = relationships.getroot()

        for rel in rels_root.findall('.//{*}Relationship'):
            if rel.get('Id') == relation_id:
                target_path = rel.get('Target')
                break

        if not target_path:
            print(f"未找到关系ID为 {relation_id} 的图片")
            return None, None

        # 处理路径格式
        if target_path.startswith('/'):
            target_path = target_path[1:]
        if not target_path.startswith('media/'):
            target_path = f"word/{target_path}"

        # 提取文件名
        image_name = target_path.split('/')[-1]

        # 尝试从media字典中获取图片数据
        for media_name, media_data in self.parts['media'].items():
            if media_name == image_name:
                return media_name, media_data

        print(f"未找到路径为 {target_path} 的图片")
        return None, None

    def save_image_by_relation_id(self, relation_id, output_path):
        """通过关系ID保存图片到指定路径

        Args:
            relation_id: 图片的关系ID (例如 'rId38')
            output_path: 输出文件路径

        Returns:
            bool: 是否成功保存
        """
        image_name, image_data = self.get_image_by_relation_id(relation_id)

        if image_data:
            try:
                # 创建输出目录（如果不存在）
                output_dir = os.path.dirname(output_path)
                if output_dir and not os.path.exists(output_dir):
                    os.makedirs(output_dir)

                # 写入图片文件
                with open(output_path, 'wb') as f:
                    f.write(image_data)

                print(f"已成功保存图片 {image_name} 到 {output_path}")
                return True
            except Exception as e:
                print(f"保存图片时出错: {e}")

        return False
    def element_to_dict(self,element_index,element_type="elements"):

        if element_type == "paragraphs":
            element = self.paragraphs[element_index]['element']
        elif element_type == "tables":
            element = self.elements[element_index]['element']
        elif element_type == "elements":
            # 获取指定索引的元素
            element = self.elements[element_index]['element']
        else:
            print(f"错误：元素类型{element_type}无效")
            return {}
        return element


    def get_run_style(self, para_index, run_index, element_type="paragraphs"):
        """获取指定元素中run的样式

        Args:
            para_index: 段落索引
            run_index: Run元素索引
            element_type: 元素类型，默认为"paragraphs"，也可以是"elements"

        Returns:
            dict: 包含样式信息的字典
        """
        # 检查索引是否有效
        if para_index < 0 or para_index >= len(self.paragraphs):
            print(f"错误：段落索引{para_index}超出范围(0-{len(self.paragraphs) - 1})")
            return {'has_style': False, 'message': '无效的段落索引'}

        # 获取段落元素
        paragraph = self.paragraphs[para_index]
        para_element = paragraph.get('element')

        # 查找所有w:r元素
        r_elements = para_element.findall(f".//{{{self.NAMESPACES['w']}}}r")
        if run_index < 0 or run_index >= len(r_elements):
            print(f"错误：Run索引{run_index}超出范围(0-{len(r_elements) - 1})")
            return {'has_style': False, 'message': '无效的Run索引'}

        # 获取指定的Run元素
        run = r_elements[run_index]

        style_info = {
            'fonts': {},  # 字体名称
            'size': None,  # 字体大小
            'bold': False,  # 是否加粗
            'italic': False,  # 是否斜体
            'underline': None,  # 下划线类型
            'color': None,  # 颜色
            'highlight': None,  # 突出显示颜色
            'strike': False,  # 是否删除线
            'caps': False,  # 是否全大写
            'small_caps': False,  # 是否小型大写字母
            'spacing': None,  # 字符间距
            'vert_align': None,  # 垂直对齐方式
            'other_properties': {}  # 其他属性
        }

        # 查找rPr元素
        rPr = run.find(f"./w:rPr", self.NAMESPACES)
        if rPr is None:
            # 尝试查找嵌套的rPr
            rPr = run.find(f".//{{{self.NAMESPACES['w']}}}rPr", self.NAMESPACES)

        if rPr is not None:
            # 提取字体信息
            rFonts = rPr.find(f"./w:rFonts", self.NAMESPACES)
            if rFonts is None:
                rFonts = rPr.find(f".//{{{self.NAMESPACES['w']}}}rFonts", self.NAMESPACES)

            if rFonts is not None:
                # 提取各种字体名称
                for font_type in ['ascii', 'hAnsi', 'eastAsia', 'cs']:
                    font_val = rFonts.get(f"{{{self.NAMESPACES['w']}}}{font_type}")
                    if font_val:
                        style_info['fonts'][font_type] = font_val

            # 提取字体大小
            sz = rPr.find(f"./w:sz", self.NAMESPACES)
            if sz is None:
                sz = rPr.find(f".//{{{self.NAMESPACES['w']}}}sz", self.NAMESPACES)

            if sz is not None:
                size_val = sz.get(f"{{{self.NAMESPACES['w']}}}val")
                if size_val:
                    # 转换为浮点数，并转换为磅值（除以2）
                    style_info['size'] = str(float(size_val) / 2)

            # 检查是否加粗
            bold = rPr.find(f"./w:b", self.NAMESPACES)
            if bold is None:
                bold = rPr.find(f".//{{{self.NAMESPACES['w']}}}b", self.NAMESPACES)

            if bold is not None:
                val = bold.get(f"{{{self.NAMESPACES['w']}}}val")
                # 如果没有val属性或val=true/1，则加粗
                style_info['bold'] = val is None or val.lower() in ['true', '1', 'on']

            # 检查是否斜体
            italic = rPr.find(f"./w:i", self.NAMESPACES)
            if italic is None:
                italic = rPr.find(f".//{{{self.NAMESPACES['w']}}}i", self.NAMESPACES)

            if italic is not None:
                val = italic.get(f"{{{self.NAMESPACES['w']}}}val")
                # 如果没有val属性或val=true/1，则斜体
                style_info['italic'] = val is None or val.lower() in ['true', '1', 'on']

            # 检查下划线
            underline = rPr.find(f"./w:u", self.NAMESPACES)
            if underline is None:
                underline = rPr.find(f".//{{{self.NAMESPACES['w']}}}u", self.NAMESPACES)

            if underline is not None:
                val = underline.get(f"{{{self.NAMESPACES['w']}}}val")
                if val:
                    style_info['underline'] = val

            # 检查颜色
            color = rPr.find(f"./w:color", self.NAMESPACES)
            if color is None:
                color = rPr.find(f".//{{{self.NAMESPACES['w']}}}color", self.NAMESPACES)

            if color is not None:
                val = color.get(f"{{{self.NAMESPACES['w']}}}val")
                if val:
                    style_info['color'] = val

            # 检查突出显示
            highlight = rPr.find(f"./w:highlight", self.NAMESPACES)
            if highlight is None:
                highlight = rPr.find(f".//{{{self.NAMESPACES['w']}}}highlight", self.NAMESPACES)

            if highlight is not None:
                val = highlight.get(f"{{{self.NAMESPACES['w']}}}val")
                if val:
                    style_info['highlight'] = val

            # 检查删除线
            strike = rPr.find(f"./w:strike", self.NAMESPACES)
            if strike is None:
                strike = rPr.find(f".//{{{self.NAMESPACES['w']}}}strike", self.NAMESPACES)

            if strike is not None:
                val = strike.get(f"{{{self.NAMESPACES['w']}}}val")
                # 如果没有val属性或val=true/1，则有删除线
                style_info['strike'] = val is None or val.lower() in ['true', '1', 'on']

            # 检查全大写
            caps = rPr.find(f"./w:caps", self.NAMESPACES)
            if caps is None:
                caps = rPr.find(f".//{{{self.NAMESPACES['w']}}}caps", self.NAMESPACES)

            if caps is not None:
                val = caps.get(f"{{{self.NAMESPACES['w']}}}val")
                # 如果没有val属性或val=true/1，则全大写
                style_info['caps'] = val is None or val.lower() in ['true', '1', 'on']

            # 检查小型大写字母
            smallCaps = rPr.find(f"./w:smallCaps", self.NAMESPACES)
            if smallCaps is None:
                smallCaps = rPr.find(f".//{{{self.NAMESPACES['w']}}}smallCaps", self.NAMESPACES)

            if smallCaps is not None:
                val = smallCaps.get(f"{{{self.NAMESPACES['w']}}}val")
                # 如果没有val属性或val=true/1，则小型大写字母
                style_info['small_caps'] = val is None or val.lower() in ['true', '1', 'on']

            # 检查字符间距
            spacing = rPr.find(f"./w:spacing", self.NAMESPACES)
            if spacing is None:
                spacing = rPr.find(f".//{{{self.NAMESPACES['w']}}}spacing", self.NAMESPACES)

            if spacing is not None:
                val = spacing.get(f"{{{self.NAMESPACES['w']}}}val")
                if val:
                    style_info['spacing'] = val

            # 检查垂直对齐方式
            vertAlign = rPr.find(f"./w:vertAlign", self.NAMESPACES)
            if vertAlign is None:
                vertAlign = rPr.find(f".//{{{self.NAMESPACES['w']}}}vertAlign", self.NAMESPACES)

            if vertAlign is not None:
                val = vertAlign.get(f"{{{self.NAMESPACES['w']}}}val")
                if val:
                    style_info['vert_align'] = val

            # 调试信息

            for key, value in style_info.items():
                if value and value != {} and value != False:
                    print(f"  {key}: {value}")

            return style_info
        else:
            # 读取run的文本内容
            t_elements = run.findall(f".//{{{self.NAMESPACES['w']}}}t")
            text_content = ""
            for t in t_elements:
                if t.text:
                    text_content += t.text

            # 没有样式信息，但有内容
            if text_content.strip():
                print(f"段落 {para_index} 的Run {run_index} 没有样式信息，但有文本内容: '{text_content}'")
                return style_info
            else:
                return {'has_style': False, 'message': 'Run无样式信息'}
    def get_run_style_form_xml(self, para, run_index, element_type="paragraphs"):
        """获取指定元素中run的样式

        Args:
            para 段落
            run_index: Run元素索引
            element_type: 元素类型，默认为"paragraphs"，也可以是"elements"

        Returns:
            dict: 包含样式信息的字典
        """


        # 获取段落元素

        para_element = para

        # 查找所有w:r元素
        r_elements = para_element.findall(f".//{{{self.NAMESPACES['w']}}}r")
        if run_index < 0 or run_index >= len(r_elements):
            print(f"错误：Run索引{run_index}超出范围(0-{len(r_elements) - 1})")
            return {'has_style': False, 'message': '无效的Run索引'}

        # 获取指定的Run元素
        run = r_elements[run_index]

        style_info = {
            'fonts': {},  # 字体名称
            'size': None,  # 字体大小
            'bold': False,  # 是否加粗
            'italic': False,  # 是否斜体
            'underline': None,  # 下划线类型
            'color': None,  # 颜色
            'highlight': None,  # 突出显示颜色
            'strike': False,  # 是否删除线
            'caps': False,  # 是否全大写
            'small_caps': False,  # 是否小型大写字母
            'spacing': None,  # 字符间距
            'vert_align': None,  # 垂直对齐方式
            'other_properties': {}  # 其他属性
        }

        # 查找rPr元素
        rPr = run.find(f"./w:rPr", self.NAMESPACES)
        if rPr is None:
            # 尝试查找嵌套的rPr
            rPr = run.find(f".//{{{self.NAMESPACES['w']}}}rPr", self.NAMESPACES)

        if rPr is not None:
            # 提取字体信息
            rFonts = rPr.find(f"./w:rFonts", self.NAMESPACES)
            if rFonts is None:
                rFonts = rPr.find(f".//{{{self.NAMESPACES['w']}}}rFonts", self.NAMESPACES)

            if rFonts is not None:
                # 提取各种字体名称
                for font_type in ['ascii', 'hAnsi', 'eastAsia', 'cs']:
                    font_val = rFonts.get(f"{{{self.NAMESPACES['w']}}}{font_type}")
                    if font_val:
                        style_info['fonts'][font_type] = font_val

            # 提取字体大小
            sz = rPr.find(f"./w:sz", self.NAMESPACES)
            if sz is None:
                sz = rPr.find(f".//{{{self.NAMESPACES['w']}}}sz", self.NAMESPACES)

            if sz is not None:
                size_val = sz.get(f"{{{self.NAMESPACES['w']}}}val")
                if size_val:
                    # 转换为浮点数，并转换为磅值（除以2）
                    style_info['size'] = str(float(size_val) / 2)

            # 检查是否加粗
            bold = rPr.find(f"./w:b", self.NAMESPACES)
            if bold is None:
                bold = rPr.find(f".//{{{self.NAMESPACES['w']}}}b", self.NAMESPACES)

            if bold is not None:
                val = bold.get(f"{{{self.NAMESPACES['w']}}}val")
                # 如果没有val属性或val=true/1，则加粗
                style_info['bold'] = val is None or val.lower() in ['true', '1', 'on']

            # 检查是否斜体
            italic = rPr.find(f"./w:i", self.NAMESPACES)
            if italic is None:
                italic = rPr.find(f".//{{{self.NAMESPACES['w']}}}i", self.NAMESPACES)

            if italic is not None:
                val = italic.get(f"{{{self.NAMESPACES['w']}}}val")
                # 如果没有val属性或val=true/1，则斜体
                style_info['italic'] = val is None or val.lower() in ['true', '1', 'on']

            # 检查下划线
            underline = rPr.find(f"./w:u", self.NAMESPACES)
            if underline is None:
                underline = rPr.find(f".//{{{self.NAMESPACES['w']}}}u", self.NAMESPACES)

            if underline is not None:
                val = underline.get(f"{{{self.NAMESPACES['w']}}}val")
                if val:
                    style_info['underline'] = val

            # 检查颜色
            color = rPr.find(f"./w:color", self.NAMESPACES)
            if color is None:
                color = rPr.find(f".//{{{self.NAMESPACES['w']}}}color", self.NAMESPACES)

            if color is not None:
                val = color.get(f"{{{self.NAMESPACES['w']}}}val")
                if val:
                    style_info['color'] = val

            # 检查突出显示
            highlight = rPr.find(f"./w:highlight", self.NAMESPACES)
            if highlight is None:
                highlight = rPr.find(f".//{{{self.NAMESPACES['w']}}}highlight", self.NAMESPACES)

            if highlight is not None:
                val = highlight.get(f"{{{self.NAMESPACES['w']}}}val")
                if val:
                    style_info['highlight'] = val

            # 检查删除线
            strike = rPr.find(f"./w:strike", self.NAMESPACES)
            if strike is None:
                strike = rPr.find(f".//{{{self.NAMESPACES['w']}}}strike", self.NAMESPACES)

            if strike is not None:
                val = strike.get(f"{{{self.NAMESPACES['w']}}}val")
                # 如果没有val属性或val=true/1，则有删除线
                style_info['strike'] = val is None or val.lower() in ['true', '1', 'on']

            # 检查全大写
            caps = rPr.find(f"./w:caps", self.NAMESPACES)
            if caps is None:
                caps = rPr.find(f".//{{{self.NAMESPACES['w']}}}caps", self.NAMESPACES)

            if caps is not None:
                val = caps.get(f"{{{self.NAMESPACES['w']}}}val")
                # 如果没有val属性或val=true/1，则全大写
                style_info['caps'] = val is None or val.lower() in ['true', '1', 'on']

            # 检查小型大写字母
            smallCaps = rPr.find(f"./w:smallCaps", self.NAMESPACES)
            if smallCaps is None:
                smallCaps = rPr.find(f".//{{{self.NAMESPACES['w']}}}smallCaps", self.NAMESPACES)

            if smallCaps is not None:
                val = smallCaps.get(f"{{{self.NAMESPACES['w']}}}val")
                # 如果没有val属性或val=true/1，则小型大写字母
                style_info['small_caps'] = val is None or val.lower() in ['true', '1', 'on']

            # 检查字符间距
            spacing = rPr.find(f"./w:spacing", self.NAMESPACES)
            if spacing is None:
                spacing = rPr.find(f".//{{{self.NAMESPACES['w']}}}spacing", self.NAMESPACES)

            if spacing is not None:
                val = spacing.get(f"{{{self.NAMESPACES['w']}}}val")
                if val:
                    style_info['spacing'] = val

            # 检查垂直对齐方式
            vertAlign = rPr.find(f"./w:vertAlign", self.NAMESPACES)
            if vertAlign is None:
                vertAlign = rPr.find(f".//{{{self.NAMESPACES['w']}}}vertAlign", self.NAMESPACES)

            if vertAlign is not None:
                val = vertAlign.get(f"{{{self.NAMESPACES['w']}}}val")
                if val:
                    style_info['vert_align'] = val

            # 调试信息

            for key, value in style_info.items():
                if value and value != {} and value != False:
                    print(f"  {key}: {value}")

            return style_info
        else:
            # 读取run的文本内容
            t_elements = run.findall(f".//{{{self.NAMESPACES['w']}}}t")
            text_content = ""
            for t in t_elements:
                if t.text:
                    text_content += t.text

            # 没有样式信息，但有内容
            if text_content.strip():

                return style_info
            else:
                return {'has_style': False, 'message': 'Run无样式信息'}
    def _get_run_style(self, para_index, run_index):
        """获取指定元素中run的样式

        Args:
            para_index: 段落索引
            run_index: Run元素索引
            element_type: 元素类型，默认为"paragraphs"，也可以是"elements"

        Returns:
            dict: 包含样式信息的字典
        """


        # 获取段落元素

        # 检查索引是否有效
        if para_index < 0 or para_index >= len(self.elements):
            print(f"错误：段落索引{para_index}超出范围(0-{len(self.elements) - 1})")
            return {'has_style': False, 'message': '无效的段落索引'}

        # 获取段落元素
        paragraph = self.elements[para_index]
        para_element = paragraph.get('element')

        # 查找所有w:r元素
        r_elements = para_element.findall(f".//{{{self.NAMESPACES['w']}}}r")
        if run_index < 0 or run_index >= len(r_elements):
            print(f"错误：Run索引{run_index}超出范围(0-{len(r_elements) - 1})")
            return {'has_style': False, 'message': '无效的Run索引'}

        # 获取指定的Run元素
        run = r_elements[run_index]

        style_info = {
            'fonts': {},  # 字体名称
            'size': None,  # 字体大小
            'bold': False,  # 是否加粗
            'italic': False,  # 是否斜体
            'underline': None,  # 下划线类型
            'color': None,  # 颜色
            'highlight': None,  # 突出显示颜色
            'strike': False,  # 是否删除线
            'caps': False,  # 是否全大写
            'small_caps': False,  # 是否小型大写字母
            'spacing': None,  # 字符间距
            'vert_align': None,  # 垂直对齐方式
            'other_properties': {}  # 其他属性
        }

        # 查找rPr元素
        rPr = run.find(f"./w:rPr", self.NAMESPACES)
        if rPr is None:
            # 尝试查找嵌套的rPr
            rPr = run.find(f".//{{{self.NAMESPACES['w']}}}rPr", self.NAMESPACES)

        if rPr is not None:
            # 提取字体信息
            rFonts = rPr.find(f"./w:rFonts", self.NAMESPACES)
            if rFonts is None:
                rFonts = rPr.find(f".//{{{self.NAMESPACES['w']}}}rFonts", self.NAMESPACES)

            if rFonts is not None:
                # 提取各种字体名称
                for font_type in ['ascii', 'hAnsi', 'eastAsia', 'cs']:
                    font_val = rFonts.get(f"{{{self.NAMESPACES['w']}}}{font_type}")
                    if font_val:
                        style_info['fonts'][font_type] = font_val

            # 提取字体大小
            sz = rPr.find(f"./w:sz", self.NAMESPACES)
            if sz is None:
                sz = rPr.find(f".//{{{self.NAMESPACES['w']}}}sz", self.NAMESPACES)

            if sz is not None:
                size_val = sz.get(f"{{{self.NAMESPACES['w']}}}val")
                if size_val:
                    # 转换为浮点数，并转换为磅值（除以2）
                    style_info['size'] = str(float(size_val) / 2)

            # 检查是否加粗
            bold = rPr.find(f"./w:b", self.NAMESPACES)
            if bold is None:
                bold = rPr.find(f".//{{{self.NAMESPACES['w']}}}b", self.NAMESPACES)

            if bold is not None:
                val = bold.get(f"{{{self.NAMESPACES['w']}}}val")
                # 如果没有val属性或val=true/1，则加粗
                style_info['bold'] = val is None or val.lower() in ['true', '1', 'on']

            # 检查是否斜体
            italic = rPr.find(f"./w:i", self.NAMESPACES)
            if italic is None:
                italic = rPr.find(f".//{{{self.NAMESPACES['w']}}}i", self.NAMESPACES)

            if italic is not None:
                val = italic.get(f"{{{self.NAMESPACES['w']}}}val")
                # 如果没有val属性或val=true/1，则斜体
                style_info['italic'] = val is None or val.lower() in ['true', '1', 'on']

            # 检查下划线
            underline = rPr.find(f"./w:u", self.NAMESPACES)
            if underline is None:
                underline = rPr.find(f".//{{{self.NAMESPACES['w']}}}u", self.NAMESPACES)

            if underline is not None:
                val = underline.get(f"{{{self.NAMESPACES['w']}}}val")
                if val:
                    style_info['underline'] = val

            # 检查颜色
            color = rPr.find(f"./w:color", self.NAMESPACES)
            if color is None:
                color = rPr.find(f".//{{{self.NAMESPACES['w']}}}color", self.NAMESPACES)

            if color is not None:
                val = color.get(f"{{{self.NAMESPACES['w']}}}val")
                if val:
                    style_info['color'] = val

            # 检查突出显示
            highlight = rPr.find(f"./w:highlight", self.NAMESPACES)
            if highlight is None:
                highlight = rPr.find(f".//{{{self.NAMESPACES['w']}}}highlight", self.NAMESPACES)

            if highlight is not None:
                val = highlight.get(f"{{{self.NAMESPACES['w']}}}val")
                if val:
                    style_info['highlight'] = val

            # 检查删除线
            strike = rPr.find(f"./w:strike", self.NAMESPACES)
            if strike is None:
                strike = rPr.find(f".//{{{self.NAMESPACES['w']}}}strike", self.NAMESPACES)

            if strike is not None:
                val = strike.get(f"{{{self.NAMESPACES['w']}}}val")
                # 如果没有val属性或val=true/1，则有删除线
                style_info['strike'] = val is None or val.lower() in ['true', '1', 'on']

            # 检查全大写
            caps = rPr.find(f"./w:caps", self.NAMESPACES)
            if caps is None:
                caps = rPr.find(f".//{{{self.NAMESPACES['w']}}}caps", self.NAMESPACES)

            if caps is not None:
                val = caps.get(f"{{{self.NAMESPACES['w']}}}val")
                # 如果没有val属性或val=true/1，则全大写
                style_info['caps'] = val is None or val.lower() in ['true', '1', 'on']

            # 检查小型大写字母
            smallCaps = rPr.find(f"./w:smallCaps", self.NAMESPACES)
            if smallCaps is None:
                smallCaps = rPr.find(f".//{{{self.NAMESPACES['w']}}}smallCaps", self.NAMESPACES)

            if smallCaps is not None:
                val = smallCaps.get(f"{{{self.NAMESPACES['w']}}}val")
                # 如果没有val属性或val=true/1，则小型大写字母
                style_info['small_caps'] = val is None or val.lower() in ['true', '1', 'on']

            # 检查字符间距
            spacing = rPr.find(f"./w:spacing", self.NAMESPACES)
            if spacing is None:
                spacing = rPr.find(f".//{{{self.NAMESPACES['w']}}}spacing", self.NAMESPACES)

            if spacing is not None:
                val = spacing.get(f"{{{self.NAMESPACES['w']}}}val")
                if val:
                    style_info['spacing'] = val

            # 检查垂直对齐方式
            vertAlign = rPr.find(f"./w:vertAlign", self.NAMESPACES)
            if vertAlign is None:
                vertAlign = rPr.find(f".//{{{self.NAMESPACES['w']}}}vertAlign", self.NAMESPACES)

            if vertAlign is not None:
                val = vertAlign.get(f"{{{self.NAMESPACES['w']}}}val")
                if val:
                    style_info['vert_align'] = val

            # 调试信息

            for key, value in style_info.items():
                if value and value != {} and value != False:
                    print(f"  {key}: {value}")

            return style_info
        else:
            # 读取run的文本内容
            t_elements = run.findall(f".//{{{self.NAMESPACES['w']}}}t")
            text_content = ""
            for t in t_elements:
                if t.text:
                    text_content += t.text

            # 没有样式信息，但有内容
            if text_content.strip():

                return style_info
            else:
                return {'has_style': False, 'message': 'Run无样式信息'}
    # 以下为单独提取特定样式的辅助函数

    def get_run_font(self, element_index, run_index,element_type="elements"):
        """提取Run的字体信息

        Args:
            element_index: self.elements的索引
            run_index: 元素中w:r的索引

        Returns:
            dict: 字体信息
        """
        # 检查元素索引是否有效
        if element_index < 0 or element_index >= len(self.elements):
            print(f"错误：元素索引{element_index}超出范围(0-{len(self.elements)-1})")
            return {'fonts': {}, 'description': '无法获取字体信息'}

        # 获取指定索引的元素
        element = self.element_to_dict(element_index, element_type)

        # 查找所有w:r元素
        r_elements = element.findall(f".//{{{self.NAMESPACES['w']}}}r")

        # 检查Run索引是否有效
        if run_index < 0 or run_index >= len(r_elements):
            print(f"错误：Run索引{run_index}超出范围(0-{len(r_elements)-1})")
            return {'fonts': {}, 'description': '无法获取字体信息'}

        # 获取指定的Run元素
        run = r_elements[run_index]

        result = {'fonts': {}, 'description': []}

        # 查找Run属性标签
        rPr = run.find(f".//{{{self.NAMESPACES['w']}}}rPr")
        if rPr is None:
            result['description'] = ['未设置字体']
            return result

        # 提取字体
        rFonts = rPr.find(f".//{{{self.NAMESPACES['w']}}}rFonts")
        if rFonts is not None:
            for font_type in ['ascii', 'hAnsi', 'eastAsia', 'cs']:
                font = rFonts.get(f"{{{self.NAMESPACES['w']}}}{font_type}")
                if font:
                    result['fonts'][font_type] = font
                    font_type_name = {
                        'ascii': '英文字体',
                        'hAnsi': '西文字体',
                        'eastAsia': '中文字体',
                        'cs': '复杂文种字体'
                    }.get(font_type, font_type)
                    result['description'].append(f"{font_type_name}: {font}")

        if not result['description']:
            result['description'] = ['未设置字体']

        return result

    def get_run_size(self, element_index, run_index,element_type="elements"):
        """提取Run的字号信息

        Args:
            element_index: self.elements的索引
            run_index: 元素中w:r的索引

        Returns:
            dict: 字号信息
        """
        # 检查元素索引是否有效
        if element_index < 0 or element_index >= len(self.elements):
            print(f"错误：元素索引{element_index}超出范围(0-{len(self.elements)-1})")
            return {'size': None, 'size_pt': None, 'description': '无法获取字号信息'}

        # 获取指定索引的元素
        element = self.element_to_dict(element_index, element_type)

        # 查找所有w:r元素
        r_elements = element.findall(f".//{{{self.NAMESPACES['w']}}}r")

        # 检查Run索引是否有效
        if run_index < 0 or run_index >= len(r_elements):
            print(f"错误：Run索引{run_index}超出范围(0-{len(r_elements)-1})")
            return {'size': None, 'size_pt': None, 'description': '无法获取字号信息'}

        # 获取指定的Run元素
        run = r_elements[run_index]

        result = {'size': None, 'size_pt': None, 'description': '未设置字号'}

        # 查找Run属性标签
        rPr = run.find(f".//{{{self.NAMESPACES['w']}}}rPr")
        if rPr is None:
            return result

        # 提取字号
        sz = rPr.find(f".//{{{self.NAMESPACES['w']}}}sz")
        if sz is not None:
            size_val = sz.get(f"{{{self.NAMESPACES['w']}}}val")
            if size_val:
                result['size'] = size_val
                # Word中的字号是实际点数的两倍
                size_pt = float(size_val) / 2
                result['size_pt'] = size_pt
                result['description'] = f"字号: {size_pt}磅 ({size_val})"

        return result

    def get_run_formatting(self, element_index, run_index,element_type="elements"):
        """提取Run的格式化信息(加粗、斜体、下划线等)

        Args:
            element_index: self.elements的索引
            run_index: 元素中w:r的索引

        Returns:
            dict: 格式化信息
        """
        # 检查元素索引是否有效
        if element_index < 0 or element_index >= len(self.elements):
            print(f"错误：元素索引{element_index}超出范围(0-{len(self.elements)-1})")
            return {'formatting': {}, 'description': []}

        # 获取指定索引的元素
        element = self.element_to_dict(element_index, element_type)

        # 查找所有w:r元素
        r_elements = element.findall(f".//{{{self.NAMESPACES['w']}}}r")

        # 检查Run索引是否有效
        if run_index < 0 or run_index >= len(r_elements):
            print(f"错误：Run索引{run_index}超出范围(0-{len(r_elements)-1})")
            return {'formatting': {}, 'description': []}

        # 获取指定的Run元素
        run = r_elements[run_index]

        result = {
            'formatting': {
                'bold': False,
                'italic': False,
                'underline': None,
                'strike': False,
                'caps': False,
                'small_caps': False
            },
            'description': []
        }

        # 查找Run属性标签
        rPr = run.find(f".//{{{self.NAMESPACES['w']}}}rPr")
        if rPr is None:
            result['description'] = ['未应用文本格式']
            return result

        # 提取加粗
        b = rPr.find(f".//{{{self.NAMESPACES['w']}}}b")
        if b is not None:
            val = b.get(f"{{{self.NAMESPACES['w']}}}val", 'true')
            is_bold = val.lower() != 'false'
            result['formatting']['bold'] = is_bold
            if is_bold:
                result['description'].append('加粗')

        # 提取斜体
        i = rPr.find(f".//{{{self.NAMESPACES['w']}}}i")
        if i is not None:
            val = i.get(f"{{{self.NAMESPACES['w']}}}val", 'true')
            is_italic = val.lower() != 'false'
            result['formatting']['italic'] = is_italic
            if is_italic:
                result['description'].append('斜体')

        # 提取下划线
        u = rPr.find(f".//{{{self.NAMESPACES['w']}}}u")
        if u is not None:
            underline_val = u.get(f"{{{self.NAMESPACES['w']}}}val", 'single')
            result['formatting']['underline'] = underline_val

            underline_types = {
                'single': '单线',
                'double': '双线',
                'thick': '粗线',
                'dotted': '点线',
                'dash': '虚线',
                'dashDotDotHeavy': '重点划线',
                'wave': '波浪线'
            }

            underline_desc = underline_types.get(underline_val, underline_val)
            result['description'].append(f'下划线({underline_desc})')

        # 提取删除线
        strike = rPr.find(f".//{{{self.NAMESPACES['w']}}}strike")
        if strike is not None:
            val = strike.get(f"{{{self.NAMESPACES['w']}}}val", 'true')
            is_strike = val.lower() != 'false'
            result['formatting']['strike'] = is_strike
            if is_strike:
                result['description'].append('删除线')

        # 提取大写
        caps = rPr.find(f".//{{{self.NAMESPACES['w']}}}caps")
        if caps is not None:
            val = caps.get(f"{{{self.NAMESPACES['w']}}}val", 'true')
            is_caps = val.lower() != 'false'
            result['formatting']['caps'] = is_caps
            if is_caps:
                result['description'].append('全大写')

        # 提取小型大写
        smallCaps = rPr.find(f".//{{{self.NAMESPACES['w']}}}smallCaps")
        if smallCaps is not None:
            val = smallCaps.get(f"{{{self.NAMESPACES['w']}}}val", 'true')
            is_small_caps = val.lower() != 'false'
            result['formatting']['small_caps'] = is_small_caps
            if is_small_caps:
                result['description'].append('小型大写')

        if not result['description']:
            result['description'] = ['常规格式(无特殊格式)']

        return result

    def get_run_color(self, element_index, run_index,element_type="elements"):
        """提取Run的颜色信息

        Args:
            element_index: self.elements的索引
            run_index: 元素中w:r的索引

        Returns:
            dict: 颜色信息
        """
        # 检查元素索引是否有效
        if element_index < 0 or element_index >= len(self.elements):
            print(f"错误：元素索引{element_index}超出范围(0-{len(self.elements)-1})")
            return {'color': None, 'highlight': None, 'description': '无法获取颜色信息'}

        # 获取指定索引的元素
        element = self.element_to_dict(element_index, element_type)

        # 查找所有w:r元素
        r_elements = element.findall(f".//{{{self.NAMESPACES['w']}}}r")

        # 检查Run索引是否有效
        if run_index < 0 or run_index >= len(r_elements):
            print(f"错误：Run索引{run_index}超出范围(0-{len(r_elements)-1})")
            return {'color': None, 'highlight': None, 'description': '无法获取颜色信息'}

        # 获取指定的Run元素
        run = r_elements[run_index]

        result = {'color': None, 'highlight': None, 'description': []}

        # 查找Run属性标签
        rPr = run.find(f".//{{{self.NAMESPACES['w']}}}rPr")
        if rPr is None:
            result['description'] = ['未设置颜色']
            return result

        # 提取文字颜色
        color = rPr.find(f".//{{{self.NAMESPACES['w']}}}color")
        if color is not None:
            color_val = color.get(f"{{{self.NAMESPACES['w']}}}val")
            result['color'] = color_val
            result['description'].append(f'文字颜色: {color_val}')

        # 提取突出显示颜色
        highlight = rPr.find(f".//{{{self.NAMESPACES['w']}}}highlight")
        if highlight is not None:
            highlight_val = highlight.get(f"{{{self.NAMESPACES['w']}}}val")
            result['highlight'] = highlight_val

            highlight_colors = {
                'yellow': '黄色',
                'green': '绿色',
                'cyan': '青色',
                'magenta': '洋红',
                'blue': '蓝色',
                'red': '红色',
                'darkBlue': '深蓝色',
                'darkCyan': '深青色',
                'darkGreen': '深绿色',
                'darkMagenta': '深洋红色',
                'darkRed': '深红色',
                'darkYellow': '深黄色',
                'darkGray': '深灰色',
                'lightGray': '浅灰色',
                'black': '黑色'
            }

            highlight_desc = highlight_colors.get(highlight_val, highlight_val)
            result['description'].append(f'突出显示: {highlight_desc}')

        if not result['description']:
            result['description'] = ['未设置颜色']

        return result

    def format_run_style(self, style_info):
        """将Run样式信息格式化为易读的字符串

        Args:
            style_info: get_run_style返回的样式信息字典

        Returns:
            str: 格式化后的样式信息字符串
        """
        if not style_info or style_info.get('has_style') is False:
            return "Run无样式信息"

        lines = []
        lines.append("Run样式信息:")

        # 格式化字体信息
        if style_info['fonts']:
            lines.append("- 字体:")
            for font_type, font in style_info['fonts'].items():
                font_type_name = {
                    'ascii': '英文字体',
                    'hAnsi': '西文字体',
                    'eastAsia': '中文字体',
                    'cs': '复杂文种字体'
                }.get(font_type, font_type)
                lines.append(f"  • {font_type_name}: {font}")

        # 格式化字号
        if style_info['size']:
            size_pt = float(style_info['size']) / 2
            lines.append(f"- 字号: {size_pt}磅 ({style_info['size']})")

        # 格式化文本格式
        format_items = []
        if style_info['bold']:
            format_items.append("加粗")
        if style_info['italic']:
            format_items.append("斜体")
        if style_info['underline']:
            underline_types = {
                'single': '单线下划线',
                'double': '双线下划线',
                'thick': '粗线下划线',
                'dotted': '点线下划线',
                'dash': '虚线下划线',
                'wave': '波浪线下划线'
            }
            underline_desc = underline_types.get(style_info['underline'], style_info['underline'])
            format_items.append(underline_desc)
        if style_info['strike']:
            format_items.append("删除线")
        if style_info['caps']:
            format_items.append("全大写")
        if style_info['small_caps']:
            format_items.append("小型大写字母")

        if format_items:
            lines.append("- 文本格式: " + ", ".join(format_items))

        # 格式化颜色信息
        if style_info['color']:
            lines.append(f"- 文字颜色: {style_info['color']}")

        if style_info['highlight']:
            highlight_colors = {
                'yellow': '黄色',
                'green': '绿色',
                'cyan': '青色',
                'magenta': '洋红',
                'blue': '蓝色',
                'red': '红色',
                'darkBlue': '深蓝色',
                'darkCyan': '深青色',
                'darkGreen': '深绿色',
                'darkMagenta': '深洋红色',
                'darkRed': '深红色',
                'darkYellow': '深黄色',
                'darkGray': '深灰色',
                'lightGray': '浅灰色',
                'black': '黑色'
            }
            highlight_desc = highlight_colors.get(style_info['highlight'], style_info['highlight'])
            lines.append(f"- 突出显示: {highlight_desc}")

        # 格式化其他特殊属性
        if style_info['spacing']:
            spacing_pt = float(style_info['spacing']) / 20
            lines.append(f"- 字符间距: {spacing_pt}磅")

        if style_info['vert_align']:
            vert_align_types = {
                'superscript': '上标',
                'subscript': '下标',
                'baseline': '基线'
            }
            vert_align_desc = vert_align_types.get(style_info['vert_align'], style_info['vert_align'])
            lines.append(f"- 垂直对齐: {vert_align_desc}")

        # 格式化其他属性
        if style_info['other_properties']:
            lines.append("- 其他属性:")
            for prop, value in style_info['other_properties'].items():
                if isinstance(value, dict):
                    attrs = [f"{k}={v}" for k, v in value.items()]
                    lines.append(f"  • {prop}: {', '.join(attrs)}")
                else:
                    lines.append(f"  • {prop}: {value}")

        return "\n".join(lines)

    def get_table_style(self, table_index,type="elements"):
        """提取表格的所有样式和属性信息，包括表格级、行级和单元格级样式

        Args:
            table_index: self.tables中的表格索引

        Returns:
            dict: 包含表格样式和属性信息的详细字典
        """
        # 获取表格元素
        # 获取表格元素
        if type == "elements":
            table = self.elements[table_index]['element']  # 转换为Element对象
        else:
            table = self.tables[table_index]['element']

        # 创建结果字典
        style_info = {
            'style_id': None,
            'style_name': None,
            'width': {'value': None, 'type': None},
            'indent': {'value': None, 'type': None},
            'borders': {
                'top': {},
                'left': {},
                'bottom': {},
                'right': {},
                'inside_h': {},
                'inside_v': {}
            },
            'shading': {
                'val': None,
                'color': None,
                'fill': None,
                'pattern': None
            },
            'layout': None,
            'alignment': None,
            'cell_margins': {
                'top': {},
                'left': {},
                'bottom': {},
                'right': {}
            },
            'look': None,
            'grid': [],
            'rows_count': 0,
            'columns_count': 0,
            'rows': [],  # 将存储每行的样式信息
            'cells': {},  # 将存储单元格样式信息，格式为 {(row_idx, col_idx): cell_info}
            'is_header_row': False,  # 表示第一行是否为标题行
            'caption': None,  # 表格标题
            'table_properties': {},  # 其他表格属性
            'description': [],
            'is_three_line_table': False
        }

        # 查找表格属性
        tblPr = table.find(f".//{{{self.NAMESPACES['w']}}}tblPr")
        if tblPr is not None:
            # 提取样式ID和名称
            style = tblPr.find(f".//{{{self.NAMESPACES['w']}}}tblStyle")
            if style is not None:
                style_info['style_id'] = style.get(f"{{{self.NAMESPACES['w']}}}val")

                # 尝试获取样式名称（如果存在样式表）
                try:
                    if hasattr(self, 'styles') and self.styles:
                        for style_def in self.styles.findall(f".//{{{self.NAMESPACES['w']}}}style"):
                            if style_def.get(f"{{{self.NAMESPACES['w']}}}styleId") == style_info['style_id']:
                                name = style_def.find(f".//{{{self.NAMESPACES['w']}}}name")
                                if name is not None:
                                    style_info['style_name'] = name.get(f"{{{self.NAMESPACES['w']}}}val")
                except Exception as e:
                    print(f"获取表格样式名称时出错: {str(e)}")

            # 提取表格宽度
            tblW = tblPr.find(f".//{{{self.NAMESPACES['w']}}}tblW")
            if tblW is not None:
                style_info['width']['value'] = tblW.get(f"{{{self.NAMESPACES['w']}}}w")
                style_info['width']['type'] = tblW.get(f"{{{self.NAMESPACES['w']}}}type")

            # 提取表格缩进
            tblInd = tblPr.find(f".//{{{self.NAMESPACES['w']}}}tblInd")
            if tblInd is not None:
                style_info['indent']['value'] = tblInd.get(f"{{{self.NAMESPACES['w']}}}w")
                style_info['indent']['type'] = tblInd.get(f"{{{self.NAMESPACES['w']}}}type")

            # 提取表格边框
            tblBorders = tblPr.find(f".//{{{self.NAMESPACES['w']}}}tblBorders")
            if tblBorders is not None:
                for border_type, border_key in [
                    ('top', 'top'),
                    ('left', 'left'),
                    ('bottom', 'bottom'),
                    ('right', 'right'),
                    ('insideH', 'inside_h'),
                    ('insideV', 'inside_v')
                ]:
                    border = tblBorders.find(f".//{{{self.NAMESPACES['w']}}}{border_type}")
                    if border is not None:
                        style_info['borders'][border_key] = {
                            'val': border.get(f"{{{self.NAMESPACES['w']}}}val"),
                            'color': border.get(f"{{{self.NAMESPACES['w']}}}color"),
                            'size': border.get(f"{{{self.NAMESPACES['w']}}}sz"),
                            'space': border.get(f"{{{self.NAMESPACES['w']}}}space"),
                            'shadow': border.get(f"{{{self.NAMESPACES['w']}}}shadow"),
                            'frame': border.get(f"{{{self.NAMESPACES['w']}}}frame")
                        }

            # 提取表格底纹
            shd = tblPr.find(f".//{{{self.NAMESPACES['w']}}}shd")
            if shd is not None:
                style_info['shading']['val'] = shd.get(f"{{{self.NAMESPACES['w']}}}val")
                style_info['shading']['color'] = shd.get(f"{{{self.NAMESPACES['w']}}}color")
                style_info['shading']['fill'] = shd.get(f"{{{self.NAMESPACES['w']}}}fill")
                style_info['shading']['pattern'] = shd.get(f"{{{self.NAMESPACES['w']}}}pattern")

            # 提取表格布局
            tblLayout = tblPr.find(f".//{{{self.NAMESPACES['w']}}}tblLayout")
            if tblLayout is not None:
                style_info['layout'] = tblLayout.get(f"{{{self.NAMESPACES['w']}}}type")

            # 提取表格对齐方式
            jc = tblPr.find(f".//{{{self.NAMESPACES['w']}}}jc")
            if jc is not None:
                style_info['alignment'] = jc.get(f"{{{self.NAMESPACES['w']}}}val")

            # 提取单元格边距
            tblCellMar = tblPr.find(f".//{{{self.NAMESPACES['w']}}}tblCellMar")
            if tblCellMar is not None:
                for margin_type in ['top', 'left', 'bottom', 'right']:
                    margin = tblCellMar.find(f".//{{{self.NAMESPACES['w']}}}{margin_type}")
                    if margin is not None:
                        style_info['cell_margins'][margin_type] = {
                            'value': margin.get(f"{{{self.NAMESPACES['w']}}}w"),
                            'type': margin.get(f"{{{self.NAMESPACES['w']}}}type")
                        }

            # 提取表格Look属性（控制表格格式应用方式）
            tblLook = tblPr.find(f".//{{{self.NAMESPACES['w']}}}tblLook")
            if tblLook is not None:
                look_val = tblLook.get(f"{{{self.NAMESPACES['w']}}}val")
                style_info['look'] = {
                    'val': look_val,
                    'first_row': look_val and (int(look_val, 16) & 0x0020) != 0,
                    'last_row': look_val and (int(look_val, 16) & 0x0040) != 0,
                    'first_column': look_val and (int(look_val, 16) & 0x0080) != 0,
                    'last_column': look_val and (int(look_val, 16) & 0x0100) != 0,
                    'no_hband': look_val and (int(look_val, 16) & 0x0200) != 0,
                    'no_vband': look_val and (int(look_val, 16) & 0x0400) != 0
                }
                # 如果first_row为True，表示第一行是标题行
                style_info['is_header_row'] = style_info['look']['first_row']

            # 提取其他表格属性
            for prop in tblPr:
                if prop.tag.endswith('}tblCaption'):
                    style_info['caption'] = prop.get(f"{{{self.NAMESPACES['w']}}}val")
                elif prop.tag.endswith('}tblDescription'):
                    style_info['table_properties']['description'] = prop.get(f"{{{self.NAMESPACES['w']}}}val")
                elif prop.tag.endswith('}tblOverlap'):
                    style_info['table_properties']['overlap'] = prop.get(f"{{{self.NAMESPACES['w']}}}val")

        # 提取表格网格（列定义）
        tblGrid = table.find(f".//{{{self.NAMESPACES['w']}}}tblGrid")
        if tblGrid is not None:
            grid_cols = tblGrid.findall(f".//{{{self.NAMESPACES['w']}}}gridCol")
            for col in grid_cols:
                col_width = col.get(f"{{{self.NAMESPACES['w']}}}w")
                style_info['grid'].append(col_width)

            style_info['columns_count'] = len(grid_cols)

        # 处理行和单元格
        rows = table.findall(f".//{{{self.NAMESPACES['w']}}}tr")
        style_info['rows_count'] = len(rows)

        for row_idx, row in enumerate(rows):
            row_info = {
                'height': {'value': None, 'rule': None},
                'is_header': row_idx == 0 and style_info['is_header_row'],
                'borders': {},
                'shading': {},
                'properties': {}
            }

            # 提取行属性
            trPr = row.find(f".//{{{self.NAMESPACES['w']}}}trPr")
            if trPr is not None:
                # 行高
                trHeight = trPr.find(f".//{{{self.NAMESPACES['w']}}}trHeight")
                if trHeight is not None:
                    row_info['height']['value'] = trHeight.get(f"{{{self.NAMESPACES['w']}}}val")
                    row_info['height']['rule'] = trHeight.get(f"{{{self.NAMESPACES['w']}}}hRule")

                # 检查行是否为标题行
                tblHeader = trPr.find(f".//{{{self.NAMESPACES['w']}}}tblHeader")
                if tblHeader is not None:
                    row_info['is_header'] = True
                    style_info['is_header_row'] = True

                # 检查行级别边框继承
                row_borders = row.find(f".//{{{self.NAMESPACES['w']}}}tblPrEx/{{{self.NAMESPACES['w']}}}tblBorders")
                if row_borders is not None:
                    for border_type, border_key in [
                        ('top', 'top'),
                        ('left', 'left'),
                        ('bottom', 'bottom'),
                        ('right', 'right'),
                        ('insideH', 'inside_h'),
                        ('insideV', 'inside_v')
                    ]:
                        border = row_borders.find(f".//{{{self.NAMESPACES['w']}}}{border_type}")
                        if border is not None:
                            row_info['borders'][border_key] = {
                                'val': border.get(f"{{{self.NAMESPACES['w']}}}val"),
                                'color': border.get(f"{{{self.NAMESPACES['w']}}}color"),
                                'size': border.get(f"{{{self.NAMESPACES['w']}}}sz"),
                                'space': border.get(f"{{{self.NAMESPACES['w']}}}space")
                            }

            # 处理单元格
            cells = row.findall(f".//{{{self.NAMESPACES['w']}}}tc")
            for cell_idx, cell in enumerate(cells):
                cell_key = (row_idx, cell_idx)
                cell_info = {
                    'width': {'value': None, 'type': None},
                    'rowspan': 1,
                    'colspan': 1,
                    'borders': {},
                    'shading': {},
                    'vertical_align': None,
                    'text_direction': None,
                    'margins': {},
                    'properties': {}
                }

                # 提取单元格属性
                tcPr = cell.find(f".//{{{self.NAMESPACES['w']}}}tcPr")
                if tcPr is not None:
                    # 单元格宽度
                    tcW = tcPr.find(f".//{{{self.NAMESPACES['w']}}}tcW")
                    if tcW is not None:
                        cell_info['width']['value'] = tcW.get(f"{{{self.NAMESPACES['w']}}}w")
                        cell_info['width']['type'] = tcW.get(f"{{{self.NAMESPACES['w']}}}type")

                    # 合并单元格
                    gridSpan = tcPr.find(f".//{{{self.NAMESPACES['w']}}}gridSpan")
                    if gridSpan is not None:
                        cell_info['colspan'] = int(gridSpan.get(f"{{{self.NAMESPACES['w']}}}val", "1"))

                    vMerge = tcPr.find(f".//{{{self.NAMESPACES['w']}}}vMerge")
                    if vMerge is not None:
                        val = vMerge.get(f"{{{self.NAMESPACES['w']}}}val")
                        # 如果val为"restart"，则是合并起始单元格
                        # 如果val不存在或为"continue"，则是被合并单元格
                        cell_info['properties']['vMerge'] = val if val else "continue"
                        if val == "restart":
                            # 标记为行合并的起始单元格
                            cell_info['rowspan'] = 2  # 默认值，实际值需要后处理

                    # 单元格边框
                    tcBorders = tcPr.find(f".//{{{self.NAMESPACES['w']}}}tcBorders")
                    if tcBorders is not None:
                        for border_type, border_key in [
                            ('top', 'top'),
                            ('left', 'left'),
                            ('bottom', 'bottom'),
                            ('right', 'right'),
                            ('insideH', 'inside_h'),
                            ('insideV', 'inside_v')
                        ]:
                            border = tcBorders.find(f".//{{{self.NAMESPACES['w']}}}{border_type}")
                            if border is not None:
                                cell_info['borders'][border_key] = {
                                    'val': border.get(f"{{{self.NAMESPACES['w']}}}val"),
                                    'color': border.get(f"{{{self.NAMESPACES['w']}}}color"),
                                    'size': border.get(f"{{{self.NAMESPACES['w']}}}sz"),
                                    'space': border.get(f"{{{self.NAMESPACES['w']}}}space")
                                }

                    # 单元格底纹
                    shd = tcPr.find(f".//{{{self.NAMESPACES['w']}}}shd")
                    if shd is not None:
                        cell_info['shading'] = {
                            'val': shd.get(f"{{{self.NAMESPACES['w']}}}val"),
                            'color': shd.get(f"{{{self.NAMESPACES['w']}}}color"),
                            'fill': shd.get(f"{{{self.NAMESPACES['w']}}}fill"),
                            'pattern': shd.get(f"{{{self.NAMESPACES['w']}}}pattern")
                        }

                    # 垂直对齐方式
                    vAlign = tcPr.find(f".//{{{self.NAMESPACES['w']}}}vAlign")
                    if vAlign is not None:
                        cell_info['vertical_align'] = vAlign.get(f"{{{self.NAMESPACES['w']}}}val")

                    # 文本方向
                    textDirection = tcPr.find(f".//{{{self.NAMESPACES['w']}}}textDirection")
                    if textDirection is not None:
                        cell_info['text_direction'] = textDirection.get(f"{{{self.NAMESPACES['w']}}}val")

                    # 单元格边距
                    tcMar = tcPr.find(f".//{{{self.NAMESPACES['w']}}}tcMar")
                    if tcMar is not None:
                        for margin_type in ['top', 'left', 'bottom', 'right']:
                            margin = tcMar.find(f".//{{{self.NAMESPACES['w']}}}{margin_type}")
                            if margin is not None:
                                cell_info['margins'][margin_type] = {
                                    'value': margin.get(f"{{{self.NAMESPACES['w']}}}w"),
                                    'type': margin.get(f"{{{self.NAMESPACES['w']}}}type")
                                }

                    # 其他单元格属性
                    noWrap = tcPr.find(f".//{{{self.NAMESPACES['w']}}}noWrap")
                    if noWrap is not None:
                        cell_info['properties']['nowrap'] = True

                    hideMark = tcPr.find(f".//{{{self.NAMESPACES['w']}}}hideMark")
                    if hideMark is not None:
                        cell_info['properties']['hidemark'] = True

                # 计算单元格包含的段落数量（但不提取内容）
                paragraphs = cell.findall(f".//{{{self.NAMESPACES['w']}}}p")
                cell_info['paragraph_count'] = len(paragraphs)

                style_info['cells'][cell_key] = cell_info

            style_info['rows'].append(row_info)

        # 处理垂直合并单元格的rowspan值
        # 查找vMerge="restart"的单元格，计算其rowspan
        for row_idx in range(style_info['rows_count']):
            for col_idx in range(style_info['columns_count']):
                cell_key = (row_idx, col_idx)
                if cell_key in style_info['cells']:
                    cell = style_info['cells'][cell_key]
                    if cell['properties'].get('vMerge') == 'restart':
                        rowspan = 1
                        for next_row in range(row_idx + 1, style_info['rows_count']):
                            next_cell_key = (next_row, col_idx)
                            if (next_cell_key in style_info['cells'] and
                                    style_info['cells'][next_cell_key]['properties'].get('vMerge') == 'continue'):
                                rowspan += 1
                            else:
                                break
                        cell['rowspan'] = rowspan

        # 检查是否为三线表
        try:
            # 1. 表头上方有线
            header_top_exists = style_info['borders']['top'].get('val') not in [None, 'none']
            if not header_top_exists and style_info['rows_count'] > 0:
                for cell_key in style_info['cells']:
                    if cell_key[0] == 0:  # 第一行
                        if style_info['cells'][cell_key]['borders'].get('top', {}).get('val') not in [None, 'none']:
                            header_top_exists = True
                            break

            # 2. 表头下方有线
            header_bottom_exists = False
            if style_info['rows_count'] > 1:
                for cell_key in style_info['cells']:
                    if cell_key[0] == 0:  # 第一行
                        if style_info['cells'][cell_key]['borders'].get('bottom', {}).get('val') not in [None, 'none']:
                            header_bottom_exists = True
                            break

            # 3. 表格底部有线
            table_bottom_exists = style_info['borders']['bottom'].get('val') not in [None, 'none']
            if not table_bottom_exists and style_info['rows_count'] > 0:
                for cell_key in style_info['cells']:
                    if cell_key[0] == style_info['rows_count'] - 1:  # 最后一行
                        if style_info['cells'][cell_key]['borders'].get('bottom', {}).get('val') not in [None, 'none']:
                            table_bottom_exists = True
                            break

            # 4. 无内部水平线
            no_inner_h_lines = style_info['borders']['inside_h'].get('val') in [None, 'none']

            # 5. 无垂直线
            no_vertical_lines = style_info['borders']['inside_v'].get('val') in [None, 'none']

            # 判断是否为三线表
            style_info['is_three_line_table'] = (
                    header_top_exists and
                    header_bottom_exists and
                    table_bottom_exists and
                    no_inner_h_lines and
                    no_vertical_lines
            )

            if style_info['is_three_line_table']:
                style_info['description'].append("符合三线表标准")
        except Exception as e:
            print(f"检查三线表特征时出错: {str(e)}")

        # 格式化描述信息
        style_info['description'].append(f"表格大小: {style_info['rows_count']}行 × {style_info['columns_count']}列")

        if style_info['style_id']:
            style_desc = f"样式ID: {style_info['style_id']}"
            if style_info['style_name']:
                style_desc += f" ({style_info['style_name']})"
            style_info['description'].append(style_desc)

        # 边框描述
        borders_desc = []
        for border_name, border_key in [
            ('上边框', 'top'),
            ('左边框', 'left'),
            ('下边框', 'bottom'),
            ('右边框', 'right'),
            ('水平内边框', 'inside_h'),
            ('垂直内边框', 'inside_v')
        ]:
            border = style_info['borders'][border_key]
            if border and border.get('val'):
                border_type = {
                    'single': '单线',
                    'double': '双线',
                    'thick': '粗线',
                    'none': '无',
                    'thin': '细线',
                    'dotted': '点线',
                    'dashed': '虚线',
                    'dashSmallGap': '短划线',
                    'dotDash': '点划线',
                    'dotDotDash': '点点划线',
                    'triple': '三线',
                    'thinThickSmallGap': '细粗线(小间隔)',
                    'thickThinSmallGap': '粗细线(小间隔)',
                    'thinThickThinSmallGap': '细粗细线(小间隔)',
                    'thinThickMediumGap': '细粗线(中间隔)',
                    'thickThinMediumGap': '粗细线(中间隔)',
                    'thinThickThinMediumGap': '细粗细线(中间隔)',
                    'thinThickLargeGap': '细粗线(大间隔)',
                    'thickThinLargeGap': '粗细线(大间隔)',
                    'thinThickThinLargeGap': '细粗细线(大间隔)',
                    'wave': '波浪线',
                    'doubleWave': '双波浪线',
                    'dashDotStroked': '实心点划线',
                    'threeDEmboss': '3D浮雕',
                    'threeDEngrave': '3D刻线',
                    'outset': '外凸',
                    'inset': '内凹'
                }.get(border.get('val'), border.get('val'))

                if border_type != '无':
                    border_size = f"{float(border.get('size', '1')) / 8:.1f}磅" if border.get('size') else ""
                    border_color = border.get('color', 'auto')
                    borders_desc.append(f"{border_name}: {border_type} {border_size} {border_color}")

        if borders_desc:
            style_info['description'].append("边框: " + ", ".join(borders_desc))

        # 底纹描述
        if style_info['shading'].get('fill'):
            fill_color = style_info['shading']['fill']
            if fill_color != 'auto' and fill_color != '000000':
                style_info['description'].append(f"底纹颜色: {fill_color}")

        # 布局描述
        if style_info['layout']:
            layout_desc = {
                'autofit': '自动适应内容',
                'fixed': '固定宽度'
            }.get(style_info['layout'], style_info['layout'])
            style_info['description'].append(f"布局: {layout_desc}")

        # 对齐方式描述
        if style_info['alignment']:
            align_desc = {
                'left': '左对齐',
                'center': '居中',
                'right': '右对齐'
            }.get(style_info['alignment'], style_info['alignment'])
            style_info['description'].append(f"表格对齐: {align_desc}")

        # 列宽描述
        if style_info['grid']:
            col_widths = []
            for i, width in enumerate(style_info['grid']):
                if width:
                    # 转换为磅
                    pt_width = float(width) / 20
                    col_widths.append(f"列{i + 1}: {pt_width:.1f}磅")
            style_info['description'].append("列宽: " + ", ".join(col_widths))

        # 单元格合并描述
        merged_cells = []
        for (row_idx, col_idx), cell in style_info['cells'].items():
            if cell['colspan'] > 1 or cell['rowspan'] > 1:
                merged_cells.append(f"单元格({row_idx + 1},{col_idx + 1}): {cell['rowspan']}行 × {cell['colspan']}列")

        if merged_cells:
            style_info['description'].append("合并单元格: " + ", ".join(merged_cells))

        return style_info
    def format_table_style(self, style_info):
        """将表格样式信息格式化为易读的字符串

        Args:
            style_info: get_table_style返回的样式信息字典

        Returns:
            str: 格式化后的样式信息字符串
        """
        if not style_info:
            return "无法获取表格样式信息"

        lines = []
        lines.append("表格样式信息:")

        # 基本信息
        lines.append(f"- 大小: {style_info['rows_count']}行 × {style_info['columns_count']}列")

        if style_info['style_id']:
            lines.append(f"- 样式ID: {style_info['style_id']}")

        # 宽度和缩进
        width_type_map = {
            'auto': '自动适应',
            'dxa': '绝对值',
            'pct': '百分比'
        }

        if style_info['width']['value']:
            width_type = width_type_map.get(style_info['width']['type'], style_info['width']['type'])
            if style_info['width']['type'] == 'pct':
                value = f"{float(style_info['width']['value']) / 50:.1f}%"
            else:
                value = f"{float(style_info['width']['value']) / 20:.1f}磅"
            lines.append(f"- 宽度: {value} ({width_type})")

        if style_info['indent']['value']:
            indent_type = width_type_map.get(style_info['indent']['type'], style_info['indent']['type'])
            indent_pt = float(style_info['indent']['value']) / 20
            lines.append(f"- 缩进: {indent_pt:.1f}磅 ({indent_type})")

        # 边框信息
        lines.append("- 边框:")
        borders_added = False
        for border_name, border_key in [
            ('上边框', 'top'),
            ('左边框', 'left'),
            ('下边框', 'bottom'),
            ('右边框', 'right'),
            ('水平内边框', 'inside_h'),
            ('垂直内边框', 'inside_v')
        ]:
            border = style_info['borders'][border_key]
            if border and border.get('val'):
                border_type = {
                    'single': '单线',
                    'double': '双线',
                    'thick': '粗线',
                    'none': '无'
                }.get(border.get('val'), border.get('val'))

                if border_type != '无':
                    border_size = f"{float(border.get('size', '1')) / 8:.1f}磅" if border.get('size') else ""
                    border_color = border.get('color', 'auto')
                    lines.append(f"  • {border_name}: {border_type} {border_size} {border_color}")
                    borders_added = True

        if not borders_added:
            lines.append("  • 无边框")

        # 布局信息
        if style_info['layout']:
            layout_desc = {
                'autofit': '自动适应内容',
                'fixed': '固定宽度'
            }.get(style_info['layout'], style_info['layout'])
            lines.append(f"- 布局方式: {layout_desc}")

        # 单元格边距
        lines.append("- 单元格边距:")
        margins_added = False
        for margin_name, margin_key in [
            ('上边距', 'top'),
            ('左边距', 'left'),
            ('下边距', 'bottom'),
            ('右边距', 'right')
        ]:
            margin = style_info['cell_margins'][margin_key]
            if margin and margin.get('value'):
                margin_pt = float(margin['value']) / 20
                lines.append(f"  • {margin_name}: {margin_pt:.1f}磅")
                margins_added = True

        if not margins_added:
            lines.append("  • 未设置边距")

        # 列宽信息
        if style_info['grid']:
            lines.append("- 列宽:")
            for i, width in enumerate(style_info['grid']):
                pt_width = float(width) / 20
                lines.append(f"  • 第{i+1}列: {pt_width:.1f}磅")

        return "\n".join(lines)

    # 以下是修改段落样式的函数

    def _get_or_create_pPr(self, paragraph_element):
        """获取或创建段落属性标签

        Args:
            paragraph_element: 段落XML元素对象

        Returns:
            ElementTree.Element: pPr元素
        """
        # 查找段落属性标签
        pPr = paragraph_element.find(f".//{{{self.NAMESPACES['w']}}}pPr")
        if pPr is None:
            # 如果不存在，则创建
            pPr = ET.Element(f"{{{self.NAMESPACES['w']}}}pPr")
            paragraph_element.insert(0, pPr)
        return pPr

    def set_paragraph_style_id(self, para_index, style_id):
        """设置段落样式ID

        Args:
            para_index: 段落索引
            style_id: 要设置的样式ID

        Returns:
            bool: 是否成功修改
        """
        # 检查索引是否有效
        if para_index < 0 or para_index >= len(self.paragraphs):
            print(f"错误：段落索引{para_index}超出范围(0-{len(self.paragraphs)-1})")
            return False

        try:
            # 获取段落元素
            paragraph = self.paragraphs[para_index]['element']

            # 获取或创建pPr元素
            pPr = self._get_or_create_pPr(paragraph)

            # 查找样式元素
            pStyle = pPr.find(f".//{{{self.NAMESPACES['w']}}}pStyle")
            if pStyle is None:
                # 如果不存在，则创建
                pStyle = ET.Element(f"{{{self.NAMESPACES['w']}}}pStyle")
                pPr.append(pStyle)

            # 设置样式ID
            pStyle.set(f"{{{self.NAMESPACES['w']}}}val", style_id)
            return True
        except Exception as e:
            print(f"设置段落样式ID时出错: {e}")
            return False

    def set_paragraph_style_id_from_xml(self, para_index, style_id):
        """设置段落样式ID

        Args:
            para_index: 段落索引
            style_id: 要设置的样式ID

        Returns:
            bool: 是否成功修改
        """

        try:
            # 获取段落元素
            paragraph =para_index

            # 获取或创建pPr元素
            pPr = self._get_or_create_pPr(paragraph)

            # 查找样式元素
            pStyle = pPr.find(f".//{{{self.NAMESPACES['w']}}}pStyle")
            if pStyle is None:
                # 如果不存在，则创建
                pStyle = ET.Element(f"{{{self.NAMESPACES['w']}}}pStyle")
                pPr.append(pStyle)

            # 设置样式ID
            pStyle.set(f"{{{self.NAMESPACES['w']}}}val", style_id)
            return True
        except Exception as e:
            print(f"设置段落样式ID时出错: {e}")
            return False
    def set_paragraph_alignment(self, para_index, alignment):
        """设置段落对齐方式

        Args:
            para_index: 段落索引
            alignment: 对齐方式 (left, right, center, both, distribute)

        Returns:
            bool: 是否成功修改
        """
        # 检查索引是否有效
        if para_index < 0 or para_index >= len(self.paragraphs):
            print(f"错误：段落索引{para_index}超出范围(0-{len(self.paragraphs)-1})")
            return False

        try:
            # 获取段落元素
            paragraph = self.paragraphs[para_index]['element']

            # 获取或创建pPr元素
            pPr = self._get_or_create_pPr(paragraph)

            # 查找对齐方式元素
            jc = pPr.find(f".//{{{self.NAMESPACES['w']}}}jc")
            if jc is None:
                # 如果不存在，则创建
                jc = ET.Element(f"{{{self.NAMESPACES['w']}}}jc")
                pPr.append(jc)

            # 设置对齐方式
            jc.set(f"{{{self.NAMESPACES['w']}}}val", alignment)
            return True
        except Exception as e:
            print(f"设置段落对齐方式时出错: {e}")
            return False

    def set_paragraph_alignment_from_xml(self, para_index, alignment):
        """设置段落对齐方式

        Args:
            para_index: 段落
            alignment: 对齐方式 (left, right, center, both, distribute)

        Returns:
            bool: 是否成功修改
        """


        try:
            # 获取段落元素
            paragraph = para_index

            # 获取或创建pPr元素
            pPr = self._get_or_create_pPr(paragraph)

            # 查找对齐方式元素
            jc = pPr.find(f".//{{{self.NAMESPACES['w']}}}jc")
            if jc is None:
                # 如果不存在，则创建
                jc = ET.Element(f"{{{self.NAMESPACES['w']}}}jc")
                pPr.append(jc)

            # 设置对齐方式
            jc.set(f"{{{self.NAMESPACES['w']}}}val", alignment)
            return True
        except Exception as e:
            print(f"设置段落对齐方式时出错: {e}")
            return False

    def set_paragraph_indentation_from_xml(self, para_index, **indentation):
        """设置段落缩进

        Args:
            para_index: 段落元素
            **indentation: 缩进设置，可包含以下参数:
                left: 左缩进
                right: 右缩进
                firstLine: 首行缩进
                hanging: 悬挂缩进

        Returns:
            bool: 是否成功修改
        """


        try:
            # 获取段落元素
            paragraph = para_index

            # 获取或创建pPr元素
            pPr = self._get_or_create_pPr(paragraph)

            # 查找缩进元素
            ind = pPr.find(f".//{{{self.NAMESPACES['w']}}}ind")
            if ind is None:
                # 如果不存在，则创建
                ind = ET.Element(f"{{{self.NAMESPACES['w']}}}ind")
                pPr.append(ind)

            # 设置各类缩进
            valid_props = ['left', 'right', 'firstLine', 'hanging']
            for prop, value in indentation.items():
                if prop in valid_props and value is not None:
                    ind.set(f"{{{self.NAMESPACES['w']}}}{prop}", str(value))

            return True
        except Exception as e:
            print(f"设置段落缩进时出错: {e}")
            return False

    def set_paragraph_indentation(self, para_index, **indentation):
        """设置段落缩进

        Args:
            para_index: 段落索引
            **indentation: 缩进设置，可包含以下参数:
                left: 左缩进
                right: 右缩进
                firstLine: 首行缩进
                hanging: 悬挂缩进

        Returns:
            bool: 是否成功修改
        """

        # 检查索引是否有效
        if para_index < 0 or para_index >= len(self.paragraphs):
            print(f"错误：段落索引{para_index}超出范围(0-{len(self.paragraphs) - 1})")
            return False

        try:
            # 获取段落元素
            paragraph = self.paragraphs[para_index]['element']

            # 获取或创建pPr元素
            pPr = self._get_or_create_pPr(paragraph)

            # 查找缩进元素
            ind = pPr.find(f".//{{{self.NAMESPACES['w']}}}ind")
            if ind is None:
                # 如果不存在，则创建
                ind = ET.Element(f"{{{self.NAMESPACES['w']}}}ind")
                pPr.append(ind)

            # 设置各类缩进
            valid_props = ['left', 'right', 'firstLine', 'hanging']
            for prop, value in indentation.items():
                if prop in valid_props and value is not None:
                    ind.set(f"{{{self.NAMESPACES['w']}}}{prop}", str(value))

            return True
        except Exception as e:
            print(f"设置段落缩进时出错: {e}")
            return False
    def set_paragraph_spacing(self, para_index, **spacing):
        """设置段落间距属性

        Args:
            para_index: 段落索引
            **spacing: 间距设置，可以包含以下键：
                before: 段前间距（磅值的1/20，如120表示6磅）
                after: 段后间距（磅值的1/20，如120表示6磅）
                beforeLines: 段前间距（行数，如150表示1.5行）
                afterLines: 段后间距（行数，如150表示1.5行）
                line: 行间距值
                lineRule: 行间距规则，可以是'auto'(多倍行距), 'exact'(固定值), 'atLeast'(最小值)

        Returns:
            bool: 是否成功设置
        """
        # 检查索引是否有效
        if para_index < 0 or para_index >= len(self.paragraphs):
            print(f"错误：段落索引{para_index}超出范围(0-{len(self.paragraphs) - 1})")
            return False

        try:
            # 获取段落元素
            paragraph = self.paragraphs[para_index]
            para_element = paragraph.get('element')

            # 获取或创建段落属性元素
            pPr = self._get_or_create_pPr(para_element)

            # 查找或创建spacing元素
            spacing_elem = pPr.find(".//w:spacing", self.NAMESPACES)
            if spacing_elem is None:
                spacing_elem = ET.SubElement(pPr, f"{{{self.NAMESPACES['w']}}}spacing")

            # 设置各种间距属性
            ns_w = self.NAMESPACES['w']

            # 段前间距 (磅值)
            if 'before' in spacing:
                spacing_elem.set(f"{{{ns_w}}}before", str(spacing['before']))

            # 段后间距 (磅值)
            if 'after' in spacing:
                spacing_elem.set(f"{{{ns_w}}}after", str(spacing['after']))

            # 段前间距 (行数)
            if 'beforeLines' in spacing:

                spacing_elem.set(f"{{{ns_w}}}beforeLines", str(spacing['beforeLines']))
                # 如果同时设置了before，清除它，避免冲突
                if f"{{{ns_w}}}before" in spacing_elem.attrib:
                    del spacing_elem.attrib[f"{{{ns_w}}}before"]

            # 段后间距 (行数)
            if 'afterLines' in spacing:
                spacing_elem.set(f"{{{ns_w}}}afterLines", str(spacing['afterLines']))
                # 如果同时设置了after，清除它，避免冲突
                if f"{{{ns_w}}}after" in spacing_elem.attrib:
                    del spacing_elem.attrib[f"{{{ns_w}}}after"]

            # 行间距值
            if 'line' in spacing:
                spacing_elem.set(f"{{{ns_w}}}line", str(spacing['line']))

            # 行间距规则
            if 'lineRule' in spacing:
                spacing_elem.set(f"{{{ns_w}}}lineRule", spacing['lineRule'])

            # 更新XML
            self.update_document_xml()



            return True
        except Exception as e:
            print(f"设置段落间距时出错: {e}")
            return False
    def set_paragraph_spacing_from_xml(self, para_index, **spacing):
        """设置段落间距属性

        Args:
            para_index: 段落
            **spacing: 间距设置，可以包含以下键：
                before: 段前间距（磅值的1/20，如120表示6磅）
                after: 段后间距（磅值的1/20，如120表示6磅）
                beforeLines: 段前间距（行数，如150表示1.5行）
                afterLines: 段后间距（行数，如150表示1.5行）
                line: 行间距值
                lineRule: 行间距规则，可以是'auto'(多倍行距), 'exact'(固定值), 'atLeast'(最小值)

        Returns:
            bool: 是否成功设置
        """


        try:
            # 获取段落元素

            para_element = para_index

            # 获取或创建段落属性元素
            pPr = self._get_or_create_pPr(para_element)

            # 查找或创建spacing元素
            spacing_elem = pPr.find(".//w:spacing", self.NAMESPACES)
            if spacing_elem is None:
                spacing_elem = ET.SubElement(pPr, f"{{{self.NAMESPACES['w']}}}spacing")

            # 设置各种间距属性
            ns_w = self.NAMESPACES['w']

            # 段前间距 (磅值)
            if 'before' in spacing:
                spacing_elem.set(f"{{{ns_w}}}before", str(spacing['before']))

            # 段后间距 (磅值)
            if 'after' in spacing:
                spacing_elem.set(f"{{{ns_w}}}after", str(spacing['after']))

            # 段前间距 (行数)
            if 'beforeLines' in spacing:
                spacing_elem.set(f"{{{ns_w}}}beforeLines", str(spacing['beforeLines']))
                # 如果同时设置了before，清除它，避免冲突
                if f"{{{ns_w}}}before" in spacing_elem.attrib:
                    del spacing_elem.attrib[f"{{{ns_w}}}before"]

            # 段后间距 (行数)
            if 'afterLines' in spacing:
                spacing_elem.set(f"{{{ns_w}}}afterLines", str(spacing['afterLines']))
                # 如果同时设置了after，清除它，避免冲突
                if f"{{{ns_w}}}after" in spacing_elem.attrib:
                    del spacing_elem.attrib[f"{{{ns_w}}}after"]

            # 行间距值
            if 'line' in spacing:
                spacing_elem.set(f"{{{ns_w}}}line", str(spacing['line']))

            # 行间距规则
            if 'lineRule' in spacing:
                spacing_elem.set(f"{{{ns_w}}}lineRule", spacing['lineRule'])

            # 更新XML
            self.update_document_xml()



            return True
        except Exception as e:
            print(f"设置段落间距时出错: {e}")
            return False
    def set_paragraph_borders(self, para_index, **borders):
        """设置段落边框

        Args:
            para_index: 段落索引
            **borders: 边框设置，可包含以下参数:
                top: 上边框字典 (val, sz, space, color)
                bottom: 下边框字典
                left: 左边框字典
                right: 右边框字典

        Returns:
            bool: 是否成功修改
        """
        # 检查索引是否有效
        if para_index < 0 or para_index >= len(self.paragraphs):
            print(f"错误：段落索引{para_index}超出范围(0-{len(self.paragraphs)-1})")
            return False

        try:
            # 获取段落元素
            paragraph = self.paragraphs[para_index]['element']

            # 获取或创建pPr元素
            pPr = self._get_or_create_pPr(paragraph)

            # 查找边框元素
            pBdr = pPr.find(f".//{{{self.NAMESPACES['w']}}}pBdr")
            if pBdr is None:
                # 如果不存在，则创建
                pBdr = ET.Element(f"{{{self.NAMESPACES['w']}}}pBdr")
                pPr.append(pBdr)

            # 设置各类边框
            valid_borders = ['top', 'bottom', 'left', 'right']
            valid_attrs = ['val', 'sz', 'space', 'color']

            for border_type, border_settings in borders.items():
                if border_type in valid_borders and isinstance(border_settings, dict):
                    # 查找特定边框元素
                    border_elem = pBdr.find(f".//{{{self.NAMESPACES['w']}}}{border_type}")
                    if border_elem is None:
                        # 如果不存在，则创建
                        border_elem = ET.Element(f"{{{self.NAMESPACES['w']}}}{border_type}")
                        pBdr.append(border_elem)

                    # 设置边框属性
                    for attr, value in border_settings.items():
                        if attr in valid_attrs and value is not None:
                            border_elem.set(f"{{{self.NAMESPACES['w']}}}{attr}", str(value))

            return True
        except Exception as e:
            print(f"设置段落边框时出错: {e}")
            return False

    def set_paragraph_borders_from_xml(self, para_index, **borders):
        """设置段落边框

        Args:
            para_index: 段落
            **borders: 边框设置，可包含以下参数:
                top: 上边框字典 (val, sz, space, color)
                bottom: 下边框字典
                left: 左边框字典
                right: 右边框字典

        Returns:
            bool: 是否成功修改
        """


        try:
            # 获取段落元素
            paragraph = para_index

            # 获取或创建pPr元素
            pPr = self._get_or_create_pPr(paragraph)

            # 查找边框元素
            pBdr = pPr.find(f".//{{{self.NAMESPACES['w']}}}pBdr")
            if pBdr is None:
                # 如果不存在，则创建
                pBdr = ET.Element(f"{{{self.NAMESPACES['w']}}}pBdr")
                pPr.append(pBdr)

            # 设置各类边框
            valid_borders = ['top', 'bottom', 'left', 'right']
            valid_attrs = ['val', 'sz', 'space', 'color']

            for border_type, border_settings in borders.items():
                if border_type in valid_borders and isinstance(border_settings, dict):
                    # 查找特定边框元素
                    border_elem = pBdr.find(f".//{{{self.NAMESPACES['w']}}}{border_type}")
                    if border_elem is None:
                        # 如果不存在，则创建
                        border_elem = ET.Element(f"{{{self.NAMESPACES['w']}}}{border_type}")
                        pBdr.append(border_elem)

                    # 设置边框属性
                    for attr, value in border_settings.items():
                        if attr in valid_attrs and value is not None:
                            border_elem.set(f"{{{self.NAMESPACES['w']}}}{attr}", str(value))

            return True
        except Exception as e:
            print(f"设置段落边框时出错: {e}")
            return False
    def set_paragraph_shading(self, para_index, val=None, color=None, fill=None):
        """设置段落背景填充

        Args:
            para_index: 段落索引
            val: 填充类型 (clear, solid)
            color: 前景色
            fill: 背景色

        Returns:
            bool: 是否成功修改
        """
        # 检查索引是否有效
        if para_index < 0 or para_index >= len(self.paragraphs):
            print(f"错误：段落索引{para_index}超出范围(0-{len(self.paragraphs)-1})")
            return False

        try:
            # 获取段落元素
            paragraph = self.paragraphs[para_index]['element']

            # 获取或创建pPr元素
            pPr = self._get_or_create_pPr(paragraph)

            # 查找背景填充元素
            shd = pPr.find(f".//{{{self.NAMESPACES['w']}}}shd")
            if shd is None:
                # 如果不存在，则创建
                shd = ET.Element(f"{{{self.NAMESPACES['w']}}}shd")
                pPr.append(shd)

            # 设置填充属性
            if val is not None:
                shd.set(f"{{{self.NAMESPACES['w']}}}val", val)

            if color is not None:
                shd.set(f"{{{self.NAMESPACES['w']}}}color", color)

            if fill is not None:
                shd.set(f"{{{self.NAMESPACES['w']}}}fill", fill)

            return True
        except Exception as e:
            print(f"设置段落背景填充时出错: {e}")
            return False
    def set_paragraph_shading_from_xml(self, para_index, val=None, color=None, fill=None):
        """设置段落背景填充

        Args:
            para_index: 段落
            val: 填充类型 (clear, solid)
            color: 前景色
            fill: 背景色

        Returns:
            bool: 是否成功修改
        """


        try:
            # 获取段落元素
            paragraph = para_index

            # 获取或创建pPr元素
            pPr = self._get_or_create_pPr(paragraph)

            # 查找背景填充元素
            shd = pPr.find(f".//{{{self.NAMESPACES['w']}}}shd")
            if shd is None:
                # 如果不存在，则创建
                shd = ET.Element(f"{{{self.NAMESPACES['w']}}}shd")
                pPr.append(shd)

            # 设置填充属性
            if val is not None:
                shd.set(f"{{{self.NAMESPACES['w']}}}val", val)

            if color is not None:
                shd.set(f"{{{self.NAMESPACES['w']}}}color", color)

            if fill is not None:
                shd.set(f"{{{self.NAMESPACES['w']}}}fill", fill)

            return True
        except Exception as e:
            print(f"设置段落背景填充时出错: {e}")
            return False
    def set_paragraph_numbering(self, para_index, num_id=None, level=None):
        """设置段落编号

        Args:
            para_index: 段落索引
            num_id: 编号ID
            level: 编号级别

        Returns:
            bool: 是否成功修改
        """
        # 检查索引是否有效
        if para_index < 0 or para_index >= len(self.paragraphs):
            print(f"错误：段落索引{para_index}超出范围(0-{len(self.paragraphs)-1})")
            return False

        try:
            # 获取段落元素
            paragraph = self.paragraphs[para_index]['element']

            # 获取或创建pPr元素
            pPr = self._get_or_create_pPr(paragraph)

            # 查找编号元素
            numPr = pPr.find(f".//{{{self.NAMESPACES['w']}}}numPr")
            if numPr is None:
                # 如果不存在，则创建
                numPr = ET.Element(f"{{{self.NAMESPACES['w']}}}numPr")
                pPr.append(numPr)

            # 设置编号ID
            if num_id is not None:
                numId = numPr.find(f".//{{{self.NAMESPACES['w']}}}numId")
                if numId is None:
                    numId = ET.Element(f"{{{self.NAMESPACES['w']}}}numId")
                    numPr.append(numId)
                numId.set(f"{{{self.NAMESPACES['w']}}}val", str(num_id))

            # 设置编号级别
            if level is not None:
                ilvl = numPr.find(f".//{{{self.NAMESPACES['w']}}}ilvl")
                if ilvl is None:
                    ilvl = ET.Element(f"{{{self.NAMESPACES['w']}}}ilvl")
                    numPr.append(ilvl)
                ilvl.set(f"{{{self.NAMESPACES['w']}}}val", str(level))

            return True
        except Exception as e:
            print(f"设置段落编号时出错: {e}")
            return False

    def set_paragraph_font(self, para_index, **font_properties):
        """设置段落级别的字体属性

        Args:
            para_index: 段落索引
            **font_properties: 字体属性设置，可包含以下参数:
                ascii: 英文字体
                hAnsi: 西文字体
                eastAsia: 中文字体
                cs: 复杂文种字体
                size: 字号(半磅值)
                bold: 是否加粗(True/False)
                italic: 是否斜体(True/False)
                underline: 下划线样式
                color: 文字颜色

        Returns:
            bool: 是否成功修改
        """
        # 检查索引是否有效
        if para_index < 0 or para_index >= len(self.paragraphs):
            print(f"错误：段落索引{para_index}超出范围(0-{len(self.paragraphs)-1})")
            return False

        try:
            # 获取段落元素
            paragraph = self.paragraphs[para_index]['element']

            # 获取或创建pPr元素
            pPr = self._get_or_create_pPr(paragraph)

            # 查找或创建rPr元素（段落级别的文本属性）
            rPr = pPr.find(f".//{{{self.NAMESPACES['w']}}}rPr")
            if rPr is None:
                rPr = ET.Element(f"{{{self.NAMESPACES['w']}}}rPr")
                pPr.append(rPr)

            # 设置字体
            font_types = ['ascii', 'hAnsi', 'eastAsia', 'cs']
            font_set = False
            for font_type in font_types:
                if font_type in font_properties:
                    font_set = True

            if font_set:
                rFonts = rPr.find(f".//{{{self.NAMESPACES['w']}}}rFonts")
                if rFonts is None:
                    rFonts = ET.Element(f"{{{self.NAMESPACES['w']}}}rFonts")
                    rPr.append(rFonts)

                for font_type in font_types:
                    if font_type in font_properties:
                        rFonts.set(f"{{{self.NAMESPACES['w']}}}{font_type}", font_properties[font_type])

            # 设置字号
            if 'size' in font_properties:
                sz = rPr.find(f".//{{{self.NAMESPACES['w']}}}sz")
                if sz is None:
                    sz = ET.Element(f"{{{self.NAMESPACES['w']}}}sz")
                    rPr.append(sz)
                sz.set(f"{{{self.NAMESPACES['w']}}}val", str(font_properties['size']))

            # 设置加粗
            if 'bold' in font_properties:
                b = rPr.find(f".//{{{self.NAMESPACES['w']}}}b")
                if font_properties['bold']:
                    if b is None:
                        b = ET.Element(f"{{{self.NAMESPACES['w']}}}b")
                        rPr.append(b)
                    b.set(f"{{{self.NAMESPACES['w']}}}val", "true")
                elif b is not None:
                    # 如果要关闭加粗，可以移除元素或设置val="false"
                    rPr.remove(b)

            # 设置斜体
            if 'italic' in font_properties:
                i = rPr.find(f".//{{{self.NAMESPACES['w']}}}i")
                if font_properties['italic']:
                    if i is None:
                        i = ET.Element(f"{{{self.NAMESPACES['w']}}}i")
                        rPr.append(i)
                    i.set(f"{{{self.NAMESPACES['w']}}}val", "true")
                elif i is not None:
                    rPr.remove(i)

            # 设置下划线
            if 'underline' in font_properties:
                u = rPr.find(f".//{{{self.NAMESPACES['w']}}}u")
                if u is None:
                    u = ET.Element(f"{{{self.NAMESPACES['w']}}}u")
                    rPr.append(u)
                u.set(f"{{{self.NAMESPACES['w']}}}val", font_properties['underline'])

            # 设置颜色
            if 'color' in font_properties:
                color = rPr.find(f".//{{{self.NAMESPACES['w']}}}color")
                if color is None:
                    color = ET.Element(f"{{{self.NAMESPACES['w']}}}color")
                    rPr.append(color)
                color.set(f"{{{self.NAMESPACES['w']}}}val", font_properties['color'])

            return True
        except Exception as e:
            print(f"设置段落字体属性时出错: {e}")
            return False
    def set_paragraph_font_from_xml(self, para_index, **font_properties):
        """设置段落级别的字体属性

        Args:
            para_index: 段落
            **font_properties: 字体属性设置，可包含以下参数:
                ascii: 英文字体
                hAnsi: 西文字体
                eastAsia: 中文字体
                cs: 复杂文种字体
                size: 字号(半磅值)
                bold: 是否加粗(True/False)
                italic: 是否斜体(True/False)
                underline: 下划线样式
                color: 文字颜色

        Returns:
            bool: 是否成功修改
        """

        try:
            # 获取段落元素
            paragraph = para_index

            # 获取或创建pPr元素
            pPr = self._get_or_create_pPr(paragraph)

            # 查找或创建rPr元素（段落级别的文本属性）
            rPr = pPr.find(f".//{{{self.NAMESPACES['w']}}}rPr")
            if rPr is None:
                rPr = ET.Element(f"{{{self.NAMESPACES['w']}}}rPr")
                pPr.append(rPr)

            # 设置字体
            font_types = ['ascii', 'hAnsi', 'eastAsia', 'cs']
            font_set = False
            for font_type in font_types:
                if font_type in font_properties:
                    font_set = True

            if font_set:
                rFonts = rPr.find(f".//{{{self.NAMESPACES['w']}}}rFonts")
                if rFonts is None:
                    rFonts = ET.Element(f"{{{self.NAMESPACES['w']}}}rFonts")
                    rPr.append(rFonts)

                for font_type in font_types:
                    if font_type in font_properties:
                        rFonts.set(f"{{{self.NAMESPACES['w']}}}{font_type}", font_properties[font_type])

            # 设置字号
            if 'size' in font_properties:
                sz = rPr.find(f".//{{{self.NAMESPACES['w']}}}sz")
                if sz is None:
                    sz = ET.Element(f"{{{self.NAMESPACES['w']}}}sz")
                    rPr.append(sz)
                sz.set(f"{{{self.NAMESPACES['w']}}}val", str(font_properties['size']))

            # 设置加粗
            if 'bold' in font_properties:
                b = rPr.find(f".//{{{self.NAMESPACES['w']}}}b")
                if font_properties['bold']:
                    if b is None:
                        b = ET.Element(f"{{{self.NAMESPACES['w']}}}b")
                        rPr.append(b)
                    b.set(f"{{{self.NAMESPACES['w']}}}val", "true")
                elif b is not None:
                    # 如果要关闭加粗，可以移除元素或设置val="false"
                    rPr.remove(b)

            # 设置斜体
            if 'italic' in font_properties:
                i = rPr.find(f".//{{{self.NAMESPACES['w']}}}i")
                if font_properties['italic']:
                    if i is None:
                        i = ET.Element(f"{{{self.NAMESPACES['w']}}}i")
                        rPr.append(i)
                    i.set(f"{{{self.NAMESPACES['w']}}}val", "true")
                elif i is not None:
                    rPr.remove(i)

            # 设置下划线
            if 'underline' in font_properties:
                u = rPr.find(f".//{{{self.NAMESPACES['w']}}}u")
                if u is None:
                    u = ET.Element(f"{{{self.NAMESPACES['w']}}}u")
                    rPr.append(u)
                u.set(f"{{{self.NAMESPACES['w']}}}val", font_properties['underline'])

            # 设置颜色
            if 'color' in font_properties:
                color = rPr.find(f".//{{{self.NAMESPACES['w']}}}color")
                if color is None:
                    color = ET.Element(f"{{{self.NAMESPACES['w']}}}color")
                    rPr.append(color)
                color.set(f"{{{self.NAMESPACES['w']}}}val", font_properties['color'])

            return True
        except Exception as e:
            print(f"设置段落字体属性时出错: {e}")
            return False
    def remove_paragraph_property(self, para_index, property_name):
        """移除段落的特定样式属性

        Args:
            para_index: 段落索引
            property_name: 要移除的属性名称(pStyle, jc, ind, spacing, pBdr, shd, numPr, rPr等)

        Returns:
            bool: 是否成功移除
        """
        # 检查索引是否有效
        if para_index < 0 or para_index >= len(self.paragraphs):
            print(f"错误：段落索引{para_index}超出范围(0-{len(self.paragraphs)-1})")
            return False

        try:
            # 获取段落元素
            paragraph = self.paragraphs[para_index]['element']

            # 查找段落属性标签
            pPr = paragraph.find(f".//{{{self.NAMESPACES['w']}}}pPr")
            if pPr is None:
                return False  # 没有样式可以移除

            # 查找指定属性
            prop = pPr.find(f".//{{{self.NAMESPACES['w']}}}{property_name}")
            if prop is not None:
                pPr.remove(prop)
                return True
            else:
                return False  # 未找到要移除的属性
        except Exception as e:
            print(f"移除段落属性时出错: {e}")
            return False

    def update_paragraph_style(self, para_index,**style_properties):
        """更新段落的多个样式属性

        Args:
            para_index: 段落索引
            **style_properties: 样式属性字典，可包含以下键：
                style_id: 样式ID
                alignment: 对齐方式
                indentation: 缩进设置字典
                spacing: 间距设置字典
                borders: 边框设置字典
                shading: 背景填充字典 (val, color, fill)
                numbering: 编号设置字典 (id, level)
                font: 字体设置字典

        Returns:
            bool: 是否成功更新所有样式
        """
        # 检查索引是否有效
        if para_index < 0 or para_index >= len(self.paragraphs):
            print(f"错误：段落索引{para_index}超出范围(0-{len(self.paragraphs)-1})")
            return False

        success = True

        # 更新样式ID
        if 'style_id' in style_properties:
            if not self.set_paragraph_style_id(para_index, style_properties['style_id']):
                success = False

        # 更新对齐方式
        if 'alignment' in style_properties:
            if not self.set_paragraph_alignment(para_index, style_properties['alignment']):
                success = False

        # 更新缩进
        if 'indentation' in style_properties and isinstance(style_properties['indentation'], dict):
            if not self.set_paragraph_indentation(para_index, **style_properties['indentation']):
                success = False

        # 更新间距
        if 'spacing' in style_properties and isinstance(style_properties['spacing'], dict):
            if not self.set_paragraph_spacing(para_index, **style_properties['spacing']):
                success = False

        # 更新边框
        if 'borders' in style_properties and isinstance(style_properties['borders'], dict):
            if not self.set_paragraph_borders(para_index, **style_properties['borders']):
                success = False

        # 更新背景填充
        if 'shading' in style_properties and isinstance(style_properties['shading'], dict):
            shading = style_properties['shading']
            if not self.set_paragraph_shading(
                para_index,
                val=shading.get('val'),
                color=shading.get('color'),
                fill=shading.get('fill')
            ):
                success = False

        # 更新编号
        if 'numbering' in style_properties and isinstance(style_properties['numbering'], dict):
            numbering = style_properties['numbering']
            if not self.set_paragraph_numbering(
                para_index,
                num_id=numbering.get('id'),
                level=numbering.get('level')
            ):
                success = False

        # 更新字体属性
        if 'fonts' in style_properties and isinstance(style_properties['fonts'], dict):
            if not self.set_paragraph_font(para_index, **style_properties['fonts']):
                success = False

        return success

    def update_paragrstyle(self, para_index,**style_properties):
        """更新段落的多个样式属性

        Args:
            para_index: 段落索引
            **style_properties: 样式属性字典，可包含以下键：
                style_id: 样式ID
                alignment: 对齐方式
                indentation: 缩进设置字典
                spacing: 间距设置字典
                borders: 边框设置字典
                shading: 背景填充字典 (val, color, fill)
                numbering: 编号设置字典 (id, level)
                font: 字体设置字典

        Returns:
            bool: 是否成功更新所有样式
        """
        # 检查索引是否有效
        if para_index < 0 or para_index >= len(self.paragraphs):
            print(f"错误：段落索引{para_index}超出范围(0-{len(self.paragraphs)-1})")
            return False

        success = True

        # 更新样式ID
        if 'style_id' in style_properties:
            if not self.set_paragraph_style_id(para_index, style_properties['style_id']):
                success = False

        # 更新对齐方式
        if 'alignment' in style_properties:
            if not self.set_paragraph_alignment(para_index, style_properties['alignment']):
                success = False

        # 更新缩进
        if 'indentation' in style_properties and isinstance(style_properties['indentation'], dict):
            if not self.set_paragraph_indentation(para_index, **style_properties['indentation']):
                success = False

        # 更新间距
        if 'spacing' in style_properties and isinstance(style_properties['spacing'], dict):
            if not self.set_paragraph_spacing(para_index, **style_properties['spacing']):
                success = False

        # 更新边框
        if 'borders' in style_properties and isinstance(style_properties['borders'], dict):
            if not self.set_paragraph_borders(para_index, **style_properties['borders']):
                success = False

        # 更新背景填充
        if 'shading' in style_properties and isinstance(style_properties['shading'], dict):
            shading = style_properties['shading']
            if not self.set_paragraph_shading(
                para_index,
                val=shading.get('val'),
                color=shading.get('color'),
                fill=shading.get('fill')
            ):
                success = False

        # 更新编号
        if 'numbering' in style_properties and isinstance(style_properties['numbering'], dict):
            numbering = style_properties['numbering']
            if not self.set_paragraph_numbering(
                para_index,
                num_id=numbering.get('id'),
                level=numbering.get('level')
            ):
                success = False

        # 更新字体属性
        if 'fonts' in style_properties and isinstance(style_properties['fonts'], dict):
            if not self.set_paragraph_font(para_index, **style_properties['fonts']):
                success = False

        return success

    def update_paragraph_style_from_xml(self, para_element, **style_properties):
        """更新段落的多个样式属性

        Args:
            para_element: 段落元素
            **style_properties: 样式属性字典，可包含以下键：
                style_id: 样式ID
                alignment: 对齐方式
                indentation: 缩进设置字典
                spacing: 间距设置字典
                borders: 边框设置字典
                shading: 背景填充字典 (val, color, fill)
                numbering: 编号设置字典 (id, level)
                font: 字体设置字典

        Returns:
            bool: 是否成功更新所有样式
        """


        success = True

        # 更新样式ID
        if 'style_id' in style_properties:
            if not self.set_paragraph_style_id_from_xml(para_element, style_properties['style_id']):
                success = False

        # 更新对齐方式
        if 'alignment' in style_properties:
            if not self.set_paragraph_alignment_from_xml(para_element, style_properties['alignment']):
                success = False

        # 更新缩进
        if 'indentation' in style_properties and isinstance(style_properties['indentation'], dict):
            if not self.set_paragraph_indentation_from_xml(para_element, **style_properties['indentation']):
                success = False

        # 更新间距
        if 'spacing' in style_properties and isinstance(style_properties['spacing'], dict):
            if not self.set_paragraph_spacing_from_xml(para_element, **style_properties['spacing']):
                success = False

        # 更新边框
        if 'borders' in style_properties and isinstance(style_properties['borders'], dict):
            if not self.set_paragraph_borders_from_xml(para_element, **style_properties['borders']):
                success = False

        # 更新背景填充
        if 'shading' in style_properties and isinstance(style_properties['shading'], dict):
            shading = style_properties['shading']
            if not self.set_paragraph_shading_from_xml(
                    para_element,
                    val=shading.get('val'),
                    color=shading.get('color'),
                    fill=shading.get('fill')
            ):
                success = False



        # 更新字体属性
        if 'fonts' in style_properties and isinstance(style_properties['fonts'], dict):
            if not self.set_paragraph_font_from_xml(para_element, **style_properties['fonts']):
                success = False

        return success
    def update_document_xml(self):
        """在保存前更新文档XML

        确保所有对XML树的修改都同步到self.parts["document"]中
        """
        try:
            # 将修改后的XML树转换为字符串
            xml_string = ET.tostring(self.root, encoding='utf-8')

            # 创建新的ElementTree对象
            updated_tree = ET.ElementTree(ET.fromstring(xml_string))

            # 更新parts中的document
            self.parts["document"] = updated_tree


            return True
        except Exception as e:
            print(f"更新文档XML时出错: {e}")
            return False

    def save(self, output_path):
        """重写父类的save方法，确保在保存前更新文档XML

        Args:
            output_path: 输出文档的路径

        Returns:
            bool: 是否成功保存
        """
        # 先确保XML树被更新到parts中
        if not self.update_document_xml():
            print("更新文档XML失败，无法保存")
            return False

        # 调用父类的save方法
        return super().save(output_path)

    def set_paragraph_runs_font(self, para_index, **font_properties):
        """设置段落中所有Run元素的字体属性

        Args:
            para_index: 段落索引
            **font_properties: 字体属性，可包含以下键：
                ascii: ASCII字体名称
                eastAsia: 东亚字体名称
                hAnsi: HANSI字体名称
                cs: 复杂脚本字体名称
                size: 字体大小（磅值）
                bold: 是否加粗 (True/False)
                italic: 是否斜体 (True/False)
                underline: 下划线类型
                color: 颜色值（如"FF0000"）

        Returns:
            bool: 是否成功修改
        """
        # 检查索引是否有效
        if para_index < 0 or para_index >= len(self.paragraphs):
            print(f"错误：段落索引{para_index}超出范围(0-{len(self.paragraphs) - 1})")
            return False

        try:
            # 获取段落元素
            paragraph = self.paragraphs[para_index]
            para_element = paragraph.get('element')

            # 查找所有w:r元素
            r_elements = para_element.findall(f".//{{{self.NAMESPACES['w']}}}r")
            if not r_elements:
                print(f"段落{para_index}中没有找到Run元素")
                return False

            # 同时修改段落级别的字体设置（可选）
            try:
                # 获取或创建段落属性元素
                pPr = self._get_or_create_pPr(para_element)

                # 获取或创建段落级别的rPr元素
                rPr_para = pPr.find(f"./w:rPr", self.NAMESPACES)
                if rPr_para is None:
                    rPr_para = ET.Element(f"{{{self.NAMESPACES['w']}}}rPr")
                    pPr.append(rPr_para)

                # 设置字体名称
                font_keys = ['ascii', 'eastAsia', 'hAnsi', 'cs']
                if any(key in font_properties for key in font_keys):
                    # 查找或创建字体元素
                    rFonts = rPr_para.find(f"./w:rFonts", self.NAMESPACES)
                    if rFonts is None:
                        rFonts = ET.Element(f"{{{self.NAMESPACES['w']}}}rFonts")
                        rPr_para.append(rFonts)

                    # 设置各种字体
                    for key in font_keys:
                        if key in font_properties and font_properties[key]:
                            rFonts.set(f"{{{self.NAMESPACES['w']}}}{key}", font_properties[key])

                # 设置字体大小
                if 'size' in font_properties and font_properties['size']:
                    # 转换为半磅单位
                    size = str(int(float(font_properties['size']) * 2))
                    # 查找或创建sz元素
                    sz = rPr_para.find(f"./w:sz", self.NAMESPACES)
                    if sz is None:
                        sz = ET.Element(f"{{{self.NAMESPACES['w']}}}sz")
                        rPr_para.append(sz)
                    sz.set(f"{{{self.NAMESPACES['w']}}}val", size)

                    # 同时设置szCs（复杂脚本字体大小）
                    szCs = rPr_para.find(f"./w:szCs", self.NAMESPACES)
                    if szCs is None:
                        szCs = ET.Element(f"{{{self.NAMESPACES['w']}}}szCs")
                        rPr_para.append(szCs)
                    szCs.set(f"{{{self.NAMESPACES['w']}}}val", size)

                # 设置加粗
                if 'bold' in font_properties:
                    # 查找或创建b元素
                    bold = rPr_para.find(f"./w:b", self.NAMESPACES)
                    if font_properties['bold']:
                        if bold is None:
                            bold = ET.Element(f"{{{self.NAMESPACES['w']}}}b")
                            rPr_para.append(bold)
                        # 移除val属性，在Word中表示启用
                        if f"{{{self.NAMESPACES['w']}}}val" in bold.attrib:
                            del bold.attrib[f"{{{self.NAMESPACES['w']}}}val"]
                    else:
                        if bold is not None:
                            rPr_para.remove(bold)

                # 设置斜体
                if 'italic' in font_properties:
                    # 查找或创建i元素
                    italic = rPr_para.find(f"./w:i", self.NAMESPACES)
                    if font_properties['italic']:
                        if italic is None:
                            italic = ET.Element(f"{{{self.NAMESPACES['w']}}}i")
                            rPr_para.append(italic)
                        # 移除val属性，在Word中表示启用
                        if f"{{{self.NAMESPACES['w']}}}val" in italic.attrib:
                            del italic.attrib[f"{{{self.NAMESPACES['w']}}}val"]
                    else:
                        if italic is not None:
                            rPr_para.remove(italic)

                # 设置颜色
                if 'color' in font_properties and font_properties['color']:
                    # 查找或创建color元素
                    color = rPr_para.find(f"./w:color", self.NAMESPACES)
                    if color is None:
                        color = ET.Element(f"{{{self.NAMESPACES['w']}}}color")
                        rPr_para.append(color)
                    color.set(f"{{{self.NAMESPACES['w']}}}val", font_properties['color'])
            except Exception as e:
                print(f"设置段落级字体属性时出错: {e}")
                # 继续处理Run元素，不中断执行

            # 修改每个Run元素的字体属性
            success = True
            for i, run in enumerate(r_elements):
                # 获取或创建rPr元素
                rPr = run.find(f"./w:rPr", self.NAMESPACES)
                if rPr is None:
                    rPr = ET.Element(f"{{{self.NAMESPACES['w']}}}rPr")
                    # 将rPr插入到Run的第一个位置
                    run.insert(0, rPr)

                # 设置字体名称
                font_keys = ['ascii', 'eastAsia', 'hAnsi', 'cs']
                if any(key in font_properties for key in font_keys):
                    # 查找或创建字体元素
                    rFonts = rPr.find(f"./w:rFonts", self.NAMESPACES)
                    if rFonts is None:
                        rFonts = ET.Element(f"{{{self.NAMESPACES['w']}}}rFonts")
                        rPr.append(rFonts)

                    # 设置各种字体
                    for key in font_keys:
                        if key in font_properties and font_properties[key]:
                            rFonts.set(f"{{{self.NAMESPACES['w']}}}{key}", font_properties[key])
                            print(f"设置Run {i} 字体属性 {key}: {font_properties[key]}")

                # 设置字体大小
                if 'size' in font_properties and font_properties['size']:
                    # 转换为半磅单位
                    size = str(int(float(font_properties['size']) * 2))
                    # 查找或创建sz元素
                    sz = rPr.find(f"./w:sz", self.NAMESPACES)
                    if sz is None:
                        sz = ET.Element(f"{{{self.NAMESPACES['w']}}}sz")
                        rPr.append(sz)
                    sz.set(f"{{{self.NAMESPACES['w']}}}val", size)


                    # 同时设置szCs（复杂脚本字体大小）
                    szCs = rPr.find(f"./w:szCs", self.NAMESPACES)
                    if szCs is None:
                        szCs = ET.Element(f"{{{self.NAMESPACES['w']}}}szCs")
                        rPr.append(szCs)
                    szCs.set(f"{{{self.NAMESPACES['w']}}}val", size)

                # 设置加粗
                if 'bold' in font_properties:
                    # 查找或创建b元素
                    bold = rPr.find(f"./w:b", self.NAMESPACES)
                    if font_properties['bold']:
                        if bold is None:
                            bold = ET.Element(f"{{{self.NAMESPACES['w']}}}b")
                            rPr.append(bold)
                        # 移除val属性，在Word中表示启用
                        if f"{{{self.NAMESPACES['w']}}}val" in bold.attrib:
                            del bold.attrib[f"{{{self.NAMESPACES['w']}}}val"]

                    else:
                        if bold is not None:
                            rPr.remove(bold)


                # 设置斜体
                if 'italic' in font_properties:
                    # 查找或创建i元素
                    italic = rPr.find(f"./w:i", self.NAMESPACES)
                    if font_properties['italic']:
                        if italic is None:
                            italic = ET.Element(f"{{{self.NAMESPACES['w']}}}i")
                            rPr.append(italic)
                        # 移除val属性，在Word中表示启用
                        if f"{{{self.NAMESPACES['w']}}}val" in italic.attrib:
                            del italic.attrib[f"{{{self.NAMESPACES['w']}}}val"]

                    else:
                        if italic is not None:
                            rPr.remove(italic)


                # 设置下划线
                if 'underline' in font_properties and font_properties['underline']:
                    # 查找或创建u元素
                    u = rPr.find(f"./w:u", self.NAMESPACES)
                    if u is None:
                        u = ET.Element(f"{{{self.NAMESPACES['w']}}}u")
                        rPr.append(u)
                    u.set(f"{{{self.NAMESPACES['w']}}}val", font_properties['underline'])


                # 设置颜色
                if 'color' in font_properties and font_properties['color']:
                    # 查找或创建color元素
                    color = rPr.find(f"./w:color", self.NAMESPACES)
                    if color is None:
                        color = ET.Element(f"{{{self.NAMESPACES['w']}}}color")
                        rPr.append(color)
                    color.set(f"{{{self.NAMESPACES['w']}}}val", font_properties['color'])


            # 更新文档XML
            self.update_document_xml()
            print(f"完成对段落 {para_index} 所有Run元素的字体设置")
            return success
        except Exception as e:
            print(f"设置段落所有Run元素的字体属性时出错: {e}")
            traceback.print_exc()
            return False

    def set_runs_bold(self, para_index, bold=True):
        """设置段落中所有文本运行的加粗格式

        Args:
            para_index: 段落索引
            bold: 是否加粗，True为加粗，False为取消加粗

        Returns:
            bool: 是否成功修改
        """
        # 检查索引是否有效
        if para_index < 0 or para_index >= len(self.paragraphs):
            print(f"错误：段落索引{para_index}超出范围(0-{len(self.paragraphs)-1})")
            return False

        try:
            # 获取段落元素
            paragraph = self.paragraphs[para_index]['element']

            # 查找所有w:r元素
            r_elements = paragraph.findall("./w:r", self.NAMESPACES)
            if not r_elements:
                print(f"段落{para_index}中没有找到文本运行")
                return False

            # 修改每个文本运行的加粗属性
            for r in r_elements:
                # 查找或创建rPr元素
                rPr = r.find(f".//{{{self.NAMESPACES['w']}}}rPr")
                if rPr is None:
                    rPr = ET.Element(f"{{{self.NAMESPACES['w']}}}rPr")
                    r.insert(0, rPr)

                # 查找加粗元素
                b = rPr.find(f".//{{{self.NAMESPACES['w']}}}b")

                # 根据参数设置或移除加粗
                if bold:
                    if b is None:
                        b = ET.Element(f"{{{self.NAMESPACES['w']}}}b")
                        rPr.append(b)
                    b.set(f"{{{self.NAMESPACES['w']}}}val", "true")
                elif b is not None:
                    rPr.remove(b)

            return True
        except Exception as e:
            print(f"设置段落文本运行加粗格式时出错: {e}")
            return False

    def set_runs_italic(self, para_index, italic=True):
        """设置段落中所有文本运行的斜体格式

        Args:
            para_index: 段落索引
            italic: 是否斜体，True为斜体，False为取消斜体

        Returns:
            bool: 是否成功修改
        """
        # 检查索引是否有效
        if para_index < 0 or para_index >= len(self.paragraphs):
            print(f"错误：段落索引{para_index}超出范围(0-{len(self.paragraphs)-1})")
            return False

        try:
            # 获取段落元素
            paragraph = self.paragraphs[para_index]['element']

            # 查找所有w:r元素
            r_elements = paragraph.findall("./w:r", self.NAMESPACES)
            if not r_elements:
                print(f"段落{para_index}中没有找到文本运行")
                return False

            # 修改每个文本运行的斜体属性
            for r in r_elements:
                # 查找或创建rPr元素
                rPr = r.find(f".//{{{self.NAMESPACES['w']}}}rPr")
                if rPr is None:
                    rPr = ET.Element(f"{{{self.NAMESPACES['w']}}}rPr")
                    r.insert(0, rPr)

                # 查找斜体元素
                i = rPr.find(f".//{{{self.NAMESPACES['w']}}}i")

                # 根据参数设置或移除斜体
                if italic:
                    if i is None:
                        i = ET.Element(f"{{{self.NAMESPACES['w']}}}i")
                        rPr.append(i)
                    i.set(f"{{{self.NAMESPACES['w']}}}val", "true")
                elif i is not None:
                    rPr.remove(i)

            return True
        except Exception as e:
            print(f"设置段落文本运行斜体格式时出错: {e}")
            return False

    def set_runs_underline(self, para_index, underline_type='single'):
        """设置段落中所有文本运行的下划线格式

        Args:
            para_index: 段落索引
            underline_type: 下划线类型，如'single'(单线)、'double'(双线)、'thick'(粗线)
                            'dotted'(点线)、'dash'(虚线)、'wave'(波浪线)，传入None表示移除下划线

        Returns:
            bool: 是否成功修改
        """
        # 检查索引是否有效
        if para_index < 0 or para_index >= len(self.paragraphs):
            print(f"错误：段落索引{para_index}超出范围(0-{len(self.paragraphs)-1})")
            return False

        try:
            # 获取段落元素
            paragraph = self.paragraphs[para_index]['element']

            # 查找所有w:r元素
            r_elements =paragraph.findall("./w:r", self.NAMESPACES)
            if not r_elements:
                print(f"段落{para_index}中没有找到文本运行")
                return False

            # 修改每个文本运行的下划线属性
            for r in r_elements:
                # 查找或创建rPr元素
                rPr = r.find(f".//{{{self.NAMESPACES['w']}}}rPr")
                if rPr is None:
                    rPr = ET.Element(f"{{{self.NAMESPACES['w']}}}rPr")
                    r.insert(0, rPr)

                # 查找下划线元素
                u = rPr.find(f".//{{{self.NAMESPACES['w']}}}u")

                # 根据参数设置或移除下划线
                if underline_type is None:
                    if u is not None:
                        rPr.remove(u)
                else:
                    if u is None:
                        u = ET.Element(f"{{{self.NAMESPACES['w']}}}u")
                        rPr.append(u)
                    u.set(f"{{{self.NAMESPACES['w']}}}val", underline_type)

            return True
        except Exception as e:
            print(f"设置段落文本运行下划线格式时出错: {e}")
            return False

    def set_runs_color(self, para_index, color):
        """设置段落中所有文本运行的颜色

        Args:
            para_index: 段落索引
            color: 颜色值，如'FF0000'表示红色

        Returns:
            bool: 是否成功修改
        """
        # 检查索引是否有效
        if para_index < 0 or para_index >= len(self.paragraphs):
            print(f"错误：段落索引{para_index}超出范围(0-{len(self.paragraphs)-1})")
            return False

        try:
            # 获取段落元素
            paragraph = self.paragraphs[para_index]['element']

            # 查找所有w:r元素
            r_elements = paragraph.findall("./w:r", self.NAMESPACES)
            if not r_elements:
                print(f"段落{para_index}中没有找到文本运行")
                return False

            # 修改每个文本运行的颜色
            for r in r_elements:
                # 查找或创建rPr元素
                rPr = r.find(f".//{{{self.NAMESPACES['w']}}}rPr")
                if rPr is None:
                    rPr = ET.Element(f"{{{self.NAMESPACES['w']}}}rPr")
                    r.insert(0, rPr)

                # 查找颜色元素
                c = rPr.find(f".//{{{self.NAMESPACES['w']}}}color")

                # 设置颜色
                if color is None:
                    if c is not None:
                        rPr.remove(c)
                else:
                    if c is None:
                        c = ET.Element(f"{{{self.NAMESPACES['w']}}}color")
                        rPr.append(c)
                    c.set(f"{{{self.NAMESPACES['w']}}}val", color)

            return True
        except Exception as e:
            print(f"设置段落文本运行颜色时出错: {e}")
            return False

    def set_runs_size(self, para_index, size):
        """设置段落中所有文本运行的字号

        Args:
            para_index: 段落索引
            size: 字号值(半磅值)，如24表示12磅

        Returns:
            bool: 是否成功修改
        """
        # 检查索引是否有效
        if para_index < 0 or para_index >= len(self.paragraphs):
            print(f"错误：段落索引{para_index}超出范围(0-{len(self.paragraphs)-1})")
            return False

        try:
            # 获取段落元素
            paragraph = self.paragraphs[para_index]['element']

            # 查找所有w:r元素
            r_elements = paragraph.findall("./w:r", self.NAMESPACES)
            if not r_elements:
                print(f"段落{para_index}中没有找到文本运行")
                return False

            # 修改每个文本运行的字号
            for r in r_elements:
                # 查找或创建rPr元素
                rPr = r.find(f".//{{{self.NAMESPACES['w']}}}rPr")
                if rPr is None:
                    rPr = ET.Element(f"{{{self.NAMESPACES['w']}}}rPr")
                    r.insert(0, rPr)

                # 查找字号元素
                sz = rPr.find(f".//{{{self.NAMESPACES['w']}}}sz")

                # 设置字号
                if size is None:
                    if sz is not None:
                        rPr.remove(sz)
                else:
                    if sz is None:
                        sz = ET.Element(f"{{{self.NAMESPACES['w']}}}sz")
                        rPr.append(sz)
                    sz.set(f"{{{self.NAMESPACES['w']}}}val", str(size))

            return True
        except Exception as e:
            print(f"设置段落文本运行字号时出错: {e}")
            return False

    def set_runs_highlight(self, para_index, highlight_color):
        """设置段落中所有文本运行的高亮颜色

        Args:
            para_index: 段落索引
            highlight_color: 高亮颜色值，如'yellow'、'green'、'red'等，传入None表示移除高亮

        Returns:
            bool: 是否成功修改
        """
        # 检查索引是否有效
        if para_index < 0 or para_index >= len(self.paragraphs):
            print(f"错误：段落索引{para_index}超出范围(0-{len(self.paragraphs)-1})")
            return False

        try:
            # 获取段落元素
            paragraph = self.paragraphs[para_index]['element']

            # 查找所有w:r元素
            r_elements =paragraph.findall("./w:r", self.NAMESPACES)
            if not r_elements:
                print(f"段落{para_index}中没有找到文本运行")
                return False

            # 修改每个文本运行的高亮颜色
            for r in r_elements:
                # 查找或创建rPr元素
                rPr = r.find(f".//{{{self.NAMESPACES['w']}}}rPr")
                if rPr is None:
                    rPr = ET.Element(f"{{{self.NAMESPACES['w']}}}rPr")
                    r.insert(0, rPr)

                # 查找高亮元素
                highlight = rPr.find(f".//{{{self.NAMESPACES['w']}}}highlight")

                # 设置高亮
                if highlight_color is None:
                    if highlight is not None:
                        rPr.remove(highlight)
                else:
                    if highlight is None:
                        highlight = ET.Element(f"{{{self.NAMESPACES['w']}}}highlight")
                        rPr.append(highlight)
                    highlight.set(f"{{{self.NAMESPACES['w']}}}val", highlight_color)

            return True
        except Exception as e:
            print(f"设置段落文本运行高亮颜色时出错: {e}")
            return False

    def set_runs_strike(self, para_index, strike=True):
        """设置段落中所有文本运行的删除线格式

        Args:
            para_index: 段落索引
            strike: 是否添加删除线，True为添加，False为移除

        Returns:
            bool: 是否成功修改
        """
        # 检查索引是否有效
        if para_index < 0 or para_index >= len(self.paragraphs):
            print(f"错误：段落索引{para_index}超出范围(0-{len(self.paragraphs)-1})")
            return False

        try:
            # 获取段落元素
            paragraph = self.paragraphs[para_index]['element']

            # 查找所有w:r元素
            r_elements = paragraph.findall("./w:r", self.NAMESPACES)
            if not r_elements:
                print(f"段落{para_index}中没有找到文本运行")
                return False

            # 修改每个文本运行的删除线属性
            for r in r_elements:
                # 查找或创建rPr元素
                rPr = r.find(f".//{{{self.NAMESPACES['w']}}}rPr")
                if rPr is None:
                    rPr = ET.Element(f"{{{self.NAMESPACES['w']}}}rPr")
                    r.insert(0, rPr)

                # 查找删除线元素
                strike_elem = rPr.find(f".//{{{self.NAMESPACES['w']}}}strike")

                # 根据参数设置或移除删除线
                if strike:
                    if strike_elem is None:
                        strike_elem = ET.Element(f"{{{self.NAMESPACES['w']}}}strike")
                        rPr.append(strike_elem)
                    strike_elem.set(f"{{{self.NAMESPACES['w']}}}val", "true")
                elif strike_elem is not None:
                    rPr.remove(strike_elem)

            return True
        except Exception as e:
            print(f"设置段落文本运行删除线格式时出错: {e}")
            return False

    def set_runs_caps(self, para_index, caps=True):
        """设置段落中所有文本运行的大写格式

        Args:
            para_index: 段落索引
            caps: 是否全部大写，True为全部大写，False为正常大小写

        Returns:
            bool: 是否成功修改
        """
        # 检查索引是否有效
        if para_index < 0 or para_index >= len(self.paragraphs):
            print(f"错误：段落索引{para_index}超出范围(0-{len(self.paragraphs)-1})")
            return False

        try:
            # 获取段落元素
            paragraph = self.paragraphs[para_index]['element']

            # 查找所有w:r元素
            r_elements = paragraph.findall("./w:r", self.NAMESPACES)
            if not r_elements:
                print(f"段落{para_index}中没有找到文本运行")
                return False

            # 修改每个文本运行的大写属性
            for r in r_elements:
                # 查找或创建rPr元素
                rPr = r.find(f".//{{{self.NAMESPACES['w']}}}rPr")
                if rPr is None:
                    rPr = ET.Element(f"{{{self.NAMESPACES['w']}}}rPr")
                    r.insert(0, rPr)

                # 查找大写元素
                caps_elem = rPr.find(f".//{{{self.NAMESPACES['w']}}}caps")

                # 根据参数设置或移除大写
                if caps:
                    if caps_elem is None:
                        caps_elem = ET.Element(f"{{{self.NAMESPACES['w']}}}caps")
                        rPr.append(caps_elem)
                    caps_elem.set(f"{{{self.NAMESPACES['w']}}}val", "true")
                elif caps_elem is not None:
                    rPr.remove(caps_elem)

            return True
        except Exception as e:
            print(f"设置段落文本运行大写格式时出错: {e}")
            return False

    def set_runs_vertical_alignment(self, para_index, alignment):
        """设置段落中所有文本运行的垂直对齐方式(上标/下标)

        Args:
            para_index: 段落索引
            alignment: 垂直对齐方式，可以是'superscript'(上标)、'subscript'(下标)、'baseline'(基线)，None表示移除设置

        Returns:
            bool: 是否成功修改
        """
        # 检查索引是否有效
        if para_index < 0 or para_index >= len(self.paragraphs):
            print(f"错误：段落索引{para_index}超出范围(0-{len(self.paragraphs)-1})")
            return False

        try:
            # 获取段落元素
            paragraph = self.paragraphs[para_index]['element']

            # 查找所有w:r元素
            r_elements = paragraph.findall("./w:r", self.NAMESPACES)
            if not r_elements:
                print(f"段落{para_index}中没有找到文本运行")
                return False

            # 修改每个文本运行的垂直对齐方式
            for r in r_elements:
                # 查找或创建rPr元素
                rPr = r.find(f".//{{{self.NAMESPACES['w']}}}rPr")
                if rPr is None:
                    rPr = ET.Element(f"{{{self.NAMESPACES['w']}}}rPr")
                    r.insert(0, rPr)

                # 查找垂直对齐元素
                vert_align = rPr.find(f".//{{{self.NAMESPACES['w']}}}vertAlign")

                # 设置垂直对齐
                if alignment is None:
                    if vert_align is not None:
                        rPr.remove(vert_align)
                else:
                    if vert_align is None:
                        vert_align = ET.Element(f"{{{self.NAMESPACES['w']}}}vertAlign")
                        rPr.append(vert_align)
                    vert_align.set(f"{{{self.NAMESPACES['w']}}}val", alignment)

            return True
        except Exception as e:
            print(f"设置段落文本运行垂直对齐方式时出错: {e}")
            return False

    def update_runs_style(self, para_index, **style_properties):
        """更新段落中所有文本运行的多个样式属性

        Args:
            para_index: 段落索引
            **style_properties: 样式属性字典，可包含以下键：
                'fonts': 字体设置字典，包含'ascii', 'eastAsia'等键
                'size': 字号值
                'bold': 是否加粗
                'italic': 是否斜体
                'underline': 下划线类型
                'color': 字体颜色
                'highlight': 高亮颜色
                'strike': 是否添加删除线
                'caps': 是否全部大写
                'vert_align': 垂直对齐方式

        Returns:
            bool: 是否成功更新所有样式
        """
        # 检查索引是否有效
        if para_index < 0 or para_index >= len(self.paragraphs):
            print(f"错误：段落索引{para_index}超出范围(0-{len(self.paragraphs)-1})")
            return False

        try:
            # 获取段落元素
            paragraph = self.paragraphs[para_index]['element']

            # 查找所有w:r元素
            r_elements = paragraph.findall("./w:r", self.NAMESPACES)
            if not r_elements:
                print(f"段落{para_index}中没有找到文本运行")
                return False

            # 对每个文本运行应用样式属性
            for r in r_elements:
                # 查找或创建rPr元素
                rPr = r.find(f".//{{{self.NAMESPACES['w']}}}rPr")
                if rPr is None:
                    rPr = ET.Element(f"{{{self.NAMESPACES['w']}}}rPr")
                    r.insert(0, rPr)

                # 设置字体
                if 'fonts' in style_properties and isinstance(style_properties['fonts'], dict):
                    fonts = style_properties['fonts']
                    if any(font_type in fonts for font_type in ['ascii', 'hAnsi', 'eastAsia', 'cs']):
                        rFonts = rPr.find(f".//{{{self.NAMESPACES['w']}}}rFonts")
                        if rFonts is None:
                            rFonts = ET.Element(f"{{{self.NAMESPACES['w']}}}rFonts")
                            rPr.append(rFonts)

                        for font_type in ['ascii', 'hAnsi', 'eastAsia', 'cs']:
                            if font_type in fonts:
                                rFonts.set(f"{{{self.NAMESPACES['w']}}}{font_type}", fonts[font_type])

                # 设置字号
                if 'size' in style_properties:
                    sz = rPr.find(f".//{{{self.NAMESPACES['w']}}}sz")
                    if sz is None:
                        sz = ET.Element(f"{{{self.NAMESPACES['w']}}}sz")
                        rPr.append(sz)
                    sz.set(f"{{{self.NAMESPACES['w']}}}val", str(style_properties['size']))

                # 设置加粗
                if 'bold' in style_properties:
                    b = rPr.find(f".//{{{self.NAMESPACES['w']}}}b")
                    if style_properties['bold']:
                        if b is None:
                            b = ET.Element(f"{{{self.NAMESPACES['w']}}}b")
                            rPr.append(b)
                        b.set(f"{{{self.NAMESPACES['w']}}}val", "true")
                    elif b is not None:
                        rPr.remove(b)

                # 设置斜体
                if 'italic' in style_properties:
                    i = rPr.find(f".//{{{self.NAMESPACES['w']}}}i")
                    if style_properties['italic']:
                        if i is None:
                            i = ET.Element(f"{{{self.NAMESPACES['w']}}}i")
                            rPr.append(i)
                        i.set(f"{{{self.NAMESPACES['w']}}}val", "true")
                    elif i is not None:
                        rPr.remove(i)

                # 设置下划线
                if 'underline' in style_properties:
                    u = rPr.find(f".//{{{self.NAMESPACES['w']}}}u")
                    if style_properties['underline'] is None:
                        if u is not None:
                            rPr.remove(u)
                    else:
                        if u is None:
                            u = ET.Element(f"{{{self.NAMESPACES['w']}}}u")
                            rPr.append(u)
                        u.set(f"{{{self.NAMESPACES['w']}}}val", style_properties['underline'])

                # 设置颜色
                if 'color' in style_properties:
                    color = rPr.find(f".//{{{self.NAMESPACES['w']}}}color")
                    if style_properties['color'] is None:
                        if color is not None:
                            rPr.remove(color)
                    else:
                        if color is None:
                            color = ET.Element(f"{{{self.NAMESPACES['w']}}}color")
                            rPr.append(color)
                        color.set(f"{{{self.NAMESPACES['w']}}}val", style_properties['color'])

                # 设置高亮
                if 'highlight' in style_properties:
                    highlight = rPr.find(f".//{{{self.NAMESPACES['w']}}}highlight")
                    if style_properties['highlight'] is None:
                        if highlight is not None:
                            rPr.remove(highlight)
                    else:
                        if highlight is None:
                            highlight = ET.Element(f"{{{self.NAMESPACES['w']}}}highlight")
                            rPr.append(highlight)
                        highlight.set(f"{{{self.NAMESPACES['w']}}}val", style_properties['highlight'])

                # 设置删除线
                if 'strike' in style_properties:
                    strike = rPr.find(f".//{{{self.NAMESPACES['w']}}}strike")
                    if style_properties['strike']:
                        if strike is None:
                            strike = ET.Element(f"{{{self.NAMESPACES['w']}}}strike")
                            rPr.append(strike)
                        strike.set(f"{{{self.NAMESPACES['w']}}}val", "true")
                    elif strike is not None:
                        rPr.remove(strike)

                # 设置大写
                if 'caps' in style_properties:
                    caps = rPr.find(f".//{{{self.NAMESPACES['w']}}}caps")
                    if style_properties['caps']:
                        if caps is None:
                            caps = ET.Element(f"{{{self.NAMESPACES['w']}}}caps")
                            rPr.append(caps)
                        caps.set(f"{{{self.NAMESPACES['w']}}}val", "true")
                    elif caps is not None:
                        rPr.remove(caps)

                # 设置垂直对齐
                if 'vert_align' in style_properties:
                    vert_align = rPr.find(f".//{{{self.NAMESPACES['w']}}}vertAlign")
                    if style_properties['vert_align'] is None:
                        if vert_align is not None:
                            rPr.remove(vert_align)
                    else:
                        if vert_align is None:
                            vert_align = ET.Element(f"{{{self.NAMESPACES['w']}}}vertAlign")
                            rPr.append(vert_align)
                        vert_align.set(f"{{{self.NAMESPACES['w']}}}val", style_properties['vert_align'])

            return True
        except Exception as e:
            print(f"更新段落文本运行样式时出错: {e}")
            return False

    def update_runs_style_from_xml(self, para_element, **style_properties):
        """更新段落中所有文本运行的多个样式属性

        Args:
            para_element: 段落元素
            **style_properties: 样式属性字典，可包含以下键：
                'fonts': 字体设置字典，包含'ascii', 'eastAsia'等键
                'size': 字号值
                'bold': 是否加粗
                'italic': 是否斜体
                'underline': 下划线类型
                'color': 字体颜色
                'highlight': 高亮颜色
                'strike': 是否添加删除线
                'caps': 是否全部大写
                'vert_align': 垂直对齐方式

        Returns:
            bool: 是否成功更新所有样式
        """

        try:
            # 获取段落元素
            paragraph = para_element

            # 查找所有w:r元素
            r_elements = paragraph.findall("./w:r", self.NAMESPACES)

            # 对每个文本运行应用样式属性
            for r in r_elements:
                # 查找或创建rPr元素
                rPr = r.find(f".//{{{self.NAMESPACES['w']}}}rPr")
                if rPr is None:
                    rPr = ET.Element(f"{{{self.NAMESPACES['w']}}}rPr")
                    r.insert(0, rPr)

                # 设置字体
                if 'fonts' in style_properties and isinstance(style_properties['fonts'], dict):
                    fonts = style_properties['fonts']
                    if any(font_type in fonts for font_type in ['ascii', 'hAnsi', 'eastAsia', 'cs']):
                        rFonts = rPr.find(f".//{{{self.NAMESPACES['w']}}}rFonts")
                        if rFonts is None:
                            rFonts = ET.Element(f"{{{self.NAMESPACES['w']}}}rFonts")
                            rPr.append(rFonts)

                        for font_type in ['ascii', 'hAnsi', 'eastAsia', 'cs']:
                            if font_type in fonts:
                                rFonts.set(f"{{{self.NAMESPACES['w']}}}{font_type}", fonts[font_type])

                # 设置字号
                if 'size' in style_properties:
                    sz = rPr.find(f".//{{{self.NAMESPACES['w']}}}sz")
                    if sz is None:
                        sz = ET.Element(f"{{{self.NAMESPACES['w']}}}sz")
                        rPr.append(sz)
                    sz.set(f"{{{self.NAMESPACES['w']}}}val", str(style_properties['size']))

                # 设置加粗
                if 'bold' in style_properties:
                    b = rPr.find(f".//{{{self.NAMESPACES['w']}}}b")
                    if style_properties['bold']:
                        if b is None:
                            b = ET.Element(f"{{{self.NAMESPACES['w']}}}b")
                            rPr.append(b)
                        b.set(f"{{{self.NAMESPACES['w']}}}val", "true")
                    elif b is not None:
                        rPr.remove(b)

                # 设置斜体
                if 'italic' in style_properties:
                    i = rPr.find(f".//{{{self.NAMESPACES['w']}}}i")
                    if style_properties['italic']:
                        if i is None:
                            i = ET.Element(f"{{{self.NAMESPACES['w']}}}i")
                            rPr.append(i)
                        i.set(f"{{{self.NAMESPACES['w']}}}val", "true")
                    elif i is not None:
                        rPr.remove(i)

                # 设置下划线
                if 'underline' in style_properties:
                    u = rPr.find(f".//{{{self.NAMESPACES['w']}}}u")
                    if style_properties['underline'] is None:
                        if u is not None:
                            rPr.remove(u)
                    else:
                        if u is None:
                            u = ET.Element(f"{{{self.NAMESPACES['w']}}}u")
                            rPr.append(u)
                        u.set(f"{{{self.NAMESPACES['w']}}}val", style_properties['underline'])

                # 设置颜色
                if 'color' in style_properties:
                    color = rPr.find(f".//{{{self.NAMESPACES['w']}}}color")
                    if style_properties['color'] is None:
                        if color is not None:
                            rPr.remove(color)
                    else:
                        if color is None:
                            color = ET.Element(f"{{{self.NAMESPACES['w']}}}color")
                            rPr.append(color)
                        color.set(f"{{{self.NAMESPACES['w']}}}val", style_properties['color'])

                # 设置高亮
                if 'highlight' in style_properties:
                    highlight = rPr.find(f".//{{{self.NAMESPACES['w']}}}highlight")
                    if style_properties['highlight'] is None:
                        if highlight is not None:
                            rPr.remove(highlight)
                    else:
                        if highlight is None:
                            highlight = ET.Element(f"{{{self.NAMESPACES['w']}}}highlight")
                            rPr.append(highlight)
                        highlight.set(f"{{{self.NAMESPACES['w']}}}val", style_properties['highlight'])

                # 设置删除线
                if 'strike' in style_properties:
                    strike = rPr.find(f".//{{{self.NAMESPACES['w']}}}strike")
                    if style_properties['strike']:
                        if strike is None:
                            strike = ET.Element(f"{{{self.NAMESPACES['w']}}}strike")
                            rPr.append(strike)
                        strike.set(f"{{{self.NAMESPACES['w']}}}val", "true")
                    elif strike is not None:
                        rPr.remove(strike)

                # 设置大写
                if 'caps' in style_properties:
                    caps = rPr.find(f".//{{{self.NAMESPACES['w']}}}caps")
                    if style_properties['caps']:
                        if caps is None:
                            caps = ET.Element(f"{{{self.NAMESPACES['w']}}}caps")
                            rPr.append(caps)
                        caps.set(f"{{{self.NAMESPACES['w']}}}val", "true")
                    elif caps is not None:
                        rPr.remove(caps)

                # 设置垂直对齐
                if 'vert_align' in style_properties:
                    vert_align = rPr.find(f".//{{{self.NAMESPACES['w']}}}vertAlign")
                    if style_properties['vert_align'] is None:
                        if vert_align is not None:
                            rPr.remove(vert_align)
                    else:
                        if vert_align is None:
                            vert_align = ET.Element(f"{{{self.NAMESPACES['w']}}}vertAlign")
                            rPr.append(vert_align)
                        vert_align.set(f"{{{self.NAMESPACES['w']}}}val", style_properties['vert_align'])

            return True
        except Exception as e:
            print(f"更新段落文本运行样式时出错: {e}")
            return False
    # 以下是修改单个文本运行的样式函数
    def get_run_element(self, para_index, run_index):
        """获取特定段落中的特定文本运行元素

        Args:
            para_index: 段落索引
            run_index: 文本运行索引

        Returns:
            Element或None: 找到的文本运行元素，未找到则返回None
        """
        # 检查段落索引是否有效
        if para_index < 0 or para_index >= len(self.paragraphs):
            print(f"错误：段落索引{para_index}超出范围(0-{len(self.paragraphs)-1})")
            return None

        # 获取段落元素
        paragraph = self.paragraphs[para_index]['element']

        result = self.get_run_element_from_xml(paragraph, run_index)

        # 返回特定的文本运行元素
        return result
    def _get_run_element(self, para_index, run_index):
        """获取特定段落中的特定文本运行元素

        Args:
            para_index: 段落索引
            run_index: 文本运行索引

        Returns:
            Element或None: 找到的文本运行元素，未找到则返回None
        """
        # 检查段落索引是否有效
        if para_index < 0 or para_index >= len(self.paragraphs):
            print(f"错误2：段落索引{para_index}超出范围(0-{len(self.paragraphs)-1})")
            return None

        # 获取段落元素
        paragraph = self.elements[para_index]['element']

        result=self.get_run_element_from_xml(paragraph,run_index)





        # 返回特定的文本运行元素
        return result

    def get_run_element_from_xml(self, para, run_index):
        """获取特定段落中的特定文本运行元素，如果不存在则创建新的run元素

        Args:
            para: 段落
            run_index: 文本运行索引

        Returns:
            Element: 找到或创建的文本运行元素
        """
        # 获取段落元素
        paragraph = para

        # 查找所有w:r元素
        r_elements = paragraph.findall("./w:r", self.NAMESPACES)

        # 如果段落中没有run元素或索引超出范围，则创建新的run元素
        if not r_elements or run_index < 0 or run_index >= len(r_elements):
            # 在创建前记录日志
            if not r_elements:
                print(f"段落中没有找到文本运行，将创建新的run元素")
            else:
                print(f"文本运行索引{run_index}超出范围(0-{len(r_elements) - 1})，将创建新的run元素")

            # 创建新的run元素
            new_run = ET.Element(f"{{{self.NAMESPACES['w']}}}r")

            # 如果段落中有其他元素，找到最后一个run元素的位置
            # 如果没有run元素，则直接追加到段落末尾
            if r_elements:
                # 如果run_index超过了现有元素数量，添加到最后
                if run_index >= len(r_elements):
                    paragraph.append(new_run)
                # 否则插入到指定位置
                else:
                    insert_position = list(paragraph).index(r_elements[0]) + max(0, min(run_index, len(r_elements) - 1))
                    paragraph.insert(insert_position, new_run)
            else:
                # 如果段落中没有任何run元素，直接添加
                paragraph.append(new_run)

            # 重新获取r_elements以确保新创建的元素也被包含
            r_elements = paragraph.findall("./w:r", self.NAMESPACES)

            print(f"已创建新的run元素，现在段落中有 {len(r_elements)} 个run元素")

            # 返回新创建的run元素
            return new_run if run_index >= len(r_elements) else r_elements[run_index]

        # 返回特定的文本运行元素
        return r_elements[run_index]

    def _get_or_create_rPr(self, r_element):
        """获取或创建文本运行属性元素

        Args:
            r_element: 文本运行元素

        Returns:
            Element: 文本运行属性元素
        """
        # 查找rPr元素
        rPr = r_element.find(f".//{{{self.NAMESPACES['w']}}}rPr")
        if rPr is None:
            # 如果不存在，则创建
            rPr = ET.Element(f"{{{self.NAMESPACES['w']}}}rPr")
            r_element.insert(0, rPr)
        return rPr

    def get_run_count(self, para_index):
        """获取段落中直接子run的数量（不包含嵌套段落的run）"""
        try:
            para_element = self.elements[para_index].get('element')
            if para_element is None:
                return 0

            # 只查找直接子run元素，不包括嵌套段落中的run
            # 注意：改用XPath表达式"./w:r"而不是".//w:r"
            runs = para_element.findall("./w:r", self.NAMESPACES)
            return len(runs)
        except Exception as e:
            print(f"获取段落{para_index}的run数量时出错: {e}")
            return 0


    def get_run_count_from_xml(self, para_index):
        """获取段落中直接子run的数量（不包含嵌套段落的run）"""
        try:
            para_element = para_index
            if para_element is None:
               return 0

            # 只查找直接子run元素，不包括嵌套段落中的run
            # 注意：改用XPath表达式"./w:r"而不是".//w:r"
            runs = para_element.findall("./w:r", self.NAMESPACES)
            return len(runs)
        except Exception as e:
            print(f"获取段落{para_index}的run数量时出错: {e}")
            return 0
    def set_table_alignment(self, table_index, alignment):
        """设置表格的对齐方式

        Args:
            table_index: 表格索引
            alignment: 对齐方式，可以是'left'、'center'或'right'

        Returns:
            bool: 是否成功设置
        """
        # 检查表格索引是否有效
        if not hasattr(self, 'tables') or table_index < 0 or table_index >= len(self.tables):
            print(f"错误：表格索引{table_index}无效")
            return False

        try:
            # 获取表格元素
            table_element = self.tables[table_index]['element']

            # 查找tblPr元素，如果不存在则创建
            tbl_pr = table_element.find(f".//{{{self.NAMESPACES['w']}}}tblPr")
            if tbl_pr is None:
                tbl_pr = ET.Element(f"{{{self.NAMESPACES['w']}}}tblPr")
                table_element.insert(0, tbl_pr)

            # 查找jc元素，如果不存在则创建
            jc = tbl_pr.find(f".//{{{self.NAMESPACES['w']}}}jc")
            if jc is None:
                jc = ET.Element(f"{{{self.NAMESPACES['w']}}}jc")
                tbl_pr.append(jc)

            # 设置对齐方式
            jc.set(f"{{{self.NAMESPACES['w']}}}val", alignment)

            # 更新XML
            self.update_document_xml()

            print(f"已将表格{table_index}的对齐方式设置为{alignment}")
            return True

        except Exception as e:
            print(f"设置表格对齐方式时出错: {e}")
            import traceback
            traceback.print_exc()
            return False

    def set_table_property(self, table_index, property_name, property_value):
        """设置表格的通用属性

        Args:
            table_index: 表格索引
            property_name: 属性名称，如'jc'（对齐）、'width'（宽度）等
            property_value: 属性值

        Returns:
            bool: 是否成功设置
        """
        # 检查表格索引是否有效
        if not hasattr(self, 'tables') or table_index < 0 or table_index >= len(self.tables):
            print(f"错误：表格索引{table_index}无效")
            return False

        try:
            # 获取表格元素
            table_element = self.tables[table_index]['element']

            # 查找tblPr元素，如果不存在则创建
            tbl_pr = table_element.find(f".//{{{self.NAMESPACES['w']}}}tblPr")
            if tbl_pr is None:
                tbl_pr = ET.Element(f"{{{self.NAMESPACES['w']}}}tblPr")
                table_element.insert(0, tbl_pr)

            # 根据属性名称处理不同类型的属性
            if property_name == 'jc':  # 对齐方式
                property_element = tbl_pr.find(f".//{{{self.NAMESPACES['w']}}}{property_name}")
                if property_element is None:
                    property_element = ET.Element(f"{{{self.NAMESPACES['w']}}}{property_name}")
                    tbl_pr.append(property_element)
                property_element.set(f"{{{self.NAMESPACES['w']}}}val", property_value)

            elif property_name == 'tblW':  # 表格宽度
                property_element = tbl_pr.find(f".//{{{self.NAMESPACES['w']}}}{property_name}")
                if property_element is None:
                    property_element = ET.Element(f"{{{self.NAMESPACES['w']}}}{property_name}")
                    tbl_pr.append(property_element)

                # 设置宽度和类型
                if isinstance(property_value, dict):
                    if 'w' in property_value:
                        property_element.set(f"{{{self.NAMESPACES['w']}}}w", str(property_value['w']))
                    if 'type' in property_value:
                        property_element.set(f"{{{self.NAMESPACES['w']}}}type", property_value['type'])
                else:
                    property_element.set(f"{{{self.NAMESPACES['w']}}}w", str(property_value))
                    property_element.set(f"{{{self.NAMESPACES['w']}}}type", "dxa")  # 默认使用dxa单位

            elif property_name == 'tblLook':  # 表格外观
                property_element = tbl_pr.find(f".//{{{self.NAMESPACES['w']}}}{property_name}")
                if property_element is None:
                    property_element = ET.Element(f"{{{self.NAMESPACES['w']}}}{property_name}")
                    tbl_pr.append(property_element)

                # 设置外观属性
                if isinstance(property_value, dict):
                    for attr, val in property_value.items():
                        property_element.set(f"{{{self.NAMESPACES['w']}}}{attr}", str(val))
                else:
                    property_element.set(f"{{{self.NAMESPACES['w']}}}val", str(property_value))

            elif property_name == 'tblStyle':  # 表格样式
                property_element = tbl_pr.find(f".//{{{self.NAMESPACES['w']}}}{property_name}")
                if property_element is None:
                    property_element = ET.Element(f"{{{self.NAMESPACES['w']}}}{property_name}")
                    tbl_pr.append(property_element)
                property_element.set(f"{{{self.NAMESPACES['w']}}}val", property_value)

            else:  # 通用属性处理
                property_element = tbl_pr.find(f".//{{{self.NAMESPACES['w']}}}{property_name}")
                if property_element is None:
                    property_element = ET.Element(f"{{{self.NAMESPACES['w']}}}{property_name}")
                    tbl_pr.append(property_element)

                # 如果属性值是字典，则设置多个属性
                if isinstance(property_value, dict):
                    for attr, val in property_value.items():
                        property_element.set(f"{{{self.NAMESPACES['w']}}}{attr}", str(val))
                else:
                    # 否则设置val属性
                    property_element.set(f"{{{self.NAMESPACES['w']}}}val", str(property_value))

            # 更新XML
            self.update_document_xml()

            print(f"已将表格{table_index}的{property_name}属性设置为{property_value}")
            return True

        except Exception as e:
            print(f"设置表格属性时出错: {e}")
            import traceback
            traceback.print_exc()
            return False
    def get_run_text(self, para_index, run_index):
        """获取特定文本运行的文本内容

        Args:
            para_index: 段落索引
            run_index: 文本运行索引

        Returns:
            str: 文本运行的文本内容，如果找不到则返回空字符串
        """
        # 获取文本运行元素
        r_element = self.get_run_element(para_index, run_index)
        if r_element is None:
            return ""

        # 查找所有w:t元素
        t_elements = r_element.findall(f".//{{{self.NAMESPACES['w']}}}t")

        # 拼接文本内容
        text = ""
        for t in t_elements:
            # 获取xml:space属性，确定是否保留空格
            preserve = t.get(f"{{{self.NAMESPACES['xml']}}}space") == "preserve"
            # 获取文本，如果需要保留空格，则不去除前后空格
            if preserve:
                text += t.text if t.text else ""
            else:
                text += t.text.strip() if t.text else ""

        return text

    def _get_run_text(self, para_index, run_index):
        """获取特定文本运行的文本内容

        Args:
            para_index: 段落索引
            run_index: 文本运行索引

        Returns:
            str: 文本运行的文本内容，如果找不到则返回空字符串
        """
        # 获取文本运行元素
        r_element = self._get_run_element(para_index, run_index)
        if r_element is None:
            return ""

        # 查找所有w:t元素
        t_elements = r_element.findall(f".//{{{self.NAMESPACES['w']}}}t")

        # 拼接文本内容
        text = ""
        for t in t_elements:
            # 获取xml:space属性，确定是否保留空格
            preserve = t.get(f"{{{self.NAMESPACES['xml']}}}space") == "preserve"
            # 获取文本，如果需要保留空格，则不去除前后空格
            if preserve:
                text += t.text if t.text else ""
            else:
                text += t.text.strip() if t.text else ""

        return text
    def set_run_font(self, para_index, run_index, **font_properties):
        """设置指定Run元素的字体属性

        Args:
            para_index: 段落索引
            run_index: Run元素索引
            **font_properties: 字体属性，可包含以下键：
                ascii: ASCII字体名称
                eastAsia: 东亚字体名称
                hAnsi: HANSI字体名称
                cs: 复杂脚本字体名称
                size: 字体大小（磅值）
                bold: 是否加粗 (True/False)
                italic: 是否斜体 (True/False)
                underline: 下划线类型
                color: 颜色值（如"FF0000"）

        Returns:
            bool: 是否成功修改
        """
        # 检查索引是否有效
        if para_index < 0 or para_index >= len(self.paragraphs):
            print(f"错误：段落索引{para_index}超出范围(0-{len(self.paragraphs) - 1})")
            return False

        # 获取段落元素
        paragraph = self.paragraphs[para_index]
        para_element = paragraph.get('element')

        # 查找所有w:r元素
        r_elements = para_element.findall(f".//{{{self.NAMESPACES['w']}}}r")
        if run_index < 0 or run_index >= len(r_elements):
            print(f"错误：Run索引{run_index}超出范围(0-{len(r_elements) - 1})")
            return False

        try:
            # 获取指定的Run元素
            run = r_elements[run_index]

            # 获取或创建rPr元素
            rPr = run.find(f"./w:rPr", self.NAMESPACES)
            if rPr is None:
                rPr = ET.Element(f"{{{self.NAMESPACES['w']}}}rPr")
                # 将rPr插入到Run的第一个位置
                run.insert(0, rPr)

            # 设置字体名称
            font_keys = ['ascii', 'eastAsia', 'hAnsi', 'cs']
            if any(key in font_properties for key in font_keys):
                # 查找或创建字体元素
                rFonts = rPr.find(f"./w:rFonts", self.NAMESPACES)
                if rFonts is None:
                    rFonts = ET.Element(f"{{{self.NAMESPACES['w']}}}rFonts")
                    rPr.append(rFonts)

                # 设置各种字体
                for key in font_keys:
                    if key in font_properties and font_properties[key]:
                        rFonts.set(f"{{{self.NAMESPACES['w']}}}{key}", font_properties[key])
                        print(f"设置字体属性 {key}: {font_properties[key]}")

            # 设置字体大小
            if 'size' in font_properties and font_properties['size']:
                size_value = font_properties['size']
                # 确保size是整数或浮点数
                if isinstance(size_value, (int, float)):
                    size = str(int(size_value * 2))  # 转换为半磅单位
                elif isinstance(size_value, str) and size_value.replace('.', '', 1).isdigit():
                    size = str(int(float(size_value) * 2))
                else:
                    size = size_value  # 保留原始值

                # 查找或创建sz元素
                sz = rPr.find(f"./w:sz", self.NAMESPACES)
                if sz is None:
                    sz = ET.Element(f"{{{self.NAMESPACES['w']}}}sz")
                    rPr.append(sz)
                sz.set(f"{{{self.NAMESPACES['w']}}}val", size)
                print(f"设置字体大小: {size} (原始值: {font_properties['size']}磅)")

                # 同时设置szCs（复杂脚本字体大小）
                szCs = rPr.find(f"./w:szCs", self.NAMESPACES)
                if szCs is None:
                    szCs = ET.Element(f"{{{self.NAMESPACES['w']}}}szCs")
                    rPr.append(szCs)
                szCs.set(f"{{{self.NAMESPACES['w']}}}val", size)

            # 设置加粗
            if 'bold' in font_properties:
                # 查找或创建b元素
                bold = rPr.find(f"./w:b", self.NAMESPACES)
                if font_properties['bold']:
                    if bold is None:
                        bold = ET.Element(f"{{{self.NAMESPACES['w']}}}b")
                        rPr.append(bold)
                    # 移除val属性，在Word中表示启用
                    if f"{{{self.NAMESPACES['w']}}}val" in bold.attrib:
                        del bold.attrib[f"{{{self.NAMESPACES['w']}}}val"]
                    print("设置加粗: 是")
                else:
                    if bold is not None:
                        rPr.remove(bold)
                    print("移除加粗")

            # 设置斜体
            if 'italic' in font_properties:
                # 查找或创建i元素
                italic = rPr.find(f"./w:i", self.NAMESPACES)
                if font_properties['italic']:
                    if italic is None:
                        italic = ET.Element(f"{{{self.NAMESPACES['w']}}}i")
                        rPr.append(italic)
                    # 移除val属性，在Word中表示启用
                    if f"{{{self.NAMESPACES['w']}}}val" in italic.attrib:
                        del italic.attrib[f"{{{self.NAMESPACES['w']}}}val"]
                    print("设置斜体: 是")
                else:
                    if italic is not None:
                        rPr.remove(italic)
                    print("移除斜体")

            # 设置下划线
            if 'underline' in font_properties and font_properties['underline']:
                # 查找或创建u元素
                u = rPr.find(f"./w:u", self.NAMESPACES)
                if u is None:
                    u = ET.Element(f"{{{self.NAMESPACES['w']}}}u")
                    rPr.append(u)
                u.set(f"{{{self.NAMESPACES['w']}}}val", font_properties['underline'])
                print(f"设置下划线: {font_properties['underline']}")

            # 设置颜色
            if 'color' in font_properties and font_properties['color']:
                # 查找或创建color元素
                color = rPr.find(f"./w:color", self.NAMESPACES)
                if color is None:
                    color = ET.Element(f"{{{self.NAMESPACES['w']}}}color")
                    rPr.append(color)
                color.set(f"{{{self.NAMESPACES['w']}}}val", font_properties['color'])
                print(f"设置颜色: {font_properties['color']}")

            # 打印XML确认
            xml_str = ET.tostring(run, encoding='utf-8').decode('utf-8')
            print(f"修改后的XML: {xml_str}")

            # 更新文档XML
            self.update_document_xml()
            return True
        except Exception as e:
            print(f"设置Run字体属性时出错: {e}")
            import traceback
            traceback.print_exc()
            return False

    def set_run_size(self, para_index, run_index, size):
        """设置特定文本运行的字号

        Args:
            para_index: 段落索引
            run_index: 文本运行索引
            size: 字号值，可以是：
                - 整数或浮点数
                - 字符串形式的数字
                - None（移除字号设置）

        Returns:
            bool: 是否成功修改
        """
        # 获取文本运行元素
        r_element = self.get_run_element(para_index, run_index)
        if r_element is None:
            print(f"错误：找不到段落{para_index}的Run{run_index}")
            return False

        try:
            # 获取或创建rPr元素
            rPr = self._get_or_create_rPr(r_element)

            # 转换size值为字符串
            if size is not None:
                if isinstance(size, (int, float)):
                    size_value = str(int(size))
                elif isinstance(size, str) and size.replace('.', '', 1).isdigit():
                    size_value = str(int(float(size)))
                else:
                    print(f"错误：无效的字号值 {size}")
                    return False
                print(f"设置字号: {size_value}")

            # 处理sz元素
            sz = rPr.find(f".//{{{self.NAMESPACES['w']}}}sz")
            if size is None:
                # 如果要移除字号设置
                if sz is not None:
                    rPr.remove(sz)
                    print("移除字号设置")
            else:
                # 如果要设置字号
                if sz is None:
                    sz = ET.Element(f"{{{self.NAMESPACES['w']}}}sz")
                    rPr.append(sz)
                sz.set(f"{{{self.NAMESPACES['w']}}}val", size_value)

            # 处理szCs元素（复杂脚本字体大小）
            szCs = rPr.find(f".//{{{self.NAMESPACES['w']}}}szCs")
            if size is None:
                if szCs is not None:
                    rPr.remove(szCs)
            else:
                if szCs is None:
                    szCs = ET.Element(f"{{{self.NAMESPACES['w']}}}szCs")
                    rPr.append(szCs)
                szCs.set(f"{{{self.NAMESPACES['w']}}}val", size_value)

            # 打印XML确认
            xml_str = ET.tostring(r_element, encoding='utf-8').decode('utf-8')
            print(f"修改后的XML: {xml_str}")

            # 更新文档XML
            self.update_document_xml()
            return True
        except Exception as e:
            print(f"设置文本运行字号时出错: {e}")
            import traceback
            traceback.print_exc()
            return False

    def set_run_bold(self, para_index, run_index, bold=True):
        """设置特定文本运行是否加粗

        Args:
            para_index: 段落索引
            run_index: 文本运行索引
            bold: 是否加粗，True为加粗，False为取消加粗

        Returns:
            bool: 是否成功修改
        """
        # 获取文本运行元素
        r_element = self.get_run_element(para_index, run_index)
        if r_element is None:
            return False

        try:
            # 获取或创建rPr元素
            rPr = self._get_or_create_rPr(r_element)

            # 查找加粗元素
            b = rPr.find(f".//{{{self.NAMESPACES['w']}}}b")

            # 根据参数设置或移除加粗
            if bold:
                if b is None:
                    b = ET.Element(f"{{{self.NAMESPACES['w']}}}b")
                    rPr.append(b)
                b.set(f"{{{self.NAMESPACES['w']}}}val", "true")
            elif b is not None:
                rPr.remove(b)

            return True
        except Exception as e:
            print(f"设置文本运行加粗格式时出错: {e}")
            return False

    def set_run_italic(self, para_index, run_index, italic=True):
        """设置特定文本运行是否斜体

        Args:
            para_index: 段落索引
            run_index: 文本运行索引
            italic: 是否斜体，True为斜体，False为取消斜体

        Returns:
            bool: 是否成功修改
        """
        # 获取文本运行元素
        r_element = self.get_run_element(para_index, run_index)
        if r_element is None:
            return False

        try:
            # 获取或创建rPr元素
            rPr = self._get_or_create_rPr(r_element)

            # 查找斜体元素
            i = rPr.find(f".//{{{self.NAMESPACES['w']}}}i")

            # 根据参数设置或移除斜体
            if italic:
                if i is None:
                    i = ET.Element(f"{{{self.NAMESPACES['w']}}}i")
                    rPr.append(i)
                i.set(f"{{{self.NAMESPACES['w']}}}val", "true")
            elif i is not None:
                rPr.remove(i)

            return True
        except Exception as e:
            print(f"设置文本运行斜体格式时出错: {e}")
            return False

    def set_run_underline(self, para_index, run_index, underline_type='single'):
        """设置特定文本运行的下划线格式

        Args:
            para_index: 段落索引
            run_index: 文本运行索引
            underline_type: 下划线类型，如'single'(单线)、'double'(双线)、'thick'(粗线)
                          'dotted'(点线)、'dash'(虚线)、'wave'(波浪线)，传入None表示移除下划线

        Returns:
            bool: 是否成功修改
        """
        # 获取文本运行元素
        r_element = self.get_run_element(para_index, run_index)
        if r_element is None:
            return False

        try:
            # 获取或创建rPr元素
            rPr = self._get_or_create_rPr(r_element)

            # 查找下划线元素
            u = rPr.find(f".//{{{self.NAMESPACES['w']}}}u")

            # 根据参数设置或移除下划线
            if underline_type is None:
                if u is not None:
                    rPr.remove(u)
            else:
                if u is None:
                    u = ET.Element(f"{{{self.NAMESPACES['w']}}}u")
                    rPr.append(u)
                u.set(f"{{{self.NAMESPACES['w']}}}val", underline_type)

            return True
        except Exception as e:
            print(f"设置文本运行下划线格式时出错: {e}")
            return False

    def set_run_color(self, para_index, run_index, color):
        """设置特定文本运行的颜色

        Args:
            para_index: 段落索引
            run_index: 文本运行索引
            color: 颜色值，如'FF0000'表示红色，传入None表示移除颜色设置

        Returns:
            bool: 是否成功修改
        """
        # 获取文本运行元素
        r_element = self.get_run_element(para_index, run_index)
        if r_element is None:
            return False

        try:
            # 获取或创建rPr元素
            rPr = self._get_or_create_rPr(r_element)

            # 查找颜色元素
            c = rPr.find(f".//{{{self.NAMESPACES['w']}}}color")

            # 根据参数设置或移除颜色
            if color is None:
                if c is not None:
                    rPr.remove(c)
            else:
                if c is None:
                    c = ET.Element(f"{{{self.NAMESPACES['w']}}}color")
                    rPr.append(c)
                c.set(f"{{{self.NAMESPACES['w']}}}val", color)

            return True
        except Exception as e:
            print(f"设置文本运行颜色时出错: {e}")
            return False

    def set_run_highlight(self, para_index, run_index, highlight_color):
        """设置特定文本运行的高亮颜色

        Args:
            para_index: 段落索引
            run_index: 文本运行索引
            highlight_color: 高亮颜色值，如'yellow'、'green'等，传入None表示移除高亮

        Returns:
            bool: 是否成功修改
        """
        # 获取文本运行元素
        r_element = self.get_run_element(para_index, run_index)
        if r_element is None:
            return False

        try:
            # 获取或创建rPr元素
            rPr = self._get_or_create_rPr(r_element)

            # 查找高亮元素
            highlight = rPr.find(f".//{{{self.NAMESPACES['w']}}}highlight")

            # 根据参数设置或移除高亮
            if highlight_color is None:
                if highlight is not None:
                    rPr.remove(highlight)
            else:
                if highlight is None:
                    highlight = ET.Element(f"{{{self.NAMESPACES['w']}}}highlight")
                    rPr.append(highlight)
                highlight.set(f"{{{self.NAMESPACES['w']}}}val", highlight_color)

            return True
        except Exception as e:
            print(f"设置文本运行高亮颜色时出错: {e}")
            return False

    def set_run_strike(self, para_index, run_index, strike=True):
        """设置特定文本运行是否有删除线

        Args:
            para_index: 段落索引
            run_index: 文本运行索引
            strike: 是否添加删除线，True为添加，False为移除

        Returns:
            bool: 是否成功修改
        """
        # 获取文本运行元素
        r_element = self.get_run_element(para_index, run_index)
        if r_element is None:
            return False

        try:
            # 获取或创建rPr元素
            rPr = self._get_or_create_rPr(r_element)

            # 查找删除线元素
            strike_elem = rPr.find(f".//{{{self.NAMESPACES['w']}}}strike")

            # 根据参数设置或移除删除线
            if strike:
                if strike_elem is None:
                    strike_elem = ET.Element(f"{{{self.NAMESPACES['w']}}}strike")
                    rPr.append(strike_elem)
                strike_elem.set(f"{{{self.NAMESPACES['w']}}}val", "true")
            elif strike_elem is not None:
                rPr.remove(strike_elem)

            return True
        except Exception as e:
            print(f"设置文本运行删除线格式时出错: {e}")
            return False

    def get_run_text_from_xml(self, para, run_index):
        """获取特定文本运行的文本内容

        Args:
            para: 段落
            run_index: 文本运行索引

        Returns:
            str: 文本运行的文本内容，如果找不到则返回空字符串
        """
        # 获取文本运行元素
        r_element = self.get_run_element_from_xml(para, run_index)
        if r_element is None:
            return ""

        # 查找所有w:t元素
        t_elements = r_element.findall(f".//{{{self.NAMESPACES['w']}}}t")

        # 拼接文本内容
        text = ""
        for t in t_elements:
            # 获取xml:space属性，确定是否保留空格
            preserve = t.get(f"{{{self.NAMESPACES['xml']}}}space") == "preserve"
            # 获取文本，如果需要保留空格，则不去除前后空格
            if preserve:
                text += t.text if t.text else ""
            else:
                text += t.text.strip() if t.text else ""

        return text

    def set_run_font_from_xml(self, para, run_index, **font_properties):
        """设置指定Run元素的字体属性

        Args:
            para: 段落
            run_index: Run元素索引
            **font_properties: 字体属性，可包含以下键：
                ascii: ASCII字体名称
                eastAsia: 东亚字体名称
                hAnsi: HANSI字体名称
                cs: 复杂脚本字体名称
                size: 字体大小（磅值）
                bold: 是否加粗 (True/False)
                italic: 是否斜体 (True/False)
                underline: 下划线类型
                color: 颜色值（如"FF0000"）

        Returns:
            bool: 是否成功修改
        """


        # 获取段落元素

        para_element =para

        # 查找所有w:r元素
        r_elements = para_element.findall(f".//{{{self.NAMESPACES['w']}}}r")
        if run_index < 0 or run_index >= len(r_elements):
            print(f"错误：Run索引{run_index}超出范围(0-{len(r_elements) - 1})")
            return False

        try:
            # 获取指定的Run元素
            run = r_elements[run_index]

            # 获取或创建rPr元素
            rPr = run.find(f"./w:rPr", self.NAMESPACES)
            if rPr is None:
                rPr = ET.Element(f"{{{self.NAMESPACES['w']}}}rPr")
                # 将rPr插入到Run的第一个位置
                run.insert(0, rPr)

            # 设置字体名称
            font_keys = ['ascii', 'eastAsia', 'hAnsi', 'cs']
            if any(key in font_properties for key in font_keys):
                # 查找或创建字体元素
                rFonts = rPr.find(f"./w:rFonts", self.NAMESPACES)
                if rFonts is None:
                    rFonts = ET.Element(f"{{{self.NAMESPACES['w']}}}rFonts")
                    rPr.append(rFonts)

                # 设置各种字体
                for key in font_keys:
                    if key in font_properties and font_properties[key]:
                        rFonts.set(f"{{{self.NAMESPACES['w']}}}{key}", font_properties[key])
                        print(f"设置字体属性 {key}: {font_properties[key]}")

            # 设置字体大小
            if 'size' in font_properties and font_properties['size']:
                size_value = font_properties['size']
                # 确保size是整数或浮点数
                if isinstance(size_value, (int, float)):
                    size = str(int(size_value * 2))  # 转换为半磅单位
                elif isinstance(size_value, str) and size_value.replace('.', '', 1).isdigit():
                    size = str(int(float(size_value) * 2))
                else:
                    size = size_value  # 保留原始值

                # 查找或创建sz元素
                sz = rPr.find(f"./w:sz", self.NAMESPACES)
                if sz is None:
                    sz = ET.Element(f"{{{self.NAMESPACES['w']}}}sz")
                    rPr.append(sz)
                sz.set(f"{{{self.NAMESPACES['w']}}}val", size)
                print(f"设置字体大小: {size} (原始值: {font_properties['size']}磅)")

                # 同时设置szCs（复杂脚本字体大小）
                szCs = rPr.find(f"./w:szCs", self.NAMESPACES)
                if szCs is None:
                    szCs = ET.Element(f"{{{self.NAMESPACES['w']}}}szCs")
                    rPr.append(szCs)
                szCs.set(f"{{{self.NAMESPACES['w']}}}val", size)

            # 设置加粗
            if 'bold' in font_properties:
                # 查找或创建b元素
                bold = rPr.find(f"./w:b", self.NAMESPACES)
                if font_properties['bold']:
                    if bold is None:
                        bold = ET.Element(f"{{{self.NAMESPACES['w']}}}b")
                        rPr.append(bold)
                    # 移除val属性，在Word中表示启用
                    if f"{{{self.NAMESPACES['w']}}}val" in bold.attrib:
                        del bold.attrib[f"{{{self.NAMESPACES['w']}}}val"]
                    print("设置加粗: 是")
                else:
                    if bold is not None:
                        rPr.remove(bold)
                    print("移除加粗")

            # 设置斜体
            if 'italic' in font_properties:
                # 查找或创建i元素
                italic = rPr.find(f"./w:i", self.NAMESPACES)
                if font_properties['italic']:
                    if italic is None:
                        italic = ET.Element(f"{{{self.NAMESPACES['w']}}}i")
                        rPr.append(italic)
                    # 移除val属性，在Word中表示启用
                    if f"{{{self.NAMESPACES['w']}}}val" in italic.attrib:
                        del italic.attrib[f"{{{self.NAMESPACES['w']}}}val"]
                    print("设置斜体: 是")
                else:
                    if italic is not None:
                        rPr.remove(italic)
                    print("移除斜体")

            # 设置下划线
            if 'underline' in font_properties and font_properties['underline']:
                # 查找或创建u元素
                u = rPr.find(f"./w:u", self.NAMESPACES)
                if u is None:
                    u = ET.Element(f"{{{self.NAMESPACES['w']}}}u")
                    rPr.append(u)
                u.set(f"{{{self.NAMESPACES['w']}}}val", font_properties['underline'])
                print(f"设置下划线: {font_properties['underline']}")

            # 设置颜色
            if 'color' in font_properties and font_properties['color']:
                # 查找或创建color元素
                color = rPr.find(f"./w:color", self.NAMESPACES)
                if color is None:
                    color = ET.Element(f"{{{self.NAMESPACES['w']}}}color")
                    rPr.append(color)
                color.set(f"{{{self.NAMESPACES['w']}}}val", font_properties['color'])
                print(f"设置颜色: {font_properties['color']}")



            # 更新文档XML
            self.update_document_xml()
            return True
        except Exception as e:
            print(f"设置Run字体属性时出错: {e}")
            import traceback
            traceback.print_exc()
            return False

    def set_run_size_from_xml(self, para, run_index, size):
        """设置特定文本运行的字号

        Args:
            para: 段落
            run_index: 文本运行索引
            size: 字号值，可以是：
                - 整数或浮点数
                - 字符串形式的数字
                - None（移除字号设置）

        Returns:
            bool: 是否成功修改
        """
        # 获取文本运行元素
        r_element = self.get_run_element_from_xml(para, run_index)
        if r_element is None:
            print(f"错误：找不到段落的Run{run_index}")
            return False

        try:
            # 获取或创建rPr元素
            rPr = self._get_or_create_rPr(r_element)

            # 转换size值为字符串
            if size is not None:
                if isinstance(size, (int, float)):
                    size_value = str(int(size))
                elif isinstance(size, str) and size.replace('.', '', 1).isdigit():
                    size_value = str(int(float(size)))
                else:
                    print(f"错误：无效的字号值 {size}")
                    return False
                print(f"设置字号: {size_value}")

            # 处理sz元素
            sz = rPr.find(f".//{{{self.NAMESPACES['w']}}}sz")
            if size is None:
                # 如果要移除字号设置
                if sz is not None:
                    rPr.remove(sz)
                    print("移除字号设置")
            else:
                # 如果要设置字号
                if sz is None:
                    sz = ET.Element(f"{{{self.NAMESPACES['w']}}}sz")
                    rPr.append(sz)
                sz.set(f"{{{self.NAMESPACES['w']}}}val", size_value)

            # 处理szCs元素（复杂脚本字体大小）
            szCs = rPr.find(f".//{{{self.NAMESPACES['w']}}}szCs")
            if size is None:
                if szCs is not None:
                    rPr.remove(szCs)
            else:
                if szCs is None:
                    szCs = ET.Element(f"{{{self.NAMESPACES['w']}}}szCs")
                    rPr.append(szCs)
                szCs.set(f"{{{self.NAMESPACES['w']}}}val", size_value)



            # 更新文档XML
            self.update_document_xml()
            return True
        except Exception as e:
            print(f"设置文本运行字号时出错: {e}")
            import traceback
            traceback.print_exc()
            return False

    def set_run_bold_from_xml(self, para, run_index, bold=True):
        """设置特定文本运行是否加粗

        Args:
            para: 段落
            run_index: 文本运行索引
            bold: 是否加粗，True为加粗，False为取消加粗

        Returns:
            bool: 是否成功修改
        """
        # 获取文本运行元素
        r_element = self.get_run_element_from_xml(para, run_index)
        if r_element is None:
            return False

        try:
            # 获取或创建rPr元素
            rPr = self._get_or_create_rPr(r_element)

            # 查找加粗元素
            b = rPr.find(f".//{{{self.NAMESPACES['w']}}}b")

            # 根据参数设置或移除加粗
            if bold:
                if b is None:
                    b = ET.Element(f"{{{self.NAMESPACES['w']}}}b")
                    rPr.append(b)
                b.set(f"{{{self.NAMESPACES['w']}}}val", "true")
            elif b is not None:
                rPr.remove(b)

            return True
        except Exception as e:
            print(f"设置文本运行加粗格式时出错: {e}")
            return False

    def set_run_italic_from_xml(self, para, run_index, italic=True):
        """设置特定文本运行是否斜体

        Args:
            para: 段落索
            run_index: 文本运行索引
            italic: 是否斜体，True为斜体，False为取消斜体

        Returns:
            bool: 是否成功修改
        """
        # 获取文本运行元素
        r_element = self.get_run_element_from_xml(para, run_index)
        if r_element is None:
            return False

        try:
            # 获取或创建rPr元素
            rPr = self._get_or_create_rPr(r_element)

            # 查找斜体元素
            i = rPr.find(f".//{{{self.NAMESPACES['w']}}}i")

            # 根据参数设置或移除斜体
            if italic:
                if i is None:
                    i = ET.Element(f"{{{self.NAMESPACES['w']}}}i")
                    rPr.append(i)
                i.set(f"{{{self.NAMESPACES['w']}}}val", "true")
            elif i is not None:
                rPr.remove(i)

            return True
        except Exception as e:
            print(f"设置文本运行斜体格式时出错: {e}")
            return False

    def set_run_underline_from_xml(self, para, run_index, underline_type='single'):
        """设置特定文本运行的下划线格式

        Args:
            para: 段落
            run_index: 文本运行索引
            underline_type: 下划线类型，如'single'(单线)、'double'(双线)、'thick'(粗线)
                          'dotted'(点线)、'dash'(虚线)、'wave'(波浪线)，传入None表示移除下划线

        Returns:
            bool: 是否成功修改
        """
        # 获取文本运行元素
        r_element = self.get_run_element_from_xml(para, run_index)
        if r_element is None:
            return False

        try:
            # 获取或创建rPr元素
            rPr = self._get_or_create_rPr(r_element)

            # 查找下划线元素
            u = rPr.find(f".//{{{self.NAMESPACES['w']}}}u")

            # 根据参数设置或移除下划线
            if underline_type is None:
                if u is not None:
                    rPr.remove(u)
            else:
                if u is None:
                    u = ET.Element(f"{{{self.NAMESPACES['w']}}}u")
                    rPr.append(u)
                u.set(f"{{{self.NAMESPACES['w']}}}val", underline_type)

            return True
        except Exception as e:
            print(f"设置文本运行下划线格式时出错: {e}")
            return False

    def set_run_color_from_xml(self, para, run_index, color):
        """设置特定文本运行的颜色

        Args:
            para: 段落
            run_index: 文本运行索引
            color: 颜色值，如'FF0000'表示红色，传入None表示移除颜色设置

        Returns:
            bool: 是否成功修改
        """
        # 获取文本运行元素
        r_element = self.get_run_element_from_xml(para, run_index)
        if r_element is None:
            return False

        try:
            # 获取或创建rPr元素
            rPr = self._get_or_create_rPr(r_element)

            # 查找颜色元素
            c = rPr.find(f".//{{{self.NAMESPACES['w']}}}color")

            # 根据参数设置或移除颜色
            if color is None:
                if c is not None:
                    rPr.remove(c)
            else:
                if c is None:
                    c = ET.Element(f"{{{self.NAMESPACES['w']}}}color")
                    rPr.append(c)
                c.set(f"{{{self.NAMESPACES['w']}}}val", color)

            return True
        except Exception as e:
            print(f"设置文本运行颜色时出错: {e}")
            return False

    def set_run_highlight_from_xml(self, para, run_index, highlight_color):
        """设置特定文本运行的高亮颜色

        Args:
            para: 段落
            run_index: 文本运行索引
            highlight_color: 高亮颜色值，如'yellow'、'green'等，传入None表示移除高亮

        Returns:
            bool: 是否成功修改
        """
        # 获取文本运行元素
        r_element = self.get_run_element_from_xml(para, run_index)
        if r_element is None:
            return False

        try:
            # 获取或创建rPr元素
            rPr = self._get_or_create_rPr(r_element)

            # 查找高亮元素
            highlight = rPr.find(f".//{{{self.NAMESPACES['w']}}}highlight")

            # 根据参数设置或移除高亮
            if highlight_color is None:
                if highlight is not None:
                    rPr.remove(highlight)
            else:
                if highlight is None:
                    highlight = ET.Element(f"{{{self.NAMESPACES['w']}}}highlight")
                    rPr.append(highlight)
                highlight.set(f"{{{self.NAMESPACES['w']}}}val", highlight_color)

            return True
        except Exception as e:
            print(f"设置文本运行高亮颜色时出错: {e}")
            return False

    def set_run_strike_from_xml(self, para, run_index, strike=True):
        """设置特定文本运行是否有删除线

        Args:
            para: 段落
            run_index: 文本运行索引
            strike: 是否添加删除线，True为添加，False为移除

        Returns:
            bool: 是否成功修改
        """
        # 获取文本运行元素
        r_element = self.get_run_element_from_xml(para, run_index)
        if r_element is None:
            return False

        try:
            # 获取或创建rPr元素
            rPr = self._get_or_create_rPr(r_element)

            # 查找删除线元素
            strike_elem = rPr.find(f".//{{{self.NAMESPACES['w']}}}strike")

            # 根据参数设置或移除删除线
            if strike:
                if strike_elem is None:
                    strike_elem = ET.Element(f"{{{self.NAMESPACES['w']}}}strike")
                    rPr.append(strike_elem)
                strike_elem.set(f"{{{self.NAMESPACES['w']}}}val", "true")
            elif strike_elem is not None:
                rPr.remove(strike_elem)

            return True
        except Exception as e:
            print(f"设置文本运行删除线格式时出错: {e}")
            return False

    def delete_runs_after_index_from_xml(self, para, run_index):
        """删除指定run索引之后的所有run元素

        Args:
            para: 段落元素
            run_index: 开始删除的Run索引（保留该索引，删除之后的）

        Returns:
            bool: 操作是否成功
        """
        try:
            # 获取段落中的所有run元素
            r_elements = para.findall(f".//{{{self.NAMESPACES['w']}}}r")

            # 检查索引是否有效
            if run_index < 0 or run_index >= len(r_elements):
                print(f"错误：Run索引{run_index}超出范围(0-{len(r_elements) - 1})")
                return False

            # 特殊处理：如果索引是0，则将第一个run的文本设置为空
            if run_index == 0:
                # 获取第一个run元素
                first_run = r_elements[0]

                # 找出所有t元素（文本元素）
                t_elements = first_run.findall(f".//{{{self.NAMESPACES['w']}}}t")

                # 如果存在t元素，设置第一个为空，删除其他的
                if t_elements:
                    # 保留第一个t元素但清空其内容
                    first_t = t_elements[0]
                    first_t.text = ""
                    # 确保space属性为preserve以保留空格
                    first_t.set(f"{{{self.NAMESPACES['xml']}}}space", "preserve")

                    # 删除其他t元素（如果有）
                    for i in range(len(t_elements) - 1, 0, -1):
                        t_parent = t_elements[i].getparent()
                        if t_parent is not None:
                            t_parent.remove(t_elements[i])
                else:
                    # 如果没有t元素，创建一个
                    t_element = ET.SubElement(first_run, f"{{{self.NAMESPACES['w']}}}t")
                    t_element.set(f"{{{self.NAMESPACES['xml']}}}space", "preserve")
                    t_element.text = ""

                print(f"已将第一个Run元素的文本内容设置为空")

            # 删除指定索引之后的所有run元素
            deleted_count = 0
            for i in range(len(r_elements) - 1, run_index, -1):  # 从后向前删除，避免索引变化
                parent = r_elements[i].getparent()
                if parent is not None:
                    parent.remove(r_elements[i])
                    deleted_count += 1

            # 打印删除数量
            print(f"已从段落中删除 {deleted_count} 个Run元素")

            # 更新XML
            self.update_document_xml()

            return True
        except Exception as e:
            print(f"删除Run元素时出错: {e}")
            import traceback
            traceback.print_exc()
            return False

    def delete_runs_after_index(self, para_index, run_index):
        """删除指定段落中指定run索引之后的所有run元素

        Args:
            para_index: 段落索引
            run_index: 开始删除的Run索引（保留该索引，删除之后的）

        Returns:
            bool: 操作是否成功
        """
        try:
            # 检查段落索引是否有效
            if para_index < 0 or para_index >= len(self.paragraphs):
                print(f"错误：段落索引{para_index}超出范围(0-{len(self.paragraphs) - 1})")
                return False

            # 获取段落元素
            para = self.paragraphs[para_index]

            # 调用XML版本的函数
            return self.delete_runs_after_index_from_xml(para, run_index)
        except Exception as e:
            print(f"删除Run元素时出错: {e}")
            import traceback
            traceback.print_exc()
            return False
    def update_run_style_from_xml(self, para, run_index, **style_properties):
        """更新Run元素的样式

                Args:
                    para: 段落
                    run_index: Run元素索引
                    **style_properties: 样式属性，可包含以下键：
                        fonts: 字体名称字典 {'ascii': 'Arial', 'eastAsia': '宋体', ...}，
                               注意：这里需要使用'fonts'而不是'font'
                        size: 字体大小
                        bold: 是否加粗
                        italic: 是否斜体
                        underline: 下划线类型
                        color: 颜色值
                        highlight: 高亮颜色
                        strike: 是否删除线
                        caps: 是否全大写
                        small_caps: 是否小型大写字母
                        spacing: 字符间距
                        vert_align: 垂直对齐方式

                Returns:
                    bool: 是否成功更新
                """



        try:
            # 获取run元素
            run = self.get_run_element_from_xml(para, run_index)
            if run is None:
                return False

            # 打印调试信息
            print(f"更新段落中Run {run_index} 的样式")
            print(f"样式属性: {style_properties}")

            # 设置文本内容
            if 'text' in style_properties:
                self.set_run_text_from_xml(para, run_index, style_properties['text'])

            # 设置字体
            if 'fonts' in style_properties:
                font_properties = {}
                # 复制字体名称
                for key, value in style_properties['fonts'].items():
                    font_properties[key] = value

                # 调用设置字体函数
                self.set_run_font_from_xml(para, run_index, **font_properties)

            # 设置字体大小
            if 'size' in style_properties:
                self.set_run_size_from_xml(para, run_index, style_properties['size'])

            # 设置加粗
            if 'bold' in style_properties:
                self.set_run_bold_from_xml(para, run_index, style_properties['bold'])

            # 设置斜体
            if 'italic' in style_properties:
                self.set_run_italic_from_xml(para, run_index, style_properties['italic'])

            # 设置下划线
            if 'underline' in style_properties:
                self.set_run_underline_from_xml(para, run_index, style_properties['underline'])

            # 设置颜色
            if 'color' in style_properties:
                self.set_run_color_from_xml(para, run_index, style_properties['color'])

            # 设置高亮
            if 'highlight' in style_properties:
                self.set_run_highlight_from_xml(para, run_index, style_properties['highlight'])

            # 设置删除线
            if 'strike' in style_properties:
                self.set_run_strike_from_xml(para, run_index, style_properties['strike'])

            # 更新XML
            self.update_document_xml()

            # 打印更新后的样式信息
            updated_style = self.get_run_style_form_xml(para, run_index)
            print(f"更新后的样式: {updated_style}")

            return True
        except Exception as e:
            print(f"更新Run样式时出错: {e}")
            import traceback
            traceback.print_exc()
            return False
    def update_run_style(self, para_index, run_index, **style_properties):
        """更新Run元素的样式

        Args:
            para_index: 段落索引
            run_index: Run元素索引
            **style_properties: 样式属性，可包含以下键：
                fonts: 字体名称字典 {'ascii': 'Arial', 'eastAsia': '宋体', ...}，
                       注意：这里需要使用'fonts'而不是'fonts'
                size: 字体大小
                bold: 是否加粗
                italic: 是否斜体
                underline: 下划线类型
                color: 颜色值
                highlight: 高亮颜色
                strike: 是否删除线
                caps: 是否全大写
                small_caps: 是否小型大写字母
                spacing: 字符间距
                vert_align: 垂直对齐方式

        Returns:
            bool: 是否成功更新
        """
        # 检查索引是否有效
        if para_index < 0 or para_index >= len(self.paragraphs):
            print(f"错误：段落索引{para_index}超出范围(0-{len(self.paragraphs) - 1})")
            return False

        try:
            # 获取run元素
            run = self.get_run_element(para_index, run_index)
            if run is None:
                return False

            # 打印调试信息
            print(f"更新段落 {para_index} 中Run {run_index} 的样式")
            print(f"样式属性: {style_properties}")
            # 设置文本内容
            if 'text' in style_properties:
                self.set_run_text(para_index, run_index, style_properties['text'])

            # 设置字体
            if 'fonts' in style_properties:
                font_properties = {}
                # 复制字体名称
                for key, value in style_properties['fonts'].items():
                    font_properties[key] = value

                # 调用设置字体函数
                self.set_run_font(para_index, run_index, **font_properties)

            # 设置字体大小
            if 'size' in style_properties:
                self.set_run_size(para_index, run_index, style_properties['size'])

            # 设置加粗
            if 'bold' in style_properties:
                self.set_run_bold(para_index, run_index, style_properties['bold'])

            # 设置斜体
            if 'italic' in style_properties:
                self.set_run_italic(para_index, run_index, style_properties['italic'])

            # 设置下划线
            if 'underline' in style_properties:
                self.set_run_underline(para_index, run_index, style_properties['underline'])

            # 设置颜色
            if 'color' in style_properties:
                self.set_run_color(para_index, run_index, style_properties['color'])

            # 设置高亮
            if 'highlight' in style_properties:
                self.set_run_highlight(para_index, run_index, style_properties['highlight'])

            # 设置删除线
            if 'strike' in style_properties:
                self.set_run_strike(para_index, run_index, style_properties['strike'])

            # 更新XML
            self.update_document_xml()

            # 打印更新后的样式信息
            updated_style = self.get_run_style(para_index, run_index)
            print(f"更新后的样式: {updated_style}")

            return True
        except Exception as e:
            print(f"更新Run样式时出错: {e}")
            import traceback
            traceback.print_exc()
            return False

    def set_run_text(self, para_index, run_index, text):
        """设置Run元素的文本内容

        Args:
            para_index: 段落索引
            run_index: Run元素索引
            text: 要设置的文本内容

        Returns:
            bool: 是否成功设置
        """
        try:
            # 获取run元素
            run = self.get_run_element(para_index, run_index)
            if run is None:
                return False

            # 查找现有的t元素
            t_element = run.find(f".//{{{self.NAMESPACES['w']}}}t")
            if t_element is None:
                # 如果不存在则创建
                t_element = ET.SubElement(run, f"{{{self.NAMESPACES['w']}}}t")
                # 设置space属性为preserve以保留空格
                t_element.set(f"{{{self.NAMESPACES['xml']}}}space", "preserve")

            # 设置文本内容
            t_element.text = text
            print(f"设置文本内容: '{text}'")

            return True
        except Exception as e:
            print(f"设置Run文本时出错: {e}")
            import traceback
            traceback.print_exc()
            return False

    def set_run_text_from_xml(self, para, run_index, text):
        """设置Run元素的文本内容

        Args:
            para: 段落元素
            run_index: Run元素索引
            text: 要设置的文本内容

        Returns:
            bool: 是否成功设置
        """
        try:
            # 获取run元素
            run = self.get_run_element_from_xml(para, run_index)
            if run is None:
                return False

            # 查找现有的t元素
            t_element = run.find(f".//{{{self.NAMESPACES['w']}}}t")
            if t_element is None:
                # 如果不存在则创建
                t_element = ET.SubElement(run, f"{{{self.NAMESPACES['w']}}}t")
                # 设置space属性为preserve以保留空格
                t_element.set(f"{{{self.NAMESPACES['xml']}}}space", "preserve")

            # 设置文本内容
            t_element.text = text
            print(f"设置文本内容: '{text}'")

            return True
        except Exception as e:
            print(f"设置Run文本时出错: {e}")
            import traceback
            traceback.print_exc()
            return False
    def insert_paragraph(self, element_index=-1, position='after', text='', **style_properties):
        """在文档中插入新段落

        Args:
            element_index: self.elements中的元素索引，支持负索引（如-1表示最后一个元素）
            position: 插入位置，'before'表示在元素前插入，'after'表示在元素后插入
            text: 要插入的段落文本
            **style_properties: 段落样式属性，可包含以下键：
                'style_id': 样式ID
                'alignment': 对齐方式
                'indentation': 缩进设置字典
                'spacing': 间距设置字典
                'fonts': 字体设置字典，包含'ascii', 'eastAsia'等键
                'size': 字号值
                'bold': 是否加粗
                'color': 字体颜色

        Returns:
            int: 新段落在self.paragraphs中的索引，失败则返回-1
        """
        # 处理负索引
        elements_count = len(self.elements)
        if element_index < 0:
            element_index = elements_count + element_index

        # 检查索引是否有效
        if element_index < 0 or element_index >= elements_count:
            print(f"错误：元素索引{element_index}超出范围(0-{elements_count-1})")
            return -1

        try:
            # 获取目标元素
            target_element = self.elements[element_index]['element']

            # 创建新段落元素
            new_para = ET.Element(f"{{{self.NAMESPACES['w']}}}p")

            # 创建段落ID (w14:paraId)
            try:
                # 检查目标元素是否有段落ID
                para_id_attr = f"{{{self.NAMESPACES['w14']}}}paraId"
                if para_id_attr in target_element.attrib:
                    # 生成新的段落ID (使用时间戳)

                    para_id = hex(int(time.time() * 1000))[2:].upper()
                    new_para.set(para_id_attr, para_id)
            except:
                # 如果无法设置段落ID，继续执行
                pass

            # 创建段落属性元素(如果有样式属性)


            # 创建文本运行元素
            if text:
                r = ET.SubElement(new_para, f"{{{self.NAMESPACES['w']}}}r")





                # 添加文本
                t = ET.SubElement(r, f"{{{self.NAMESPACES['w']}}}t")
                # 如果文本包含空格或特殊字符，设置xml:space="preserve"
                if text.startswith(' ') or text.endswith(' ') or '  ' in text:
                    t.set(f"{{{self.NAMESPACES['xml']}}}space", "preserve")
                t.text = text
            self.update_paragraph_style_from_xml(new_para, **style_properties)
            if text:
               self.update_run_style_from_xml(new_para, 0, **style_properties)
            # 直接在文档树中插入新段落
            # 获取文档体(body)
            body = self.root.find(f".//{{{self.NAMESPACES['w']}}}body")
            if body is None:
                print("错误：无法找到文档体(body)元素")
                return -1

            # 查找目标元素在body中的位置
            body_children = list(body)
            target_index = -1
            for i, child in enumerate(body_children):
                if child == target_element:
                    target_index = i
                    break

            if target_index == -1:
                # 如果找不到目标元素，可能是因为它不是body的直接子元素
                # 尝试使用elements中的信息找到正确的位置
                target_info = self.elements[element_index]
                if 'index' in target_info:
                    # 使用索引信息定位
                    target_index = target_info['index']

            if target_index == -1:
                print("错误：无法在文档树中定位目标元素")
                return -1

            # 根据position参数插入段落
            if position.lower() == 'before':
                body.insert(target_index, new_para)
            else:  # 默认在后面插入
                body.insert(target_index + 1, new_para)

            # 重新解析文档结构，更新self.elements和self.paragraphs
            self.get_structured_body_elements()

            # 查找插入的段落在self.paragraphs中的索引
            for i, para in enumerate(self.paragraphs):
                # 由于ElementTree不保证对象相等比较有效，使用XML字符串比较
                if self._elements_equal(para['element'], new_para):
                    return i

            # 如果找不到插入的段落，说明something happened
            print("警告：段落已插入，但无法在self.paragraphs中找到")
            return -1

        except Exception as e:

            print(f"插入段落时出错: {e}")
            traceback.print_exc()
            return -1

    def _elements_equal(self, elem1, elem2):
        """比较两个XML元素是否相等（内容相同）

        Args:
            elem1: 第一个XML元素
            elem2: 第二个XML元素

        Returns:
            bool: 如果元素相等返回True，否则返回False
        """
        try:
            # 比较标签
            if elem1.tag != elem2.tag:
                return False

            # 比较属性
            if elem1.attrib != elem2.attrib:
                return False

            # 比较文本内容
            if (elem1.text or "").strip() != (elem2.text or "").strip():
                return False

            # 比较尾部文本
            if (elem1.tail or "").strip() != (elem2.tail or "").strip():
                return False

            # 比较子元素数量
            if len(elem1) != len(elem2):
                return False

            # 递归比较子元素
            for child1, child2 in zip(elem1, elem2):
                if not self._elements_equal(child1, child2):
                    return False

            return True
        except:
            return False

    def insert_image(self, ele_index, run_index=-1, position='after', image_path='',
                     width=None, height=None, description=None, wrap_text='inline',
                     new_page=False, line_spacing=240, line_rule='auto'):
        """在文档中指定位置插入图片

        Args:
            ele_index: self.elements的段落索引
            run_index: 段落中文本运行的索引，-1表示段落末尾
            position: 插入位置，'before'表示在运行前插入，'after'表示在运行后插入
            image_path: 图片文件的路径
            width: 图片宽度(厘米)，不指定则使用原始大小
            height: 图片高度(厘米)，不指定则使用原始大小
            description: 图片描述
            wrap_text: 图片环绕文字的方式，可选值:
                      'inline'(嵌入式)、'square'(四周型)、'tight'(紧密型)、
                      'through'(穿越型)、'topAndBottom'(上下型)、'behind'(衬于文字下方)、
                      'inFront'(浮于文字上方)
            new_page: 是否在图片前插入分页符，确保图片在新页开始
      line_spacing: 行距值，如240
        line_rule: 行距规则，如'auto', 'exact', 'atLeast'
        Returns:
            str: 新创建的图片关系ID，失败则返回None
        """
        # 检查图片文件是否存在
        if not os.path.exists(image_path):
            print(f"错误：图片文件 {image_path} 不存在")
            return None

        # 获取图片信息
        try:
            img = Image.open(image_path)
            img_format = img.format.lower()
            img_width, img_height = img.size

            # 如果没有指定宽高，使用原始尺寸（转换为EMU单位，1厘米=360000 EMU）
            if width is None:
                # 默认分辨率为96 DPI，即96像素/英寸
                # 1英寸 = 2.54厘米，所以1厘米 = 96/2.54 像素
                # 因此，像素到厘米的转换：厘米 = 像素 * 2.54 / 96
                width_cm = img_width * 2.54 / 96
                width_emu = int(width_cm * 360000)
            else:
                width_emu = int(width * 360000)

            if height is None:
                height_cm = img_height * 2.54 / 96
                height_emu = int(height_cm * 360000)
            else:
                height_emu = int(height * 360000)

        except Exception as e:
            print(f"获取图片信息时出错: {e}")
            return None

        # 处理索引为paragraphs索引还是elements索引
        try:
            absolute_index = abs(ele_index)
            if absolute_index <= len(self.elements):
                paragraph= self.elements[absolute_index]['element']
            else:
                    print(f"错误：索引{ele_index}不是有效的索引")
                    return None
        except Exception as e:
            print(f"获取段落元素时出错: {e}")
            return None

        # 如果需要在新页面显示图片，先在该位置插入分页符
        if new_page:
            try:
                # 获取段落的pPr元素，如果不存在则创建
                pPr = paragraph.find(f".//{{{self.NAMESPACES['w']}}}pPr")
                if pPr is None:
                    pPr = ET.SubElement(paragraph, f"{{{self.NAMESPACES['w']}}}pPr")

                # 添加分页符设置
                page_break = ET.SubElement(pPr, f"{{{self.NAMESPACES['w']}}}pageBreakBefore")
                page_break.set(f"{{{self.NAMESPACES['w']}}}val", "1")
            except Exception as e:
                print(f"插入分页符时出错: {e}")
                # 继续执行，不因分页符失败而中断整个操作
        if line_spacing is not None:
            try:
                # 获取段落的pPr元素，如果不存在则创建
                pPr = paragraph.find(f".//{{{self.NAMESPACES['w']}}}pPr")
                if pPr is None:
                    pPr = ET.SubElement(paragraph, f"{{{self.NAMESPACES['w']}}}pPr")

                # 检查是否已存在spacing元素
                spacing = pPr.find(f".//{{{self.NAMESPACES['w']}}}spacing")
                if spacing is None:
                    spacing = ET.SubElement(pPr, f"{{{self.NAMESPACES['w']}}}spacing")

                # 设置行距值
                spacing.set(f"{{{self.NAMESPACES['w']}}}line", str(line_spacing))

                # 如果提供了行距规则，也设置它
                if line_rule:
                    spacing.set(f"{{{self.NAMESPACES['w']}}}lineRule", line_rule)
            except Exception as e:
                print(f"设置行距时出错: {e}")
        # 获取段落中的文本运行
        try:
            r_elements =paragraph.findall("./w:r", self.NAMESPACES)
            if not r_elements:
                # 如果段落中没有文本运行，创建一个空的文本运行
                run_index = 0
                r = ET.SubElement(paragraph, f"{{{self.NAMESPACES['w']}}}r")
                r_elements = [r]
            elif run_index < 0:
                # 负索引表示从末尾计数
                run_index = len(r_elements) + run_index
                if run_index < 0:
                    run_index = 0

            # 检查运行索引是否有效
            if run_index >= len(r_elements):
                run_index = len(r_elements) - 1

            # 获取目标文本运行
            target_run = r_elements[run_index]
        except Exception as e:
            print(f"获取文本运行时出错: {e}")
            return None

        # 读取图片文件
        try:
            with open(image_path, 'rb') as img_file:
                img_data = img_file.read()

            # 生成图片ID和文件名
            img_id = str(uuid.uuid4())
            img_name = os.path.basename(image_path)
            img_ext = os.path.splitext(img_name)[1].lower()

            # 生成关系ID
            rel_id = f"rId{int(time.time() * 1000)}"

            # 创建图片关系
            # 检查是否已经存在media文件夹
            if 'media' not in self.parts:
                self.parts['media'] = {}

            # 将图片添加到media文件夹
            self.parts['media'][img_name] = img_data

            # 添加关系到document.xml.rels
            if 'relationships' not in self.parts:
                print("错误：找不到document.xml.rels文件")
                return None

            # 获取关系文件
            rels_tree = self.parts['relationships']
            rels_root = rels_tree.getroot()

            # 创建新的关系元素
            new_rel = ET.Element("Relationship")
            new_rel.set("Id", rel_id)
            new_rel.set("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image")
            new_rel.set("Target", f"media/{img_name}")

            # 添加到关系文件
            rels_root.append(new_rel)

            # 更新关系文件
            self.parts['relationships'] = rels_tree

            # 创建图片XML结构
            new_run = ET.Element(f"{{{self.NAMESPACES['w']}}}r")
            drawing = ET.SubElement(new_run, f"{{{self.NAMESPACES['w']}}}drawing")

            # 根据环绕方式选择不同的XML结构
            if wrap_text == 'inline' or wrap_text == 'none':
                # 嵌入式图片
                graphic_container = ET.SubElement(drawing, f"{{{self.NAMESPACES['wp']}}}inline")
            else:
                # 浮动式图片(四周环绕、紧密型、上下型等)
                graphic_container = ET.SubElement(drawing, f"{{{self.NAMESPACES['wp']}}}anchor")
                graphic_container.set("simplePos", "0")
                graphic_container.set("relativeHeight", "251658240")
                graphic_container.set("behindDoc", "1" if wrap_text == 'behind' else "0")
                graphic_container.set("locked", "0")
                graphic_container.set("layoutInCell", "1")
                graphic_container.set("allowOverlap", "1")

                # 添加环绕文字设置
                wrap_element = None
                if wrap_text == 'square':
                    wrap_element = ET.SubElement(graphic_container, f"{{{self.NAMESPACES['wp']}}}wrapSquare")
                    wrap_element.set("wrapText", "bothSides")
                elif wrap_text == 'tight':
                    wrap_element = ET.SubElement(graphic_container, f"{{{self.NAMESPACES['wp']}}}wrapTight")
                    wrap_element.set("wrapText", "bothSides")
                elif wrap_text == 'through':
                    wrap_element = ET.SubElement(graphic_container, f"{{{self.NAMESPACES['wp']}}}wrapThrough")
                    wrap_element.set("wrapText", "bothSides")
                elif wrap_text == 'topAndBottom':
                    wrap_element = ET.SubElement(graphic_container, f"{{{self.NAMESPACES['wp']}}}wrapTopAndBottom")
                elif wrap_text == 'behind':
                    # 衬于文字下方不需要特殊的wrap元素
                    pass
                elif wrap_text == 'inFront':
                    # 浮于文字上方不需要特殊的wrap元素
                    graphic_container.set("behindDoc", "0")

                # 添加定位信息
                simple_pos = ET.SubElement(graphic_container, f"{{{self.NAMESPACES['wp']}}}simplePos")
                simple_pos.set("x", "0")
                simple_pos.set("y", "0")

                # 添加水平定位
                pos_h = ET.SubElement(graphic_container, f"{{{self.NAMESPACES['wp']}}}positionH")
                pos_h.set("relativeFrom", "column")
                pos_h_align = ET.SubElement(pos_h, f"{{{self.NAMESPACES['wp']}}}align")
                pos_h_align.text = "center"

                # 添加垂直定位
                pos_v = ET.SubElement(graphic_container, f"{{{self.NAMESPACES['wp']}}}positionV")
                pos_v.set("relativeFrom", "paragraph")
                pos_v_align = ET.SubElement(pos_v, f"{{{self.NAMESPACES['wp']}}}align")
                pos_v_align.text = "center"

            # 设置图片大小
            extent = ET.SubElement(graphic_container, f"{{{self.NAMESPACES['wp']}}}extent")
            extent.set("cx", str(width_emu))
            extent.set("cy", str(height_emu))

            # 设置效果范围
            effect_extent = ET.SubElement(graphic_container, f"{{{self.NAMESPACES['wp']}}}effectExtent")
            effect_extent.set("l", "0")
            effect_extent.set("t", "0")
            effect_extent.set("r", "0")
            effect_extent.set("b", "0")

            # 设置DOC PROPS
            doc_pr = ET.SubElement(graphic_container, f"{{{self.NAMESPACES['wp']}}}docPr")
            doc_pr.set("id", img_id)
            doc_pr.set("name", img_name)
            if description:
                doc_pr.set("descr", description)
            # 添加图形框架属性 - 这是缺少的部分
            cnv_graphic_frame_pr = ET.SubElement(graphic_container, f"{{{self.NAMESPACES['wp']}}}cNvGraphicFramePr")
            graphic_frame_locks = ET.SubElement(cnv_graphic_frame_pr, f"{{{self.NAMESPACES['a']}}}graphicFrameLocks")
            graphic_frame_locks.set("noChangeAspect", "1")
            # 添加图片数据
            graphic = ET.SubElement(graphic_container, f"{{{self.NAMESPACES['a']}}}graphic")
            graphic_data = ET.SubElement(graphic, f"{{{self.NAMESPACES['a']}}}graphicData")
            graphic_data.set("uri", "http://schemas.openxmlformats.org/drawingml/2006/picture")

            pic = ET.SubElement(graphic_data, f"{{{self.NAMESPACES['pic']}}}pic")

            # 图片非视觉属性
            nvpic_pr = ET.SubElement(pic, f"{{{self.NAMESPACES['pic']}}}nvPicPr")

            # 图片非视觉绘图属性
            cnvpr = ET.SubElement(nvpic_pr, f"{{{self.NAMESPACES['pic']}}}cNvPr")
            cnvpr.set("id", "0")
            cnvpr.set("name", img_name)
            if description:
                cnvpr.set("descr", description)

            # 图片非视觉图片属性
            cnvpic_pr = ET.SubElement(nvpic_pr, f"{{{self.NAMESPACES['pic']}}}cNvPicPr")
            pic_locks = ET.SubElement(cnvpic_pr, f"{{{self.NAMESPACES['a']}}}picLocks")
            pic_locks.set("noChangeAspect", "1")
            # 图片填充
            blip_fill = ET.SubElement(pic, f"{{{self.NAMESPACES['pic']}}}blipFill")
            blip = ET.SubElement(blip_fill, f"{{{self.NAMESPACES['a']}}}blip")
            blip.set(f"{{{self.NAMESPACES['r']}}}embed", rel_id)

            # 源矩形
            src_rect = ET.SubElement(blip_fill, f"{{{self.NAMESPACES['a']}}}srcRect")

            # 拉伸
            stretch = ET.SubElement(blip_fill, f"{{{self.NAMESPACES['a']}}}stretch")
            fill_rect = ET.SubElement(stretch, f"{{{self.NAMESPACES['a']}}}fillRect")

            # 图片形状属性
            sppr = ET.SubElement(pic, f"{{{self.NAMESPACES['pic']}}}spPr")

            # 预设几何形状
            xfrm = ET.SubElement(sppr, f"{{{self.NAMESPACES['a']}}}xfrm")
            off = ET.SubElement(xfrm, f"{{{self.NAMESPACES['a']}}}off")
            off.set("x", "0")
            off.set("y", "0")
            ext = ET.SubElement(xfrm, f"{{{self.NAMESPACES['a']}}}ext")
            ext.set("cx", str(width_emu))
            ext.set("cy", str(height_emu))

            # 预设几何形状
            prst_geom = ET.SubElement(sppr, f"{{{self.NAMESPACES['a']}}}prstGeom")
            prst_geom.set("prst", "rect")
            av_lst = ET.SubElement(prst_geom, f"{{{self.NAMESPACES['a']}}}avLst")

            # 添加无填充 - 这是缺少的部分
            no_fill = ET.SubElement(sppr, f"{{{self.NAMESPACES['a']}}}noFill")

            # 添加线条属性 - 这是缺少的部分
            ln = ET.SubElement(sppr, f"{{{self.NAMESPACES['a']}}}ln")
            ln_no_fill = ET.SubElement(ln, f"{{{self.NAMESPACES['a']}}}noFill")
            # 根据position参数插入图片
            if position.lower() == 'before':
                paragraph.insert(list(paragraph).index(target_run), new_run)
            else:  # 默认在后面插入
                paragraph.insert(list(paragraph).index(target_run) + 1, new_run)

            # 更新文档XML
            self.update_document_xml()

            # 成功添加图片
            return rel_id

        except Exception as e:
            print(f"插入图片时出错: {e}")
            traceback.print_exc()
            return None

    def replace_image(self, rel_id, image_path, width=None, height=None, description=None):
        """替换文档中已有的图片

        Args:
            rel_id: 要替换的图片关系ID
            image_path: 新图片文件的路径
            width: 新图片宽度(厘米)，不指定则使用原始大小
            height: 新图片高度(厘米)，不指定则使用原始大小
            description: 新的图片描述，不指定则保留原描述

        Returns:
            bool: 是否成功替换
        """
        # 检查图片文件是否存在
        if not os.path.exists(image_path):
            print(f"错误：图片文件 {image_path} 不存在")
            return False

        # 获取新图片信息
        try:
            img = Image.open(image_path)
            img_format = img.format.lower()
            img_width, img_height = img.size

            # 如果没有指定宽高，使用原始尺寸（转换为EMU单位，1厘米=360000 EMU）
            if width is None:
                # 默认分辨率为96 DPI，即96像素/英寸
                # 1英寸 = 2.54厘米，所以1厘米 = 96/2.54 像素
                # 因此，像素到厘米的转换：厘米 = 像素 * 2.54 / 96
                width_cm = img_width * 2.54 / 96
                width_emu = int(width_cm * 360000)
            else:
                width_emu = int(width * 360000)

            if height is None:
                height_cm = img_height * 2.54 / 96
                height_emu = int(height_cm * 360000)
            else:
                height_emu = int(height * 360000)
        except Exception as e:
            print(f"获取新图片信息时出错: {e}")
            return False

        try:
            # 验证关系ID格式
            if not rel_id.startswith("rId"):
                print(f"错误：无效的关系ID格式: {rel_id}")
                return False

            # 获取关系文件
            if 'relationships' not in self.parts:
                print("错误：找不到document.xml.rels文件")
                return False

            rels_tree = self.parts['relationships']
            rels_root = rels_tree.getroot()

            # 查找要替换的图片关系
            rel_element = None
            for rel in rels_root.findall("./Relationship",
                                         {'': "http://schemas.openxmlformats.org/package/2006/relationships"}):
                if rel.get("Id") == rel_id:
                    rel_element = rel
                    break

            if rel_element is None:
                print(f"错误：找不到ID为 {rel_id} 的图片关系")
                return False

            # 获取原图片的Target路径
            old_target = rel_element.get("Target")
            if not old_target or not old_target.startswith("media/"):
                print(f"错误：图片关系的Target不是有效的media路径: {old_target}")
                return False

            old_image_name = old_target.replace("media/", "")

            # 生成新图片文件名（保留扩展名）
            img_name = os.path.basename(image_path)

            # 读取新图片文件
            with open(image_path, 'rb') as img_file:
                img_data = img_file.read()

            # 更新media文件夹中的图片
            if 'media' not in self.parts:
                self.parts['media'] = {}

            # 替换媒体文件
            self.parts['media'][old_image_name] = img_data

            # 查找文档中引用该图片的所有位置
            updated_count = 0
            for element in self.document_root.findall(
                    f".//{{{self.NAMESPACES['a']}}}blip[@{{{self.NAMESPACES['r']}}}embed='{rel_id}']", self.NAMESPACES):
                # 查找该图片的尺寸元素
                try:
                    # 遍历向上找到pic元素
                    pic_elem = element
                    while pic_elem is not None and pic_elem.tag != f"{{{self.NAMESPACES['pic']}}}pic":
                        pic_elem = pic_elem.getparent()

                    if pic_elem is not None:
                        # 更新描述信息
                        if description is not None:
                            # 更新cNvPr的descr属性
                            cnvpr = pic_elem.find(f".//{{{self.NAMESPACES['pic']}}}cNvPr", self.NAMESPACES)
                            if cnvpr is not None:
                                cnvpr.set("descr", description)

                        # 更新尺寸
                        # 查找ext元素
                        ext_elements = pic_elem.findall(f".//{{{self.NAMESPACES['a']}}}ext", self.NAMESPACES)
                        for ext in ext_elements:
                            ext.set("cx", str(width_emu))
                            ext.set("cy", str(height_emu))

                        # 查找extent元素
                        extent_elements = pic_elem.findall(f".//{{{self.NAMESPACES['wp']}}}extent", self.NAMESPACES)
                        for extent in extent_elements:
                            extent.set("cx", str(width_emu))
                            extent.set("cy", str(height_emu))

                        updated_count += 1
                except Exception as e:
                    print(f"更新图片尺寸时出错: {e}")
                    # 继续处理其他图片引用

            # 更新文档XML
            self.update_document_xml()

            print(f"成功替换图片 {rel_id}，更新了 {updated_count} 个图片引用")
            return True
        except Exception as e:
            print(f"替换图片时出错: {e}")
            traceback.print_exc()
            return False

    def insert_image_with_caption(self, para_index, image_path, caption_text, chapter_num="1",
                                  width=None, height=None, description=None, wrap_text='inline',
                                  new_page=False, caption_style=None):
        """插入图片并添加标题

        Args:
            para_index: 段落索引
            image_path: 图片文件路径
            caption_text: 图片标题文本
            chapter_num: 章节编号，默认为1"
            width: 图片宽度(厘米)
            height: 图片高度(厘米)
            description: 图片描述
            wrap_text: 图片环绕文字方式
            new_page: 是否在新页开始
            caption_style: 标题样式属性字典，如{'fonts': {'eastAsia': '黑体'}, 'size': 24, 'bold': True}

        Returns:
            tuple: (图片关系ID, 标题段落索引)，如果失败，相应的值为None
        """
        try:
            # 插入图片
            rel_id = self.insert_image(
                para_index=para_index,
                image_path=image_path,
                width=width,
                height=height,
                description=description,
                wrap_text=wrap_text,
                new_page=new_page
            )

            if rel_id is None:
                return None, None

            # 确定要添加标题的段落索引
            # 如果para_index是负数，需要转换为实际索引
            if para_index < 0:
                para_index = len(self.get_all_paragraphs()) + para_index

            # 检查caption_style参数
            if caption_style is None:
                caption_style = {}

            # 插入图片标题
            caption_para_idx = self.insert_figure_caption(
                para_index=para_index,
                chapter_num=chapter_num,
                caption_text=caption_text,
                **caption_style
            )

            return rel_id, caption_para_idx

        except Exception as e:
            print(f"插入图片及标题时出错: {e}")
            traceback.print_exc()
            return None, None

    def insert_run(self, para_index, run_index=-1, position='after', text='', **style_properties):
        """在段落中插入新的文本运行(run)

        Args:
            para_index: self.elements或self.paragraphs中的段落索引
            run_index: 段落中文本运行的索引，-1表示段落末尾
            position: 插入位置，'before'表示在运行前插入，'after'表示在运行后插入
            text: 要插入的文本内容
            **style_properties: 文本运行的样式属性，可包含以下键：
                'fonts': 字体设置字典，包含'ascii', 'eastAsia'等键
                'size': 字号值(半磅值)
                'bold': 是否加粗
                'italic': 是否斜体
                'underline': 下划线类型
                'color': 字体颜色
                'highlight': 高亮颜色
                'strike': 是否添加删除线
                'caps': 是否全部大写
                'vert_align': 垂直对齐方式(上标/下标)

        Returns:
            bool: 是否成功插入
        """
        # 处理索引为paragraphs索引还是elements索引
        try:
            if para_index >= 0 and para_index < len(self.paragraphs):
                # 是段落索引
                paragraph = self.paragraphs[para_index]['element']
            elif para_index >= 0 and para_index < len(self.elements) and self.elements[para_index]['type'] == 'paragraph':
                # 是elements索引，且为段落类型
                paragraph = self.elements[para_index]['element']
            else:
                # 处理负索引
                if para_index < 0:
                    elements_count = len(self.elements)
                    para_index = elements_count + para_index
                    if para_index >= 0 and para_index < elements_count and self.elements[para_index]['type'] == 'paragraph':
                        paragraph = self.elements[para_index]['element']
                    else:
                        print(f"错误：索引{para_index}不是有效的段落索引")
                        return False
                else:
                    print(f"错误：索引{para_index}不是有效的段落索引")
                    return False
        except Exception as e:
            print(f"获取段落元素时出错: {e}")
            return False

        # 获取段落中的文本运行
        try:
            r_elements = paragraph.findall("./w:r", self.NAMESPACES)
            if not r_elements:
                # 如果段落中没有文本运行，创建一个空的文本运行
                run_index = 0
                position = 'before'  # 没有现有运行，只能在前面插入
                r_elements = []
            elif run_index < 0:
                # 负索引表示从末尾计数
                run_index = len(r_elements) + run_index
                if run_index < 0:
                    run_index = 0

            # 检查运行索引是否有效
            if r_elements and run_index >= len(r_elements):
                run_index = len(r_elements) - 1
                position = 'after'  # 超出索引范围，只能在最后一个后面插入

            # 创建新的文本运行元素
            new_run = ET.Element(f"{{{self.NAMESPACES['w']}}}r")

            # 添加文本内容
            if text:
                t = ET.SubElement(new_run, f"{{{self.NAMESPACES['w']}}}t")
                if text.startswith(' ') or text.endswith(' ') or '  ' in text:
                    t.set(f"{{{self.NAMESPACES['xml']}}}space", "preserve")
                t.text = text

            # 插入新 run
            if r_elements:
                target_run = r_elements[run_index]
                if position.lower() == 'before':
                    paragraph.insert(list(paragraph).index(target_run), new_run)
                else:
                    paragraph.insert(list(paragraph).index(target_run) + 1, new_run)
            else:
                paragraph.append(new_run)

            # 统一用专用方法设置 run 样式
            all_runs = paragraph.findall("./w:r", self.NAMESPACES)

            for idx in range(len(all_runs)):
                self.update_run_style_from_xml(paragraph, idx, **style_properties)
            # 更新文档XML
            self.update_document_xml()

            return True
        except Exception as e:
            print(f"插入文本运行时出错: {e}")
            return False

    def set_table_style(self, table_index, **style_properties):
        """设置表格的样式和属性

        Args:
            table_index: self.tables中的表格索引
            **style_properties: 可以包含以下属性:
                - style_id: 表格样式ID
                - width: 表格宽度(dict): {'value': '值', 'type': '类型'}
                - indent: 表格缩进(dict): {'value': '值', 'type': '类型'}
                - borders: 表格边框(dict): {
                    'top': {'val': '类型', 'color': '颜色', 'sz': '粗细', 'space': '间距'},
                    'left': {...},
                    'bottom': {...},
                    'right': {...},
                    'inside_h': {...},
                    'inside_v': {...}
                  }
                - layout: 表格布局类型('autofit' or 'fixed')
                - cell_margins: 单元格边距(dict): {
                    'top': {'value': '值', 'type': '类型'},
                    'left': {...},
                    'bottom': {...},
                    'right': {...}
                  }

        Returns:
            bool: 操作是否成功
        """
        # 检查索引是否有效
        absolute_index = abs(table_index)
        if absolute_index >= len(self.tables):
            print(f"错误：表格索引{table_index}超出范围(0-{len(self.tables)-1})")
            return False

        # 获取表格元素
        table = self.tables[table_index]['element']

        # 获取或创建tblPr元素
        tblPr = table.find(f".//{{{self.NAMESPACES['w']}}}tblPr")
        if tblPr is None:
            tblPr = ET.Element(f"{{{self.NAMESPACES['w']}}}tblPr")
            table.insert(0, tblPr)

        # 设置样式ID
        if 'style_id' in style_properties:
            style_element = tblPr.find(f".//{{{self.NAMESPACES['w']}}}tblStyle")
            if style_element is None:
                style_element = ET.SubElement(tblPr, f"{{{self.NAMESPACES['w']}}}tblStyle")
            style_element.set(f"{{{self.NAMESPACES['w']}}}val", style_properties['style_id'])

        # 设置表格宽度
        if 'width' in style_properties:
            width_info = style_properties['width']
            tblW = tblPr.find(f".//{{{self.NAMESPACES['w']}}}tblW")
            if tblW is None:
                tblW = ET.SubElement(tblPr, f"{{{self.NAMESPACES['w']}}}tblW")

            if 'value' in width_info:
                tblW.set(f"{{{self.NAMESPACES['w']}}}w", str(width_info['value']))
            if 'type' in width_info:
                tblW.set(f"{{{self.NAMESPACES['w']}}}type", width_info['type'])

        # 设置表格缩进
        if 'indent' in style_properties:
            indent_info = style_properties['indent']
            tblInd = tblPr.find(f".//{{{self.NAMESPACES['w']}}}tblInd")
            if tblInd is None:
                tblInd = ET.SubElement(tblPr, f"{{{self.NAMESPACES['w']}}}tblInd")

            if 'value' in indent_info:
                tblInd.set(f"{{{self.NAMESPACES['w']}}}w", str(indent_info['value']))
            if 'type' in indent_info:
                tblInd.set(f"{{{self.NAMESPACES['w']}}}type", indent_info['type'])

        # 设置表格边框
        if 'borders' in style_properties:
            borders_info = style_properties['borders']
            tblBorders = tblPr.find(f".//{{{self.NAMESPACES['w']}}}tblBorders")
            if tblBorders is None:
                tblBorders = ET.SubElement(tblPr, f"{{{self.NAMESPACES['w']}}}tblBorders")

            border_mapping = {
                'top': 'top',
                'left': 'left',
                'bottom': 'bottom',
                'right': 'right',
                'inside_h': 'insideH',
                'inside_v': 'insideV'
            }

            for border_key, border_xml_name in border_mapping.items():
                if border_key in borders_info:
                    border_info = borders_info[border_key]
                    border_element = tblBorders.find(f".//{{{self.NAMESPACES['w']}}}{border_xml_name}")
                    if border_element is None:
                        border_element = ET.SubElement(tblBorders, f"{{{self.NAMESPACES['w']}}}{border_xml_name}")

                    for attr_name, xml_attr in [
                        ('val', 'val'),
                        ('color', 'color'),
                        ('sz', 'sz'),
                        ('space', 'space')
                    ]:
                        if attr_name in border_info:
                            border_element.set(f"{{{self.NAMESPACES['w']}}}{xml_attr}", str(border_info[attr_name]))

        # 设置表格布局
        if 'layout' in style_properties:
            tblLayout = tblPr.find(f".//{{{self.NAMESPACES['w']}}}tblLayout")
            if tblLayout is None:
                tblLayout = ET.SubElement(tblPr, f"{{{self.NAMESPACES['w']}}}tblLayout")
            tblLayout.set(f"{{{self.NAMESPACES['w']}}}type", style_properties['layout'])

        # 设置单元格边距
        if 'cell_margins' in style_properties:
            margin_info = style_properties['cell_margins']
            tblCellMar = tblPr.find(f".//{{{self.NAMESPACES['w']}}}tblCellMar")
            if tblCellMar is None:
                tblCellMar = ET.SubElement(tblPr, f"{{{self.NAMESPACES['w']}}}tblCellMar")

            for margin_type in ['top', 'left', 'bottom', 'right']:
                if margin_type in margin_info:
                    margin_element = tblCellMar.find(f".//{{{self.NAMESPACES['w']}}}{margin_type}")
                    if margin_element is None:
                        margin_element = ET.SubElement(tblCellMar, f"{{{self.NAMESPACES['w']}}}{margin_type}")

                    margin_data = margin_info[margin_type]
                    if 'value' in margin_data:
                        margin_element.set(f"{{{self.NAMESPACES['w']}}}w", str(margin_data['value']))
                    if 'type' in margin_data:
                        margin_element.set(f"{{{self.NAMESPACES['w']}}}type", margin_data['type'])



        return True
    def set_table_style_from_xml(self, table, **style_properties):
        """设置表格的样式和属性

        Args:
            table: table元素
            **style_properties: 可以包含以下属性:
                - style_id: 表格样式ID
                - width: 表格宽度(dict): {'value': '值', 'type': '类型'}
                - indent: 表格缩进(dict): {'value': '值', 'type': '类型'}
                - borders: 表格边框(dict): {
                    'top': {'val': '类型', 'color': '颜色', 'sz': '粗细', 'space': '间距'},
                    'left': {...},
                    'bottom': {...},
                    'right': {...},
                    'inside_h': {...},
                    'inside_v': {...}
                  }
                - layout: 表格布局类型('autofit' or 'fixed')
                - cell_margins: 单元格边距(dict): {
                    'top': {'value': '值', 'type': '类型'},
                    'left': {...},
                    'bottom': {...},
                    'right': {...}
                  }

        Returns:
            bool: 操作是否成功
        """


        # 获取表格元素


        # 获取或创建tblPr元素
        tblPr = table.find(f".//{{{self.NAMESPACES['w']}}}tblPr")
        if tblPr is None:
            tblPr = ET.Element(f"{{{self.NAMESPACES['w']}}}tblPr")
            table.insert(0, tblPr)

        # 设置样式ID
        if 'style_id' in style_properties:
            style_element = tblPr.find(f".//{{{self.NAMESPACES['w']}}}tblStyle")
            if style_element is None:
                style_element = ET.SubElement(tblPr, f"{{{self.NAMESPACES['w']}}}tblStyle")
            style_element.set(f"{{{self.NAMESPACES['w']}}}val", style_properties['style_id'])

        # 设置表格宽度
        if 'width' in style_properties:
            width_info = style_properties['width']
            tblW = tblPr.find(f".//{{{self.NAMESPACES['w']}}}tblW")
            if tblW is None:
                tblW = ET.SubElement(tblPr, f"{{{self.NAMESPACES['w']}}}tblW")

            if 'value' in width_info:
                tblW.set(f"{{{self.NAMESPACES['w']}}}w", str(width_info['value']))
            if 'type' in width_info:
                tblW.set(f"{{{self.NAMESPACES['w']}}}type", width_info['type'])

        # 设置表格缩进
        if 'indent' in style_properties:
            indent_info = style_properties['indent']
            tblInd = tblPr.find(f".//{{{self.NAMESPACES['w']}}}tblInd")
            if tblInd is None:
                tblInd = ET.SubElement(tblPr, f"{{{self.NAMESPACES['w']}}}tblInd")

            if 'value' in indent_info:
                tblInd.set(f"{{{self.NAMESPACES['w']}}}w", str(indent_info['value']))
            if 'type' in indent_info:
                tblInd.set(f"{{{self.NAMESPACES['w']}}}type", indent_info['type'])

        # 设置表格边框
        if 'borders' in style_properties:
            borders_info = style_properties['borders']
            tblBorders = tblPr.find(f".//{{{self.NAMESPACES['w']}}}tblBorders")
            if tblBorders is None:
                tblBorders = ET.SubElement(tblPr, f"{{{self.NAMESPACES['w']}}}tblBorders")

            border_mapping = {
                'top': 'top',
                'left': 'left',
                'bottom': 'bottom',
                'right': 'right',
                'inside_h': 'insideH',
                'inside_v': 'insideV'
            }

            for border_key, border_xml_name in border_mapping.items():
                if border_key in borders_info:
                    border_info = borders_info[border_key]
                    border_element = tblBorders.find(f".//{{{self.NAMESPACES['w']}}}{border_xml_name}")
                    if border_element is None:
                        border_element = ET.SubElement(tblBorders, f"{{{self.NAMESPACES['w']}}}{border_xml_name}")

                    for attr_name, xml_attr in [
                        ('val', 'val'),
                        ('color', 'color'),
                        ('sz', 'sz'),
                        ('space', 'space')
                    ]:
                        if attr_name in border_info:
                            border_element.set(f"{{{self.NAMESPACES['w']}}}{xml_attr}", str(border_info[attr_name]))

        # 设置表格布局
        if 'layout' in style_properties:
            tblLayout = tblPr.find(f".//{{{self.NAMESPACES['w']}}}tblLayout")
            if tblLayout is None:
                tblLayout = ET.SubElement(tblPr, f"{{{self.NAMESPACES['w']}}}tblLayout")
            tblLayout.set(f"{{{self.NAMESPACES['w']}}}type", style_properties['layout'])

        # 设置单元格边距
        if 'cell_margins' in style_properties:
            margin_info = style_properties['cell_margins']
            tblCellMar = tblPr.find(f".//{{{self.NAMESPACES['w']}}}tblCellMar")
            if tblCellMar is None:
                tblCellMar = ET.SubElement(tblPr, f"{{{self.NAMESPACES['w']}}}tblCellMar")

            for margin_type in ['top', 'left', 'bottom', 'right']:
                if margin_type in margin_info:
                    margin_element = tblCellMar.find(f".//{{{self.NAMESPACES['w']}}}{margin_type}")
                    if margin_element is None:
                        margin_element = ET.SubElement(tblCellMar, f"{{{self.NAMESPACES['w']}}}{margin_type}")

                    margin_data = margin_info[margin_type]
                    if 'value' in margin_data:
                        margin_element.set(f"{{{self.NAMESPACES['w']}}}w", str(margin_data['value']))
                    if 'type' in margin_data:
                        margin_element.set(f"{{{self.NAMESPACES['w']}}}type", margin_data['type'])



        return True

    def set_table_row_style_from_xml(self, table, row_index, **style_properties):
        """设置表格行的样式和属性

        Args:
            table: 表格元素
            row_index: 行索引
            **style_properties: 可以包含以下属性:
                - height: 行高设置(dict): {'value': '值', 'rule': '规则'}
                    rule可选值: 'auto', 'atLeast', 'exact'
                - cannot_split: 布尔值，是否禁止跨页分割
                - is_header: 布尔值，是否为表头行
                - borders: 行边框(dict): 结构同表格边框

        Returns:
            bool: 操作是否成功
        """


        # 获取所有行
        tr_elements = table.findall(f".//{{{self.NAMESPACES['w']}}}tr")

        # 检查行索引是否有效
        if row_index < 0 or row_index >= len(tr_elements):
            print(f"错误：行索引{row_index}超出范围(0-{len(tr_elements) - 1})")
            return False

        # 获取目标行
        tr = tr_elements[row_index]

        # 获取或创建trPr元素（行属性）
        trPr = tr.find(f".//{{{self.NAMESPACES['w']}}}trPr")
        if trPr is None:
            trPr = ET.Element(f"{{{self.NAMESPACES['w']}}}trPr")
            tr.insert(0, trPr)

        # 设置行高
        if 'height' in style_properties:
            height_info = style_properties['height']
            trHeight = trPr.find(f".//{{{self.NAMESPACES['w']}}}trHeight")
            if trHeight is None:
                trHeight = ET.SubElement(trPr, f"{{{self.NAMESPACES['w']}}}trHeight")

            if 'value' in height_info:
                trHeight.set(f"{{{self.NAMESPACES['w']}}}val", str(height_info['value']))
            if 'rule' in height_info:
                # 规则可以是: 'auto', 'atLeast', 'exact'
                trHeight.set(f"{{{self.NAMESPACES['w']}}}hRule", height_info['rule'])

        # 设置是否允许跨页分割
        if 'cannot_split' in style_properties:
            cantSplit = trPr.find(f".//{{{self.NAMESPACES['w']}}}cantSplit")
            if style_properties['cannot_split']:
                if cantSplit is None:
                    cantSplit = ET.SubElement(trPr, f"{{{self.NAMESPACES['w']}}}cantSplit")
                # Word中不需要属性值，仅标记存在即可
            else:
                # 移除不能分割标记
                if cantSplit is not None:
                    trPr.remove(cantSplit)

        # 设置是否为表头行（重复行）
        if 'is_header' in style_properties:
            tblHeader = trPr.find(f".//{{{self.NAMESPACES['w']}}}tblHeader")
            if style_properties['is_header']:
                if tblHeader is None:
                    tblHeader = ET.SubElement(trPr, f"{{{self.NAMESPACES['w']}}}tblHeader")
                # Word中不需要属性值，仅标记存在即可
            else:
                # 移除表头行标记
                if tblHeader is not None:
                    trPr.remove(tblHeader)

        # 设置行边框（如果需要）
        if 'borders' in style_properties:
            # 行级边框通常应用到每个单元格
            # 这里需要遍历行中的每个单元格，为每个单元格设置边框
            td_elements = tr.findall(f".//{{{self.NAMESPACES['w']}}}tc")
            for cell_index, td in enumerate(td_elements):
                self.set_table_cell_borders(table, row_index, cell_index, **style_properties['borders'])

        # 更新XML
        self.update_document_xml()
        return True

    def set_table_cell_style_from_xml(self, table, row_index, cell_index, **style_properties):
        """设置表格单元格的样式和属性

        Args:
            table: 的表格
            row_index: 行索引
            cell_index: 单元格索引
            **style_properties: 可以包含以下属性:
                - width: 单元格宽度(dict): {'value': '值', 'type': '类型'}
                - vertical_align: 垂直对齐方式，可选值: 'top', 'center', 'bottom'
                - text_direction: 文本方向，可选值: 'lr', 'rl', 'tb', 'bt'
                - shading: 背景填充(dict): {'fill': '填充颜色', 'color': '文本颜色', 'val': '填充类型'}
                - borders: 单元格边框(dict): 结构同表格边框
                - margins: 单元格内边距(dict): 结构同表格单元格边距
                - rowspan: 跨行数量
                - colspan: 跨列数量

        Returns:
            bool: 操作是否成功
        """


        # 获取所有行
        tr_elements = table.findall(f".//{{{self.NAMESPACES['w']}}}tr")

        # 检查行索引是否有效
        if row_index < 0 or row_index >= len(tr_elements):
            print(f"错误：行索引{row_index}超出范围(0-{len(tr_elements) - 1})")
            return False

        # 获取目标行
        tr = tr_elements[row_index]

        # 获取行中的所有单元格
        tc_elements = tr.findall(f".//{{{self.NAMESPACES['w']}}}tc")

        # 检查单元格索引是否有效
        if cell_index < 0 or cell_index >= len(tc_elements):
            print(f"错误：单元格索引{cell_index}超出范围(0-{len(tc_elements) - 1})")
            return False

        # 获取目标单元格
        tc = tc_elements[cell_index]

        # 获取或创建tcPr元素（单元格属性）
        tcPr = tc.find(f".//{{{self.NAMESPACES['w']}}}tcPr")
        if tcPr is None:
            tcPr = ET.Element(f"{{{self.NAMESPACES['w']}}}tcPr")
            tc.insert(0, tcPr)

        # 设置单元格宽度
        if 'width' in style_properties:
            width_info = style_properties['width']
            tcW = tcPr.find(f".//{{{self.NAMESPACES['w']}}}tcW")
            if tcW is None:
                tcW = ET.SubElement(tcPr, f"{{{self.NAMESPACES['w']}}}tcW")

            if 'value' in width_info:
                tcW.set(f"{{{self.NAMESPACES['w']}}}w", str(width_info['value']))
            if 'type' in width_info:
                tcW.set(f"{{{self.NAMESPACES['w']}}}type", width_info['type'])

        # 设置垂直对齐方式
        if 'vertical_align' in style_properties:
            vAlign = tcPr.find(f".//{{{self.NAMESPACES['w']}}}vAlign")
            if vAlign is None:
                vAlign = ET.SubElement(tcPr, f"{{{self.NAMESPACES['w']}}}vAlign")

            # 可选值: 'top', 'center', 'bottom'
            vAlign.set(f"{{{self.NAMESPACES['w']}}}val", style_properties['vertical_align'])

        # 设置文本方向
        if 'text_direction' in style_properties:
            textDirection = tcPr.find(f".//{{{self.NAMESPACES['w']}}}textDirection")
            if textDirection is None:
                textDirection = ET.SubElement(tcPr, f"{{{self.NAMESPACES['w']}}}textDirection")

            # 可选值: 'lr'(左到右), 'rl'(右到左), 'tb'(上到下), 'bt'(下到上)
            textDirection.set(f"{{{self.NAMESPACES['w']}}}val", style_properties['text_direction'])

        # 设置背景填充
        if 'shading' in style_properties:
            shading_info = style_properties['shading']
            shd = tcPr.find(f".//{{{self.NAMESPACES['w']}}}shd")
            if shd is None:
                shd = ET.SubElement(tcPr, f"{{{self.NAMESPACES['w']}}}shd")

            if 'val' in shading_info:
                shd.set(f"{{{self.NAMESPACES['w']}}}val", shading_info['val'])
            if 'color' in shading_info:
                shd.set(f"{{{self.NAMESPACES['w']}}}color", shading_info['color'])
            if 'fill' in shading_info:
                shd.set(f"{{{self.NAMESPACES['w']}}}fill", shading_info['fill'])

        # 设置单元格边框
        if 'borders' in style_properties:
            borders_info = style_properties['borders']
            tcBorders = tcPr.find(f".//{{{self.NAMESPACES['w']}}}tcBorders")
            if tcBorders is None:
                tcBorders = ET.SubElement(tcPr, f"{{{self.NAMESPACES['w']}}}tcBorders")

            border_mapping = {
                'top': 'top',
                'left': 'left',
                'bottom': 'bottom',
                'right': 'right',
                'inside_h': 'insideH',
                'inside_v': 'insideV',
                'tl2br': 'tl2br',  # 左上到右下的对角线
                'tr2bl': 'tr2bl'  # 右上到左下的对角线
            }

            for border_key, border_xml_name in border_mapping.items():
                if border_key in borders_info:
                    border_info = borders_info[border_key]
                    border_element = tcBorders.find(f".//{{{self.NAMESPACES['w']}}}{border_xml_name}")
                    if border_element is None:
                        border_element = ET.SubElement(tcBorders, f"{{{self.NAMESPACES['w']}}}{border_xml_name}")

                    for attr_name, xml_attr in [
                        ('val', 'val'),
                        ('color', 'color'),
                        ('sz', 'sz'),
                        ('space', 'space')
                    ]:
                        if attr_name in border_info:
                            border_element.set(f"{{{self.NAMESPACES['w']}}}{xml_attr}", str(border_info[attr_name]))

        # 设置单元格内边距
        if 'margins' in style_properties:
            margin_info = style_properties['margins']
            tcMar = tcPr.find(f".//{{{self.NAMESPACES['w']}}}tcMar")
            if tcMar is None:
                tcMar = ET.SubElement(tcPr, f"{{{self.NAMESPACES['w']}}}tcMar")

            for margin_type in ['top', 'left', 'bottom', 'right']:
                if margin_type in margin_info:
                    margin_element = tcMar.find(f".//{{{self.NAMESPACES['w']}}}{margin_type}")
                    if margin_element is None:
                        margin_element = ET.SubElement(tcMar, f"{{{self.NAMESPACES['w']}}}{margin_type}")

                    margin_data = margin_info[margin_type]
                    if 'value' in margin_data:
                        margin_element.set(f"{{{self.NAMESPACES['w']}}}w", str(margin_data['value']))
                    if 'type' in margin_data:
                        margin_element.set(f"{{{self.NAMESPACES['w']}}}type", margin_data['type'])

        # 设置跨行和跨列
        # 注意: 这些属性通常在表格创建时设置，修改现有表格的合并单元格需要更复杂的处理

        # 设置跨列(水平合并)
        if 'colspan' in style_properties:
            gridSpan = tcPr.find(f".//{{{self.NAMESPACES['w']}}}gridSpan")
            if style_properties['colspan'] > 1:
                if gridSpan is None:
                    gridSpan = ET.SubElement(tcPr, f"{{{self.NAMESPACES['w']}}}gridSpan")
                gridSpan.set(f"{{{self.NAMESPACES['w']}}}val", str(style_properties['colspan']))
            else:
                # 移除跨列标记
                if gridSpan is not None:
                    tcPr.remove(gridSpan)

        # 设置跨行(垂直合并)
        if 'rowspan' in style_properties:
            if style_properties['rowspan'] > 1:
                # 对于起始单元格，需要设置vMerge="restart"
                vMerge = tcPr.find(f".//{{{self.NAMESPACES['w']}}}vMerge")
                if vMerge is None:
                    vMerge = ET.SubElement(tcPr, f"{{{self.NAMESPACES['w']}}}vMerge")
                vMerge.set(f"{{{self.NAMESPACES['w']}}}val", "restart")

                # 对于后续的跨行单元格，需要设置vMerge而不指定值
                # 这需要在后续行的相应单元格上设置
                for i in range(1, style_properties['rowspan']):
                    next_row_index = row_index + i
                    if next_row_index < len(tr_elements):
                        next_row = tr_elements[next_row_index]
                        next_cells = next_row.findall(f".//{{{self.NAMESPACES['w']}}}tc")

                        if cell_index < len(next_cells):
                            next_cell = next_cells[cell_index]
                            next_tcPr = next_cell.find(f".//{{{self.NAMESPACES['w']}}}tcPr")
                            if next_tcPr is None:
                                next_tcPr = ET.Element(f"{{{self.NAMESPACES['w']}}}tcPr")
                                next_cell.insert(0, next_tcPr)

                            next_vMerge = next_tcPr.find(f".//{{{self.NAMESPACES['w']}}}vMerge")
                            if next_vMerge is None:
                                next_vMerge = ET.SubElement(next_tcPr, f"{{{self.NAMESPACES['w']}}}vMerge")
                            # 不设置值，表示这是被合并的单元格
            else:
                # 移除垂直合并标记
                vMerge = tcPr.find(f".//{{{self.NAMESPACES['w']}}}vMerge")
                if vMerge is not None:
                    tcPr.remove(vMerge)

        # 更新XML
        self.update_document_xml()
        return True

    def create_paragraph_in_cell(self, table, row_index, cell_index, **paragraph_properties):
        """在指定表格单元格中创建新的段落元素

        Args:
            table: 表格元素
            row_index: 行索引
            cell_index: 单元格索引
            **paragraph_properties: 可选的段落属性，可包含以下参数:
                - text: 段落文本内容
                - alignment: 对齐方式，可选值: 'left', 'center', 'right', 'both', 'justify'
                - indent_left: 左缩进值
                - indent_right: 右缩进值
                - indent_first_line: 首行缩进值
                - spacing_before: 段前间距
                - spacing_after: 段后间距
                - line_spacing: 行间距
                - line_rule: 行距规则，可选值: 'auto', 'atLeast', 'exact'
                - style_id: 段落样式ID，如'1'表示正文，'2'表示标题1等

        Returns:
            tuple: (bool, paragraph_element)，表示操作是否成功及创建的段落元素
        """
        try:
            # 获取所有行
            tr_elements = table.findall(f".//{{{self.NAMESPACES['w']}}}tr")

            # 检查行索引是否有效
            if row_index < 0 or row_index >= len(tr_elements):
                print(f"错误：行索引{row_index}超出范围(0-{len(tr_elements) - 1})")
                return False, None

            # 获取目标行
            tr = tr_elements[row_index]

            # 获取行中的所有单元格
            tc_elements = tr.findall(f".//{{{self.NAMESPACES['w']}}}tc")

            # 检查单元格索引是否有效
            if cell_index < 0 or cell_index >= len(tc_elements):
                print(f"错误：单元格索引{cell_index}超出范围(0-{len(tc_elements) - 1})")
                return False, None

            # 获取目标单元格
            tc = tc_elements[cell_index]

            # 创建新段落元素
            paragraph = ET.SubElement(tc, f"{{{self.NAMESPACES['w']}}}p")

            # 如果有提供段落样式属性，创建段落属性元素
            if any(key in paragraph_properties for key in
                   ['alignment', 'indent_left', 'indent_right', 'indent_first_line',
                    'spacing_before', 'spacing_after', 'line_spacing', 'line_rule', 'style_id']):
                pPr = ET.SubElement(paragraph, f"{{{self.NAMESPACES['w']}}}pPr")

                # 设置段落样式ID
                if 'style_id' in paragraph_properties:
                    style = ET.SubElement(pPr, f"{{{self.NAMESPACES['w']}}}pStyle")
                    style.set(f"{{{self.NAMESPACES['w']}}}val", paragraph_properties['style_id'])

                # 设置对齐方式
                if 'alignment' in paragraph_properties:
                    jc = ET.SubElement(pPr, f"{{{self.NAMESPACES['w']}}}jc")
                    jc.set(f"{{{self.NAMESPACES['w']}}}val", paragraph_properties['alignment'])

                # 设置缩进
                if any(key in paragraph_properties for key in ['indent_left', 'indent_right', 'indent_first_line']):
                    ind = ET.SubElement(pPr, f"{{{self.NAMESPACES['w']}}}ind")

                    if 'indent_left' in paragraph_properties:
                        ind.set(f"{{{self.NAMESPACES['w']}}}left", str(paragraph_properties['indent_left']))

                    if 'indent_right' in paragraph_properties:
                        ind.set(f"{{{self.NAMESPACES['w']}}}right", str(paragraph_properties['indent_right']))

                    if 'indent_first_line' in paragraph_properties:
                        ind.set(f"{{{self.NAMESPACES['w']}}}firstLine", str(paragraph_properties['indent_first_line']))

                # 设置段落间距
                if any(key in paragraph_properties for key in
                       ['spacing_before', 'spacing_after', 'line_spacing', 'line_rule']):
                    spacing = ET.SubElement(pPr, f"{{{self.NAMESPACES['w']}}}spacing")

                    if 'spacing_before' in paragraph_properties:
                        spacing.set(f"{{{self.NAMESPACES['w']}}}before", str(paragraph_properties['spacing_before']))

                    if 'spacing_after' in paragraph_properties:
                        spacing.set(f"{{{self.NAMESPACES['w']}}}after", str(paragraph_properties['spacing_after']))

                    if 'line_spacing' in paragraph_properties:
                        spacing.set(f"{{{self.NAMESPACES['w']}}}line", str(paragraph_properties['line_spacing']))

                    if 'line_rule' in paragraph_properties:
                        spacing.set(f"{{{self.NAMESPACES['w']}}}lineRule", paragraph_properties['line_rule'])

            # 如果提供了文本内容，添加一个文本运行
            if 'text' in paragraph_properties:
                run = ET.SubElement(paragraph, f"{{{self.NAMESPACES['w']}}}r")
                t = ET.SubElement(run, f"{{{self.NAMESPACES['w']}}}t")
                # 设置space属性为preserve以保留空格
                t.set(f"{{{self.NAMESPACES['xml']}}}space", "preserve")
                t.text = paragraph_properties['text']

                # 如果有其他运行样式属性，可以在这里设置

            # 更新XML
            self.update_document_xml()

            print(f"在表格单元格({row_index}, {cell_index})中创建了新段落")
            return True, paragraph

        except Exception as e:
            print(f"在表格单元格中创建段落时出错: {e}")
            import traceback
            traceback.print_exc()
            return False, None
    def set_table_grid(self, table_index, column_widths):
        """设置表格的列宽

        Args:
            table_index: self.tables中的表格索引
            column_widths: 列宽列表，每个元素是一个数值，表示列的宽度(单位：二十分之一点)

        Returns:
            bool: 操作是否成功
        """
        # 检查索引是否有效
        absolute_index = abs(table_index)
        if absolute_index >= len(self.tables):
            print(f"错误：表格索引{table_index}超出范围(0-{len(self.tables)-1})")
            return False

        # 获取表格元素
        table = self.tables[table_index]['element']

        # 获取或创建tblGrid元素
        tblGrid = table.find(f".//{{{self.NAMESPACES['w']}}}tblGrid")
        if tblGrid is None:
            tblGrid = ET.Element(f"{{{self.NAMESPACES['w']}}}tblGrid")
            # 插入在tblPr之后
            tblPr = table.find(f".//{{{self.NAMESPACES['w']}}}tblPr")
            if tblPr is not None:
                tblPr_index = list(table).index(tblPr)
                table.insert(tblPr_index + 1, tblGrid)
            else:
                table.insert(0, tblGrid)

        # 清除现有的gridCol元素
        for gridCol in tblGrid.findall(f".//{{{self.NAMESPACES['w']}}}gridCol"):
            tblGrid.remove(gridCol)

        # 添加新的gridCol元素
        for width in column_widths:
            gridCol = ET.SubElement(tblGrid, f"{{{self.NAMESPACES['w']}}}gridCol")
            gridCol.set(f"{{{self.NAMESPACES['w']}}}w", str(width))



        return True

    def set_table_borders(self, table_index, **borders):
        """设置表格的边框样式

        Args:
            table_index: self.tables中的表格索引
            **borders: 可以包含以下属性:
                - top: 上边框 {'val': 类型, 'color': 颜色, 'sz': 粗细, 'space': 间距}
                - left: 左边框
                - bottom: 下边框
                - right: 右边框
                - inside_h: 水平内边框
                - inside_v: 垂直内边框

                边框类型(val)可选值: 'single', 'double', 'thick', 'none' 等
                颜色(color)格式: 'auto' 或 六位十六进制颜色值如 '000000'
                粗细(sz)单位: 1/8点，常用值: 4(0.5pt), 8(1pt), 12(1.5pt), 16(2pt)等

        Returns:
            bool: 操作是否成功
        """
        # 简化调用set_table_style，只传入borders参数
        return self.set_table_style(table_index, borders=borders)

    def set_table_cell_margins(self, table_index, **margins):
        """设置表格的单元格边距

        Args:
            table_index: self.tables中的表格索引
            **margins: 可以包含以下属性:
                - top: 上边距 {'value': 值, 'type': 单位类型}
                - left: 左边距
                - bottom: 下边距
                - right: 右边距

                单位类型(type)通常为'dxa'(二十分之一点)

        Returns:
            bool: 操作是否成功
        """
        # 简化调用set_table_style，只传入cell_margins参数
        return self.set_table_style(table_index, cell_margins=margins)

    def set_table_width(self, table_index, width, width_type='dxa'):
        """设置表格的宽度

        Args:
            table_index: self.tables中的表格索引
            width: 宽度值
            width_type: 宽度单位类型，默认'dxa'(二十分之一点)，
                        可选: 'auto'(自动), 'pct'(百分比)

        Returns:
            bool: 操作是否成功
        """
        width_info = {'value': str(width), 'type': width_type}
        return self.set_table_style(table_index, width=width_info)

    def set_table_row_borders(self, table_index, row_index, **borders):
        """设置表格特定行的边框样式（同时设置行级和单元格级边框）

        Args:
            table_index: self.tables中的表格索引
            row_index: 行索引，从0开始
            **borders: 可以包含以下属性:
                - top: 上边框 {'val': 类型, 'color': 颜色, 'sz': 粗细, 'space': 间距}
                - left: 左边框
                - bottom: 下边框
                - right: 右边框
                - inside_h: 水平内边框
                - inside_v: 垂直内边框

                边框类型(val)可选值: 'single', 'double', 'thick', 'none' 等
                颜色(color)格式: 'auto' 或 六位十六进制颜色值如 '000000'
                粗细(sz)单位: 1/8点，常用值: 4(0.5pt), 8(1pt), 12(1.5pt), 16(2pt)等

        Returns:
            bool: 操作是否成功
        """
        # 检查表格索引是否有效
        absolute_index = abs(table_index)
        if absolute_index >= len(self.tables):
            print(f"错误：表格索引{table_index}超出范围(0-{len(self.tables) - 1})")
            return False

        # 获取表格元素
        table = self.tables[table_index]['element']

        # 查找所有行
        rows = table.findall(f".//{{{self.NAMESPACES['w']}}}tr")

        # 检查行索引是否有效
        if row_index < 0 or row_index >= len(rows):
            print(f"错误：行索引{row_index}超出范围(0-{len(rows) - 1})")
            return False

        # 获取目标行
        row = rows[row_index]

        try:
            # 1. 设置行级边框（保留原功能）
            # 获取或创建tblPrEx元素 (表格属性例外，应用于特定行)
            tblPrEx = row.find(f".//{{{self.NAMESPACES['w']}}}tblPrEx")
            if tblPrEx is None:
                tblPrEx = ET.Element(f"{{{self.NAMESPACES['w']}}}tblPrEx")

                # 插入到行的第一个位置或trPr之后
                trPr = row.find(f".//{{{self.NAMESPACES['w']}}}trPr")
                if trPr is not None:
                    trPr_index = list(row).index(trPr)
                    row.insert(trPr_index + 1, tblPrEx)
                else:
                    row.insert(0, tblPrEx)

            # 获取或创建tblBorders元素
            tblBorders = tblPrEx.find(f".//{{{self.NAMESPACES['w']}}}tblBorders")
            if tblBorders is None:
                tblBorders = ET.SubElement(tblPrEx, f"{{{self.NAMESPACES['w']}}}tblBorders")

            # 设置行级边框
            border_mapping = {
                'top': 'top',
                'left': 'left',
                'bottom': 'bottom',
                'right': 'right',
                'inside_h': 'insideH',
                'inside_v': 'insideV'
            }

            for border_key, border_xml_name in border_mapping.items():
                if border_key in borders:
                    border_info = borders[border_key]
                    border_element = tblBorders.find(f".//{{{self.NAMESPACES['w']}}}{border_xml_name}")
                    if border_element is None:
                        border_element = ET.SubElement(tblBorders, f"{{{self.NAMESPACES['w']}}}{border_xml_name}")

                    for attr_name, xml_attr in [
                        ('val', 'val'),
                        ('color', 'color'),
                        ('sz', 'sz'),
                        ('space', 'space')
                    ]:
                        if attr_name in border_info:
                            border_element.set(f"{{{self.NAMESPACES['w']}}}{xml_attr}", str(border_info[attr_name]))

            # 2. 添加设置单元格级边框（新功能）
            # 找到这一行中的所有单元格
            cells = row.findall(f".//{{{self.NAMESPACES['w']}}}tc")
            for cell in cells:
                # 获取或创建tcPr元素
                tcPr = cell.find(f".//{{{self.NAMESPACES['w']}}}tcPr")
                if tcPr is None:
                    tcPr = ET.SubElement(cell, f"{{{self.NAMESPACES['w']}}}tcPr")

                # 获取或创建tcBorders元素
                tcBorders = tcPr.find(f".//{{{self.NAMESPACES['w']}}}tcBorders")
                if tcBorders is None:
                    tcBorders = ET.SubElement(tcPr, f"{{{self.NAMESPACES['w']}}}tcBorders")

                # 设置单元格边框 - 仅设置top和bottom，这是三线表的关键
                cell_border_keys = ['top', 'bottom']  # 主要关注这两个边框

                for border_key in cell_border_keys:
                    if border_key in borders:
                        border_info = borders[border_key]
                        border_element = tcBorders.find(f".//{{{self.NAMESPACES['w']}}}{border_key}")
                        if border_element is None:
                            border_element = ET.SubElement(tcBorders, f"{{{self.NAMESPACES['w']}}}{border_key}")

                        for attr_name, xml_attr in [
                            ('val', 'val'),
                            ('color', 'color'),
                            ('sz', 'sz'),
                            ('space', 'space')
                        ]:
                            if attr_name in border_info:
                                border_element.set(f"{{{self.NAMESPACES['w']}}}{xml_attr}", str(border_info[attr_name]))
            self.update_document_xml()

            return True

        except Exception as e:
            print(f"设置表格行边框时出错: {e}")
            traceback.print_exc()
            return False

    def set_table_cell_borders(self, table_index, row_index, cell_index, **borders):
        """设置表格特定单元格的边框样式

        Args:
            table_index: self.tables中的表格索引
            row_index: 行索引，从0开始
            cell_index: 单元格索引，从0开始
            **borders: 可以包含以下属性:
                - top: 上边框 {'val': 类型, 'color': 颜色, 'sz': 粗细, 'space': 间距}
                - left: 左边框
                - bottom: 下边框
                - right: 右边框

                边框类型(val)可选值: 'single', 'double', 'thick', 'none' 等
                颜色(color)格式: 'auto' 或 六位十六进制颜色值如 '000000'
                粗细(sz)单位: 1/8点，常用值: 4(0.5pt), 8(1pt), 12(1.5pt), 16(2pt)等

        Returns:
            bool: 操作是否成功
        """
        # 检查表格索引是否有效
        absolute_index = abs(table_index)
        if absolute_index >= len(self.tables):
            print(f"错误：表格索引{table_index}超出范围(0-{len(self.tables)-1})")
            return False

        # 获取表格元素
        table = self.tables[table_index]['element']

        # 查找所有行
        rows = table.findall(f".//{{{self.NAMESPACES['w']}}}tr")

        # 检查行索引是否有效
        if row_index < 0 or row_index >= len(rows):
            print(f"错误：行索引{row_index}超出范围(0-{len(rows)-1})")
            return False

        # 获取目标行
        row = rows[row_index]

        # 查找所有单元格
        cells = row.findall(f".//{{{self.NAMESPACES['w']}}}tc")

        # 检查单元格索引是否有效
        if cell_index < 0 or cell_index >= len(cells):
            print(f"错误：单元格索引{cell_index}超出范围(0-{len(cells)-1})")
            return False

        # 获取目标单元格
        cell = cells[cell_index]

        try:
            # 获取或创建tcPr元素
            tcPr = cell.find(f".//{{{self.NAMESPACES['w']}}}tcPr")
            if tcPr is None:
                tcPr = ET.Element(f"{{{self.NAMESPACES['w']}}}tcPr")
                cell.insert(0, tcPr)

            # 获取或创建tcBorders元素
            tcBorders = tcPr.find(f".//{{{self.NAMESPACES['w']}}}tcBorders")
            if tcBorders is None:
                tcBorders = ET.SubElement(tcPr, f"{{{self.NAMESPACES['w']}}}tcBorders")

            # 设置边框
            for border_key, border_xml_name in [
                ('top', 'top'),
                ('left', 'left'),
                ('bottom', 'bottom'),
                ('right', 'right')
            ]:
                if border_key in borders:
                    border_info = borders[border_key]
                    border_element = tcBorders.find(f".//{{{self.NAMESPACES['w']}}}{border_xml_name}")
                    if border_element is None:
                        border_element = ET.SubElement(tcBorders, f"{{{self.NAMESPACES['w']}}}{border_xml_name}")

                    for attr_name, xml_attr in [
                        ('val', 'val'),
                        ('color', 'color'),
                        ('sz', 'sz'),
                        ('space', 'space')
                    ]:
                        if attr_name in border_info:
                            border_element.set(f"{{{self.NAMESPACES['w']}}}{xml_attr}", str(border_info[attr_name]))


            return True

        except Exception as e:
            print(f"设置表格单元格边框时出错: {e}")
            traceback.print_exc()
            return False
    def set_table_cell_borders_from_xml(self, table, row_index, cell_index, **borders):
        """设置表格特定单元格的边框样式

        Args:
            table: 表格元素
            row_index: 行索引，从0开始
            cell_index: 单元格索引，从0开始
            **borders: 可以包含以下属性:
                - top: 上边框 {'val': 类型, 'color': 颜色, 'sz': 粗细, 'space': 间距}
                - left: 左边框
                - bottom: 下边框
                - right: 右边框

                边框类型(val)可选值: 'single', 'double', 'thick', 'none' 等
                颜色(color)格式: 'auto' 或 六位十六进制颜色值如 '000000'
                粗细(sz)单位: 1/8点，常用值: 4(0.5pt), 8(1pt), 12(1.5pt), 16(2pt)等

        Returns:
            bool: 操作是否成功
        """
        # 检查表格索引是否有效


        # 查找所有行
        rows = table.findall(f".//{{{self.NAMESPACES['w']}}}tr")

        # 检查行索引是否有效
        if row_index < 0 or row_index >= len(rows):
            print(f"错误：行索引{row_index}超出范围(0-{len(rows)-1})")
            return False

        # 获取目标行
        row = rows[row_index]

        # 查找所有单元格
        cells = row.findall(f".//{{{self.NAMESPACES['w']}}}tc")

        # 检查单元格索引是否有效
        if cell_index < 0 or cell_index >= len(cells):
            print(f"错误：单元格索引{cell_index}超出范围(0-{len(cells)-1})")
            return False

        # 获取目标单元格
        cell = cells[cell_index]

        try:
            # 获取或创建tcPr元素
            tcPr = cell.find(f".//{{{self.NAMESPACES['w']}}}tcPr")
            if tcPr is None:
                tcPr = ET.Element(f"{{{self.NAMESPACES['w']}}}tcPr")
                cell.insert(0, tcPr)

            # 获取或创建tcBorders元素
            tcBorders = tcPr.find(f".//{{{self.NAMESPACES['w']}}}tcBorders")
            if tcBorders is None:
                tcBorders = ET.SubElement(tcPr, f"{{{self.NAMESPACES['w']}}}tcBorders")

            # 设置边框
            for border_key, border_xml_name in [
                ('top', 'top'),
                ('left', 'left'),
                ('bottom', 'bottom'),
                ('right', 'right')
            ]:
                if border_key in borders:
                    border_info = borders[border_key]
                    border_element = tcBorders.find(f".//{{{self.NAMESPACES['w']}}}{border_xml_name}")
                    if border_element is None:
                        border_element = ET.SubElement(tcBorders, f"{{{self.NAMESPACES['w']}}}{border_xml_name}")

                    for attr_name, xml_attr in [
                        ('val', 'val'),
                        ('color', 'color'),
                        ('sz', 'sz'),
                        ('space', 'space')
                    ]:
                        if attr_name in border_info:
                            border_element.set(f"{{{self.NAMESPACES['w']}}}{xml_attr}", str(border_info[attr_name]))


            return True

        except Exception as e:
            print(f"设置表格单元格边框时出错: {e}")
            traceback.print_exc()
            return False
    def create_three_line_table(self, table_index):
        """将表格样式设置为标准三线表

        Args:
            table_index: self.tables中的表格索引

        Returns:
            bool: 操作是否成功
        """
        try:
            # 检查索引是否有效
            absolute_index = abs(table_index)
            if absolute_index >= len(self.tables):
                print(f"错误：表格索引{table_index}超出范围(0-{len(self.tables) - 1})")
                return False

            # 获取表格行数
            row_count = self.get_table_dimensions(table_index)[0]
            if row_count < 2:
                print(f"警告：表格至少需要两行才能设置为标准三线表，当前只有 {row_count} 行")
                return False


            rows, cols = self.get_table_dimensions(table_index)

            # 1. 清空所有边框
            self.set_table_borders(
                table_index,
                top={"val": "none"},
                bottom={"val": "none"},
                left={"val": "none"},
                right={"val": "none"},
                inside_h={"val": "none"},
                inside_v={"val": "none"}
            )
            for row in range(rows):
                self.set_table_row_borders(
                    table_index,
                    row_index=row,
                    top={"val": "none"},
                    bottom={"val": "none"},
                    left={"val": "none"},
                    right={"val": "none"},
                    inside_h={"val": "none"},
                    inside_v={"val": "none"}
                )
                for col in range(cols):
                    self.set_table_cell_borders(
                        table_index, row, col,
                        top={"val": "none"},
                        bottom={"val": "none"},
                        left={"val": "none"},
                        right={"val": "none"}
                    )

            # 2. 设置三线
            # 顶线
            # 2. 设置三线
            # 顶线：最上面一行所有单元格加顶边框
            for col in range(cols):
                self.set_table_cell_borders(
                    table_index, 0, col,
                    top={"val": "single", "sz": 12, "color": "000000", "space": "0"}
                )
            # 表头下边框
            for col in range(cols):
                self.set_table_cell_borders(
                    table_index, 0, col,
                    bottom={"val": "single", "sz": 4, "color": "000000", "space": "0"}
                )
            # 底线：最下面一行所有单元格加底边框
            for col in range(cols):
                self.set_table_cell_borders(
                    table_index, rows - 1, col,
                    bottom={"val": "single", "sz": 12, "color": "000000", "space": "0"}
                )
            print("三线表设置完成")
            return True

        except Exception as e:
            print(f"创建三线表时出错: {e}")
            import traceback
            traceback.print_exc()
            return False

    def insert_table(self, element_index=-1, position='after', rows=2, cols=2, data=None, **style_properties):
        """在文档中插入新表格

        Args:
            element_index: self.elements中的元素索引，支持负索引（如-1表示最后一个元素）
            position: 插入位置，'before'表示在元素前插入，'after'表示在元素后插入
            rows: 表格行数
            cols: 表格列数
            data: 表格数据，二维列表，每个元素为单元格文本。如果为None，创建空表格
            **style_properties: 表格样式属性，可包含以下键：
                'style_id': 样式ID
                'width': 表格宽度(dict): {'value': '值', 'type': '类型'}
                'borders': 边框设置
                'layout': 布局类型('autofit'或'fixed')
                'cell_margins': 单元格边距
                'grid': 列宽列表
                'three_line_style': 是否设置为三线表格式(True/False)

        Returns:
            int: 新表格在self.tables中的索引，失败则返回-1
        """
        # 处理负索引
        elements_count = len(self.elements)
        if element_index < 0:
            element_index = elements_count + element_index

        # 检查索引是否有效
        if element_index < 0 or element_index >= elements_count:
            print(f"错误：元素索引{element_index}超出范围(0-{elements_count-1})")
            return -1

        try:
            # 获取目标元素
            target_element = self.elements[element_index]['element']

            # 创建新表格元素
            table = ET.Element(f"{{{self.NAMESPACES['w']}}}tbl")

            # 创建表格属性元素
            tblPr = ET.SubElement(table, f"{{{self.NAMESPACES['w']}}}tblPr")

            # 设置表格样式
            if 'style_id' in style_properties:
                tblStyle = ET.SubElement(tblPr, f"{{{self.NAMESPACES['w']}}}tblStyle")
                tblStyle.set(f"{{{self.NAMESPACES['w']}}}val", style_properties['style_id'])

            # 设置表格宽度
            if 'width' in style_properties:
                width_info = style_properties['width']
                tblW = ET.SubElement(tblPr, f"{{{self.NAMESPACES['w']}}}tblW")
                if 'value' in width_info:
                    tblW.set(f"{{{self.NAMESPACES['w']}}}w", str(width_info['value']))
                if 'type' in width_info:
                    tblW.set(f"{{{self.NAMESPACES['w']}}}type", width_info['type'])
            else:
                # 默认宽度设置
                tblW = ET.SubElement(tblPr, f"{{{self.NAMESPACES['w']}}}tblW")
                tblW.set(f"{{{self.NAMESPACES['w']}}}w", "0")
                tblW.set(f"{{{self.NAMESPACES['w']}}}type", "auto")

            # 设置表格边框
            if 'borders' in style_properties:
                borders_info = style_properties['borders']
                tblBorders = ET.SubElement(tblPr, f"{{{self.NAMESPACES['w']}}}tblBorders")

                border_mapping = {
                    'top': 'top',
                    'left': 'left',
                    'bottom': 'bottom',
                    'right': 'right',
                    'inside_h': 'insideH',
                    'inside_v': 'insideV'
                }

                for border_key, border_xml_name in border_mapping.items():
                    if border_key in borders_info:
                        border_info = borders_info[border_key]
                        border_element = ET.SubElement(tblBorders, f"{{{self.NAMESPACES['w']}}}{border_xml_name}")

                        for attr_name, xml_attr in [
                            ('val', 'val'),
                            ('color', 'color'),
                            ('sz', 'sz'),
                            ('space', 'space')
                        ]:
                            if attr_name in border_info:
                                border_element.set(f"{{{self.NAMESPACES['w']}}}{xml_attr}", str(border_info[attr_name]))
            else:
                # 默认边框设置
                tblBorders = ET.SubElement(tblPr, f"{{{self.NAMESPACES['w']}}}tblBorders")
                for border_name in ['top', 'left', 'bottom', 'right', 'insideH', 'insideV']:
                    border = ET.SubElement(tblBorders, f"{{{self.NAMESPACES['w']}}}{border_name}")
                    border.set(f"{{{self.NAMESPACES['w']}}}val", "single")
                    border.set(f"{{{self.NAMESPACES['w']}}}sz", "4")
                    border.set(f"{{{self.NAMESPACES['w']}}}space", "0")
                    border.set(f"{{{self.NAMESPACES['w']}}}color", "auto")

            # 设置表格布局
            if 'layout' in style_properties:
                tblLayout = ET.SubElement(tblPr, f"{{{self.NAMESPACES['w']}}}tblLayout")
                tblLayout.set(f"{{{self.NAMESPACES['w']}}}type", style_properties['layout'])

            # 设置单元格边距
            if 'cell_margins' in style_properties:
                margin_info = style_properties['cell_margins']
                tblCellMar = ET.SubElement(tblPr, f"{{{self.NAMESPACES['w']}}}tblCellMar")

                for margin_type in ['top', 'left', 'bottom', 'right']:
                    if margin_type in margin_info:
                        margin_element = ET.SubElement(tblCellMar, f"{{{self.NAMESPACES['w']}}}{margin_type}")
                        margin_data = margin_info[margin_type]
                        if 'value' in margin_data:
                            margin_element.set(f"{{{self.NAMESPACES['w']}}}w", str(margin_data['value']))
                        if 'type' in margin_data:
                            margin_element.set(f"{{{self.NAMESPACES['w']}}}type", margin_data['type'])
            else:
                # 默认单元格边距
                tblCellMar = ET.SubElement(tblPr, f"{{{self.NAMESPACES['w']}}}tblCellMar")
                margin_types = {'top': '100', 'left': '100', 'bottom': '100', 'right': '100'}
                for m_type, m_val in margin_types.items():
                    margin = ET.SubElement(tblCellMar, f"{{{self.NAMESPACES['w']}}}{m_type}")
                    margin.set(f"{{{self.NAMESPACES['w']}}}w", m_val)
                    margin.set(f"{{{self.NAMESPACES['w']}}}type", "dxa")

            # 创建表格网格(列定义)
            tblGrid = ET.SubElement(table, f"{{{self.NAMESPACES['w']}}}tblGrid")

            # 如果提供了列宽，使用提供的值
            if 'grid' in style_properties and isinstance(style_properties['grid'], list):
                grid_widths = style_properties['grid']
                for width in grid_widths[:cols]:  # 确保不超过需要的列数
                    gridCol = ET.SubElement(tblGrid, f"{{{self.NAMESPACES['w']}}}gridCol")
                    gridCol.set(f"{{{self.NAMESPACES['w']}}}w", str(width))
            else:
                # 默认均等列宽
                default_width = "2000"  # 默认列宽
                for _ in range(cols):
                    gridCol = ET.SubElement(tblGrid, f"{{{self.NAMESPACES['w']}}}gridCol")
                    gridCol.set(f"{{{self.NAMESPACES['w']}}}w", default_width)

            # 创建表格行和单元格
            for row_idx in range(rows):
                tr = ET.SubElement(table, f"{{{self.NAMESPACES['w']}}}tr")

                for col_idx in range(cols):
                    tc = ET.SubElement(tr, f"{{{self.NAMESPACES['w']}}}tc")

                    # 创建单元格属性
                    tcPr = ET.SubElement(tc, f"{{{self.NAMESPACES['w']}}}tcPr")

                    # 单元格宽度
                    tcW = ET.SubElement(tcPr, f"{{{self.NAMESPACES['w']}}}tcW")
                    tcW.set(f"{{{self.NAMESPACES['w']}}}w", "0")
                    tcW.set(f"{{{self.NAMESPACES['w']}}}type", "auto")

                    # 创建单元格内容段落
                    p = ET.SubElement(tc, f"{{{self.NAMESPACES['w']}}}p")

                    # 如果提供了数据，填充单元格内容
                    if data is not None and row_idx < len(data) and col_idx < len(data[row_idx]):
                        cell_text = data[row_idx][col_idx]
                        if cell_text:
                            r = ET.SubElement(p, f"{{{self.NAMESPACES['w']}}}r")
                            t = ET.SubElement(r, f"{{{self.NAMESPACES['w']}}}t")
                            if cell_text.startswith(' ') or cell_text.endswith(' ') or '  ' in cell_text:
                                t.set(f"{{{self.NAMESPACES['xml']}}}space", "preserve")
                            t.text = cell_text

            # 直接在文档树中插入新表格
            # 获取文档体(body)
            body = self.root.find(f".//{{{self.NAMESPACES['w']}}}body")
            if body is None:
                print("错误：无法找到文档体(body)元素")
                return -1

            # 查找目标元素在body中的位置
            body_children = list(body)
            target_index = -1
            for i, child in enumerate(body_children):
                if child == target_element:
                    target_index = i
                    break

            if target_index == -1:
                # 如果找不到目标元素，可能是因为它不是body的直接子元素
                # 尝试使用elements中的信息找到正确的位置
                target_info = self.elements[element_index]
                if 'index' in target_info:
                    # 使用索引信息定位
                    target_index = target_info['index']

            if target_index == -1:
                print("错误：无法在文档树中定位目标元素")
                return -1

            # 根据position参数插入表格
            if position.lower() == 'before':
                body.insert(target_index, table)
            else:  # 默认在后面插入
                body.insert(target_index + 1, table)

            # 重新解析文档结构，更新self.elements和self.tables
            self.get_structured_body_elements()

            # 查找插入的表格在self.tables中的索引
            for i, tbl in enumerate(self.tables):
                # 使用比较字符串内容的方式确定是否是同一张表格
                if self._elements_equal(tbl['element'], table):
                    # 如果需要设置为三线表格式
                    if style_properties.get('three_line_style', False):
                        self.create_three_line_table(i)
                    return i

            # 如果找不到插入的表格，说明出了问题
            print("警告：表格已插入，但无法在self.tables中找到")
            return -1

        except Exception as e:
            print(f"插入表格时出错: {e}")
            traceback.print_exc()
            return -1

    def update_table_text_style(self, table_index, **style_properties):
        """修改表格中所有单元格的文本样式"""
        try:
            absolute_index = abs(table_index)
            if absolute_index >= len(self.tables):
                print(f"错误：表格索引{table_index}超出范围(0-{len(self.tables) - 1})")
                return False

            table = self.tables[table_index]['element']
            header_row_different = style_properties.pop('header_row_different', True)
            header_style = style_properties.pop('header_style', {})
            replace_text = 'text' in style_properties
            new_text = style_properties.pop('text', None)
            rows = table.findall(f".//{{{self.NAMESPACES['w']}}}tr")

            for row_idx, row in enumerate(rows):
                current_style = header_style if (row_idx == 0 and header_row_different) else style_properties
                cells = row.findall(f".//{{{self.NAMESPACES['w']}}}tc")
                for cell in cells:
                    if 'vertical_alignment' in current_style:
                        tcPr = cell.find(f".//{{{self.NAMESPACES['w']}}}tcPr")
                        if tcPr is None:
                            tcPr = ET.SubElement(cell, f"{{{self.NAMESPACES['w']}}}tcPr")
                        vAlign = tcPr.find(f".//{{{self.NAMESPACES['w']}}}vAlign")
                        if vAlign is None:
                            vAlign = ET.SubElement(tcPr, f"{{{self.NAMESPACES['w']}}}vAlign")
                        vAlign.set(f"{{{self.NAMESPACES['w']}}}val", current_style['vertical_alignment'])

                    paragraphs = cell.findall(f".//{{{self.NAMESPACES['w']}}}p")
                    for p in paragraphs:
                        # 统一用 update_paragraph_style_from_xml 设置段落样式
                        self.update_paragraph_style_from_xml(p, **current_style)
                        if replace_text:
                            for r in list(p.findall(f".//{{{self.NAMESPACES['w']}}}r")):
                                p.remove(r)
                            r = ET.SubElement(p, f"{{{self.NAMESPACES['w']}}}r")
                            self.update_run_style_from_xml(p, 0, **current_style)
                            t = r.find(f".//{{{self.NAMESPACES['w']}}}t")
                            if t is None:
                                t = ET.SubElement(r, f"{{{self.NAMESPACES['w']}}}t")
                            if new_text and (new_text.startswith(' ') or new_text.endswith(' ') or '  ' in new_text):
                                t.set(f"{{{self.NAMESPACES['xml']}}}space", "preserve")
                            t.text = new_text if new_text is not None else ""
                        else:
                            runs = p.findall(f".//{{{self.NAMESPACES['w']}}}r")
                            for r_index, r in enumerate(runs):
                                self.update_run_style_from_xml(p, r_index, **current_style)
            return True
        except Exception as e:
            print(f"修改表格文本样式时出错: {e}")
            traceback.print_exc()
            return False

    def set_table_text_alignment(self, table_index, alignment='center', header_alignment=None):
        """设置表格中所有单元格的文本对齐方式

        Args:
            table_index: 表格索引
            alignment: 对齐方式，可选值: 'left', 'center', 'right', 'both'(两端对齐)
            header_alignment: 表头对齐方式，如果为None则使用alignment

        Returns:
            bool: 操作是否成功
        """
        try:
            # 检查索引是否有效
            absolute_index = abs(table_index)
            if absolute_index >= len(self.tables):
                print(f"错误：表格索引{table_index}超出范围(0-{len(self.tables)-1})")
                return False

            # 获取表格元素
            table = self.tables[table_index]['element']

            # 使用header_alignment如果提供，否则使用alignment
            if header_alignment is None:
                header_alignment = alignment

            # 查找所有行
            rows = table.findall(f".//{{{self.NAMESPACES['w']}}}tr")

            # 遍历每一行
            for row_idx, row in enumerate(rows):
                # 确定当前行使用的对齐方式
                current_alignment = header_alignment if row_idx == 0 else alignment

                # 查找行中的所有单元格
                cells = row.findall(f".//{{{self.NAMESPACES['w']}}}tc")

                # 遍历每个单元格
                for cell in cells:
                    # 查找单元格中的所有段落
                    paragraphs = cell.findall(f".//{{{self.NAMESPACES['w']}}}p")

                    # 遍历每个段落
                    for p in paragraphs:
                        # 获取或创建pPr元素
                        pPr = p.find(f".//{{{self.NAMESPACES['w']}}}pPr")
                        if pPr is None:
                            pPr = ET.Element(f"{{{self.NAMESPACES['w']}}}pPr")
                            p.insert(0, pPr)

                        # 设置对齐方式
                        jc = pPr.find(f".//{{{self.NAMESPACES['w']}}}jc")
                        if jc is None:
                            jc = ET.SubElement(pPr, f"{{{self.NAMESPACES['w']}}}jc")
                        jc.set(f"{{{self.NAMESPACES['w']}}}val", current_alignment)

            return True

        except Exception as e:
            print(f"设置表格文本对齐方式时出错: {e}")
            traceback.print_exc()
            return False

    def insert_caption(self, element_index, caption_type, chapter_num, caption_text, position='after', auto_num=True, style_id=None, **style_properties):
        """
        插入表格或图片的标题

        Args:
            element_index (int/Element): 元素索引（表格或段落的索引）或元素对象
            caption_type (str): 标题类型，'table'表示表格标题，'figure'表示图片标题
            chapter_num (str): 章节编号，如'4'
            caption_text (str): 标题描述文本
            position (str, optional): 插入位置，'before'表示在元素前，'after'表示在元素后。默认为'after'
            auto_num (bool, optional): 是否使用自动编号。默认为True
            style_id (str, optional): 标题段落样式ID。如果为None，则根据类型自动选择样式
            **style_properties: 其他样式属性，包括font(字体)、size(大小)、bold(粗体)、italic(斜体)、
                              color(颜色)、highlight(高亮)、underline(下划线)等

        Returns:
            int: 新插入的标题段落的索引
        """
        # 验证标题类型
        if caption_type not in ['table', 'figure']:
            raise ValueError("caption_type 必须是 'table' 或 'figure'")

        # 确定样式ID
        if style_id is None:
            style_id = "Caption"  # 默认使用Caption样式

        # 确定标题前缀文本
        prefix_text = "表 " if caption_type == 'table' else "图 "

        # 创建标题段落
        caption_para_index = self.insert_paragraph(
            element_index=element_index,
            position=position,
            text="",  # 先创建空段落，后面添加文本运行
            style_id=style_id,
            alignment="center"
        )
        print(1111)
        # 添加段落样式属性
        if style_properties:
            self.update_paragraph_style(caption_para_index, **style_properties)
        print(2222)
        # 添加标题文本部分
        self.insert_run(para_index=caption_para_index, text=prefix_text, **style_properties)
        print(123)
        self.insert_run(para_index=caption_para_index, text=f"{chapter_num}-", **style_properties)
        print(3333)
        if auto_num:
            # 插入自动编号字段
            self._insert_seq_field(caption_para_index, caption_type)
        else:
            # 直接插入编号
            self.insert_run(para_index=caption_para_index, text="1", **style_properties)

        # 添加描述文本
        self.insert_run(para_index=caption_para_index, text=f" {caption_text}", **style_properties)

        # 更新文档
        self.update_document_xml()
        return caption_para_index

    def insert_table_caption(self, table_index, chapter_num, caption_text, auto_num=True, style_id=None, **style_properties):
        """
        插入表格标题

        Args:
            table_index (int): 表格索引
            chapter_num (str): 章节编号，如'4'
            caption_text (str): 标题描述文本
            auto_num (bool, optional): 是否使用自动编号。默认为True
            style_id (str, optional): 标题段落样式ID。如果为None，则使用默认样式
            **style_properties: 其他样式属性，包括font(字体)、size(大小)、bold(粗体)、italic(斜体)等

        Returns:
            int: 新插入的标题段落的索引

        Example:
            doc.insert_table_caption(0, "4", "用户信息表",
                                     font={'eastAsia': '宋体'},
                                     size=24,
                                     bold=True)
        """
        # 获取表格元素
        tables = self.get_all_tables()
        if table_index >= len(tables):
            raise ValueError(f"表格索引超出范围: {table_index}")

        table_element = tables[table_index]['index']  # 获取索引而非元素对象

        # 在表格前插入标题
        return self.insert_caption(
            element_index=table_element,
            caption_type='table',
            chapter_num=chapter_num,
            caption_text=caption_text,
            position='before',  # 表格标题通常在表格上方
            auto_num=auto_num,
            style_id=style_id,
            **style_properties
        )

    def insert_figure_caption(self, para_index, chapter_num, caption_text, auto_num=True, style_id=None, **style_properties):
        """
        插入图片标题

        Args:
            para_index (int): 包含图片的段落索引
            chapter_num (str): 章节编号，如'4'
            caption_text (str): 标题描述文本
            auto_num (bool, optional): 是否使用自动编号。默认为True
            style_id (str, optional): 标题段落样式ID。如果为None，则使用默认样式

        Returns:
            int: 新插入的标题段落的索引

        Example:
            # 先插入图片
            doc.insert_image(para_index=-1, image_path="example.jpg", width=10, height=8)
            # 然后插入图片标题
            doc.insert_figure_caption(-1, "4", "登录页面")
        """
        # 验证段落索引
        if not isinstance(para_index, int):
            raise TypeError("段落索引必须是整数")

        # 在图片段落后插入标题
        return self.insert_caption(
            element_index=para_index,  # 直接传递段落索引
            caption_type='figure',
            chapter_num=chapter_num,
            caption_text=caption_text,
            position='after',  # 图片标题通常在图片下方
            auto_num=auto_num,
            style_id=style_id,
            **style_properties
        )

    def _insert_seq_field(self, para_index, seq_name):
        """
        在段落中插入SEQ字段(自动编号字段)

        Args:
            para_index (int): 段落索引
            seq_name (str): 序列名称，如'table'或'figure'
        """
        # 获取段落元素
        paras = self.get_all_paragraphs()
        if para_index < 0:
            para_index = len(paras) + para_index

        if para_index < 0 or para_index >= len(paras):
            raise ValueError(f"段落索引超出范围: {para_index}")

        # 获取实际的段落element，而不是字典
        para = paras[para_index]
        if isinstance(para, dict) and 'element' in para:
            para = para['element']
        elif isinstance(para, dict):
            raise TypeError(f"无法获取段落元素，获取到的是字典: {para}")

        # 确定序列类型的中文名称
        seq_type = "表" if seq_name == 'table' else "图"

        # 创建字段开始标记
        r_begin = ET.SubElement(para, f"{{{self.NAMESPACES['w']}}}r")
        rpr_begin = ET.SubElement(r_begin, f"{{{self.NAMESPACES['w']}}}rPr")
        ET.SubElement(r_begin, f"{{{self.NAMESPACES['w']}}}fldChar", attrib={f"{{{self.NAMESPACES['w']}}}fldCharType": "begin"})

        # 创建字段指令文本
        r_instr = ET.SubElement(para, f"{{{self.NAMESPACES['w']}}}r")
        rpr_instr = ET.SubElement(r_instr, f"{{{self.NAMESPACES['w']}}}rPr")
        instr_text = ET.SubElement(r_instr, f"{{{self.NAMESPACES['w']}}}instrText")
        instr_text.text = f" SEQ {seq_type} \\* ARABIC "
        # 使用字符串拼接而不是f-string以避免KeyError
        instr_text.set("{" + self.NAMESPACES['xml'] + "}space", "preserve")

        # 创建字段分隔符
        r_sep = ET.SubElement(para, f"{{{self.NAMESPACES['w']}}}r")
        rpr_sep = ET.SubElement(r_sep, f"{{{self.NAMESPACES['w']}}}rPr")
        ET.SubElement(r_sep, f"{{{self.NAMESPACES['w']}}}fldChar", attrib={f"{{{self.NAMESPACES['w']}}}fldCharType": "separate"})

        # 创建字段结果文本
        r_result = ET.SubElement(para, f"{{{self.NAMESPACES['w']}}}r")
        rpr_result = ET.SubElement(r_result, f"{{{self.NAMESPACES['w']}}}rPr")
        result_text = ET.SubElement(r_result, f"{{{self.NAMESPACES['w']}}}t")
        result_text.text = "1"  # 默认初始编号

        # 创建字段结束标记
        r_end = ET.SubElement(para, f"{{{self.NAMESPACES['w']}}}r")
        rpr_end = ET.SubElement(r_end, f"{{{self.NAMESPACES['w']}}}rPr")
        ET.SubElement(r_end, f"{{{self.NAMESPACES['w']}}}fldChar", attrib={f"{{{self.NAMESPACES['w']}}}fldCharType": "end"})

        # 更新文档
        self.update_document_xml()

    def insert_table_of_contents(self, element_index=-1, position='after', title="目 录",
                                 heading_levels=(1, 2, 3), style_id="10", hyperlinks=True,
                                 show_page_numbers=True, right_align_page_numbers=True,
                                 leader_char="dot", title_font=None, title_style=None,
                                 headings=None, title_style_id="TOC Heading"):
        """
        在文档中指定位置插入目录

        Args:
            element_index (int): 要插入目录的参考元素索引，默认为-1（文档末尾）
            position (str): 插入位置，'before'、'after'或'replace'
            title (str): 目录标题，默认为"目 录"
            heading_levels (tuple): 包含的标题级别范围，默认为(1,2,3)表示包含标题1-3级
            style_id (str): 目录使用的样式ID，默认为"10"（与Word标准目录一致）
            hyperlinks (bool): 是否添加超链接，默认为True
            show_page_numbers (bool): 是否显示页码，默认为True
            right_align_page_numbers (bool): 是否右对齐页码，默认为True
            leader_char (str): 页码前的引导符类型，可选"dot"、"hyphen"、"underscore"或"none"
            title_font (dict): 目录标题字体属性
            title_style (dict): 目录标题其他样式属性
            headings (list): 可选的标题列表，格式为[(索引, 文本, 级别)...]
            title_style_id (str): 目录标题使用的样式ID，默认为"TOC Heading"
                                 中文版Word可能是"目录标题"或"目录 1"等

        Returns:
            int: 新插入的目录标题段落的索引
        """
        # 在修改文档前保存当前段落样式
        original_styles = {}
        for idx, element in enumerate([i.get('element') for i in self.elements]):
            if element.tag.endswith('}p'):
                pPr = element.find(f".//{{{self.NAMESPACES['w']}}}pPr")
                if pPr is not None:
                    pStyle = pPr.find(f".//{{{self.NAMESPACES['w']}}}pStyle")
                    if pStyle is not None:
                        style_val = pStyle.get(f"{{{self.NAMESPACES['w']}}}val", "")
                        original_styles[idx] = style_val

        # 默认字体设置
        if title_font is None:
            title_font = {
                'eastAsia': '黑体',
                'size': 32,  # 16磅
                'kern': 0,
                'snapToGrid': 0
            }

        # 默认样式设置
        if title_style is None:
            title_style = {
                'alignment': 'center',
                'spacing': {
                    'line': 420,
                    'lineRule': 'exact'
                },
                'adjustRightInd': 0,
                'snapToGrid': 0
            }

        # 记录开始修改前元素的总数
        original_element_count = len([i.get('element') for i in self.elements])

        # 1. 插入目录标题
        title_index = self.insert_paragraph(
            element_index=element_index,
            position=position,
            text='',  # 空文本，稍后添加运行
            **title_style
        )

        # 重新获取元素列表
        elements = [i.get('element') for i in self.elements]

        # 设置标题字体属性
        self.set_paragraph_font(title_index, **title_font)

        # 设置目录标题样式ID - 这是新增的关键部分
        # 尝试使用不同可能的目录标题样式名称（支持中英文Word版本）
        possible_toc_heading_styles = [
            title_style_id,  # 用户指定的样式ID
            "TOC Heading",  # 英文版Word的标准目录标题
            "目录标题",  # 中文版Word常用的目录标题
            "目录 1",  # 可能的中文版变体
            "TOCHeading"  # 无空格变体
        ]

        # 尝试设置目录标题样式
        toc_heading_set = False
        for style_name in possible_toc_heading_styles:
            try:
                self.set_paragraph_style_id(title_index, style_name)
                p_element = elements[title_index]
                pPr = self._get_or_create_pPr(p_element)
                pStyle = pPr.find(f".//{{{self.NAMESPACES['w']}}}pStyle")
                if pStyle is not None and pStyle.get(f"{{{self.NAMESPACES['w']}}}val") == style_name:
                    toc_heading_set = True
                    break
            except Exception:
                # 如果样式不存在或设置失败，尝试下一个
                continue

        # 如果所有标准目录标题样式都失败，使用标题1样式作为备选
        if not toc_heading_set:
            try:
                fallback_styles = ["Heading1", "标题1", "1"]
                for style_name in fallback_styles:
                    try:
                        self.set_paragraph_style_id(title_index, style_name)
                        toc_heading_set = True
                        break
                    except Exception:
                        continue
            except Exception:
                # 如果设置失败，继续使用当前样式
                pass

        # 添加标题文本（分散排列）
        if title == "目 录":
            # 标准"目 录"形式
            self.insert_run(title_index, text="目 ")
            self.insert_run(title_index, text=" ")
            self.insert_run(title_index, text="录")
        else:
            # 直接添加自定义标题
            self.insert_run(title_index, text=title)

        # 2. 创建目录内容主段落（仅包含TOC域代码）
        toc_index = self.insert_paragraph(
            element_index=title_index,
            position='after',
            text=''
        )

        # 重新获取元素列表
        elements = [i.get('element') for i in self.elements]

        # 设置基本段落属性
        p_element = elements[toc_index]
        pPr = self._get_or_create_pPr(p_element)

        # 将样式ID设置为"10"（与Word标准目录一致）
        style_element = ET.SubElement(pPr, f"{{{self.NAMESPACES['w']}}}pStyle")
        style_element.set(f"{{{self.NAMESPACES['w']}}}val", style_id)

        # 3. 添加TOC域代码

        # 第一个运行 - 开始域
        r_element = ET.SubElement(p_element, f"{{{self.NAMESPACES['w']}}}r")
        rPr_element = ET.SubElement(r_element, f"{{{self.NAMESPACES['w']}}}rPr")
        sz_element = ET.SubElement(rPr_element, f"{{{self.NAMESPACES['w']}}}sz")
        sz_element.set(f"{{{self.NAMESPACES['w']}}}val", "28")
        szCs_element = ET.SubElement(rPr_element, f"{{{self.NAMESPACES['w']}}}szCs")
        szCs_element.set(f"{{{self.NAMESPACES['w']}}}val", "28")
        fldChar_element = ET.SubElement(r_element, f"{{{self.NAMESPACES['w']}}}fldChar")
        fldChar_element.set(f"{{{self.NAMESPACES['w']}}}fldCharType", "begin")

        # 第二个运行 - 指令文本
        r_element = ET.SubElement(p_element, f"{{{self.NAMESPACES['w']}}}r")
        rPr_element = ET.SubElement(r_element, f"{{{self.NAMESPACES['w']}}}rPr")
        sz_element = ET.SubElement(rPr_element, f"{{{self.NAMESPACES['w']}}}sz")
        sz_element.set(f"{{{self.NAMESPACES['w']}}}val", "28")
        szCs_element = ET.SubElement(rPr_element, f"{{{self.NAMESPACES['w']}}}szCs")
        szCs_element.set(f"{{{self.NAMESPACES['w']}}}val", "28")
        instrText_element = ET.SubElement(r_element, f"{{{self.NAMESPACES['w']}}}instrText")
        instrText_element.set(f"{{{self.NAMESPACES['xml']}}}space", "preserve")

        # 构建TOC域代码
        toc_options = f' TOC \\o "{heading_levels[0]}-{heading_levels[1]}"'
        if hyperlinks:
            toc_options += " \\h"
        if not show_page_numbers:
            toc_options += " \\n"
        toc_options += " \\z \\u"  # 标准选项：隐藏制表符和使用段落样式中的大纲级别

        instrText_element.text = toc_options

        # 第三个运行 - 分隔符
        r_element = ET.SubElement(p_element, f"{{{self.NAMESPACES['w']}}}r")
        rPr_element = ET.SubElement(r_element, f"{{{self.NAMESPACES['w']}}}rPr")
        sz_element = ET.SubElement(rPr_element, f"{{{self.NAMESPACES['w']}}}sz")
        sz_element.set(f"{{{self.NAMESPACES['w']}}}val", "28")
        szCs_element = ET.SubElement(rPr_element, f"{{{self.NAMESPACES['w']}}}szCs")
        szCs_element.set(f"{{{self.NAMESPACES['w']}}}val", "28")
        fldChar_element = ET.SubElement(r_element, f"{{{self.NAMESPACES['w']}}}fldChar")
        fldChar_element.set(f"{{{self.NAMESPACES['w']}}}fldCharType", "separate")

        # 4. 如果提供了标题列表，为每个标题创建单独的段落作为目录项
        last_index = toc_index

        if headings:
            # 过滤适合当前目录级别的标题
            filtered_headings = [h for h in headings if
                                 h[2] is not None and heading_levels[0] <= h[2] <= heading_levels[1]]

            for idx, (para_index, text, level) in enumerate(filtered_headings):
                # 为每个标题创建书签ID
                bookmark_id = f"_Toc{22700 + idx}"

                # 创建段落样式
                item_style = {}

                # 根据级别设置缩进（一级标题不缩进，二级以上缩进）
                if level > 1:
                    indent_value = (level - 1) * 420  # 每级缩进约0.21英寸
                    item_style['indentation'] = {'left': indent_value}

                # 为每个目录项创建单独的段落（关键改动）
                item_index = self.insert_paragraph(
                    element_index=last_index,
                    position='after',
                    text='',
                    **item_style
                )
                last_index = item_index

                # 重新获取元素列表（每次插入后更新）
                elements = [i.get('element') for i in self.elements]

                # 获取目录项段落元素
                p_element = elements[item_index]

                # 设置目录项样式
                pPr = self._get_or_create_pPr(p_element)

                # 使用对应级别的TOC样式
                toc_level_style = f"TOC{level}" if level <= 9 else "TOC9"
                try:
                    style_element = ET.SubElement(pPr, f"{{{self.NAMESPACES['w']}}}pStyle")
                    style_element.set(f"{{{self.NAMESPACES['w']}}}val", toc_level_style)
                except Exception:
                    # 如果样式不存在，使用默认样式
                    style_element = ET.SubElement(pPr, f"{{{self.NAMESPACES['w']}}}pStyle")
                    style_element.set(f"{{{self.NAMESPACES['w']}}}val", style_id)

                # 设置制表位（如果需要页码）
                if show_page_numbers and right_align_page_numbers:
                    tabs_element = ET.SubElement(pPr, f"{{{self.NAMESPACES['w']}}}tabs")
                    tab_element = ET.SubElement(tabs_element, f"{{{self.NAMESPACES['w']}}}tab")
                    tab_element.set(f"{{{self.NAMESPACES['w']}}}val", "right")

                    # 设置引导符
                    if leader_char == "dot":
                        tab_element.set(f"{{{self.NAMESPACES['w']}}}leader", "dot")
                    elif leader_char == "hyphen":
                        tab_element.set(f"{{{self.NAMESPACES['w']}}}leader", "hyphen")
                    elif leader_char == "underscore":
                        tab_element.set(f"{{{self.NAMESPACES['w']}}}leader", "underscore")

                    # 设置制表位位置
                    tab_element.set(f"{{{self.NAMESPACES['w']}}}pos", "8312")

                # 构建目录项内容

                # 1. 添加标题文本
                # 分析标题文本部分
                text_parts = []
                # 检测是否有章节号模式（如"第一章"）
                chapter_match = re.match(r'^(第[一二三四五六七八九十百千万\d]+[章节篇])(.*)', text)
                if chapter_match:
                    text_parts.append(chapter_match.group(1))  # 章节号
                    if chapter_match.group(2):
                        text_parts.append(chapter_match.group(2).strip())  # 标题内容
                else:
                    # 按空格拆分
                    text_parts = text.split()
                    if not text_parts:
                        text_parts = [text]

                # 为每部分创建文本运行
                for i, part in enumerate(text_parts):
                    r_element = ET.SubElement(p_element, f"{{{self.NAMESPACES['w']}}}r")
                    rPr_element = ET.SubElement(r_element, f"{{{self.NAMESPACES['w']}}}rPr")

                    # 设置字体大小
                    sz_element = ET.SubElement(rPr_element, f"{{{self.NAMESPACES['w']}}}sz")
                    sz_element.set(f"{{{self.NAMESPACES['w']}}}val", "28")
                    szCs_element = ET.SubElement(rPr_element, f"{{{self.NAMESPACES['w']}}}szCs")
                    szCs_element.set(f"{{{self.NAMESPACES['w']}}}val", "28")

                    if i == 0:  # 为第一部分添加东亚字体提示
                        rFonts_element = ET.SubElement(rPr_element, f"{{{self.NAMESPACES['w']}}}rFonts")
                        rFonts_element.set(f"{{{self.NAMESPACES['w']}}}hint", "eastAsia")

                    t_element = ET.SubElement(r_element, f"{{{self.NAMESPACES['w']}}}t")
                    t_element.text = part

                    # 在部分之间添加空格
                    if i < len(text_parts) - 1:
                        r_element = ET.SubElement(p_element, f"{{{self.NAMESPACES['w']}}}r")
                        t_element = ET.SubElement(r_element, f"{{{self.NAMESPACES['w']}}}t")
                        t_element.set(f"{{{self.NAMESPACES['xml']}}}space", "preserve")
                        t_element.text = " "

                # 2. 添加制表符（如果需要页码）
                if show_page_numbers:
                    r_element = ET.SubElement(p_element, f"{{{self.NAMESPACES['w']}}}r")
                    tab_element = ET.SubElement(r_element, f"{{{self.NAMESPACES['w']}}}tab")

                    # 3. 添加页码占位符
                    r_element = ET.SubElement(p_element, f"{{{self.NAMESPACES['w']}}}r")
                    t_element = ET.SubElement(r_element, f"{{{self.NAMESPACES['w']}}}t")
                    t_element.text = str(idx + 1)  # 使用索引+1作为页码占位符

                # 4. 为对应的标题段落添加书签（用于超链接）
                try:
                    # 重新获取元素列表（可能已经变化）
                    elements = [i.get('element') for i in self.elements]

                    # 获取对应的标题段落
                    heading_p_element = elements[para_index]

                    # 添加书签开始
                    bookmark_start = ET.Element(f"{{{self.NAMESPACES['w']}}}bookmarkStart")
                    bookmark_start.set(f"{{{self.NAMESPACES['w']}}}id", str(idx))
                    bookmark_start.set(f"{{{self.NAMESPACES['w']}}}name", bookmark_id)

                    # 添加书签结束
                    bookmark_end = ET.Element(f"{{{self.NAMESPACES['w']}}}bookmarkEnd")
                    bookmark_end.set(f"{{{self.NAMESPACES['w']}}}id", str(idx))

                    # 插入书签标记
                    heading_p_element.insert(0, bookmark_start)
                    heading_p_element.append(bookmark_end)
                except Exception as e:
                    print(f"无法为标题添加书签: {e}")

        # 更新元素列表（最后一次）
        elements = [i.get('element') for i in self.elements]

        # 创建TOC域结束运行
        p_element = elements[toc_index]
        r_element = ET.SubElement(p_element, f"{{{self.NAMESPACES['w']}}}r")
        fldChar_element = ET.SubElement(r_element, f"{{{self.NAMESPACES['w']}}}fldChar")
        fldChar_element.set(f"{{{self.NAMESPACES['w']}}}fldCharType", "end")

        # 确保新插入段落的索引正确偏移
        offset = len(elements) - original_element_count

        # 恢复原有段落的样式（仅对原始段落）
        for idx, style_val in original_styles.items():
            # 计算新的索引，考虑插入段落导致的偏移
            if idx >= title_index:
                new_idx = idx + offset
            else:
                new_idx = idx

            # 确保索引有效
            if 0 <= new_idx < len(elements):
                try:
                    # 获取段落元素
                    p_element = elements[new_idx]
                    if p_element.tag.endswith('}p'):
                        # 获取或创建段落属性
                        pPr = p_element.find(f".//{{{self.NAMESPACES['w']}}}pPr")
                        if pPr is None:
                            pPr = ET.SubElement(p_element, f"{{{self.NAMESPACES['w']}}}pPr")

                        # 查找样式元素
                        pStyle = pPr.find(f".//{{{self.NAMESPACES['w']}}}pStyle")

                        # 如果样式元素不存在，创建它
                        if pStyle is None:
                            pStyle = ET.SubElement(pPr, f"{{{self.NAMESPACES['w']}}}pStyle")

                        # 设置样式值
                        pStyle.set(f"{{{self.NAMESPACES['w']}}}val", style_val)
                except Exception as e:
                    print(f"恢复段落样式时出错 (索引 {new_idx}): {e}")

        # 更新文档
        self.update_document_xml()

        return title_index
    def create_table_of_contents_from_headings(self, element_index=-1, position='after',
                                               max_level=3, **toc_options):
        """
        基于现有标题创建目录

        Args:
            element_index (int): 要插入目录的参考元素索引
            position (str): 插入位置，'before'、'after'或'replace'
            max_level (int): 包括的最大标题级别
            **toc_options: 传递给insert_table_of_contents的其他选项

        Returns:
            int: 新插入的目录标题段落的索引
        """
        # 获取文档中的标题
        headings = self.get_heading_paragraphs()

        if not headings:
            print("警告：未在文档中找到标题")

        # 设置包含级别（至少是2级）
        min_level = 2
        if headings:
            available_levels = [h[2] for h in headings if h[2] is not None]
            if available_levels:
                min_level = min(available_levels)
                actual_max_level = min(max(available_levels), max_level)
            else:
                actual_max_level = max_level
        else:
            actual_max_level = max_level

        # 创建目录
        heading_levels = (min_level, actual_max_level+1)

        # 重要：传递收集到的标题列表给insert_table_of_contents
        return self.insert_table_of_contents(
            element_index,
            position,
            heading_levels=heading_levels,
            headings=headings,  # 传递标题列表
            **toc_options
        )

    def update_toc_field(self, toc_para_index):
        """
        更新目录域代码。
        注意：这只更新域代码，不会刷新目录内容。
        用户需要在Word中打开文档并更新域。

        Args:
            toc_para_index (int): 包含TOC域的段落索引
        """
        p_element = [ i.get('element') for i in self.elements ][toc_para_index]

        # 查找instrText元素
        instr_text = p_element.find(f".//{{{self.NAMESPACES['w']}}}instrText")
        if instr_text is not None and "TOC" in instr_text.text:
            # 域代码已存在，不需要更新
            return

        # 如果找不到域代码，可能是结构不符合预期
        print("警告：未找到目录域代码或目录结构不符合预期")


    def get_heading_paragraphs(self):
        """
        获取文档中所有标题段落

        Returns:
            list: 包含(索引, 标题文本, 级别)的元组列表
        """
        headings = []
        elements = [i .get('element') for i in self.elements]

        for i, element in enumerate(elements):
            if element.tag.endswith('}p'):
                # 查找段落样式
                pPr = element.find(f".//{{{self.NAMESPACES['w']}}}pPr")
                if pPr is not None:
                    pStyle = pPr.find(f".//{{{self.NAMESPACES['w']}}}pStyle")
                    if pStyle is not None:
                        style_val = pStyle.get(f"{{{self.NAMESPACES['w']}}}val", "")

                        # 检查是否是标题样式
                        level = None

                        # 匹配英文标题样式 (Heading1, Heading2...)
                        if style_val.lower().startswith("heading"):
                            try:
                                level = int(style_val[7:])
                            except ValueError:
                                # 不是数字后缀，尝试其他匹配
                                pass

                        # 匹配中文标题样式 (标题1, 标题2...)
                        elif style_val.startswith("标题"):
                            try:
                                level = int(style_val[2:])
                            except ValueError:
                                # 不是数字后缀，尝试其他匹配
                                pass

                        # 匹配纯数字样式 (1, 2, 3...)，通常用于标题
                        elif style_val.isdigit():
                            level = int(style_val)
                            # 通常只有1-9是标题级别
                            if level > 9:
                                level = None

                        # 其他可能的标题样式模式
                        elif "title" in style_val.lower() or "heading" in style_val.lower():
                            # 尝试从样式名称中提取级别
                            for char in style_val:
                                if char.isdigit():
                                    try:
                                        level = int(char)
                                        break
                                    except ValueError:
                                        pass

                        # 也可以检查outlineLvl属性确定大纲级别
                        if level is None:
                            outline_level = None
                            # 检查pStyle的父元素pPr中是否有outlineLvl设置
                            outline_pr = pPr.find(f".//{{{self.NAMESPACES['w']}}}outlineLvl")
                            if outline_pr is not None:
                                try:
                                    level = int(outline_pr.get(f"{{{self.NAMESPACES['w']}}}val", ""))
                                except ValueError:
                                    # 不是有效的数字值
                                    pass

                        # 如果确定是标题，添加到结果中
                        if level is not None:
                            # 获取标题文本
                            text = self.get_paragraph_text(element)
                            headings.append((i, text, level))

        return headings


    def get_outline_level(self, para_index):
        """
        获取段落的大纲级别（用于标题）

        Args:
            para_index (int): 段落索引

        Returns:
            int: 大纲级别，如果不是大纲则返回None
        """
        element =[i .get('element') for i in self.elements][para_index]
        if not element.tag.endswith('}p'):
            return None

        # 检查段落样式
        pPr = element.find(f".//{{{self.NAMESPACES['w']}}}pPr")
        if pPr is None:
            return None

        # 1. 直接检查outlineLvl属性
        outline_pr = pPr.find(f".//{{{self.NAMESPACES['w']}}}outlineLvl")
        if outline_pr is not None:
            try:
                return int(outline_pr.get(f"{{{self.NAMESPACES['w']}}}val", ""))
            except ValueError:
                pass

        # 2. 检查样式ID
        pStyle = pPr.find(f".//{{{self.NAMESPACES['w']}}}pStyle")
        if pStyle is not None:
            style_val = pStyle.get(f"{{{self.NAMESPACES['w']}}}val", "")

            # 英文标题样式
            if style_val.lower().startswith("heading"):
                try:
                    return int(style_val[7:]) - 1  # 通常Heading1对应级别0
                except ValueError:
                    pass

            # 中文标题样式
            elif style_val.startswith("标题"):
                try:
                    return int(style_val[2:]) - 1  # 标题1对应级别0
                except ValueError:
                    pass

            # 纯数字样式
            elif style_val.isdigit():
                level = int(style_val)
                if 1 <= level <= 9:
                    return level - 1

        # 无法确定大纲级别
        return None

    def remove_element(self, element_index):
        """
        从文档中删除指定索引的元素（段落、表格等）并更新XML树

        Args:
            element_index (int): 要删除的元素索引

        Returns:
            bool: 是否成功删除
        """
        try:
            # 获取元素列表
            elements = [i.get('element') for i in self.elements]

            # 检查索引是否有效
            if element_index < 0:
                element_index = len(elements) + element_index

            if element_index < 0 or element_index >= len(elements):
                print(f"错误：元素索引 {element_index} 超出范围")
                return False

            # 获取要删除的元素
            element_to_remove = elements[element_index]

            # 找到父元素
            parent = None
            for parent_elem in self.root.iter():
                for child in list(parent_elem):
                    if child is element_to_remove:
                        parent = parent_elem
                        break
                if parent:
                    break

            if parent is None:
                print("错误：找不到元素的父节点")
                return False

            # 从父元素中移除该元素
            parent.remove(element_to_remove)

            # 更新self.elements
            self.elements.pop(element_index)

            # 更新XML树
            self.update_document_xml()

            return True
        except Exception as e:
            print(f"删除元素时出错: {e}")
            return False

    def remove_paragraph(self, para_index):
        """
        从文档中删除指定索引的段落并更新XML树

        Args:
            para_index (int): 要删除的段落索引

        Returns:
            bool: 是否成功删除
        """
        try:
            # 获取所有段落
            paragraphs = []
            para_indices = []

            # 获取元素列表
            for idx, item in enumerate(self.elements):
                element = item.get('element')
                if element.tag.endswith('}p'):
                    paragraphs.append(element)
                    para_indices.append(idx)

            # 检查索引是否有效
            if para_index < 0:
                para_index = len(paragraphs) + para_index

            if para_index < 0 or para_index >= len(paragraphs):
                print(f"错误：段落索引 {para_index} 超出范围")
                return False

            # 获取段落对应的元素索引
            element_index = para_indices[para_index]

            # 调用remove_element删除段落
            return self.remove_element(element_index)
        except Exception as e:
            print(f"删除段落时出错: {e}")
            return False

    def remove_table(self, table_index):
        """
        从文档中删除指定索引的表格并更新XML树

        Args:
            table_index (int): 要删除的表格索引

        Returns:
            bool: 是否成功删除
        """
        try:
            # 获取所有表格
            tables = []
            table_indices = []

            # 获取元素列表
            for idx, item in enumerate(self.elements):
                element = item.get('element')
                if element.tag.endswith('}tbl'):
                    tables.append(element)
                    table_indices.append(idx)

            # 检查索引是否有效
            if table_index < 0:
                table_index = len(tables) + table_index

            if table_index < 0 or table_index >= len(tables):
                print(f"错误：表格索引 {table_index} 超出范围")
                return False

            # 获取表格对应的元素索引
            element_index = table_indices[table_index]

            # 调用remove_element删除表格
            return self.remove_element(element_index)
        except Exception as e:
            print(f"删除表格时出错: {e}")
            return False

    def remove_elements_between(self, start_index, end_index):
        """
        从文档中删除指定范围的元素（包括首尾）并更新XML树

        Args:
            start_index (int): 起始元素索引
            end_index (int): 结束元素索引

        Returns:
            bool: 是否成功删除
        """
        try:
            # 获取元素列表
            elements = [i.get('element') for i in self.elements]

            # 检查索引是否有效
            if start_index < 0:
                start_index = len(elements) + start_index
            if end_index < 0:
                end_index = len(elements) + end_index

            if start_index < 0 or start_index >= len(elements) or end_index < 0 or end_index >= len(elements):
                print(f"错误：索引范围 {start_index}-{end_index} 超出有效范围")
                return False

            if start_index > end_index:
                start_index, end_index = end_index, start_index

            # 从后向前删除（避免索引变化）
            for idx in range(end_index, start_index - 1, -1):
                if not self.remove_element(idx):
                    print(f"删除索引 {idx} 的元素失败")
                    return False

            return True
        except Exception as e:
            print(f"删除元素范围时出错: {e}")
            return False

    def remove_content_between_paragraphs(self, start_para_index, end_para_index):
        """
        从文档中删除两个段落之间的所有内容（包括指定的段落）并更新XML树

        Args:
            start_para_index (int): 起始段落索引
            end_para_index (int): 结束段落索引

        Returns:
            bool: 是否成功删除
        """
        try:
            # 获取所有段落
            paragraphs = []
            para_indices = []

            # 获取元素列表
            for idx, item in enumerate(self.elements):
                element = item.get('element')
                if element.tag.endswith('}p'):
                    paragraphs.append(element)
                    para_indices.append(idx)

            # 检查索引是否有效
            if start_para_index < 0:
                start_para_index = len(paragraphs) + start_para_index
            if end_para_index < 0:
                end_para_index = len(paragraphs) + end_para_index

            if start_para_index < 0 or start_para_index >= len(
                    paragraphs) or end_para_index < 0 or end_para_index >= len(paragraphs):
                print(f"错误：段落范围 {start_para_index}-{end_para_index} 超出有效范围")
                return False

            if start_para_index > end_para_index:
                start_para_index, end_para_index = end_para_index, start_para_index

            # 获取对应的元素索引
            start_element_index = para_indices[start_para_index]
            end_element_index = para_indices[end_para_index]

            # 删除元素范围
            return self.remove_elements_between(start_element_index, end_element_index)
        except Exception as e:
            print(f"删除段落范围时出错: {e}")
            return False


    def get_image_paragraphs_indices(self):
        """
        获取文档中所有包含图片的段落在self.elements中的索引

        Returns:
            list: 包含图片的段落索引列表，每项为 (元素索引, 关系ID列表)
        """
        try:
            image_paragraphs = []

            # 获取元素列表
            for idx, item in enumerate(self.elements):
                element = item.get('element')

                # 检查是否是段落
                if element.tag.endswith('}p'):
                    # 查找drawing元素（内联图片）
                    drawings = element.findall(f".//{{{self.NAMESPACES['w']}}}drawing")

                    # 查找object元素（嵌入对象，可能是图片）
                    objects = element.findall(f".//{{{self.NAMESPACES['w']}}}object")

                    # 查找pict元素（VML图形）
                    picts = element.findall(f".//{{{self.NAMESPACES['o']}}}pict")

                    # 收集所有图片关系ID
                    rel_ids = []

                    # 检查drawing中的blip元素（实际图片引用）
                    for drawing in drawings:
                        blips = drawing.findall(f".//{{{self.NAMESPACES['a']}}}blip")
                        for blip in blips:
                            # 获取图片的关系ID
                            rel_id = blip.get(f"{{{self.NAMESPACES['r']}}}embed")
                            if rel_id:
                                rel_ids.append(rel_id)

                    # 检查object中的imagedata元素
                    for obj in objects:
                        imagedata = obj.findall(f".//{{{self.NAMESPACES['v']}}}imagedata")
                        for img in imagedata:
                            rel_id = img.get(f"{{{self.NAMESPACES['r']}}}id")
                            if rel_id:
                                rel_ids.append(rel_id)

                    # 检查pict中的imagedata元素
                    for pict in picts:
                        imagedata = pict.findall(f".//{{{self.NAMESPACES['v']}}}imagedata")
                        for img in imagedata:
                            rel_id = img.get(f"{{{self.NAMESPACES['r']}}}id")
                            if rel_id:
                                rel_ids.append(rel_id)

                    # 如果段落包含图片，添加到结果中
                    if rel_ids:
                        image_paragraphs.append((idx, rel_ids))

            return image_paragraphs
        except Exception as e:
            print(f"获取图片段落索引时出错: {e}")
            return []

    def get_element_index_from_paragraph_index(self, paragraph_index):
        """
        将段落索引转换为self.elements中的索引

        Args:
            paragraph_index (int): 段落索引 (在所有段落中的位置)

        Returns:
            int: 对应的元素索引，如果找不到则返回-1
        """
        try:


            # 遍历所有元素，查找段落
            for idx, item in enumerate(self.elements):
                element = item.get('element')
                if element==self.paragraphs[paragraph_index].get('element'):
                    return idx







        except Exception as e:
            print(f"转换段落索引时出错: {e}")
            return -1

    def get_element_index_from_table_index(self, table_index):
        """
        将表格索引转换为self.elements中的索引

        Args:
            table_index (int): 表格索引 (在所有表格中的位置)

        Returns:
            int: 对应的元素索引，如果找不到则返回-1
        """
        try:
            # 获取所有表格及其对应的元素索引
            table_indices = []

            # 遍历所有元素，查找表格
            for idx, item in enumerate(self.elements):
                element = item.get('element')
                if element.tag.endswith('}tbl'):
                    table_indices.append(idx)

            # 处理负索引
            if table_index < 0:
                table_index = len(table_indices) + table_index

            # 检查索引是否有效
            if table_index < 0 or table_index >= len(table_indices):
                print(f"错误：表格索引 {table_index} 超出有效范围")
                return -1

            # 返回对应的元素索引
            return table_indices[table_index]
        except Exception as e:
            print(f"转换表格索引时出错: {e}")
            return -1

    def get_paragraph_index_from_element_index(self, element_index):
        """
        将元素索引转换为段落索引（如果该元素是段落）

        Args:
            element_index (int): 元素索引

        Returns:
            int: 对应的段落索引，如果元素不是段落或找不到则返回-1
        """
        try:
            # 检查索引是否有效
            if element_index < 0:
                element_index = len(self.elements) + element_index

            if element_index < 0 or element_index >= len(self.elements):
                print(f"错误：元素索引 {element_index} 超出有效范围")
                return -1

            # 获取元素
            element = self.elements[element_index].get('element')

            # 检查是否是段落
            if not element.tag.endswith('}p'):
                print(f"警告：索引 {element_index} 对应的元素不是段落")
                return -1

            # 直接在paragraphs列表中查找匹配的元素
            # 使用enumerate高效获取索引和元素
            for idx, para in enumerate(self.paragraphs):
                if para.get('element') == element:
                    return idx

            # 如果没找到匹配的段落
            print(f"警告：未找到与元素索引 {element_index} 对应的段落")
            return -1
        except Exception as e:
            print(f"转换元素索引时出错: {e}")
            return -1

    def get_table_index_from_element_index(self, element_index):
        """
        将元素索引转换为表格索引（如果该元素是表格）

        Args:
            element_index (int): 元素索引

        Returns:
            int: 对应的表格索引，如果元素不是表格或找不到则返回-1
        """
        try:
            # 检查索引是否有效
            if element_index < 0:
                element_index = len(self.elements) + element_index

            if element_index < 0 or element_index >= len(self.elements):
                print(f"错误：元素索引 {element_index} 超出有效范围")
                return -1

            # 获取元素
            element = self.elements[element_index].get('element')

            # 检查是否是表格
            if not element.tag.endswith('}tbl'):
                print(f"警告：索引 {element_index} 对应的元素不是表格")
                return -1

            # 计算在此之前有多少个表格
            table_count = 0
            for idx in range(element_index):
                if self.elements[idx].get('element').tag.endswith('}tbl'):
                    table_count += 1

            return table_count
        except Exception as e:
            print(f"转换元素索引时出错: {e}")
            return -1

    def get_element_indices_by_type(self, element_type):
        """
        获取指定类型的所有元素的索引

        Args:
            element_type (str): 元素类型，可选值：'paragraph', 'table', 'section', 'all'

        Returns:
            dict: 包含不同类型元素索引的字典，如{"paragraph": [0, 2, ...], "table": [1, 5, ...]}
        """
        try:
            result = {
                "paragraph": [],
                "table": [],
                "section": [],
                "other": []
            }

            # 遍历所有元素
            for idx, item in enumerate(self.elements):
                element = item.get('element')

                # 根据标签判断元素类型
                if element.tag.endswith('}p'):
                    result["paragraph"].append(idx)
                elif element.tag.endswith('}tbl'):
                    result["table"].append(idx)
                elif element.tag.endswith('}sectPr'):
                    result["section"].append(idx)
                else:
                    result["other"].append(idx)

            # 根据请求的类型返回结果
            if element_type == 'all':
                return result
            elif element_type in result:
                return {element_type: result[element_type]}
            else:
                print(f"警告：未知的元素类型 '{element_type}'")
                return {}
        except Exception as e:
            print(f"获取元素索引时出错: {e}")
            return {}

    def get_document_structure(self):
        """
        获取文档结构的详细描述

        Returns:
            list: 包含文档结构信息的列表，每项为字典：
                - type: 元素类型 ('paragraph', 'table', 'section', 等)
                - index: 元素在self.elements中的索引
                - type_index: 元素在其类型中的索引 (如段落索引、表格索引)
                - content_summary: 内容摘要 (对于段落为文本前30个字符，对于表格为尺寸)
                - style: 元素的样式ID (如果有)
        """
        try:
            structure = []
            para_count = 0
            table_count = 0
            section_count = 0
            other_count = 0

            # 遍历所有元素
            for idx, item in enumerate(self.elements):
                element = item.get('element')
                element_info = {
                    "index": idx,
                    "style": None,
                    "content_summary": ""
                }

                # 根据标签判断元素类型
                if element.tag.endswith('}p'):
                    element_info["type"] = "paragraph"
                    element_info["type_index"] = para_count
                    para_count += 1

                    # 获取段落文本摘要
                    text = self.get_paragraph_text(element)
                    element_info["content_summary"] = text[:30] + ("..." if len(text) > 30 else "")

                    # 获取段落样式
                    pPr = element.find(f".//{{{self.NAMESPACES['w']}}}pPr")
                    if pPr is not None:
                        pStyle = pPr.find(f".//{{{self.NAMESPACES['w']}}}pStyle")
                        if pStyle is not None:
                            style_val = pStyle.get(f"{{{self.NAMESPACES['w']}}}val", "")
                            element_info["style"] = style_val

                elif element.tag.endswith('}tbl'):
                    element_info["type"] = "table"
                    element_info["type_index"] = table_count
                    table_count += 1

                    # 获取表格尺寸
                    rows = element.findall(f".//{{{self.NAMESPACES['w']}}}tr")
                    if rows:
                        cols = rows[0].findall(f".//{{{self.NAMESPACES['w']}}}tc")
                        element_info["content_summary"] = f"{len(rows)}行 x {len(cols)}列"

                    # 获取表格样式
                    tblPr = element.find(f".//{{{self.NAMESPACES['w']}}}tblPr")
                    if tblPr is not None:
                        tblStyle = tblPr.find(f".//{{{self.NAMESPACES['w']}}}tblStyle")
                        if tblStyle is not None:
                            style_val = tblStyle.get(f"{{{self.NAMESPACES['w']}}}val", "")
                            element_info["style"] = style_val

                elif element.tag.endswith('}sectPr'):
                    element_info["type"] = "section"
                    element_info["type_index"] = section_count
                    section_count += 1

                    # 获取节属性摘要
                    pgSz = element.find(f".//{{{self.NAMESPACES['w']}}}pgSz")
                    if pgSz is not None:
                        w = pgSz.get(f"{{{self.NAMESPACES['w']}}}w", "")
                        h = pgSz.get(f"{{{self.NAMESPACES['w']}}}h", "")
                        element_info["content_summary"] = f"页面尺寸: {w}x{h}"

                else:
                    element_info["type"] = "other"
                    element_info["type_index"] = other_count
                    other_count += 1
                    element_info["content_summary"] = element.tag.split('}')[-1]

                structure.append(element_info)

            return structure
        except Exception as e:
            print(f"获取文档结构时出错: {e}")
            return []

    def get_image_details(self):
        """
        获取文档中所有图片的详细信息

        Returns:
            list: 包含图片详细信息的列表，每项为字典，包含:
                - paragraph_index: 段落在self.elements中的索引
                - relation_id: 图片关系ID
                - content_type: 图片类型 (如 'image/jpeg', 'image/png')
                - file_name: 原始文件名（如果可用）
                - size: 图片尺寸（字节数）
                - dimensions: 图片尺寸（宽x高，如果可用）
        """
        try:
            image_details = []

            # 获取所有包含图片的段落
            image_paragraphs = self.get_image_paragraphs_indices()

            # 获取文档关系
            rels = {}
            if self.parts['relationships'] is not None:
                root = self.parts['relationships'].getroot()
                for rel in root.findall('.//{*}Relationship'):
                    rel_id = rel.get('Id')
                    rel_type = rel.get('Type')
                    rel_target = rel.get('Target')

                    # 只处理图片关系
                    if rel_type and 'image' in rel_type:
                        rels[rel_id] = {
                            'type': rel_type,
                            'target': rel_target
                        }

            # 处理每个包含图片的段落
            for para_idx, rel_ids in image_paragraphs:
                for rel_id in rel_ids:
                    image_info = {
                        'paragraph_index': para_idx,
                        'relation_id': rel_id,
                        'content_type': 'unknown',
                        'file_name': 'unknown',
                        'size': 0,
                        'dimensions': 'unknown'
                    }

                    # 获取关系信息
                    if rel_id in rels:
                        rel_target = rels[rel_id]['target']

                        # 提取文件名
                        if rel_target:
                            file_name = rel_target.split('/')[-1]
                            image_info['file_name'] = file_name

                        # 获取媒体内容
                        media_path = f"word/{rel_target}"
                        if media_path in self.parts['media']:
                            media_content = self.parts['media'][media_path]

                            # 获取文件大小
                            image_info['size'] = len(media_content)

                            # 判断内容类型
                            if media_content.startswith(b'\xFF\xD8'):
                                image_info['content_type'] = 'image/jpeg'
                            elif media_content.startswith(b'\x89PNG'):
                                image_info['content_type'] = 'image/png'
                            elif media_content.startswith(b'GIF8'):
                                image_info['content_type'] = 'image/gif'
                            elif media_content.startswith(b'BM'):
                                image_info['content_type'] = 'image/bmp'

                            # 尝试获取图片尺寸
                            try:
                                from io import BytesIO
                                from PIL import Image

                                img = Image.open(BytesIO(media_content))
                                image_info['dimensions'] = f"{img.width}x{img.height}"
                            except:
                                # 如果PIL不可用或图片解析失败，忽略尺寸信息
                                pass

                    image_details.append(image_info)

            return image_details
        except Exception as e:
            print(f"获取图片详细信息时出错: {e}")
            return []


    def remove_image_at_paragraph(self, paragraph_index, image_index=None):
        """
        删除指定段落中的图片

        Args:
            paragraph_index (int): 段落索引
            image_index (int, optional): 如果段落中有多个图片，指定要删除的图片索引，None表示删除所有图片

        Returns:
            bool: 操作是否成功
        """
        try:
            # 获取元素列表
            elements = [i.get('element') for i in self.elements]

            # 检查索引是否有效
            if paragraph_index < 0:
                paragraph_index = len(elements) + paragraph_index

            if paragraph_index < 0 or paragraph_index >= len(elements):
                print(f"错误：段落索引 {paragraph_index} 超出有效范围")
                return False

            # 获取段落元素
            paragraph = elements[paragraph_index]

            if not paragraph.tag.endswith('}p'):
                print(f"错误：索引 {paragraph_index} 对应的元素不是段落")
                return False

            # 查找图片元素
            drawing_elements = paragraph.findall(f".//{{{self.NAMESPACES['w']}}}drawing")
            object_elements = paragraph.findall(f".//{{{self.NAMESPACES['w']}}}object")
            pict_elements = paragraph.findall(f".//{{{self.NAMESPACES['v']}}}pict")

            # 合并所有图片元素
            all_image_elements = []

            # 添加drawing元素
            for drawing in drawing_elements:
                # 查找具体的图片引用（blip）
                blips = drawing.findall(f".//{{{self.NAMESPACES['a']}}}blip")
                if blips:
                    all_image_elements.append(drawing)

            # 添加object元素
            for obj in object_elements:
                # 查找具体的图片数据
                imagedata = obj.findall(f".//{{{self.NAMESPACES['v']}}}imagedata")
                if imagedata:
                    all_image_elements.append(obj)

            # 添加pict元素
            for pict in pict_elements:
                # 查找具体的图片数据
                imagedata = pict.findall(f".//{{{self.NAMESPACES['v']}}}imagedata")
                if imagedata:
                    all_image_elements.append(pict)

            # 检查是否找到图片
            if not all_image_elements:
                print(f"警告：段落 {paragraph_index} 中未找到图片")
                return False

            # 确定要删除的图片
            if image_index is not None:
                # 删除特定索引的图片
                if image_index < 0:
                    image_index = len(all_image_elements) + image_index

                if image_index < 0 or image_index >= len(all_image_elements):
                    print(f"错误：图片索引 {image_index} 超出有效范围")
                    return False

                elements_to_remove = [all_image_elements[image_index]]
            else:
                # 删除所有图片
                elements_to_remove = all_image_elements

            # 从段落中移除图片元素
            for img_elem in elements_to_remove:
                # 查找图片元素的父元素
                parent = None
                for p in paragraph.iter():
                    if img_elem in list(p):
                        parent = p
                        break

                if parent is not None:
                    # 如果图片在run中，仅移除图片元素而不是整个run
                    if parent.tag.endswith('}r'):
                        parent.remove(img_elem)

                        # 如果run现在为空，且不包含文本，可以移除这个run
                        if len(parent) == 0 or (len(parent) == 1 and parent[0].tag.endswith('}rPr')):
                            parent_of_run = None
                            for p in paragraph.iter():
                                if parent in list(p):
                                    parent_of_run = p
                                    break

                            if parent_of_run is not None:
                                parent_of_run.remove(parent)
                    else:
                        # 其他情况直接移除图片元素
                        parent.remove(img_elem)
                else:
                    print(f"警告：找不到图片元素的父元素")

            # 更新XML树
            self.update_document_xml()

            return True
        except Exception as e:
            print(f"删除图片时出错: {e}")
            return False

    def extract_comments(self):
        """
        提取文档中的所有批注及其相关信息

        Returns:
            list: 包含所有批注信息的列表，每个批注表示为字典，包含ID、作者、日期、内容和引用文本等信息
        """
        # 初始化结果列表
        comments = []

        # 首先尝试从self.parts获取comments
        if hasattr(self, 'parts') and 'comments' in self.parts and self.parts['comments'] is not None:
            comments_root = self.parts['comments'].getroot()
        # 如果失败，尝试从self.docx_parts获取word/comments.xml
        elif hasattr(self, 'docx_parts') and 'word/comments.xml' in self.docx_parts:
            # 获取comments.xml内容
            comments_content = self.docx_parts['word/comments.xml']

            # 检查内容是否为None
            if comments_content is None:
                print("批注内容为空")
                return comments

            # 如果保存的是字节内容，需要解析为XML
            if isinstance(comments_content, bytes):
                try:
                    import xml.etree.ElementTree as ET
                    from io import BytesIO
                    comments_root = ET.parse(BytesIO(comments_content)).getroot()
                except Exception as e:
                    print(f"解析批注XML时出错: {e}")
                    return comments
            else:
                # 如果已经是ElementTree，直接获取根元素
                try:
                    comments_root = comments_content.getroot()
                except AttributeError:
                    print("批注内容格式不正确")
                    return comments
        else:
            print("文档中没有批注内容")
            return comments

        # 获取所有批注元素
        comment_elements = comments_root.findall('.//{%s}comment' % self.NAMESPACES['w'])

        # 如果没有找到批注元素
        if not comment_elements:
            print("文档中没有批注元素")
            return comments

        # 创建批注ID到批注内容的映射
        comments_map = {}

        # 提取每个批注的信息
        for comment_elem in comment_elements:
            comment_id = comment_elem.get('{%s}id' % self.NAMESPACES['w'])
            author = comment_elem.get('{%s}author' % self.NAMESPACES['w'], '未知作者')
            date = comment_elem.get('{%s}date' % self.NAMESPACES['w'], '未知日期')

            # 提取批注文本内容
            comment_text = ""
            for paragraph in comment_elem.findall('.//{%s}p' % self.NAMESPACES['w']):
                for run in paragraph.findall('.//{%s}r' % self.NAMESPACES['w']):
                    for text in run.findall('.//{%s}t' % self.NAMESPACES['w']):
                        comment_text += text.text if text.text else ""

            # 创建批注信息字典
            comment_info = {
                'id': comment_id,
                'author': author,
                'date': date,
                'text': comment_text,
                'referenced_text': '',  # 将在后续步骤中填充
                'paragraph_index': -1,  # 将在后续步骤中填充
                'element_index': -1  # 将在后续步骤中填充
            }

            # 将批注添加到映射中
            comments_map[comment_id] = comment_info

        # 查找文档中的批注引用及其位置
        self._find_comment_references(comments_map)

        # 将批注映射转换为列表
        return list(comments_map.values())

    def _find_comment_references(self, comments_map):
        """
        在文档中查找批注引用，并将引用文本和位置信息添加到批注信息中

        Args:
            comments_map (dict): 批注ID到批注信息的映射
        """
        # 遍历文档中的所有段落
        for elem_index, elem_info in enumerate(self.elements):

            element = elem_info['element']

            # 检查是否为段落
            if element.tag.endswith('}p'):
                # 获取段落索引
                para_indices = [i for i, p in enumerate(self.paragraphs) if p == element]
                para_index = para_indices[0] if para_indices else -1

                # 查找批注开始和结束标记
                comment_starts = element.findall('.//{%s}commentRangeStart' % self.NAMESPACES['w'])
                comment_ends = element.findall('.//{%s}commentRangeEnd' % self.NAMESPACES['w'])
                comment_refs = element.findall('.//{%s}commentReference' % self.NAMESPACES['w'])

                # 处理批注范围开始和结束
                comment_ranges = {}

                # 记录批注范围开始
                for start in comment_starts:
                    comment_id = start.get('{%s}id' % self.NAMESPACES['w'])
                    if comment_id in comments_map:
                        if comment_id not in comment_ranges:
                            comment_ranges[comment_id] = {'start': start, 'end': None, 'text': ""}

                # 记录批注范围结束
                for end in comment_ends:
                    comment_id = end.get('{%s}id' % self.NAMESPACES['w'])
                    if comment_id in comments_map and comment_id in comment_ranges:
                        comment_ranges[comment_id]['end'] = end

                # 获取段落文本
                para_text = self.get_paragraph_text(element)

                # 处理批注引用
                for ref in comment_refs:
                    comment_id = ref.get('{%s}id' % self.NAMESPACES['w'])
                    if comment_id in comments_map:
                        # 更新批注位置信息
                        comments_map[comment_id]['paragraph_index'] = para_index
                        comments_map[comment_id]['element_index'] = elem_index

                        # 如果我们有范围信息，尝试提取引用的文本
                        if comment_id in comment_ranges:
                            # 简单处理：将整个段落作为引用文本
                            # 注意：实际提取准确的范围文本需要更复杂的逻辑
                            comments_map[comment_id]['referenced_text'] = para_text

    def get_comment_at_paragraph(self, para_index):
        """
        获取特定段落中的所有批注

        Args:
            para_index (int): 段落索引

        Returns:
            list: 该段落中的所有批注
        """
        all_comments = self.extract_comments()
        return [comment for comment in all_comments if comment['paragraph_index'] == para_index]

    def get_comment_by_id(self, comment_id):
        """
        通过ID获取特定批注

        Args:
            comment_id (str): 批注ID

        Returns:
            dict: 批注信息，如果未找到则返回None
        """
        all_comments = self.extract_comments()
        for comment in all_comments:
            if comment['id'] == comment_id:
                return comment
        return None

    def get_table_dimensions(self, table_index):
        try:
            if table_index >= len(self.tables):
                return None

            table = self.tables[table_index]['element']

            # 获取所有行，使用直接子元素查询
            rows = table.findall('./w:tr', self.NAMESPACES)
            if not rows:
                return (0, 0)

            row_count = len(rows)

            # 使用tblGrid获取列数（更可靠）
            tblGrid = table.find('./w:tblGrid', self.NAMESPACES)
            if tblGrid is not None:
                gridCols = tblGrid.findall('./w:gridCol', self.NAMESPACES)
                if gridCols:
                    print(f'gridCols count: {len(gridCols)}')
                    return (row_count, len(gridCols))

            # 如果没有tblGrid信息，则从第一行单元格计算
            first_row = rows[0]
            cells = first_row.findall('./w:tc', self.NAMESPACES)  # 直接子元素
            col_count = len(cells)
            print(f'首行单元格数: {col_count}')

            return (row_count, col_count)
        except Exception as e:
            print(f"获取表格尺寸时出错: {e}")
            return None

    def get_all_tables_dimensions(self):
        """
        获取文档中所有表格的尺寸信息

        返回:
            list: 包含所有表格尺寸的列表，格式为 [(表格索引, 行数, 列数), ...]
        """
        try:
            tables = self.get_all_tables()
            if not tables:
                return []

            dimensions = []
            for i in range(len(tables)):
                dim = self.get_table_dimensions(i)
                if dim:
                    dimensions.append((i, dim[0], dim[1]))

            return dimensions

        except Exception as e:
            print(f"获取所有表格尺寸时出错: {e}")
            return []



    def get_paragraph_style_from_element(self, paragraph_element):
        """
        获取段落元素的样式信息

        参数:
            paragraph_element (Element): w:p XML元素

        返回:
            dict: 包含段落样式的字典
        """
        try:
            if paragraph_element is None or not paragraph_element.tag.endswith('}p'):
                print("输入不是有效的段落元素")
                return {}

            style_info = {}

            # 获取段落属性元素
            pPr = paragraph_element.find(f".//{{{self.NAMESPACES['w']}}}pPr")
            if pPr is None:
                return {'message': '段落没有显式样式属性'}

            # 获取样式ID
            pStyle = pPr.find(f".//{{{self.NAMESPACES['w']}}}pStyle")
            if pStyle is not None:
                style_id = pStyle.get(f"{{{self.NAMESPACES['w']}}}val", "")
                style_info['style_id'] = style_id

            # 获取对齐方式
            jc = pPr.find(f".//{{{self.NAMESPACES['w']}}}jc")
            if jc is not None:
                alignment = jc.get(f"{{{self.NAMESPACES['w']}}}val", "")
                style_info['alignment'] = alignment

            # 获取缩进
            ind = pPr.find(f".//{{{self.NAMESPACES['w']}}}ind")
            if ind is not None:
                indentation = {}
                for attr in ['left', 'right', 'firstLine', 'hanging']:
                    val = ind.get(f"{{{self.NAMESPACES['w']}}}{attr}", None)
                    if val is not None:
                        indentation[attr] = val
                if indentation:
                    style_info['indentation'] = indentation

            # 获取段落间距
            spacing = pPr.find(f".//{{{self.NAMESPACES['w']}}}spacing")
            if spacing is not None:
                spacing_info = {}
                for attr in ['before', 'after', 'line', 'lineRule','beforeLines', 'afterLines']:
                    val = spacing.get(f"{{{self.NAMESPACES['w']}}}{attr}", None)
                    if val is not None:
                        spacing_info[attr] = val
                if spacing_info:
                    style_info['spacing'] = spacing_info

            # 获取边框
            pBdr = pPr.find(f".//{{{self.NAMESPACES['w']}}}pBdr")
            if pBdr is not None:
                borders = {}
                for border_type in ['top', 'left', 'bottom', 'right']:
                    border = pBdr.find(f".//{{{self.NAMESPACES['w']}}}{border_type}")
                    if border is not None:
                        border_info = {}
                        for attr in ['val', 'sz', 'color', 'space']:
                            val = border.get(f"{{{self.NAMESPACES['w']}}}{attr}", None)
                            if val is not None:
                                border_info[attr] = val
                        if border_info:
                            borders[border_type] = border_info
                if borders:
                    style_info['borders'] = borders

            # 获取底纹
            shd = pPr.find(f".//{{{self.NAMESPACES['w']}}}shd")
            if shd is not None:
                shading = {}
                for attr in ['val', 'color', 'fill']:
                    val = shd.get(f"{{{self.NAMESPACES['w']}}}{attr}", None)
                    if val is not None:
                        shading[attr] = val
                if shading:
                    style_info['shading'] = shading

            # 获取段落级别字体设置
            rPr = pPr.find(f".//{{{self.NAMESPACES['w']}}}rPr")
            if rPr is not None:
                run_props = self._extract_run_properties_from_element(rPr)
                if run_props:
                    style_info['run_properties'] = run_props

            return style_info

        except Exception as e:
            print(f"获取段落样式时出错: {e}")
            return {'error': str(e)}

    def get_runs_from_paragraph(self, paragraph_element):
        """
        从段落元素中获取所有run元素

        参数:
            paragraph_element (Element): w:p XML元素

        返回:
            list: run元素列表
        """
        try:
            if paragraph_element is None or not paragraph_element.tag.endswith('}p'):
                print("输入不是有效的段落元素")
                return []

            # 获取所有run元素
            runs = paragraph_element.findall(f".//{{{self.NAMESPACES['w']}}}r")
            return runs

        except Exception as e:
            print(f"获取run元素时出错: {e}")
            return []

    def get_run_style_from_element(self, run_element):
        """
        获取run元素的样式信息

        参数:
            run_element (Element): w:r XML元素

        返回:
            dict: 包含run样式的字典
        """
        try:
            if run_element is None or not run_element.tag.endswith('}r'):
                print("输入不是有效的run元素")
                return {}

            # 获取run属性元素
            rPr = run_element.find(f".//{{{self.NAMESPACES['w']}}}rPr")
            if rPr is None:
                return {'message': 'run没有显式样式属性'}

            return self._extract_run_properties_from_element(rPr)

        except Exception as e:
            print(f"获取run样式时出错: {e}")
            return {'error': str(e)}

    def _extract_run_properties_from_element(self, rPr):
        """
        从rPr元素中提取run属性

        参数:
            rPr (Element): w:rPr XML元素

        返回:
            dict: 包含run属性的字典
        """
        run_props = {}

        # 获取字体
        rFonts = rPr.find(f".//{{{self.NAMESPACES['w']}}}rFonts")
        if rFonts is not None:
            fonts = {}
            for attr in ['ascii', 'hAnsi', 'eastAsia', 'cs']:
                val = rFonts.get(f"{{{self.NAMESPACES['w']}}}{attr}", None)
                if val is not None:
                    fonts[attr] = val
            if fonts:
                run_props['fonts'] = fonts

        # 获取字号
        sz = rPr.find(f".//{{{self.NAMESPACES['w']}}}sz")
        if sz is not None:
            size_val = sz.get(f"{{{self.NAMESPACES['w']}}}val", None)
            if size_val is not None:
                run_props['size'] = size_val

        # 获取字体颜色
        color = rPr.find(f".//{{{self.NAMESPACES['w']}}}color")
        if color is not None:
            color_val = color.get(f"{{{self.NAMESPACES['w']}}}val", None)
            if color_val is not None:
                run_props['color'] = color_val

        # 获取粗体
        b = rPr.find(f".//{{{self.NAMESPACES['w']}}}b")
        if b is not None:
            val = b.get(f"{{{self.NAMESPACES['w']}}}val", "true")
            run_props['bold'] = val != "false"

        # 获取斜体
        i = rPr.find(f".//{{{self.NAMESPACES['w']}}}i")
        if i is not None:
            val = i.get(f"{{{self.NAMESPACES['w']}}}val", "true")
            run_props['italic'] = val != "false"

        # 获取下划线
        u = rPr.find(f".//{{{self.NAMESPACES['w']}}}u")
        if u is not None:
            val = u.get(f"{{{self.NAMESPACES['w']}}}val", None)
            if val is not None:
                run_props['underline'] = val

        # 获取删除线
        strike = rPr.find(f".//{{{self.NAMESPACES['w']}}}strike")
        if strike is not None:
            val = strike.get(f"{{{self.NAMESPACES['w']}}}val", "true")
            run_props['strike'] = val != "false"

        # 获取突出显示
        highlight = rPr.find(f".//{{{self.NAMESPACES['w']}}}highlight")
        if highlight is not None:
            val = highlight.get(f"{{{self.NAMESPACES['w']}}}val", None)
            if val is not None:
                run_props['highlight'] = val

        # 获取大小写
        caps = rPr.find(f".//{{{self.NAMESPACES['w']}}}caps")
        if caps is not None:
            val = caps.get(f"{{{self.NAMESPACES['w']}}}val", "true")
            run_props['caps'] = val != "false"

        # 获取小型大写字母
        smallCaps = rPr.find(f".//{{{self.NAMESPACES['w']}}}smallCaps")
        if smallCaps is not None:
            val = smallCaps.get(f"{{{self.NAMESPACES['w']}}}val", "true")
            run_props['smallCaps'] = val != "false"

        # 获取垂直对齐
        vertAlign = rPr.find(f".//{{{self.NAMESPACES['w']}}}vertAlign")
        if vertAlign is not None:
            val = vertAlign.get(f"{{{self.NAMESPACES['w']}}}val", None)
            if val is not None:
                run_props['vertAlign'] = val

        return run_props

    def get_table_cell_style(self, table_index, row_idx, col_idx):
        """
        获取表格中特定单元格内段落的元素

        参数:
            table_index (int): 表格索引
            row_idx (int): 行索引
            col_idx (int): 列索引

        返回:
            Element: 单元格中的段落元素
            None: 如果单元格不存在或没有段落元素
        """
        try:
            # 获取所有表格
            if table_index >= len(self.tables):
                return None

            table = self.tables[table_index]['element']

            # 获取行元素
            rows = table.findall('.//w:tr', self.NAMESPACES)
            if not rows or row_idx >= len(rows):
                return None

            row = rows[row_idx]

            # 获取单元格元素
            cells = row.findall('.//w:tc', self.NAMESPACES)
            if not cells or col_idx >= len(cells):
                return None

            cell = cells[col_idx]

            # 获取单元格中的段落元素
            paragraphs = cell.findall('.//w:p', self.NAMESPACES)
            if not paragraphs:
                return None

            # 返回第一个段落元素
            return paragraphs[0]

        except Exception as e:
            print(f"获取表格单元格样式时出错: {e}")
            return None

    def get_table_cell_paragraphs(self, table_index, row_idx, col_idx):
        try:
            print(f"获取表格 {table_index} 的第 {row_idx} 行 第 {col_idx} 列的段落")
            # 获取所有表格
            if table_index >= len(self.tables):
                print(f"表格索引 {table_index} 超出范围，总表格数: {len(self.tables)}")
                return None

            table = self.tables[table_index]['element']

            # 获取行元素
            rows = table.findall('.//w:tr', self.NAMESPACES)
            print(f"找到 {len(rows)} 行")
            if not rows or row_idx >= len(rows):
                print(f"行索引 {row_idx} 超出范围，总行数: {len(rows)}")
                return None

            row = rows[row_idx]

            # 获取单元格元素
            cells = row.findall('.//w:tc', self.NAMESPACES)
            print(f"第 {row_idx} 行找到 {len(cells)} 个单元格")
            if not cells or col_idx >= len(cells):
                print(f"列索引 {col_idx} 超出范围，总列数: {len(cells)}")
                return None

            cell = cells[col_idx]

            # 获取单元格中的段落元素
            paragraphs = cell.findall('.//w:p', self.NAMESPACES)
            print(f"找到 {len(paragraphs)} 个段落")
            return paragraphs
        except Exception as e:
            print(f"获取表格单元格段落时出错: {e}")
            return None
    def get_table_cell_text(self, table_index, row_idx, col_idx):
        """
        获取表格中特定单元格的文本内容

        参数:
            table_index (int): 表格索引
            row_idx (int): 行索引
            col_idx (int): 列索引

        返回:
            str: 单元格的文本内容
            None: 如果单元格不存在
        """
        paragraphs = self.get_table_cell_paragraphs(table_index, row_idx, col_idx)
        if not paragraphs:
            return None

        text = ""
        for p in paragraphs:
            runs = p.findall('.//w:t', self.NAMESPACES)
            for run in runs:
                if run.text:
                    text += run.text

            # 在段落之间添加换行符
            if p != paragraphs[-1]:
                text += "\n"

        return text

    def add_comment(self, element_index, author="", comment_text="", element_type="paragraph", run_index=None):
        """
        为段落、运行(run)或表格添加批注

        参数:
            element_index (int): 要添加批注的元素索引
            author (str): 批注作者名称，默认为空
            comment_text (str): 批注内容文本
            element_type (str): 元素类型，可选值: "paragraph", "run", "table"
            run_index (int): 如果element_type为"run"，则指定run的索引，默认为None

        返回:
            str: 新创建的批注ID
        """
        import datetime
        import uuid



        # 获取XML命名空间
        ns_w = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
        ns_w14 = "{http://schemas.microsoft.com/office/word/2010/wordml}"
        ns_w15 = "{http://schemas.microsoft.com/office/word/2012/wordml}"
        ns_rel = "{http://schemas.openxmlformats.org/package/2006/relationships}"









        # 获取或创建comments.xml的etree
        if self.parts['comments'] is  None:


           self.create_comments_file()


        # 获取或创建commentsExtended.xml的etree

        if self.parts['commentsExtended'] is  None:


                self.create_comments_extended_file()

        # 获取或创建people.xml的etree
        if self.parts['people'] is  None:


            self.create_people_file()
        people_root = self.docx_parts['people']
        comments_root = self.parts['comments']
        comments_extended_root = self.parts['commentsExtended']

        # 3. 为目标元素创建批注范围标记
        # 计算新的批注ID
        comment_id = "0"
        comments = comments_root.getroot().findall(f".//{ns_w}comment")
        if comments:
            # 找到最大ID并加1
            max_id = max([int(comment.get(f"{ns_w}id", "0")) for comment in comments])
            comment_id = str(max_id + 1)

        # 3.1 创建commentRangeStart和commentRangeEnd元素
        comment_start = ET.Element(f"{ns_w}commentRangeStart")
        comment_start.set(f"{ns_w}id", comment_id)

        comment_end = ET.Element(f"{ns_w}commentRangeEnd")
        comment_end.set(f"{ns_w}id", comment_id)

        comment_ref = ET.Element(f"{ns_w}commentReference")
        comment_ref.set(f"{ns_w}id", comment_id)

        # 根据元素类型添加批注范围
        # 根据元素类型添加批注范围
        if element_type == "paragraph":
            target_element = self.elements[element_index]['element']
            # 获取段落中的所有run元素
            runs = self.get_runs_from_paragraph(target_element)

            if runs:
                # 进行操作前获取所有子元素的副本
                children = list(target_element)

                # 找到第一个和最后一个run在children中的位置
                first_run_pos = -1
                last_run_pos = -1

                for i, child in enumerate(children):
                    if child is runs[0]:  # 使用is而不是==来比较对象身份
                        first_run_pos = i
                    if child is runs[-1]:
                        last_run_pos = i

                # 插入批注标记
                if first_run_pos != -1 and last_run_pos != -1:
                    # 在第一个run前插入commentRangeStart
                    children.insert(first_run_pos, comment_start)
                    # 因为插入了元素，后面元素的位置向后移动
                    last_run_pos += 1
                    # 在最后一个run后插入commentRangeEnd
                    children.insert(last_run_pos + 1, comment_end)
                    comment_ref_run = ET.Element(f"{ns_w}r")
                    comment_ref_run.append(comment_ref)


                    # 添加包含commentReference的r元素
                    children.append(comment_ref_run)

                    # 重建段落内容
                    for child in list(target_element):
                        target_element.remove(child)

                    for child in children:
                        target_element.append(child)

        elif element_type == "table":
            target_element = self.tables[element_index]['element']
            # 对于表格，我们需要找到表格的第一个单元格和最后一个单元格
            first_cell = target_element.find(f".//{ns_w}tc")
            last_cell = list(target_element.findall(f".//{ns_w}tc"))[-1]

            if first_cell is not None and last_cell is not None:
                # 在第一个单元格的第一个段落前添加commentRangeStart
                first_para = first_cell.find(f".//{ns_w}p")
                if first_para is not None:
                    first_para_runs = list(first_para.findall(f".//{ns_w}r"))
                    if first_para_runs:
                        # 获取段落的所有子元素
                        children = list(first_para)
                        # 找到第一个run的位置
                        first_run_pos = -1
                        for i, child in enumerate(children):
                            if child is first_para_runs[0]:
                                first_run_pos = i
                                break

                        if first_run_pos != -1:
                            # 在第一个run前插入commentRangeStart
                            children.insert(first_run_pos, comment_start)

                            # 重建段落内容
                            for child in list(first_para):
                                first_para.remove(child)

                            for child in children:
                                first_para.append(child)
                        else:
                            # 如果找不到run的位置，就插入到段落开始
                            first_para.insert(0, comment_start)
                    else:
                        # 如果没有run元素，直接添加到段落开始
                        first_para.insert(0, comment_start)

                # 在最后一个单元格的最后一个段落后添加commentRangeEnd和commentReference
                last_para = list(last_cell.findall(f".//{ns_w}p"))[-1]
                if last_para is not None:
                    # 直接使用append()添加到段落末尾
                    last_para.append(comment_end)

                    # 创建包含commentReference的r元素
                    comment_ref_run = ET.Element(f"{ns_w}r")
                    comment_ref_run.append(comment_ref)

                    # 添加包含commentReference的r元素
                    last_para.append(comment_ref_run)
        elif element_type == "run":
            # 获取要批注的run元素
            target_run = self._get_run_element(element_index, run_index)
            parent_paragraph = self.elements[element_index]['element']

            # 找到run在段落中的位置
            children = list(parent_paragraph)
            run_pos = -1

            # 首先尝试在段落的直接子元素中查找
            for i, child in enumerate(children):
                if child is target_run:
                    run_pos = i
                    break

            # 如果在直接子元素中找不到，查找所有run
            if run_pos == -1:
                runs_in_paragraph = list(parent_paragraph.findall(f".//{ns_w}r"))
                for i, run in enumerate(runs_in_paragraph):
                    if run is target_run:
                        # 找到了目标run，但我们需要找到它在父元素中的位置
                        # 获取该run的父元素
                        parent_of_run = None
                        for elem in parent_paragraph.iter():
                            if target_run in list(elem):
                                parent_of_run = elem
                                break

                        if parent_of_run:
                            # 在父元素中找到run的位置
                            children_of_parent = list(parent_of_run)
                            for j, child in enumerate(children_of_parent):
                                if child is target_run:
                                    # 插入批注标记
                                    # 插入批注标记
                                    children_of_parent.insert(j, comment_start)
                                    children_of_parent.insert(j + 2, comment_end)

                                    # 创建包含commentReference的r元素
                                    comment_ref_run = ET.Element(f"{ns_w}r")
                                    comment_ref_run.append(comment_ref)

                                    # 添加包含commentReference的r元素
                                    children_of_parent.insert(j + 3, comment_ref_run)

                                    # 重建父元素内容
                                    for child in list(parent_of_run):
                                        parent_of_run.remove(child)

                                    for child in children_of_parent:
                                        parent_of_run.append(child)

                                    run_pos = j  # 更新位置
                                    break
                            break

            if run_pos == -1:
                raise ValueError("无法确定run元素在段落中的位置")
            else:
                # 如果run是段落的直接子元素
                if run_pos >= 0:
                    # 插入批注标记
                    children.insert(run_pos, comment_start)
                    run_pos += 1  # 因为我们插入了一个元素，run的位置向后移动
                    children.insert(run_pos + 1, comment_end)

                    # 创建包含commentReference的r元素
                    comment_ref_run = ET.Element(f"{ns_w}r")
                    comment_ref_run.append(comment_ref)

                    # 添加包含commentReference的r元素
                    children.insert(run_pos + 2, comment_ref_run)

                    # 重建段落内容
                    for child in list(parent_paragraph):
                        parent_paragraph.remove(child)

                    for child in children:
                        parent_paragraph.append(child)




        # 4. 创建批注内容
        comment = ET.Element(f"{ns_w}comment")
        comment.set(f"{ns_w}id", comment_id)

        if author:
            comment.set(f"{ns_w}author", author)
        else:
            comment.set(f"{ns_w}author", "Anonymous")

        # 设置批注日期为当前时间
        current_time = datetime.datetime.now().strftime("%Y-%m-%dT%H:%M:%SZ")
        comment.set(f"{ns_w}date", current_time)

        # 初始化批注为作者的首字母
        initials = author[0] if author and author else "A"
        comment.set(f"{ns_w}initials", initials)

        # 创建批注内容段落
        para_id = str(uuid.uuid4())[:8].upper()
        comment_para = ET.SubElement(comment, f"{ns_w}p")
        comment_para.set(f"{ns_w14}paraId", para_id)

        # 添加段落属性
        pPr = ET.SubElement(comment_para, f"{ns_w}pPr")
        pStyle = ET.SubElement(pPr, f"{ns_w}pStyle")
        pStyle.set(f"{ns_w}val", "CommentText")

        # 添加运行(run)元素和文本
        r = ET.SubElement(comment_para, f"{ns_w}r")
        rPr = ET.SubElement(r, f"{ns_w}rPr")
        rStyle = ET.SubElement(rPr, f"{ns_w}rStyle")
        rStyle.set(f"{ns_w}val", "CommentReference")

        # 添加批注文本
        r = ET.SubElement(comment_para, f"{ns_w}r")
        t = ET.SubElement(r, f"{ns_w}t")
        t.text = comment_text

        # 5. 将批注添加到comments.xml
        comments_root.getroot().append(comment)

        # 6. 添加commentsExtended条目
        comment_ex = ET.SubElement(comments_extended_root.getroot(), f"{ns_w15}commentEx")
        comment_ex.set(f"{ns_w15}paraId", para_id)
        comment_ex.set(f"{ns_w15}done", "0")

        # 7. 添加或更新people条目
        person_found = False
        for person in people_root.getroot().findall(f".//{ns_w15}person"):
            if person.get(f"{ns_w15}author") == author:
                person_found = True
                break

        if not person_found and author:
            person = ET.SubElement(people_root.getroot(), f"{ns_w15}person")
            person.set(f"{ns_w15}author", author)
            presence_info = ET.SubElement(person, f"{ns_w15}presenceInfo")
            presence_info.set(f"{ns_w15}providerId", "None")
            presence_info.set(f"{ns_w15}userId", author)

        # 8. 更新document.xml.rels以确保relationship存在


        rels_root = self.docx_parts['relationships']

        # 检查并添加comments.xml关系
        comments_rel_exists = False
        comments_rel_id = None
        for rel in rels_root.findall(f".//{ns_rel}Relationship"):
            if rel.get("Target") == "comments.xml":
                comments_rel_exists = True
                comments_rel_id = rel.get("Id")
                break

        if not comments_rel_exists:
            new_rel = ET.SubElement(rels_root, f"{ns_rel}Relationship")
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

            comments_rel_id = f"rId{max_rid + 1}"
            new_rel.set("Id", comments_rel_id)
            new_rel.set("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments")
            new_rel.set("Target", "comments.xml")

        # 检查并添加commentsExtended.xml关系
        comments_ex_rel_exists = False
        for rel in rels_root.findall(f".//{ns_rel}Relationship"):
            if rel.get("Target") == "commentsExtended.xml":
                comments_ex_rel_exists = True
                break

        if not comments_ex_rel_exists:
            new_rel = ET.SubElement(rels_root, f"{ns_rel}Relationship")
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

            new_rel.set("Id", f"rId{max_rid + 1}")
            new_rel.set("Type", "http://schemas.microsoft.com/office/2011/relationships/commentsExtended")
            new_rel.set("Target", "commentsExtended.xml")

        # 检查并添加people.xml关系
        people_rel_exists = False
        for rel in rels_root.findall(f".//{ns_rel}Relationship"):
            if rel.get("Target") == "people.xml":
                people_rel_exists = True
                break

        if not people_rel_exists:
            new_rel = ET.SubElement(rels_root, f"{ns_rel}Relationship")
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

            new_rel.set("Id", f"rId{max_rid + 1}")
            new_rel.set("Type", "http://schemas.microsoft.com/office/2011/relationships/people")
            new_rel.set("Target", "people.xml")

        # 9. 更新docx_parts中的XML树
        self.docx_parts['comments'] = comments_root
        self.docx_parts['commentsExtended'] = comments_extended_root
        self.docx_parts['people'] = people_root

        self.update_document_xml()

        return comment_id

    def insert_page_break_before_paragraph(self, para_index):
        """在指定段落前插入换页符

        Args:
            para_index (int): 段落索引，换页符将在此段落前插入

        Returns:
            bool: 操作是否成功
        """
        try:
            # 检查段落索引是否有效
            if para_index < 0 or para_index >= len(self.paragraphs):
                print(f"错误：段落索引 {para_index} 超出范围 (0-{len(self.paragraphs) - 1})")
                return False

            # 获取段落元素
            paragraph = self.paragraphs[para_index]['element']

            # 获取或创建段落属性元素 (pPr)
            pPr = self._get_or_create_pPr(paragraph)

            # 查找现有的 pageBreakBefore 元素
            page_break = pPr.find(f".//{{{self.NAMESPACES['w']}}}pageBreakBefore")

            # 如果已存在，则更新，否则创建新的
            if page_break is None:
                page_break = ET.SubElement(pPr, f"{{{self.NAMESPACES['w']}}}pageBreakBefore")

            # 设置值为 "1" 表示启用换页符
            page_break.set(f"{{{self.NAMESPACES['w']}}}val", "1")

            # 更新文档 XML
            self.update_document_xml()

            print(f"在段落 {para_index} 前成功插入换页符")
            return True

        except Exception as e:
            print(f"在段落前插入换页符时出错: {e}")
            import traceback
            traceback.print_exc()
            return False

    def insert_page_break(self, para_index, position="before"):
        """在指定段落前或后插入换页符

        这个方法支持两种类型的换页符：
        1. 段落前分页：使用 pageBreakBefore 属性，让段落在新页开始（position="before"）
        2. 手动分页：在段落中插入显式的分页符（position="after"）

        Args:
            para_index (int): 段落索引
            position (str): 'before' 在段落前插入分页，'after' 在段落后插入分页

        Returns:
            bool: 操作是否成功
        """
        if position.lower() == "before":
            # 使用段落的 pageBreakBefore 属性
            return self.insert_page_break_before_paragraph(para_index)
        else:
            try:
                # 检查段落索引是否有效
                if para_index < 0 or para_index >= len(self.paragraphs):
                    print(f"错误：段落索引 {para_index} 超出范围 (0-{len(self.paragraphs) - 1})")
                    return False

                # 创建一个新的段落元素来包含显式分页符
                new_para = ET.Element(f"{{{self.NAMESPACES['w']}}}p")

                # 创建一个 run 元素
                run = ET.SubElement(new_para, f"{{{self.NAMESPACES['w']}}}r")

                # 添加分页符 <w:br w:type="page"/>
                br = ET.SubElement(run, f"{{{self.NAMESPACES['w']}}}br")
                br.set(f"{{{self.NAMESPACES['w']}}}type", "page")

                # 获取文档体
                body = self.root.find(f".//{{{self.NAMESPACES['w']}}}body")
                if body is None:
                    print("错误：无法找到文档体(body)元素")
                    return False

                # 获取目标段落在body中的位置
                target_element = self.paragraphs[para_index]['element']
                body_children = list(body)
                target_index = -1
                for i, child in enumerate(body_children):
                    if child == target_element:
                        target_index = i
                        break

                if target_index == -1:
                    print("错误：无法在文档树中定位目标段落")
                    return False

                # 在段落后插入新的分页段落
                body.insert(target_index + 1, new_para)

                # 重新解析文档结构
                self.get_structured_body_elements()

                # 更新文档 XML
                self.update_document_xml()

                print(f"在段落 {para_index} 后成功插入换页符")
                return True

            except Exception as e:
                print(f"在段落后插入换页符时出错: {e}")
                import traceback
                traceback.print_exc()
                return False

    def get_image_from_pra(self, element):
        """从段落元素中提取图片及相关信息。

        此函数分析段落元素中的图片(drawing或pict元素)，提取图片的嵌入ID、尺寸、描述信息，
        以及确定图片与文本的相对位置关系。

        Args:
            element: XML段落元素，通常是w:p元素

        Returns:
            dict: 包含图片详细信息的字典，包括:
                - drawings_count: drawing元素数量
                - picts_count: pict元素数量
                - has_text: 是否包含文本
                - text_content: 段落中的文本内容(截断至100字符)
                - standalone_image: 图片是否独立存在(无文本)
                - text_before_image: 文本是否在图片前
                - text_after_image: 文本是否在图片后
                - text_surrounds_image: 图片是否被文本包围
                - image_position: 图片位置描述('standalone','surrounded_by_text',
                                  'after_text','before_text'或'unknown')
                - image_descriptions: 图片描述信息列表，包含name和description
                - embed_ids: 图片嵌入关系ID列表
                - dimensions: 图片尺寸信息列表，包含原始EMU单位和转换为厘米的尺寸
                - embed_types: 图片嵌入类型列表('inline', 'anchor'等)
                - wrap_types: 图片的文本环绕方式列表
        """
        # 检查图片与文本的相对位置
        text_elements = element.findall(f".//{{{self.NAMESPACES['w']}}}t")
        text_content = "".join([t.text for t in text_elements if t.text]) if text_elements else ""

        # 分析图片位置上下文
        drawings_positions = []
        picts_positions = []
        all_children = list(element.iter())
        text_before_image = False
        text_after_image = False
        standalone_image = len(text_content.strip()) == 0

        # 收集所有 t 元素的位置和 drawing/pict 元素的位置
        t_positions = []
        drawing_positions = []
        pict_positions = []

        for i, child in enumerate(all_children):
            tag_with_ns = child.tag
            tag_name = tag_with_ns.split('}')[-1] if '}' in tag_with_ns else tag_with_ns

            if tag_name == 't' and child.text and child.text.strip():
                t_positions.append(i)
            elif tag_name == 'drawing':
                drawing_positions.append(i)
            elif tag_name == 'pict':
                pict_positions.append(i)

        # 确定文本和图片的相对位置关系
        image_positions = drawing_positions + pict_positions
        if image_positions and t_positions:
            min_image_pos = min(image_positions)
            max_image_pos = max(image_positions)
            min_text_pos = min(t_positions)
            max_text_pos = max(t_positions)

            text_before_image = min_text_pos < min_image_pos
            text_after_image = max_text_pos > max_image_pos
            text_surrounds_image = text_before_image and text_after_image
        drawings = element.findall(f".//{{{self.NAMESPACES['w']}}}drawing") or []
        picts = element.findall(f".//{{{self.NAMESPACES['w']}}}pict") or []
        # 获取图片描述信息
        image_descriptions = []
        embed_ids = []
        image_dimensions = []
        embed_types = []
        wrap_types = []

        for drawing in drawings:
            docPr = drawing.find(f".//wp:docPr", namespaces=self.NAMESPACES)
            if docPr is not None:
                name = docPr.get('name', '')
                desc = docPr.get('descr', '')
                image_descriptions.append({"name": name, "description": desc})

        for drawing in drawings:
            # 提取嵌入关系ID
            blip = drawing.find(f".//a:blip", namespaces=self.NAMESPACES)
            embed_id = blip.get(f"{{{self.NAMESPACES['r']}}}embed") if blip is not None else None

            # 提取图片尺寸
            extent = drawing.find(f".//wp:extent", namespaces=self.NAMESPACES)
            xfrm_ext = drawing.find(f".//a:ext", namespaces=self.NAMESPACES)

            width = None
            height = None

            # 尝试从extent获取尺寸
            if extent is not None:
                width = extent.get('cx')
                height = extent.get('cy')
            # 如果没有找到，尝试从xfrm中获取
            elif xfrm_ext is not None:
                width = xfrm_ext.get('cx')
                height = xfrm_ext.get('cy')

            # 将EMU单位转换为厘米（1厘米 = 360000 EMU）
            if width is not None and height is not None:
                try:
                    width_cm = float(width) / 360000
                    height_cm = float(height) / 360000
                    dimensions = {'width_emu': width, 'height_emu': height,
                                  'width_cm': round(width_cm, 2), 'height_cm': round(height_cm, 2)}
                except (ValueError, TypeError):
                    dimensions = {'width_emu': width, 'height_emu': height}
            else:
                dimensions = None

            # 提取嵌入类型
            inline_elem = drawing.find(f".//wp:inline", namespaces=self.NAMESPACES)
            anchor_elem = drawing.find(f".//wp:anchor", namespaces=self.NAMESPACES)

            if inline_elem is not None:
                embed_type = "inline"
                wrap_type = "inline"
            elif anchor_elem is not None:
                embed_type = "anchor"

                # 提取环绕方式
                wrap_none = anchor_elem.find(f".//wp:wrapNone", namespaces=self.NAMESPACES)
                wrap_square = anchor_elem.find(f".//wp:wrapSquare", namespaces=self.NAMESPACES)
                wrap_tight = anchor_elem.find(f".//wp:wrapTight", namespaces=self.NAMESPACES)
                wrap_through = anchor_elem.find(f".//wp:wrapThrough", namespaces=self.NAMESPACES)
                wrap_top_and_bottom = anchor_elem.find(f".//wp:wrapTopAndBottom", namespaces=self.NAMESPACES)

                if wrap_none is not None:
                    wrap_type = "none"
                elif wrap_square is not None:
                    wrap_type = "square"
                elif wrap_tight is not None:
                    wrap_type = "tight"
                elif wrap_through is not None:
                    wrap_type = "through"
                elif wrap_top_and_bottom is not None:
                    wrap_type = "topAndBottom"
                else:
                    wrap_type = "unknown"

                # 检查是否在文本前面或后面
                behind_doc = anchor_elem.get('behindDoc')
                if behind_doc == '1':
                    wrap_type += "-behind"

            else:
                embed_type = "unknown"
                wrap_type = "unknown"

            embed_ids.append(embed_id)
            image_dimensions.append(dimensions)
            embed_types.append(embed_type)
            wrap_types.append(wrap_type)

        # 处理pict元素中的图片（如果有）
        for pict in picts:
            # 提取嵌入关系ID
            imagedata = pict.find(f".//*[@src]")
            embed_id = imagedata.get('src') if imagedata is not None else None

            # pict元素的尺寸通常在shape元素中
            shape = pict.find(f".//v:shape", namespaces=self.NAMESPACES)
            width = None
            height = None

            if shape is not None:
                style = shape.get('style', '')
                # 尝试从style属性中提取宽度和高度
                for style_part in style.split(';'):
                    if 'width:' in style_part:
                        width = style_part.split('width:')[1].strip()
                    elif 'height:' in style_part:
                        height = style_part.split('height:')[1].strip()

            dimensions = {'width': width, 'height': height} if width or height else None

            # pict元素通常是旧版的内嵌图片
            embed_type = "pict"
            wrap_type = "inline"  # 默认为内嵌

            embed_ids.append(embed_id)
            image_dimensions.append(dimensions)
            embed_types.append(embed_type)
            wrap_types.append(wrap_type)

        image_info = {
            'drawings_count': len(drawings),
            'picts_count': len(picts),
            'has_text': len(text_content.strip()) > 0,
            'text_content': text_content[:100] + '...' if len(text_content) > 100 else text_content,
            'standalone_image': standalone_image,
            'text_before_image': text_before_image,
            'text_after_image': text_after_image,
            'text_surrounds_image': text_before_image and text_after_image if not standalone_image else False,
            'image_position': 'standalone' if standalone_image else
            'surrounded_by_text' if text_before_image and text_after_image else
            'after_text' if text_before_image else
            'before_text' if text_after_image else 'unknown',
            'image_descriptions': image_descriptions,
            'embed_ids': embed_ids,
            'dimensions': image_dimensions,
            'embed_types': embed_types,
            'wrap_types': wrap_types
        }
        return image_info
# 使用方法示例
# if __name__ == "__main__":
#
#         # 创建一个临时文档对象
#         doc = DocxElementParser("test.docx")
#
#         print(doc.get_image_from_pra(doc.paragraphs[0]['element']) )

