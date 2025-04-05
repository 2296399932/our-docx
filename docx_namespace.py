import xml.etree.ElementTree as ET
from io import BytesIO
import re
import os
import shutil
from docx_parser import DocxFile
import traceback
import xml.dom.minidom as minidom
import xml.etree.ElementTree as ET
import pandas as pd
import time
import os
import uuid
import base64
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
        'wpsCustomData': 'http://www.wps.cn/officeDocument/2013/wpsCustomData'
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
                self.paragraphs.append(elem_info)
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
            'other_properties': {}
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
            for key in ['before', 'after', 'line', 'lineRule']:
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
        """获取段落间距信息
        
        Args:
            paragraph_element: 段落XML元素对象
            
        Returns:
            dict: 包含间距信息的字典
        """
        result = {
            'before': None,
            'after': None,
            'line': None,
            'lineRule': None,
            'description': []
        }
        
        # 查找段落属性标签
        pPr = self.paragraphs[num]['element'].find(f".//{{{self.NAMESPACES['w']}}}pPr")
        if pPr is None:
            return result
            
        # 提取间距信息
        spacing = pPr.find(f".//{{{self.NAMESPACES['w']}}}spacing")
        if spacing is not None:
            for key in ['before', 'after', 'line', 'lineRule']:
                val = spacing.get(f"{{{self.NAMESPACES['w']}}}{key}")
                if val:
                    result[key] = val
                    
                    if key in ['before', 'after']:
                        # Word间距单位是1/20磅
                        space_pt = float(val) / 20
                        space_name = '段前距' if key == 'before' else '段后距'
                        result['description'].append(f"{space_name}: {val} (约 {space_pt:.2f}磅)")
                    elif key == 'line':
                        line_rule = result.get('lineRule', 'auto')
                        if line_rule == 'exact' or line_rule == 'atLeast':
                            # 行距是以1/20磅为单位
                            line_pt = float(val) / 20
                            rule_text = '固定值' if line_rule == 'exact' else '最小值'
                            result['description'].append(f"行距: {line_pt:.2f}磅 ({rule_text})")
                        else:
                            # 行距是以百分比为单位，240 = 2倍行距
                            line_percent = float(val) / 240 * 100
                            result['description'].append(f"行距: {line_percent:.0f}% (约 {line_percent/100:.2f}倍)")
                    elif key == 'lineRule':
                        rule_map = {
                            'auto': '自动',
                            'exact': '固定值',
                            'atLeast': '最小值'
                        }
                        rule_text = rule_map.get(val, val)
                        result['description'].append(f"行距规则: {rule_text}")
                        
        if not result['description']:
            result['description'] = ['未设置间距']
            
        return result
        
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
        """获取段落字体信息
        
        Args:
            paragraph_element: 段落XML元素对象
            
        Returns:
            dict: 包含字体信息的字典
        """
        result = {
            'fonts': {},
            'size': None,
            'attributes': {},
            'color': None,
            'description': []
        }
        
        # 查找段落属性标签
        pPr = self.paragraphs[num]['element'].find(f".//{{{self.NAMESPACES['w']}}}pPr")
        if pPr is None:
            return {'description': ['未设置段落级字体属性']}
            
        # 提取字体属性
        rPr = pPr.find(f".//{{{self.NAMESPACES['w']}}}rPr")
        if rPr is not None:
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
            
            # 提取字号            
            sz = rPr.find(f".//{{{self.NAMESPACES['w']}}}sz")
            if sz is not None:
                size_val = sz.get(f"{{{self.NAMESPACES['w']}}}val")
                result['size'] = size_val
                # Word中的字号是实际点数的两倍
                size_pt = float(size_val) / 2
                result['description'].append(f"字号: {size_pt}磅 ({size_val})")
                
            # 提取加粗、倾斜、下划线等格式
            style_names = {
                'b': '加粗',
                'i': '倾斜',
                'u': '下划线',
                'strike': '删除线',
                'caps': '全大写',
                'smallCaps': '小型大写字母'
            }
            
            for style_tag, style_name in style_names.items():
                tag = rPr.find(f".//{{{self.NAMESPACES['w']}}}{style_tag}")
                if tag is not None:
                    val = tag.get(f"{{{self.NAMESPACES['w']}}}val", 'true')
                    result['attributes'][style_tag] = val
                    is_on = val.lower() != 'false' if isinstance(val, str) else bool(val)
                    if is_on:
                        result['description'].append(f"{style_name}")
                    
            # 提取文字颜色
            color = rPr.find(f".//{{{self.NAMESPACES['w']}}}color")
            if color is not None:
                color_val = color.get(f"{{{self.NAMESPACES['w']}}}val")
                result['color'] = color_val
                result['description'].append(f"颜色: {color_val}")
                
        if not result['description']:
            result['description'] = ['未设置字体属性']
            
        return result
    
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
            'font': self.get_paragraph_font(num)
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
    def get_run_style(self, element_index, run_index,element_type="elements"):
        """提取指定元素中特定Run的所有样式信息
        
        Args:
            element_index: self.elements的索引
            run_index: 元素中w:r的索引
            
        Returns:
            dict: 包含Run样式信息的字典
        """
        # 检查元素索引是否有效
        if element_index < 0 or element_index >= len(self.elements):
            print(f"错误：元素索引{element_index}超出范围(0-{len(self.elements) - 1})")
            return {}
        element= self.element_to_dict(element_index, element_type)
        # 查找所有w:r元素
        r_elements = element.findall(f".//{{{self.NAMESPACES['w']}}}r")
        
        # 检查Run索引是否有效
        if run_index < 0 or run_index >= len(r_elements):
            print(f"错误：Run索引{run_index}超出范围(0-{len(r_elements)-1})")
            return {}
            
        # 获取指定的Run元素
        run = r_elements[run_index]
        
        # 提取Run样式信息
        style_info = {
            'fonts': {},
            'size': None,
            'bold': False,
            'italic': False,
            'underline': None,
            'color': None,
            'highlight': None,
            'strike': False,
            'caps': False,
            'small_caps': False,
            'spacing': None,
            'vert_align': None,
            'other_properties': {}
        }
        
        # 查找Run属性标签
        rPr = run.find(f".//{{{self.NAMESPACES['w']}}}rPr")
        if rPr is None:
            return {'has_style': False, 'message': 'Run无样式信息'}
            
        # 1. 提取字体
        rFonts = rPr.find(f".//{{{self.NAMESPACES['w']}}}rFonts")
        if rFonts is not None:
            for font_type in ['ascii', 'hAnsi', 'eastAsia', 'cs']:
                font = rFonts.get(f"{{{self.NAMESPACES['w']}}}{font_type}")
                if font:
                    style_info['fonts'][font_type] = font
                    
        # 2. 提取字号
        sz = rPr.find(f".//{{{self.NAMESPACES['w']}}}sz")
        if sz is not None:
            style_info['size'] = sz.get(f"{{{self.NAMESPACES['w']}}}val")
            
        # 3. 提取加粗
        b = rPr.find(f".//{{{self.NAMESPACES['w']}}}b")
        if b is not None:
            val = b.get(f"{{{self.NAMESPACES['w']}}}val", 'true')
            style_info['bold'] = val.lower() != 'false'
            
        # 4. 提取斜体
        i = rPr.find(f".//{{{self.NAMESPACES['w']}}}i")
        if i is not None:
            val = i.get(f"{{{self.NAMESPACES['w']}}}val", 'true')
            style_info['italic'] = val.lower() != 'false'
            
        # 5. 提取下划线
        u = rPr.find(f".//{{{self.NAMESPACES['w']}}}u")
        if u is not None:
            style_info['underline'] = u.get(f"{{{self.NAMESPACES['w']}}}val", 'single')
            
        # 6. 提取文字颜色
        color = rPr.find(f".//{{{self.NAMESPACES['w']}}}color")
        if color is not None:
            style_info['color'] = color.get(f"{{{self.NAMESPACES['w']}}}val")
            
        # 7. 提取突出显示
        highlight = rPr.find(f".//{{{self.NAMESPACES['w']}}}highlight")
        if highlight is not None:
            style_info['highlight'] = highlight.get(f"{{{self.NAMESPACES['w']}}}val")
            
        # 8. 提取删除线
        strike = rPr.find(f".//{{{self.NAMESPACES['w']}}}strike")
        if strike is not None:
            val = strike.get(f"{{{self.NAMESPACES['w']}}}val", 'true')
            style_info['strike'] = val.lower() != 'false'
            
        # 9. 提取大小写格式
        caps = rPr.find(f".//{{{self.NAMESPACES['w']}}}caps")
        if caps is not None:
            val = caps.get(f"{{{self.NAMESPACES['w']}}}val", 'true')
            style_info['caps'] = val.lower() != 'false'
            
        # 10. 提取小型大写字母
        smallCaps = rPr.find(f".//{{{self.NAMESPACES['w']}}}smallCaps")
        if smallCaps is not None:
            val = smallCaps.get(f"{{{self.NAMESPACES['w']}}}val", 'true')
            style_info['small_caps'] = val.lower() != 'false'
            
        # 11. 提取字符间距
        spacing = rPr.find(f".//{{{self.NAMESPACES['w']}}}spacing")
        if spacing is not None:
            style_info['spacing'] = spacing.get(f"{{{self.NAMESPACES['w']}}}val")
            
        # 12. 提取上下标
        vertAlign = rPr.find(f".//{{{self.NAMESPACES['w']}}}vertAlign")
        if vertAlign is not None:
            style_info['vert_align'] = vertAlign.get(f"{{{self.NAMESPACES['w']}}}val")
            
        # 13. 提取其他属性
        for child in rPr:
            tag_name = child.tag.split('}')[-1] if '}' in child.tag else child.tag
            # 跳过已经处理过的标签
            if tag_name in ['rFonts', 'sz', 'b', 'i', 'u', 'color', 'highlight', 'strike', 'caps', 'smallCaps', 'spacing', 'vertAlign']:
                continue
                
            # 处理其他标签
            attrs = {}
            for key, value in child.attrib.items():
                # 简化命名空间
                attr_name = key.split('}')[-1] if '}' in key else key
                attrs[attr_name] = value
                
            style_info['other_properties'][tag_name] = attrs
            
        return style_info
        
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

    def get_table_style(self, table_index):
        """提取表格的所有样式和属性信息
        
        Args:
            table_index: self.tables中的表格索引
            
        Returns:
            dict: 包含表格样式和属性信息的字典
        """
        # 检查索引是否有效
        if table_index < 0 or table_index >= len(self.tables):
            print(f"错误：表格索引{table_index}超出范围(0-{len(self.tables)-1})")
            return {}
            
        # 获取表格元素
        table = self.tables[table_index]['element']
        
        # 创建结果字典
        style_info = {
            'style_id': None,
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
            'layout': None,
            'cell_margins': {
                'top': {},
                'left': {},
                'bottom': {},
                'right': {}
            },
            'grid': [],
            'rows_count': 0,
            'columns_count': 0,
            'description': []
        }
        
        # 查找表格属性
        tblPr = table.find(f".//{{{self.NAMESPACES['w']}}}tblPr")
        if tblPr is not None:
            # 提取样式ID
            style = tblPr.find(f".//{{{self.NAMESPACES['w']}}}tblStyle")
            if style is not None:
                style_info['style_id'] = style.get(f"{{{self.NAMESPACES['w']}}}val")
                
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
                            'space': border.get(f"{{{self.NAMESPACES['w']}}}space")
                        }
                        
            # 提取表格布局
            tblLayout = tblPr.find(f".//{{{self.NAMESPACES['w']}}}tblLayout")
            if tblLayout is not None:
                style_info['layout'] = tblLayout.get(f"{{{self.NAMESPACES['w']}}}type")
                
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
        
        # 提取表格网格（列定义）
        tblGrid = table.find(f".//{{{self.NAMESPACES['w']}}}tblGrid")
        if tblGrid is not None:
            grid_cols = tblGrid.findall(f".//{{{self.NAMESPACES['w']}}}gridCol")
            for col in grid_cols:
                col_width = col.get(f"{{{self.NAMESPACES['w']}}}w")
                style_info['grid'].append(col_width)
            
            style_info['columns_count'] = len(grid_cols)
                
        # 统计行数
        rows = table.findall(f".//{{{self.NAMESPACES['w']}}}tr")
        style_info['rows_count'] = len(rows)
        
        # 格式化描述信息
        style_info['description'].append(f"表格大小: {style_info['rows_count']}行 × {style_info['columns_count']}列")
        
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
                    'none': '无'
                }.get(border.get('val'), border.get('val'))
                
                if border_type != '无':
                    border_size = f"{float(border.get('size', '1')) / 8:.1f}磅" if border.get('size') else ""
                    border_color = border.get('color', 'auto')
                    borders_desc.append(f"{border_name}: {border_type} {border_size} {border_color}")
                    
        if borders_desc:
            style_info['description'].append("边框: " + ", ".join(borders_desc))
            
        # 布局描述
        if style_info['layout']:
            layout_desc = {
                'autofit': '自动适应内容',
                'fixed': '固定宽度'
            }.get(style_info['layout'], style_info['layout'])
            style_info['description'].append(f"布局: {layout_desc}")
            
        # 列宽描述
        if style_info['grid']:
            col_widths = []
            for i, width in enumerate(style_info['grid']):
                # 转换为磅
                pt_width = float(width) / 20
                col_widths.append(f"列{i+1}: {pt_width:.1f}磅")
            style_info['description'].append("列宽: " + ", ".join(col_widths))
            
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
            print(f"错误：段落索引{para_index}超出范围(0-{len(self.paragraphs)-1})")
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
        """设置段落间距
        
        Args:
            para_index: 段落索引
            **spacing: 间距设置，可包含以下参数:
                before: 段前距
                after: 段后距
                line: 行距值
                lineRule: 行距规则 (auto, exact, atLeast)
            
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
            
            # 查找间距元素
            spacing_elem = pPr.find(f".//{{{self.NAMESPACES['w']}}}spacing")
            if spacing_elem is None:
                # 如果不存在，则创建
                spacing_elem = ET.Element(f"{{{self.NAMESPACES['w']}}}spacing")
                pPr.append(spacing_elem)
                
            # 设置各类间距
            valid_props = ['before', 'after', 'line', 'lineRule']
            for prop, value in spacing.items():
                if prop in valid_props and value is not None:
                    spacing_elem.set(f"{{{self.NAMESPACES['w']}}}{prop}", str(value))
                    
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
    
    def update_paragraph_style(self, para_index, **style_properties):
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
        if 'font' in style_properties and isinstance(style_properties['font'], dict):
            if not self.set_paragraph_font(para_index, **style_properties['font']):
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
            print(1111111)
            
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
        """修改段落中所有文本运行的字体属性
        
        Args:
            para_index: 段落索引
            **font_properties: 字体属性设置
            
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
            r_elements = paragraph.findall(f".//{{{self.NAMESPACES['w']}}}r")
            if not r_elements:
                print(f"段落{para_index}中没有找到文本运行")
                return False
            
            # 修改每个文本运行的字体属性
            for r in r_elements:
                # 查找或创建rPr元素
                rPr = r.find(f".//{{{self.NAMESPACES['w']}}}rPr")
                if rPr is None:
                    rPr = ET.Element(f"{{{self.NAMESPACES['w']}}}rPr")
                    # 插入到r的第一个位置
                    r.insert(0, rPr)
                
                # 设置字体
                if any(font_type in font_properties for font_type in ['ascii', 'hAnsi', 'eastAsia', 'cs']):
                    rFonts = rPr.find(f".//{{{self.NAMESPACES['w']}}}rFonts")
                    if rFonts is None:
                        rFonts = ET.Element(f"{{{self.NAMESPACES['w']}}}rFonts")
                        rPr.append(rFonts)
                    
                    for font_type in ['ascii', 'hAnsi', 'eastAsia', 'cs']:
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
                        rPr.remove(b)
                
                # 设置颜色
                if 'color' in font_properties:
                    color = rPr.find(f".//{{{self.NAMESPACES['w']}}}color")
                    if color is None:
                        color = ET.Element(f"{{{self.NAMESPACES['w']}}}color")
                        rPr.append(color)
                    color.set(f"{{{self.NAMESPACES['w']}}}val", font_properties['color'])
            
            return True
        except Exception as e:
            print(f"修改段落文本运行字体时出错: {e}")
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
            r_elements = paragraph.findall(f".//{{{self.NAMESPACES['w']}}}r")
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
            r_elements = paragraph.findall(f".//{{{self.NAMESPACES['w']}}}r")
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
            r_elements = paragraph.findall(f".//{{{self.NAMESPACES['w']}}}r")
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
            r_elements = paragraph.findall(f".//{{{self.NAMESPACES['w']}}}r")
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
            r_elements = paragraph.findall(f".//{{{self.NAMESPACES['w']}}}r")
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
            r_elements = paragraph.findall(f".//{{{self.NAMESPACES['w']}}}r")
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
            r_elements = paragraph.findall(f".//{{{self.NAMESPACES['w']}}}r")
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
            r_elements = paragraph.findall(f".//{{{self.NAMESPACES['w']}}}r")
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
            r_elements = paragraph.findall(f".//{{{self.NAMESPACES['w']}}}r")
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
            r_elements = paragraph.findall(f".//{{{self.NAMESPACES['w']}}}r")
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

    # 以下是修改单个文本运行的样式函数
    
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
            print(f"错误：段落索引{para_index}超出范围(0-{len(self.paragraphs)-1})")
            return None
            
        # 获取段落元素
        paragraph = self.paragraphs[para_index]['element']
        
        # 查找所有w:r元素
        r_elements = paragraph.findall(f".//{{{self.NAMESPACES['w']}}}r")
        if not r_elements:
            print(f"段落{para_index}中没有找到文本运行")
            return None
            
        # 检查文本运行索引是否有效
        if run_index < 0 or run_index >= len(r_elements):
            print(f"错误：文本运行索引{run_index}超出范围(0-{len(r_elements)-1})")
            return None
            
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
        """获取段落中文本运行的数量
        
        Args:
            para_index: 段落索引
            
        Returns:
            int: 文本运行的数量，如果段落索引无效则返回-1
        """
        # 检查段落索引是否有效
        if para_index < 0 or para_index >= len(self.paragraphs):
            print(f"错误：段落索引{para_index}超出范围(0-{len(self.paragraphs)-1})")
            return -1
            
        # 获取段落元素
        paragraph = self.paragraphs[para_index]['element']
        
        # 查找所有w:r元素
        r_elements = paragraph.findall(f".//{{{self.NAMESPACES['w']}}}r")
        return len(r_elements)
        
    def get_run_text(self, para_index, run_index):
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
        """设置特定文本运行的字体属性
        
        Args:
            para_index: 段落索引
            run_index: 文本运行索引
            **font_properties: 字体属性设置，可包含以下参数:
                ascii: 英文字体
                hAnsi: 西文字体
                eastAsia: 中文字体
                cs: 复杂文种字体
            
        Returns:
            bool: 是否成功修改
        """
        # 获取文本运行元素
        r_element = self._get_run_element(para_index, run_index)
        if r_element is None:
            return False
            
        try:
            # 获取或创建rPr元素
            rPr = self._get_or_create_rPr(r_element)
            
            # 如果有设置字体
            if any(font_type in font_properties for font_type in ['ascii', 'hAnsi', 'eastAsia', 'cs']):
                # 查找字体元素
                rFonts = rPr.find(f".//{{{self.NAMESPACES['w']}}}rFonts")
                if rFonts is None:
                    # 如果不存在，则创建
                    rFonts = ET.Element(f"{{{self.NAMESPACES['w']}}}rFonts")
                    rPr.append(rFonts)
                    
                # 设置各类字体
                for font_type in ['ascii', 'hAnsi', 'eastAsia', 'cs']:
                    if font_type in font_properties:
                        rFonts.set(f"{{{self.NAMESPACES['w']}}}{font_type}", font_properties[font_type])
                        
            return True
        except Exception as e:
            print(f"设置文本运行字体时出错: {e}")
            return False
            
    def set_run_size(self, para_index, run_index, size):
        """设置特定文本运行的字号
        
        Args:
            para_index: 段落索引
            run_index: 文本运行索引
            size: 字号值(半磅值)，如24表示12磅
            
        Returns:
            bool: 是否成功修改
        """
        # 获取文本运行元素
        r_element = self._get_run_element(para_index, run_index)
        if r_element is None:
            return False
            
        try:
            # 获取或创建rPr元素
            rPr = self._get_or_create_rPr(r_element)
            
            # 查找字号元素
            sz = rPr.find(f".//{{{self.NAMESPACES['w']}}}sz")
            if size is None:
                # 如果要移除字号设置
                if sz is not None:
                    rPr.remove(sz)
            else:
                # 如果要设置字号
                if sz is None:
                    sz = ET.Element(f"{{{self.NAMESPACES['w']}}}sz")
                    rPr.append(sz)
                sz.set(f"{{{self.NAMESPACES['w']}}}val", str(size))
                
            return True
        except Exception as e:
            print(f"设置文本运行字号时出错: {e}")
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
        r_element = self._get_run_element(para_index, run_index)
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
        r_element = self._get_run_element(para_index, run_index)
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
        r_element = self._get_run_element(para_index, run_index)
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
        r_element = self._get_run_element(para_index, run_index)
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
        r_element = self._get_run_element(para_index, run_index)
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
        r_element = self._get_run_element(para_index, run_index)
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
            
    def update_run_style(self, para_index, run_index, **style_properties):
        """更新特定文本运行的多个样式属性
        
        Args:
            para_index: 段落索引
            run_index: 文本运行索引
            **style_properties: 样式属性字典，可包含以下键：
                'fonts': 字体设置字典，包含'ascii', 'eastAsia'等键
                'size': 字号值
                'bold': 是否加粗
                'italic': 是否斜体
                'underline': 下划线类型
                'color': 字体颜色
                'highlight': 高亮颜色
                'strike': 是否添加删除线
                
        Returns:
            bool: 是否成功更新所有样式
        """
        # 获取文本运行元素
        r_element = self._get_run_element(para_index, run_index)
        if r_element is None:
            return False
            
        try:
            # 获取或创建rPr元素
            rPr = self._get_or_create_rPr(r_element)
            
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
                    
            return True
        except Exception as e:
            print(f"更新文本运行样式时出错: {e}")
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
                'font': 字体设置字典，包含'ascii', 'eastAsia'等键
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
            if any(key in style_properties for key in ['style_id', 'alignment', 'indentation', 'spacing']):
                pPr = ET.SubElement(new_para, f"{{{self.NAMESPACES['w']}}}pPr")
                
                # 设置样式ID
                if 'style_id' in style_properties:
                    pStyle = ET.SubElement(pPr, f"{{{self.NAMESPACES['w']}}}pStyle")
                    pStyle.set(f"{{{self.NAMESPACES['w']}}}val", style_properties['style_id'])
                    
                # 设置对齐方式
                if 'alignment' in style_properties:
                    jc = ET.SubElement(pPr, f"{{{self.NAMESPACES['w']}}}jc")
                    jc.set(f"{{{self.NAMESPACES['w']}}}val", style_properties['alignment'])
                    
                # 设置缩进
                if 'indentation' in style_properties and isinstance(style_properties['indentation'], dict):
                    ind = ET.SubElement(pPr, f"{{{self.NAMESPACES['w']}}}ind")
                    for ind_type, value in style_properties['indentation'].items():
                        ind.set(f"{{{self.NAMESPACES['w']}}}{ind_type}", str(value))
                        
                # 设置间距
                if 'spacing' in style_properties and isinstance(style_properties['spacing'], dict):
                    spacing = ET.SubElement(pPr, f"{{{self.NAMESPACES['w']}}}spacing")
                    for spacing_type, value in style_properties['spacing'].items():
                        spacing.set(f"{{{self.NAMESPACES['w']}}}{spacing_type}", str(value))
                        
                # 设置段落级别的字体属性
                if any(key in style_properties for key in ['font', 'size', 'bold', 'color']):
                    rPr = ET.SubElement(pPr, f"{{{self.NAMESPACES['w']}}}rPr")
                    
                    # 字体
                    if 'font' in style_properties and isinstance(style_properties['font'], dict):
                        font = style_properties['font']
                        if any(font_type in font for font_type in ['ascii', 'hAnsi', 'eastAsia', 'cs']):
                            rFonts = ET.SubElement(rPr, f"{{{self.NAMESPACES['w']}}}rFonts")
                            for font_type, font_name in font.items():
                                rFonts.set(f"{{{self.NAMESPACES['w']}}}{font_type}", font_name)
                                
                    # 字号
                    if 'size' in style_properties:
                        sz = ET.SubElement(rPr, f"{{{self.NAMESPACES['w']}}}sz")
                        sz.set(f"{{{self.NAMESPACES['w']}}}val", str(style_properties['size']))
                        
                    # 加粗
                    if 'bold' in style_properties and style_properties['bold']:
                        b = ET.SubElement(rPr, f"{{{self.NAMESPACES['w']}}}b")
                        b.set(f"{{{self.NAMESPACES['w']}}}val", "true")
                        
                    # 颜色
                    if 'color' in style_properties:
                        color = ET.SubElement(rPr, f"{{{self.NAMESPACES['w']}}}color")
                        color.set(f"{{{self.NAMESPACES['w']}}}val", style_properties['color'])
            
            # 创建文本运行元素
            if text:
                r = ET.SubElement(new_para, f"{{{self.NAMESPACES['w']}}}r")
                
                # 如果有运行级样式属性
                if any(key in style_properties for key in ['font', 'size', 'bold', 'color']):
                    rPr = ET.SubElement(r, f"{{{self.NAMESPACES['w']}}}rPr")
                    
                    # 字体
                    if 'font' in style_properties and isinstance(style_properties['font'], dict):
                        font = style_properties['font']
                        if any(font_type in font for font_type in ['ascii', 'hAnsi', 'eastAsia', 'cs']):
                            rFonts = ET.SubElement(rPr, f"{{{self.NAMESPACES['w']}}}rFonts")
                            for font_type, font_name in font.items():
                                rFonts.set(f"{{{self.NAMESPACES['w']}}}{font_type}", font_name)
                                
                    # 字号
                    if 'size' in style_properties:
                        sz = ET.SubElement(rPr, f"{{{self.NAMESPACES['w']}}}sz")
                        sz.set(f"{{{self.NAMESPACES['w']}}}val", str(style_properties['size']))
                        
                    # 加粗
                    if 'bold' in style_properties and style_properties['bold']:
                        b = ET.SubElement(rPr, f"{{{self.NAMESPACES['w']}}}b")
                        b.set(f"{{{self.NAMESPACES['w']}}}val", "true")
                        
                    # 颜色
                    if 'color' in style_properties:
                        color = ET.SubElement(rPr, f"{{{self.NAMESPACES['w']}}}color")
                        color.set(f"{{{self.NAMESPACES['w']}}}val", style_properties['color'])
                
                # 添加文本
                t = ET.SubElement(r, f"{{{self.NAMESPACES['w']}}}t")
                # 如果文本包含空格或特殊字符，设置xml:space="preserve"
                if text.startswith(' ') or text.endswith(' ') or '  ' in text:
                    t.set(f"{{{self.NAMESPACES['xml']}}}space", "preserve")
                t.text = text
            
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

    def insert_image(self, para_index, run_index=-1, position='after', image_path='', 
                     width=None, height=None, description=None):
        """在文档中指定位置插入图片
        
        Args:
            para_index: self.elements或self.paragraphs中的段落索引
            run_index: 段落中文本运行的索引，-1表示段落末尾
            position: 插入位置，'before'表示在运行前插入，'after'表示在运行后插入
            image_path: 图片文件的路径
            width: 图片宽度(厘米)，不指定则使用原始大小
            height: 图片高度(厘米)，不指定则使用原始大小
            description: 图片描述
            
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
                        return None
                else:
                    print(f"错误：索引{para_index}不是有效的段落索引")
                    return None
        except Exception as e:
            print(f"获取段落元素时出错: {e}")
            return None
            
        # 获取段落中的文本运行
        try:
            r_elements = paragraph.findall(f".//{{{self.NAMESPACES['w']}}}r")
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
            rel_id = f"rId{int(time.time()*1000)}"
            
            # 创建图片关系
            # 检查是否已经存在media文件夹
            if 'media' not in self.parts:
                self.parts['media'] = {}
                
            # 将图片添加到media文件夹
            self.parts['media'][img_name] = img_data
            print(self.parts['other'].keys())
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
            self.parts['relationships']= rels_tree
            
            # 创建图片XML结构
            new_run = ET.Element(f"{{{self.NAMESPACES['w']}}}r")
            drawing = ET.SubElement(new_run, f"{{{self.NAMESPACES['w']}}}drawing")
            inline = ET.SubElement(drawing, f"{{{self.NAMESPACES['wp']}}}inline")
            
            # 设置图片大小
            extent = ET.SubElement(inline, f"{{{self.NAMESPACES['wp']}}}extent")
            extent.set("cx", str(width_emu))
            extent.set("cy", str(height_emu))
            
            # 设置效果范围
            effect_extent = ET.SubElement(inline, f"{{{self.NAMESPACES['wp']}}}effectExtent")
            effect_extent.set("l", "0")
            effect_extent.set("t", "0")
            effect_extent.set("r", "0")
            effect_extent.set("b", "0")
            
            # 设置DOC PROPS
            doc_pr = ET.SubElement(inline, f"{{{self.NAMESPACES['wp']}}}docPr")
            doc_pr.set("id", img_id)
            doc_pr.set("name", img_name)
            if description:
                doc_pr.set("descr", description)
                
            # 添加图片数据
            graphic = ET.SubElement(inline, f"{{{self.NAMESPACES['a']}}}graphic")
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
            
            # 根据position参数插入图片
            if position.lower() == 'before':
                paragraph.insert(list(paragraph).index(target_run), new_run)
            else:  # 默认在后面插入
                paragraph.insert(list(paragraph).index(target_run) + 1, new_run)
                
            # 成功添加图片
            return rel_id
            
        except Exception as e:

            print(f"插入图片时出错: {e}")
            traceback.print_exc()
            return None

    def insert_run(self, para_index, run_index=-1, position='after', text='', **style_properties):
        """在段落中插入新的文本运行(run)
        
        Args:
            para_index: self.elements或self.paragraphs中的段落索引
            run_index: 段落中文本运行的索引，-1表示段落末尾
            position: 插入位置，'before'表示在运行前插入，'after'表示在运行后插入
            text: 要插入的文本内容
            **style_properties: 文本运行的样式属性，可包含以下键：
                'font': 字体设置字典，包含'ascii', 'eastAsia'等键
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
            r_elements = paragraph.findall(f".//{{{self.NAMESPACES['w']}}}r")
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
            
            # 设置运行属性
            if style_properties:
                rPr = ET.SubElement(new_run, f"{{{self.NAMESPACES['w']}}}rPr")
                
                # 设置字体
                if 'font' in style_properties and isinstance(style_properties['font'], dict):
                    font = style_properties['font']
                    if any(font_type in font for font_type in ['ascii', 'hAnsi', 'eastAsia', 'cs']):
                        rFonts = ET.SubElement(rPr, f"{{{self.NAMESPACES['w']}}}rFonts")
                        for font_type, font_name in font.items():
                            if font_type in ['ascii', 'hAnsi', 'eastAsia', 'cs']:
                                rFonts.set(f"{{{self.NAMESPACES['w']}}}{font_type}", font_name)
                                
                # 设置字号
                if 'size' in style_properties:
                    sz = ET.SubElement(rPr, f"{{{self.NAMESPACES['w']}}}sz")
                    sz.set(f"{{{self.NAMESPACES['w']}}}val", str(style_properties['size']))
                    
                # 设置加粗
                if 'bold' in style_properties and style_properties['bold']:
                    b = ET.SubElement(rPr, f"{{{self.NAMESPACES['w']}}}b")
                    b.set(f"{{{self.NAMESPACES['w']}}}val", "true")
                    
                # 设置斜体
                if 'italic' in style_properties and style_properties['italic']:
                    i = ET.SubElement(rPr, f"{{{self.NAMESPACES['w']}}}i")
                    i.set(f"{{{self.NAMESPACES['w']}}}val", "true")
                    
                # 设置下划线
                if 'underline' in style_properties and style_properties['underline']:
                    u = ET.SubElement(rPr, f"{{{self.NAMESPACES['w']}}}u")
                    u.set(f"{{{self.NAMESPACES['w']}}}val", style_properties['underline'])
                    
                # 设置颜色
                if 'color' in style_properties:
                    color = ET.SubElement(rPr, f"{{{self.NAMESPACES['w']}}}color")
                    color.set(f"{{{self.NAMESPACES['w']}}}val", style_properties['color'])
                    
                # 设置高亮
                if 'highlight' in style_properties:
                    highlight = ET.SubElement(rPr, f"{{{self.NAMESPACES['w']}}}highlight")
                    highlight.set(f"{{{self.NAMESPACES['w']}}}val", style_properties['highlight'])
                    
                # 设置删除线
                if 'strike' in style_properties and style_properties['strike']:
                    strike = ET.SubElement(rPr, f"{{{self.NAMESPACES['w']}}}strike")
                    strike.set(f"{{{self.NAMESPACES['w']}}}val", "true")
                    
                # 设置大写
                if 'caps' in style_properties and style_properties['caps']:
                    caps = ET.SubElement(rPr, f"{{{self.NAMESPACES['w']}}}caps")
                    caps.set(f"{{{self.NAMESPACES['w']}}}val", "true")
                    
                # 设置垂直对齐
                if 'vert_align' in style_properties:
                    vert_align = ET.SubElement(rPr, f"{{{self.NAMESPACES['w']}}}vertAlign")
                    vert_align.set(f"{{{self.NAMESPACES['w']}}}val", style_properties['vert_align'])
            
            # 添加文本内容
            if text:
                t = ET.SubElement(new_run, f"{{{self.NAMESPACES['w']}}}t")
                # 如果文本包含空格或特殊字符，设置xml:space="preserve"
                if text.startswith(' ') or text.endswith(' ') or '  ' in text:
                    t.set(f"{{{self.NAMESPACES['xml']}}}space", "preserve")
                t.text = text
            
            # 根据position参数和现有文本运行插入新的文本运行
            if r_elements:
                # 有现有的文本运行
                target_run = r_elements[run_index]
                if position.lower() == 'before':
                    paragraph.insert(list(paragraph).index(target_run), new_run)
                else:  # 默认在后面插入
                    paragraph.insert(list(paragraph).index(target_run) + 1, new_run)
            else:
                # 没有现有的文本运行，直接添加到段落
                paragraph.append(new_run)
                
            return True
            
        except Exception as e:
            import traceback
            print(f"插入文本运行时出错: {e}")
            traceback.print_exc()
            return False


# 使用方法示例
if __name__ == "__main__":
    # 创建一个解析器对象
    parser = DocxElementParser('output.docx')
    
    # 调用方法
    # print(parser.extract_images_simple('./output'))
    # 或
    print( parser.extract_paragraph_style(parser.elements[200]['element']))
    # parser.set_paragraph_font(200,
    #                           eastAsia="黑体",
    #                           ascii="Times New Roman",
    #                           size=28,  # 14磅
    #                           bold=True,
    #                           color="FF0000"  # 红色
    #                           )
    # print( parser.extract_paragraph_style(parser.paragraphs[200]['element']))
    # parser.save('output.docx')
    #

    # 修改段落和其中所有文本运行的字体

    # 同时修改所有文本运行的字体
    parser.set_paragraph_spacing(200, 
                           line=600,  # 行距值
                           lineRule="auto",  # 行距规则：auto(倍数)/exact(精确值)/atLeast(最小值)
                           before=400,  # 段前距
                           after=400   # 段后距
                           )
    print( parser.extract_paragraph_style(parser.paragraphs[200]['element']))

    # parser.save('output2.docx')
    #
    parser.insert_paragraph(text="这是插入的段落")
    # 先保存当前的样式ID
    current_style = parser.extract_paragraph_style(parser.paragraphs[200]['element'])
    style_id = current_style.get('style_id')

    # 设置段落样式
    parser.set_paragraph_style_id(200, '3')  # 设置基本样式

    # 然后覆盖特定属性
    parser.set_paragraph_spacing(200, 
                           before=400,
                           after=400,
                           line=600,
                           lineRule="auto")

    # 如果之前有样式ID，重新设置
    if style_id:
        parser.set_paragraph_style_id(200, style_id)
    
    # # 保存文档
    # parser.save('output2.docx')



    # 示例3：在文档最后一个段落插入图片
    relation_id = parser.insert_image(
        para_index=-1,        # 最后一个段落
        image_path="image_21.png",
        description="签名"
    )

    # 保存文档
    parser.save('output_with_images.docx')