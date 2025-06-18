import os
from typing import Dict, Any, List
import uuid
from fastapi import UploadFile, HTTPException
from models.document_models import DocumentResponse, ParagraphModel, RunModel, TableModel, ImageModel, DocElement, DocumentModel
from util.style_analyzer import StyleAnalyzer
import logging
from datetime import datetime

# 配置日志
logger = logging.getLogger(__name__)

# 添加中文字号到Canvas-Editor字号的映射表
# 映射关系基于Word中字号（磅值的2倍）到Canvas-Editor要求的数值
CHINESE_FONT_SIZE_MAPPING = {
    # Word字号值: Canvas-Editor字号值
    '84': 56,  # 初号
    '72': 48,  # 小初
    '52': 34,  # 一号
    '48': 32,  # 小一
    '36': 29,  # 二号
    '32': 24,  # 小二
    '30': 21,  # 三号
    '28': 20,  # 小三
    '24': 18,  # 四号
    '21': 16,  # 小四
    '18': 14,  # 五号
    '15': 12,  # 小五
    '13': 10,  # 六号
    '11': 8,   # 小六
    '10': 7,   # 七号
    '8': 6,    # 八号
}

# 字体映射，将Word文档中的字体名称映射到Canvas-Editor支持的字体
FONT_NAME_MAPPING = {
    '宋体': 'SimSun',
    '黑体': 'SimHei',
    '楷体': 'KaiTi',
    '仿宋': 'FangSong',
    '微软雅黑': 'Microsoft YaHei',
    '华文宋体': 'STSong',
    '华文黑体': 'STHeiti',
    '华文楷体': 'STKaiti',
    '华文仿宋': 'STFangsong',
    'Times New Roman': 'Times New Roman',
    'Arial': 'Arial',
}


class DocumentService:
    """文档处理服务"""

    @staticmethod
    async def save_uploaded_file(file_content: bytes, filename: str, upload_dir: str = "uploads") -> DocumentResponse:
        """保存上传的文件并返回响应"""
        # 确保上传目录存在
        os.makedirs(upload_dir, exist_ok=True)

        # 生成唯一文件名避免冲突
        file_extension = os.path.splitext(filename)[1]
        unique_filename = f"{uuid.uuid4()}{file_extension}"
        file_path = os.path.join(upload_dir, unique_filename)

        # 保存文件
        with open(file_path, "wb") as file:
            file.write(file_content)

        # 返回文件上传信息，不进行文档处理
        return DocumentResponse(
            filename=filename,
            status="文件上传成功",
            file_path=file_path,
            upload_time=datetime.now()
        )

    @staticmethod
    def convert_to_canvas_editor_format(content: List[Any]) -> List[Dict]:
        """
        将内部文档格式转换为Canvas-Editor期望的IElement[]格式
        """
        canvas_elements = []

        for item in content:
            # 处理Pydantic模型对象，将其转换为字典
            if hasattr(item, "dict"):
                item_dict = item.dict()
            elif hasattr(item, "__dict__"):
                item_dict = item.__dict__
            else:
                item_dict = item  # 假设已经是字典

            # 获取元素类型
            element_type = item_dict.get('type', 'paragraph')

            if element_type == 'paragraph' or element_type == 'title':
                # 转换段落为Canvas-Editor格式
                paragraph_element = {
                    'type': 'paragraph',
                    'value': []
                }

                # 如果是标题段落
                if element_type == 'title':
                    level = item_dict.get('level', 1)
                    paragraph_element['level'] = level

                # 设置对齐方式
                if item_dict.get('align'):
                    paragraph_element['textAlign'] = item_dict['align']

                # 设置间距
                if item_dict.get('spacing_before') is not None:
                    paragraph_element['marginTop'] = item_dict['spacing_before']
                if item_dict.get('spacing_after') is not None:
                    paragraph_element['marginBottom'] = item_dict['spacing_after']

                # 设置缩进
                if item_dict.get('indent_left') is not None:
                    paragraph_element['marginLeft'] = item_dict['indent_left']
                if item_dict.get('indent_right') is not None:
                    paragraph_element['marginRight'] = item_dict['indent_right']
                if item_dict.get('indent_first_line') is not None:
                    paragraph_element['firstLineIndent'] = item_dict['indent_first_line']

                # 处理段落中的文本运行
                runs = item_dict.get('runs', [])
                # 如果没有文本运行，至少添加一个空文本值
                if not runs:
                    paragraph_element['value'].append({'value': ''})
                else:
                    for run in runs:
                        # 确保run也是字典格式
                        if hasattr(run, "dict"):
                            run_dict = run.dict()
                        elif hasattr(run, "__dict__"):
                            run_dict = run.__dict__
                        else:
                            run_dict = run

                        text_value = {
                            'value': run_dict.get('text', '')
                        }

                        # 复制文本样式属性
                        if run_dict.get('font_name'):
                            text_value['fontFamily'] = run_dict['font_name']

                        if run_dict.get('font_size'):
                            text_value['size'] = run_dict['font_size']

                        if run_dict.get('bold'):
                            text_value['bold'] = True

                        if run_dict.get('italic'):
                            text_value['italic'] = True

                        if run_dict.get('underline'):
                            text_value['underline'] = True

                        if run_dict.get('strike'):
                            text_value['strikeout'] = True

                        if run_dict.get('color'):
                            text_value['color'] = run_dict['color']

                        if run_dict.get('highlight'):
                            text_value['highlight'] = run_dict['highlight']

                        if run_dict.get('superscript'):
                            text_value['type'] = 'superscript'

                        if run_dict.get('subscript'):
                            text_value['type'] = 'subscript'

                        paragraph_element['value'].append(text_value)

                canvas_elements.append(paragraph_element)

            elif element_type == 'table':
                # 转换表格为Canvas-Editor格式
                table_element = {
                    'type': 'table',
                    'value': {
                        'width': 0,
                        'colgroup': [],
                        'trList': []
                    }
                }

                # 获取表格样式中的宽度
                style = item_dict.get('style', {})
                if isinstance(style, dict) and style.get('width'):
                    table_element['value']['width'] = style['width']

                # 设置列组信息
                dimensions = item_dict.get('dimensions', {})
                cols_count = 0

                # 处理dimensions根据实际类型
                if isinstance(dimensions, dict):
                    cols_count = dimensions.get('columns', 0)
                elif isinstance(dimensions, (list, tuple)) and len(dimensions) >= 2:
                    # 假设第二个元素是列数
                    cols_count = dimensions[1]

                # 默认每列宽度相等
                col_width = 100
                for i in range(cols_count):
                    table_element['value']['colgroup'].append({
                        'width': col_width
                    })

                # 处理表格行
                rows = item_dict.get('rows', [])
                for row in rows:
                    tr_element = {
                        'height': 30,  # 默认行高
                        'tdList': []
                    }

                    for cell in row:
                        # 确保cell是字典
                        if hasattr(cell, "dict"):
                            cell_dict = cell.dict()
                        elif hasattr(cell, "__dict__"):
                            cell_dict = cell.__dict__
                        else:
                            cell_dict = cell

                        td_element = {
                            'colspan': 1,
                            'rowspan': 1,
                            'value': []
                        }

                        # 提取单元格内容，转换为段落元素
                        cell_content = cell_dict.get('content', [])
                        for content_item in cell_content:
                            if isinstance(content_item, dict) and content_item.get('type') == 'paragraph':
                                p_element = {
                                    'type': 'paragraph',
                                    'value': []
                                }

                                for run in content_item.get('runs', []):
                                    run_dict = run if isinstance(run, dict) else run.__dict__ if hasattr(run, "__dict__") else {}
                                    text_value = {
                                        'value': run_dict.get('text', '')
                                    }

                                    # 应用样式
                                    run_style = run_dict.get('style', {})
                                    if isinstance(run_style, dict):
                                        if 'bold' in run_style:
                                            text_value['bold'] = run_style['bold']
                                        if 'italic' in run_style:
                                            text_value['italic'] = run_style['italic']

                                    p_element['value'].append(text_value)

                                td_element['value'].append(p_element)

                        # 如果没有复杂内容，则使用简单文本
                        if not td_element['value'] and cell_dict.get('text'):
                            p_element = {
                                'type': 'paragraph',
                                'value': [{
                                    'value': cell_dict.get('text', '')
                                }]
                            }
                            td_element['value'].append(p_element)

                        tr_element['tdList'].append(td_element)

                    table_element['value']['trList'].append(tr_element)

                canvas_elements.append(table_element)

            elif element_type == 'image':
                # 转换图片为Canvas-Editor格式
                image_element = {
                    'type': 'image',
                    'width': 100,  # 默认宽度
                    'height': 100,  # 默认高度
                }

                # 设置图片属性
                if item_dict.get('width'):
                    image_element['width'] = item_dict['width']
                if item_dict.get('height'):
                    image_element['height'] = item_dict['height']
                if item_dict.get('file_name'):
                    image_element['value'] = f"images/{item_dict['file_name']}"

                canvas_elements.append(image_element)

        return canvas_elements

    @staticmethod
    async def process_document(file_path: str) -> Dict[str, Any]:
        """处理Word文档并提取内容"""
        try:
            logger.info(f"开始处理文档: {file_path}")

            # 检查文件是否存在
            if not os.path.exists(file_path):
                logger.error(f"文件不存在: {file_path}")
                raise HTTPException(status_code=404, detail=f"文件不存在: {file_path}")

            # 使用StyleAnalyzer处理文档
            try:
                document = StyleAnalyzer(file_path)
                logger.info(f"成功加载文档: {file_path}")
            except Exception as e:
                logger.error(f"加载文档失败: {str(e)}", exc_info=True)
                raise HTTPException(status_code=500, detail=f"无法解析文档: {str(e)}")

            # 存储文档内容元素
            content = []

            # 标题级别映射 - 将数字级别映射为Canvas-Editor使用的枚举字符串
            title_level_mapping = {
                2: 'first',
                3: 'second',
                4: 'third',
                5: 'fourth',
                6: 'fifth',
                7: 'sixth'
            }

            # 获取标题段落索引
            try:
                heading_paragraphs = document.get_heading_paragraphs() if hasattr(document, 'get_heading_paragraphs') else []
                # 修正：根据函数返回的是元组列表(索引, 标题文本, 级别)，正确创建字典
                heading_indices = {index: level for index, _, level in heading_paragraphs} if heading_paragraphs else {}
                logger.debug(f"找到 {len(heading_indices)} 个标题段落")
            except Exception as e:
                logger.warning(f"提取标题段落失败: {str(e)}", exc_info=True)
                heading_indices = {}

            # 遍历所有元素并处理
            element_count = 0
            paragraph_count = 0
            image_count = 0
            table_count = 0

            logger.info(f"文档总计 {len(document.elements)} 个元素")

            for elem_info in document.elements:
                element_count += 1
                element_type = elem_info.get('type')
                element = elem_info.get('element')
                try:
                    if element_type == 'paragraph':
                        paragraph_count += 1

                        # 提取段落样式信息
                        try:
                            style_info = document.get_paragraph_complete_style_info(element)
                            effective_style = style_info.get('effective_style', {})
                            para_props = effective_style.get('paragraph_properties', {})
                        except Exception as e:
                            logger.warning(f"处理段落 {elem_info["index"]} 样式失败: {str(e)}", exc_info=True)
                            effective_style = {}
                            para_props = {}

                        # 获取段落中的所有run
                        try:
                            runs = document.get_runs_from_paragraph(element)
                            logger.debug(f"段落 {elem_info["index"]} 包含 {len(runs)} 个文本运行")
                        except Exception as e:
                            logger.warning(f"获取段落 {elem_info["index"]} 文本运行失败: {str(e)}", exc_info=True)
                            runs = []

                        run_models = []

                        # 处理每个run
                        for run_index, run in enumerate(runs):
                            try:
                                run_style_info = document.get_run_complete_style_info(element, run, run_index)
                                run_props = run_style_info.get('effective_style', {}).get('run_properties', {})
                                print(run_style_info)
                                # 提取run的文本内容
                                run_text = document.get_run_text_from_xml(element, run_index)
                                
                                # 字体信息
                                fonts = run_props.get('font', {})
                                font_name = fonts.get('eastAsia') or fonts.get('ascii') or fonts.get('hAnsi')
                                
                                # 如果字体名称在映射中存在，使用映射后的名称
                                if font_name in FONT_NAME_MAPPING:
                                    font_name = FONT_NAME_MAPPING[font_name]

                                # 字体大小 - 转换中文字号
                                font_size = None
                                if run_props.get('size'):
                                    size_str = run_props.get('size')
                                    # 先检查是否是中文字号（在映射表中）
                                    if size_str in CHINESE_FONT_SIZE_MAPPING:
                                        font_size = CHINESE_FONT_SIZE_MAPPING[size_str]
                                    else:
                                        try:
                                            # Word中字号是磅值的2倍
                                            raw_size = float(size_str)
                                            # 转换为Canvas-Editor使用的字号
                                            # 按比例近似匹配最接近的Canvas-Editor字号
                                            font_size = round(raw_size / 2)
                                        except (ValueError, TypeError):
                                            font_size = 16  # 默认字号

                                # 处理颜色值，确保正确的格式（带有#前缀）
                                color = run_props.get('color')
                                if color and not color.startswith('#'):
                                    # 如果颜色值不是以#开头，添加#前缀
                                    color = f"#{color}"
                                    
                                # 同样处理高亮颜色
                                highlight = run_props.get('highlight')
                             
                                
                                run_model = RunModel(
                                    value=run_text,
                                    font=font_name,
                                    size=font_size,
                                    bold=run_props.get('bold') == 'true',
                                    italic=run_props.get('italic') == 'true',
                                    underline=bool(run_props.get('underline')),
                                    strike=run_props.get('strike') == 'true',
                                    color=color,
                                    highlight=highlight,
                                )
                                run_models.append(run_model)
                            except Exception as e:
                                logger.warning(f"处理段落 {elem_info["index"]} 文本运行 {run_index} 失败: {str(e)}", exc_info=True)
                                # 添加一个空运行，避免完全跳过
                                run_models.append(RunModel(value="[处理错误]"))

                        # 获取缩进信息
                        indentation = para_props.get('indentation', {})
                        indent_left = indentation.get('left')
                        indent_right = indentation.get('right')
                        indent_first_line = indentation.get('firstLine')

                        # 获取间距信息
                        spacing = para_props.get('spacing', {})
                        spacing_before = spacing.get('before')
                        spacing_after = spacing.get('after')
                        line_spacing = spacing.get('line')

                        # 获取对齐方式
                        align_map = {
                            'left': 'left',
                            'right': 'right',
                            'center': 'center',
                            'both': 'justify',
                            'justify': 'justify'
                        }
                        align = align_map.get(para_props.get('alignment'), 'left')

                        # 创建段落模型
                        paragraph_model = {
                            'type': 'paragraph',
                            'value': '',
                            'paragraphId': str(uuid.uuid4()),
                            'valueList': run_models,
                            'rowFlex': align,
                            'indent': float(indent_left) if indent_left else None,
                            'spaceBefore': {
                                'value': float(spacing_before) if spacing_before else 0,
                                'unit': 'pt'
                            },
                            'spaceAfter': {
                                'value': float(spacing_after) if spacing_after else 0,
                                'unit': 'pt'
                            },
                            'lineSpacingMode': 'multiple',
                            'lineSpacingValue': {
                                'value': float(line_spacing) if line_spacing else 1,
                                'unit': 'multiple'
                            }
                        }

                        # 判断是否为标题段落，并映射为枚举字符串格式
                        if elem_info["index"] in heading_indices:
                            numerical_level = heading_indices[elem_info["index"]]
                            # 调整级别值（如果需要），然后映射为字符串枚举
                            adjusted_level = max(2, min(7, numerical_level))  # 确保级别在1-6范围内
                            paragraph_model['type'] = 'title'
                            paragraph_model['level'] = title_level_mapping[adjusted_level]

                        content.append(paragraph_model)

                    # 这里应该添加对其他类型元素的处理，如表格、图片等
                    # 暂时省略，请根据需要补充

                except Exception as e:
                    logger.error(f"处理元素 {element_count} 失败: {str(e)}", exc_info=True)
            
            return {
                "filename": os.path.basename(file_path),
                "content": content,
                "stats": {
                    "total_elements": element_count,
                    "paragraphs": paragraph_count,
                    "images": image_count,
                    "tables": table_count
                }
            }
        
        except Exception as e:
            logger.error(f"文档处理失败: {str(e)}", exc_info=True)
            raise HTTPException(status_code=500, detail=f"文档处理失败: {str(e)}")
