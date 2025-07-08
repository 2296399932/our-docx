import json
import os
from typing import Dict, Any, List, Optional
import uuid
from fastapi import UploadFile, HTTPException
from models.document_models import DocumentResponse, ParagraphModel, RunModel, TableModel, ImageModel, DocElement, DocumentModel, TableCellModel, TableRowModel, Border, VerticalAlign
from util.style_analyzer import StyleAnalyzer
import logging
from datetime import datetime
import traceback
import base64

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
    '44': 29,  # 二号
    '36': 24,  # 小二
    '32': 21,  # 三号
    '30': 20,  # 小三
    '28': 18,  # 四号
    '24': 16,  # 小四

    '21': 14,  # 五号
    '18': 12,  # 小五
    '15': 10,  # 六号
    '13': 8,   # 小六
    '11': 7,   # 七号
    '10': 6,    # 八号
}
indent_num=760/7
line_spacing_num=240
width_num=20  # 表格宽度转换常量
px_ch_width=37 # 1px对应多少ch
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
        try:
            # 确保上传目录存在
            os.makedirs(upload_dir, exist_ok=True)
            logger.info(f"确保上传目录存在: {upload_dir}")

            # 生成唯一文件名避免冲突
            file_extension = os.path.splitext(filename)[1]
            unique_filename = f"{uuid.uuid4()}{file_extension}"
            file_path = os.path.join(upload_dir, unique_filename)
            logger.info(f"将保存文件到: {file_path}")

            # 保存文件
            try:
                with open(file_path, "wb") as file:
                    file.write(file_content)
                logger.info(f"文件保存成功，大小: {len(file_content)} 字节")
            except Exception as e:
                logger.error(f"保存文件失败: {str(e)}", exc_info=True)
                print(f"ERROR: 保存文件失败: {str(e)}")
                print(f"ERROR: 错误详情:\n{traceback.format_exc()}")
                raise HTTPException(status_code=500, detail=f"保存文件失败: {str(e)}")

            # 返回文件上传信息，不进行文档处理
            return DocumentResponse(
                filename=filename,
                status="文件上传成功",
                file_path=file_path,
                upload_time=datetime.now()
            )
        except Exception as e:
            if not isinstance(e, HTTPException):
                logger.error(f"文件上传过程出错: {str(e)}", exc_info=True)
                print(f"ERROR: 文件上传过程出错: {str(e)}")
                print(f"ERROR: 错误详情:\n{traceback.format_exc()}")
                raise HTTPException(status_code=500, detail=f"文件上传失败: {str(e)}")
            raise

    @staticmethod
    def _convert_models_to_dict(obj):
        """将模型对象转换为可序列化的字典

        Args:
            obj: 要转换的对象
            
        Returns:
            转换后的可序列化对象
        """
        if hasattr(obj, 'dict') and callable(getattr(obj, 'dict')):
            # Pydantic模型对象
            return obj.dict()
        elif hasattr(obj, '__dict__'):
            # 一般的类对象
            result = {}
            for key, value in obj.__dict__.items():
                if not key.startswith('_'):  # 排除私有属性
                    result[key] = DocumentService._convert_models_to_dict(value)
            return result
        elif isinstance(obj, list):
            # 列表
            return [DocumentService._convert_models_to_dict(item) for item in obj]
        elif isinstance(obj, dict):
            # 字典
            return {key: DocumentService._convert_models_to_dict(value) for key, value in obj.items()}
        elif isinstance(obj, (str, int, float, bool, type(None))):
            # 基本类型
            return obj
        else:
            # 其他类型，尝试转换为字符串
            try:
                return str(obj)
            except Exception:
                return None

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
                print(f"ERROR: 加载文档失败: {str(e)}")
                print(f"ERROR: 错误详情:\n{traceback.format_exc()}")
                raise HTTPException(status_code=500, detail=f"无法解析文档: {str(e)}")

            # 存储文档内容元素
            content = []

            # 标题级别映射 - 将数字级别映射为Canvas-Editor使用的枚举字符串
            title_level_mapping = {
                2: 'first',
                3: 'second',
                4: 'third',
                5: 'fourth',

            }
            
            # 初始化索引映射数组
            result_json = []

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
            # 上传文档预存的数据标记对应于element的index属性，所以我们需要根据index来获取对应的元素信息
            result_json=[]
            for elem_info in document.elements:
                element_count += 1
                element_type = elem_info.get('type')
                element = elem_info.get('element')
                try:
                    if element_type == 'paragraph':
                        paragraph_count += 1

                        # 检查段落是否包含图片信息
                        if 'image_info' in elem_info:
                            image_info = elem_info['image_info']
                            has_text = image_info.get('has_text', False)

                            # 如果段落不包含文本，仅处理为图片
                            if not has_text:
                                # 图片处理逻辑
                                image_count += 1
                                try:
                                    # 获取图片元数据
                                    image_dimensions = image_info.get('dimensions', [{}])[0] if image_info.get(
                                        'dimensions') else {}
                                    embed_ids = image_info.get('embed_ids', [])
                                    image_descriptions = image_info.get('image_descriptions', [])
                                    wrap_types = image_info.get('wrap_types', [])

                                    # 提取第一个图片的信息(假设只取一个)
                                    embed_id = embed_ids[0] if embed_ids else None
                                    width = image_dimensions.get('width_cm')
                                    height = image_dimensions.get('height_cm')
                                    description = image_descriptions[0].get('description',
                                                                            '') if image_descriptions else ''
                                    original_filename = image_descriptions[0].get('name',
                                                                                  f'image_{image_count}.png') if image_descriptions else f'image_{image_count}.png'
                                    wrap_type = wrap_types[0] if wrap_types else 'inline'

                                    # 生成图片文件名
                                    file_extension = os.path.splitext(original_filename)[1] or '.png'


                                    # 获取图片数据和保存图片
                                    if embed_id:


                                            # 获取图片元组(图片名称, 图片二进制数据)
                                            image_result = document.get_image_by_relation_id(embed_id)
                                            if image_result and isinstance(image_result, tuple) and len(
                                                    image_result) == 2:
                                                image_name, image_data = image_result

                                                if image_data and isinstance(image_data, bytes):
                                                    mime_type = "image/jpeg"  # 默认MIME类型
                                                    if file_extension.lower() in ['.png']:
                                                        mime_type = "image/png"
                                                    elif file_extension.lower() in ['.jpg', '.jpeg']:
                                                        mime_type = "image/jpeg"
                                                    elif file_extension.lower() in ['.gif']:
                                                        mime_type = "image/gif"
                                                    elif file_extension.lower() in ['.bmp']:
                                                        mime_type = "image/bmp"
                                                    elif file_extension.lower() in ['.tiff', '.tif']:
                                                        mime_type = "image/tiff"
                                                    elif file_extension.lower() in ['.svg']:
                                                        mime_type = "image/svg+xml"

                                                    # 编码为Base64
                                                    encoded_image = base64.b64encode(image_data).decode('utf-8')
                                                    data_url = f"data:{mime_type};base64,{encoded_image}"
                                                else:
                                                    data_url = ""
                                            else:
                                                data_url = ""

                                    # 获取环绕方式
                                    img_display_map = {
                                        'inline': 'inline',
                                        'square': 'surround',
                                        'tight': 'surround',
                                        'through': 'surround',
                                        'topAndBottom': 'block',
                                        'none-behind': 'float-bottom',
                                        'none': 'float-top'
                                    }
                                    img_display = img_display_map.get(wrap_type, 'inline')
                                    # 提取段落样式信息
                                    try:
                                        style_info = document.get_paragraph_complete_style_info(element)
                                        effective_style = style_info.get('effective_style', {})
                                        para_props = effective_style.get('paragraph_properties', {})
                                    except Exception as e:
                                        logger.warning(f"处理段落 {elem_info['index']} 样式失败: {str(e)}",
                                                       exc_info=True)

                                        para_props = {}

                                    # 获取段落中的所有run
                                    try:
                                        runs = document.get_runs_from_paragraph(element)
                                        logger.debug(f"段落 {elem_info['index']} 包含 {len(runs)} 个文本运行")
                                    except Exception as e:
                                        logger.warning(f"获取段落 {elem_info['index']} 文本运行失败: {str(e)}",
                                                       exc_info=True)



                                    # 获取间距信息
                                    spacing = para_props.get('spacing', {})
                                    spacing_before = spacing.get('before')
                                    spacing_after = spacing.get('after')
                                    line_spacing = spacing.get('line')
                                    lineRule_spacing = spacing.get('lineRule')

                                    # 获取对齐方式
                                    align_map = {
                                        'left': 'left',
                                        'right': 'right',
                                        'center': 'center',
                                        'both': 'alignment',
                                        'justify': 'justify'
                                    }
                                    align = align_map.get(para_props.get('alignment'), 'left')
                                    print('alignment', para_props.get('alignment'))
                                    # 创建图片模型
                                    image_model = ImageModel(
                                        id=f"img-{image_info['embed_ids'][0]}-{image_count}",

                                        description=description,
                                        width=width*px_ch_width,
                                        height=height* px_ch_width,
                                        value= data_url,
                                        imgDisplay=img_display,
                                        type="image",
                                        rowFlex=align,
                                        rowMargin=float(
                                            line_spacing) / line_spacing_num if line_spacing else None
                                    )

                                    # 添加图片到文档内容

                                    content.append(image_model)

                                except Exception as e:
                                    logger.error(f"处理图片 {elem_info['index']} 失败: {str(e)}", exc_info=True)
                                    result_json.append({'id': f"img-{image_info['embed_ids'][0]}-error", "index": elem_info['index']})
                                    content.append(ImageModel(
                                        id=f"img-{image_info['embed_ids'][0]}-error",
                                        file_name="error.png",
                                        type="image",
                                        description=f"图片处理失败: {str(e)}",
                                        value=""
                                    ))

                            else:
                                    # 段落包含文本和图片，需要处理两者
                                    text_before_image = image_info.get('text_before_image', False)
                                    text_after_image = image_info.get('text_after_image', False)

                                    # 提取段落样式信息
                                    try:
                                        style_info = document.get_paragraph_complete_style_info(element)
                                        effective_style = style_info.get('effective_style', {})
                                        para_props = effective_style.get('paragraph_properties', {})
                                    except Exception as e:
                                        logger.warning(f"处理段落 {elem_info['index']} 样式失败: {str(e)}",
                                                       exc_info=True)
                                        effective_style = {}
                                        para_props = {}

                                    # 获取段落中的所有run
                                    try:
                                        runs = document.get_runs_from_paragraph(element)
                                        logger.debug(f"段落 {elem_info['index']} 包含 {len(runs)} 个文本运行")
                                    except Exception as e:
                                        logger.warning(f"获取段落 {elem_info['index']} 文本运行失败: {str(e)}",
                                                       exc_info=True)
                                        runs = []

                                    run_models = []

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
                                    lineRule_spacing = spacing.get('lineRule')

                                    # 获取对齐方式
                                    align_map = {
                                        'left': 'left',
                                        'right': 'right',
                                        'center': 'center',
                                        'both': 'alignment',
                                        'justify': 'justify'
                                    }
                                    align = align_map.get(para_props.get('alignment'), 'left')

                                    # 处理每个run
                                    for run_index, run in enumerate(runs):
                                        try:
                                            run_style_info = document.get_run_complete_style_info(element, run,
                                                                                                  run_index)
                                            run_props = run_style_info.get('effective_style', {}).get('run_properties',
                                                                                                      {})

                                            # 提取run的文本内容
                                            run_text = document.get_run_text_from_xml(element, run_index)

                                            # 字体信息
                                            fonts = run_props.get('fonts', {})
                                            font_name = fonts.get('eastAsia') or fonts.get('ascii') or fonts.get(
                                                'hAnsi')

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
                                                lineRule=lineRule_spacing,
                                                line=line_spacing,
                                                paragraphId=f"para-{paragraph_count}",
                                                superscript=run_props.get('superscript') == 'true',
                                                subscript=run_props.get('subscript') == 'true',
                                                indent=float(
                                                    indent_first_line) / indent_num if indent_first_line else None,
                                                rowFlex=align,
                                                rowMargin=float(
                                                    line_spacing) / line_spacing_num if line_spacing else None
                                            )
                                            run_models.append(run_model)
                                        except Exception as e:
                                            logger.warning(
                                                f"处理段落 {elem_info['index']} 文本运行 {run_index} 失败: {str(e)}",
                                                exc_info=True)
                                            # 添加一个空运行，避免完全跳过
                                            run_models.append(RunModel(value="[处理错误]"))

                                    # 创建段落模型之前，确保每个段落末尾都有换行符
                                    if run_models and run_models[-1].value != "\u200B":
                                        # 添加换行符作为最后一个run
                                        run_models.append(RunModel(
                                            value="\u200B",  # ZERO常量
                                            paragraphId=f"para-{paragraph_count}"
                                        ))

                                    # 创建段落模型
                                    paragraph_model = {
                                        'type': 'paragraph',
                                        'value': '',
                                        'paragraphId': f"para-{paragraph_count}",  # 使用顺序数字作为ID
                                        'valueList': run_models,
                                        'rowFlex': align,
                                        'rowMargin': float(
                                            line_spacing) / line_spacing_num if line_spacing else None,
                                        'indent': float(indent_first_line) / indent_num if indent_first_line else None,
                                        'lineRule': lineRule_spacing,
                                        'line': line_spacing,

                                    }

                                    # 判断是否为标题段落，并映射为枚举字符串格式
                                    if elem_info["index"] in heading_indices:
                                        numerical_level = heading_indices[elem_info["index"]]
                                        # 调整级别值（如果需要），然后映射为字符串枚举
                                        adjusted_level = max(2, min(7, numerical_level))  # 确保级别在1-6范围内
                                        paragraph_model['type'] = 'title'
                                        paragraph_model['level'] = title_level_mapping[adjusted_level]

                                    # 图片处理逻辑
                                    image_count += 1
                                    try:
                                        # 获取图片元数据
                                        image_dimensions = image_info.get('dimensions', [{}])[0] if image_info.get(
                                            'dimensions') else {}
                                        print("image_dimensions:",image_dimensions)
                                        embed_ids = image_info.get('embed_ids', [])
                                        image_descriptions = image_info.get('image_descriptions', [])
                                        wrap_types = image_info.get('wrap_types', [])

                                        # 提取第一个图片的信息(假设只取一个)
                                        embed_id = embed_ids[0] if embed_ids else None
                                        width = image_dimensions.get('width')
                                        height = image_dimensions.get('height')
                                        description = image_descriptions[0].get('description',
                                                                                '') if image_descriptions else ''
                                        original_filename = image_descriptions[0].get('name',
                                                                                      f'image_{image_count}.png') if image_descriptions else f'image_{image_count}.png'
                                        wrap_type = wrap_types[0] if wrap_types else 'inline'

                                        # 生成图片文件名
                                        file_extension = os.path.splitext(original_filename)[1] or '.png'


                                        # 获取图片数据和保存图片
                                        if embed_id:

                                                # 获取图片元组(图片名称, 图片二进制数据)
                                                image_result = document.get_image_by_relation_id(embed_id)
                                                if image_result and isinstance(image_result, tuple) and len(
                                                        image_result) == 2:
                                                    image_name, image_data = image_result

                                                    if image_data and isinstance(image_data, bytes):
                                                        mime_type = "image/jpeg"  # 默认MIME类型
                                                        if file_extension.lower() in ['.png']:
                                                            mime_type = "image/png"
                                                        elif file_extension.lower() in ['.jpg', '.jpeg']:
                                                            mime_type = "image/jpeg"
                                                        elif file_extension.lower() in ['.gif']:
                                                            mime_type = "image/gif"
                                                        elif file_extension.lower() in ['.bmp']:
                                                            mime_type = "image/bmp"
                                                        elif file_extension.lower() in ['.tiff', '.tif']:
                                                            mime_type = "image/tiff"
                                                        elif file_extension.lower() in ['.svg']:
                                                            mime_type = "image/svg+xml"

                                                        # 编码为Base64
                                                        encoded_image = base64.b64encode(image_data).decode('utf-8')
                                                        data_url = f"data:{mime_type};base64,{encoded_image}"
                                                    else:
                                                        data_url = ""
                                                else:
                                                    data_url = ""

                                        # 获取环绕方式
                                        img_display_map = {
                                            'inline': 'inline',
                                            'square': 'surround',
                                            'tight': 'surround',
                                            'through': 'surround',
                                            'topAndBottom': 'block',
                                            'none-behind': 'float-bottom',
                                            'none': 'float-top'
                                        }
                                        img_display = img_display_map.get(wrap_type, 'inline')
                                        print("width:",width)
                                        # 创建图片模型
                                        image_model = ImageModel(
                                            id=f"img-{image_info['embed_ids'][0]}-{image_count}",
                                            type="image",
                                            description=description,
                                            width=width * px_ch_width,
                                            height=height* px_ch_width,
                                            value=data_url,
                                            imgDisplay=img_display,
                                            rowFlex=align,
                                            rowMargin=float(
                                                line_spacing) / line_spacing_num if line_spacing else None
                                        )

                                        # 根据文本和图片的位置关系决定添加顺序
                                        if text_before_image:
                                            # 先添加段落，再添加图片
                                            result_json.append({'id': paragraph_model['paragraphId'], "index": elem_info['index']})
                                            content.append(paragraph_model)
                                            result_json.append({'id': f"img-{image_info['embed_ids'][0]}-{image_count}", "index": elem_info['index']})
                                            content.append(image_model)
                                        else:
                                            # 先添加图片，再添加段落
                                            result_json.append({'id': f"img-{image_info['embed_ids'][0]}-{image_count}", "index": elem_info['index']})
                                            content.append(image_model)
                                            result_json.append({'id': paragraph_model['paragraphId'], "index": elem_info['index']})
                                            content.append(paragraph_model)

                                    except Exception as e:
                                        logger.error(f"处理图片 {elem_info['index']} 失败: {str(e)}", exc_info=True)
                                        # 出错时只添加段落，图片添加错误占位符
                                        result_json.append({'id': paragraph_model['paragraphId'], "index": elem_info['index']})
                                        content.append(paragraph_model)
                                        result_json.append({'id': f"img-{image_info['embed_ids'][0]}-error", "index": elem_info['index']})
                                        content.append(ImageModel(
                                            id=f"img-{image_info['embed_ids'][0]}-error",
                                            file_name="error.png",
                                            type="image",
                                            description=f"图片处理失败: {str(e)}",
                                            value=""
                                        ))
                        else:
                                    # 普通段落处理(没有图片)
                                    try:
                                        style_info = document.get_paragraph_complete_style_info(element)
                                        effective_style = style_info.get('effective_style', {})
                                        para_props = effective_style.get('paragraph_properties', {})
                                    except Exception as e:
                                        logger.warning(f"处理段落 {elem_info['index']} 样式失败: {str(e)}",
                                                       exc_info=True)
                                        effective_style = {}
                                        para_props = {}

                                    # 获取段落中的所有run
                                    try:
                                        runs = document.get_runs_from_paragraph(element)
                                        logger.debug(f"段落 {elem_info['index']} 包含 {len(runs)} 个文本运行")
                                    except Exception as e:
                                        logger.warning(f"获取段落 {elem_info['index']} 文本运行失败: {str(e)}",
                                                       exc_info=True)
                                        runs = []

                                    run_models = []

                                    # 获取缩进信息
                                    indentation = para_props.get('indentation', {})
                                    indent_left = indentation.get('left')
                                    indent_right = indentation.get('right')
                                    indent_first_line = indentation.get('firstLine')

                                    # 获取间距信息

                                    spacing = para_props.get('spacing', {})
                                    print('alignment', para_props.get('alignment'))
                                    spacing_before = spacing.get('before')
                                    spacing_after = spacing.get('after')
                                    line_spacing = spacing.get('line')
                                    lineRule_spacing = spacing.get('lineRule')

                                    # 获取对齐方式
                                    align_map = {
                                        'left': 'left',
                                        'right': 'right',
                                        'center': 'center',
                                        'both': 'alignment',
                                        'justify': 'justify'
                                    }
                                    align = align_map.get(para_props.get('alignment'), 'left')

                                    # 判断是否为标题段落
                                    is_heading = elem_info["index"] in heading_indices
                                    # 生成适当的ID
                                    element_id = f"title-{uuid.uuid4()}" if is_heading else f"para-{paragraph_count}"
                                    id_field_name = "titleId" if is_heading else "paragraphId"

                                    # 处理每个run
                                    for run_index, run in enumerate(runs):
                                        try:
                                            run_style_info = document.get_run_complete_style_info(element, run,
                                                                                                  run_index)
                                            run_props = run_style_info.get('effective_style', {}).get('run_properties',
                                                                                                      {})

                                            # 提取run的文本内容
                                            run_text = document.get_run_text_from_xml(element, run_index)

                                            # 字体信息
                                            fonts = run_props.get('fonts', {})
                                            font_name = fonts.get('eastAsia') or fonts.get('ascii') or fonts.get(
                                                'hAnsi')

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

                                            # 创建run模型，使用正确的ID字段
                                            run_model_data = {
                                                "value": run_text,
                                                "font": font_name,
                                                "size": font_size,
                                                "bold": run_props.get('bold') == 'true',
                                                "italic": run_props.get('italic') == 'true',
                                                "underline": bool(run_props.get('underline')),
                                                "strike": run_props.get('strike') == 'true',
                                                "color": color,
                                                "highlight": highlight,
                                                "lineRule": lineRule_spacing,
                                                "line": line_spacing,
                                                "superscript": run_props.get('superscript') == 'true',
                                                "subscript": run_props.get('subscript') == 'true',
                                                "indent": float(indent_first_line) / indent_num if indent_first_line else None,
                                                "rowFlex": align,
                                                "rowMargin": float(line_spacing) / line_spacing_num if line_spacing else None
                                            }
                                            
                                            # 根据是否为标题添加正确的ID字段
                                            run_model_data[id_field_name] = element_id
                                            
                                            # 如果是标题，添加level属性
                                            if is_heading:
                                                numerical_level = heading_indices[elem_info["index"]]
                                                adjusted_level = max(2, min(7, numerical_level))
                                                run_model_data["level"] = title_level_mapping[adjusted_level]
                                            
                                            run_models.append(RunModel(**run_model_data))
                                            
                                        except Exception as e:
                                            logger.warning(
                                                f"处理段落 {elem_info['index']} 文本运行 {run_index} 失败: {str(e)}",
                                                exc_info=True)
                                            # 添加一个空运行，避免完全跳过
                                            run_models.append(RunModel(value="[处理错误]"))

                                    # 创建段落模型之前，确保每个段落末尾都有换行符
                                    if run_models and run_models[-1].value != "\u200B":
                                        # 添加换行符作为最后一个run，使用正确的ID字段
                                        end_run_data = {
                                            "value": "\u200B"  # ZERO常量
                                        }
                                        end_run_data[id_field_name] = element_id
                                        run_models.append(RunModel(**end_run_data))

                                    # 创建段落/标题模型
                                    element_model = {
                                        'value': '',
                                        'valueList': run_models,
                                        'rowFlex': align,
                                        'indent': float(indent_first_line) / indent_num if indent_first_line else None,
                                        'lineRule': lineRule_spacing,
                                        'line': line_spacing,
                                        'rowMargin': float(line_spacing) / line_spacing_num if line_spacing else None
                                    }
                                    
                                    # 根据是否为标题设置不同的属性
                                    if is_heading:
                                        numerical_level = heading_indices[elem_info["index"]]
                                        adjusted_level = max(2, min(7, numerical_level))
                                        element_model['type'] = 'title'
                                        element_model['level'] = title_level_mapping[adjusted_level]
                                        element_model['titleId'] = element_id
                                    else:
                                        element_model['type'] = 'paragraph'
                                        element_model['paragraphId'] = element_id

                                    # 添加元素索引映射
                                    element_id = element_model.get('titleId') or element_model.get('paragraphId')
                                    result_json.append({'id': element_id, "index": elem_info['index']})
                                    content.append(element_model)

                    elif element_type == 'table':
                        table_count += 1
                        # 遍历所有表格并处理
                        try:

                            tab_index= document.get_table_index_from_element_index(elem_info["index"])
                            table_style = document.get_table_style(elem_info["index"])

                            # 获取表格尺寸
                            dimensions = document.get_table_dimensions(tab_index)
                            rows_count, cols_count = dimensions if dimensions else (0, 0)

                            # 创建表格行列表
                            table_rows = []
                            
                            # 遍历表格的所有行
                            for row_idx, row_info in enumerate(table_style.get('rows', [])):
                                # 创建单元格列表
                                cells = []
                                
                                # 处理每一行中的单元格
                                for col_idx in range(cols_count):
                                    cell_key = (row_idx, col_idx)
                                    cell_style = table_style.get('cells', {}).get(cell_key, {})
                                    
                                    # 获取单元格内容（段落）
                                    cell_paragraphs = document.get_table_cell_paragraphs(tab_index, row_idx, col_idx)

                                    cell_content = []
                                    
                                    # 处理单元格中的段落
                                    for para in cell_paragraphs:
                                        # 提取段落样式
                                        try:
                                            style_info = document.get_paragraph_complete_style_info(para)
                                            effective_style = style_info.get('effective_style', {})
                                            para_props = effective_style.get('paragraph_properties', {})


                                        except Exception as e:
                                            logger.warning(f"处理段落 样式失败: {str(e)}",
                                                           exc_info=True)
                                            effective_style = {}
                                            para_props = {}

                                            # 获取段落中的所有run
                                        try:
                                            runs = document.get_runs_from_paragraph(para)
                                            logger.debug(f"段落 包含 {len(runs)} 个文本运行")
                                        except Exception as e:
                                            logger.warning(f"获取段落  文本运行失败: {str(e)}",
                                                           exc_info=True)
                                            runs = []

                                        run_models = []

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
                                        lineRule_spacing = spacing.get('lineRule')

                                        # 获取对齐方式
                                        align_map = {
                                            'left': 'left',
                                            'right': 'right',
                                            'center': 'center',
                                            'both': 'alignment',
                                            'justify': 'justify'
                                        }
                                        align = align_map.get(para_props.get('alignment'), 'left')
                                        # 处理每个run
                                        for run_index, run in enumerate(runs):
                                            try:
                                                run_style_info = document.get_run_complete_style_info(para, run,
                                                                                                      run_index)
                                                run_props = run_style_info.get('effective_style', {}).get(
                                                    'run_properties', {})

                                                # 提取run的文本内容
                                                run_text = document.get_run_text_from_xml(para, run_index)

                                                # 字体信息
                                                fonts = run_props.get('fonts', {})
                                                font_name = fonts.get('eastAsia') or fonts.get('ascii') or fonts.get(
                                                    'hAnsi')

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
                                                    lineRule=lineRule_spacing,
                                                    line=line_spacing,
                                                    paragraphId=f"para-{paragraph_count}",
                                                    superscript=run_props.get('superscript') == 'true',
                                                    subscript=run_props.get('subscript') == 'true',
                                                    indent=float(
                                                        indent_first_line) / indent_num if indent_first_line else None,
                                                    rowFlex=align,
                                                    rowMargin=float(
                                                        line_spacing) / line_spacing_num if line_spacing else None

                                                )
                                                run_models.append(run_model)
                                            except Exception as e:
                                                logger.warning(
                                                    f"处理段落 {elem_info["index"]} 文本运行 {run_index} 失败: {str(e)}",
                                                    exc_info=True)
                                                # 添加一个空运行，避免完全跳过
                                                run_models.append(RunModel(value="[处理错误]"))

                                        # 创建段落模型之前，确保每个段落末尾都有换行符
                                        if run_models and run_models[-1].value != "\u200B":
                                            # 添加换行符作为最后一个run
                                            run_models.append(RunModel(
                                                value="\u200B",  # ZERO常量
                                                paragraphId=f"para-{paragraph_count}"
                                            ))

                                        # 创建段落模型
                                        paragraph_model = {
                                            'type': 'paragraph',
                                            'value': '',
                                            'paragraphId': f"para-{paragraph_count}",  # 使用顺序数字作为ID
                                            'valueList': run_models,
                                            'rowFlex': align,
                                            'indent': float(
                                                indent_first_line) / indent_num if indent_first_line else None,
                                            'lineRule': lineRule_spacing,
                                            'line': line_spacing
                                        }

                                        # 判断是否为标题段落，并映射为枚举字符串格式
                                        if elem_info["index"] in heading_indices:
                                            numerical_level = heading_indices[elem_info["index"]]
                                            # 调整级别值（如果需要），然后映射为字符串枚举
                                            adjusted_level = max(2, min(5, numerical_level))  # 确保级别在1-4范围内
                                            paragraph_model['type'] = 'title'
                                            paragraph_model['level'] = title_level_mapping[adjusted_level]

                                        cell_content.append(paragraph_model)

                                    # 创建单元格模型
                                    vertical_align = cell_style.get('vertical_align', 'center')
                                    if vertical_align is None:
                                        vertical_align = 'center'  # 默认值
                                    
                                    # 确保verticalAlign是有效的枚举值
                                    try:
                                        # 检查是否是有效的VerticalAlign枚举值
                                        valid_values = set(item.value for item in VerticalAlign)
                                        if vertical_align not in valid_values:
                                            vertical_align = 'center'  # 如果不是有效值，使用默认值
                                        vertical_align_enum = VerticalAlign(vertical_align)
                                    except ValueError:
                                        # 如果转换失败，使用默认值
                                        vertical_align_enum = VerticalAlign.CENTER
                                    
                                    # 处理边框
                                    borders = []
                                    cell_borders = cell_style.get('borders', {})
                                    has_borders = False
                                    
                                    # 检查是否有边框
                                    for border_key, border_data in cell_borders.items():
                                        if border_data and border_data.get('val') != 'nil':
                                            has_borders = True
                                            if border_key == 'top':
                                                borders.append('top')
                                            elif border_key == 'right':
                                                borders.append('right')
                                            elif border_key == 'bottom':
                                                borders.append('bottom')
                                            elif border_key == 'left':
                                                borders.append('left')
                                    
                                    # 处理背景颜色，确保有#前缀
                                    background_color = cell_style.get('shading', {}).get('fill')
                                    if background_color and not background_color.startswith('#'):
                                        background_color = f"#{background_color}"
                                    
                                    cell = TableCellModel(
                                        rowspan=cell_style.get('rowspan', 1),
                                        colspan=cell_style.get('colspan', 1),
                                        verticalAlign=vertical_align_enum,
                                        backgroundColor=background_color,
                                        borderTypes=borders,
                                        value=cell_content
                                    )
                                    cells.append(cell)
                                
                                # 创建行模型
                                height_value = row_info.get('height', {}).get('value')
                                # min_height = float(height_value) if height_value else None
                                
                                row = TableRowModel(
                                    # minHeight=min_height,
                                    tdList=cells
                                )
                                table_rows.append(row)
                            
                            # 创建表格模型
                            table_width = None
                            width_info = table_style.get('width', {})
                            if width_info.get('value') and width_info.get('type') == 'dxa':
                                try:
                                    table_width = float(width_info.get('value')) / width_num
                                except (ValueError, TypeError):
                                    pass
                            
                            # 处理表格边框类型和颜色
                            border_type = None
                            border_color = None
                            
                            # 检查表格边框
                            table_borders = table_style.get('borders', {})
                            has_border = False
                            border_sides = []
                            
                            # 检查表格边框情况 - 只有val不为"none"时才计入有效边框
                            for border_key, border_data in table_borders.items():
                                if border_data and border_data.get('val') and border_data.get('val') != 'nil' and border_data.get('val') != 'none':
                                    has_border = True
                                    if border_key == 'top':
                                        border_sides.append('top')
                                    elif border_key == 'right':
                                        border_sides.append('right')
                                    elif border_key == 'bottom':
                                        border_sides.append('bottom')
                                    elif border_key == 'left':
                                        border_sides.append('left')
                                    elif border_key == 'insideH':
                                        border_sides.append('insideH')
                                    elif border_key == 'insideV':
                                        border_sides.append('insideV')
                            
                            # 根据边框情况决定边框类型
                            # if table_style.get('is_three_line_table'):
                            #     border_type = ['three-line']
                            if not has_border:
                                border_type = 'empty'
                            elif len(border_sides) == 6:  # 所有边框都存在
                                border_type = 'all'
                            elif 'top' in border_sides and 'right' in border_sides and 'bottom' in border_sides and 'left' in border_sides:
                                if 'insideH' not in border_sides and 'insideV' not in border_sides:
                                    border_type ='external'
                                else:
                                    border_type = 'all'
                            elif 'insideH' in border_sides or 'insideV' in border_sides:
                                border_type = 'internal'
                            else:
                                border_type = 'external'
                            
                            # 检查是否为虚线边框
                            for border_key, border_data in table_borders.items():
                                if border_data and border_data.get('val') == 'dashed':
                                    border_type = 'dash'
                                    break
                            
                            # 获取表格边框颜色
                            for border_key, border_data in table_borders.items():
                                if border_data and border_data.get('color'):
                                    border_color = f"#{border_data.get('color')}"
                                    break
                            
                            # 创建列组
                            colgroup = []
                            grid_widths = table_style.get('grid', [])
                            for width in grid_widths:
                                try:
                                    col_width = float(width) / width_num
                                    colgroup.append({'width': col_width})
                                except (ValueError, TypeError):
                                    colgroup.append({'width': 100})  # 默认宽度
                            
                            table_model = TableModel(
                                id=f"table-{table_count}",
                                width=table_width,
                                trList=table_rows,
                                borderType=border_type,
                                colgroup=colgroup,
                                borderColor=border_color,
                                type="table"
                            )
                            
                            # 添加表格到文档内容
                            result_json.append({'id': table_model.id, "index": elem_info['index']})
                            content.append(table_model)
                            
                        except Exception as e:
                            logger.error(f"处理表格 {elem_info['index']} 失败: {str(e)}", exc_info=True)
                            # 添加一个空表格，避免完全跳过
                            error_table_id = f"table-{table_count}-error"
                            result_json.append({'id': error_table_id, "index": elem_info['index']})
                            content.append(TableModel(id=error_table_id))

                    # 这里应该添加对其他类型元素的处理，如表格、图片等
                    # 暂时省略，请根据需要补充

                except Exception as e:
                    logger.error(f"处理元素 {element_count} 失败: {str(e)}", exc_info=True)
       
            # 更新图片计数和总元素计数
            file_path_json = f'{file_path[:-5]}.json'
            # 保存原始数据
            with open(file_path_json, 'w', encoding='utf-8') as f:
                # 转换为可序列化的字典
                serializable_content = DocumentService._convert_models_to_dict(content)
                json.dump(serializable_content, f, ensure_ascii=False, indent=4)
            
            # 保存索引映射关系
            index_mapping_path = f"{file_path[:-5]}_index_mapping.json"
            with open(index_mapping_path, 'w', encoding='utf-8') as f:
                json.dump(result_json, f, ensure_ascii=False, indent=4)
                
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

    @staticmethod
    async def export_document(content: Dict[str, Any], format_type: str = "docx", output_dir: str = "output") -> str:
        """导出文档内容到指定格式

        Args:
            content: 文档内容数据
            format_type: 导出格式，默认为"docx"
            output_dir: 输出目录

        Returns:
            输出文件的路径
        """
        try:
            logger.info(f"开始导出文档，格式: {format_type}")
            
            # 确保输出目录存在
            os.makedirs(output_dir, exist_ok=True)
            
            # 生成唯一文件名
            file_id = uuid.uuid4()
            
            if format_type.lower() == "docx":
                # TODO: 实现docx格式导出
                # 目前先保存为JSON，后续可以实现转换为DOCX的功能
                output_path = os.path.join(output_dir, f"{file_id}.json")
                
                # 转换为可序列化的字典
                serializable_content = DocumentService._convert_models_to_dict(content)
                with open(output_path, 'w', encoding='utf-8') as f:
                    json.dump(serializable_content, f, ensure_ascii=False, indent=4)
                    
                logger.info(f"文档导出成功: {output_path}")
                return output_path
            else:
                # 其他格式导出为JSON
                output_path = os.path.join(output_dir, f"{file_id}.json")
                
                # 转换为可序列化的字典
                serializable_content = DocumentService._convert_models_to_dict(content)
                with open(output_path, 'w', encoding='utf-8') as f:
                    json.dump(serializable_content, f, ensure_ascii=False, indent=4)
                    
                logger.info(f"文档导出成功: {output_path}")
                return output_path
        except Exception as e:
            logger.error(f"导出文档失败: {str(e)}", exc_info=True)
            print(f"ERROR: 导出文档失败: {str(e)}")
            print(f"ERROR: 错误详情:\n{traceback.format_exc()}")
            raise HTTPException(status_code=500, detail=f"导出文档失败: {str(e)}")
