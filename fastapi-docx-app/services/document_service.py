import json
import os
from typing import Dict, Any, List, Optional, Tuple
import uuid
from fastapi import UploadFile, HTTPException
from models.document_models import DocumentResponse, ParagraphModel, RunModel, TableModel, ImageModel, DocElement, DocumentModel, TableCellModel, TableRowModel, Border, VerticalAlign
from util.style_analyzer import StyleAnalyzer
import logging
from datetime import datetime
import traceback
import base64
import shutil

# 配置日志
logger = logging.getLogger(__name__)

# API服务器基础URL配置
# 可以根据环境变量或配置文件进行设置
API_BASE_URL = "http://localhost:8000"  # 默认值

# 图片服务器域名配置
# 如果图片需要通过不同的域名访问，可以单独配置
IMAGE_SERVER_URL = None  # 默认与API_BASE_URL相同

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
    def set_api_base_url(base_url: str):
        """设置API基础URL
        
        Args:
            base_url: API服务器的基础URL
        """
        global API_BASE_URL
        API_BASE_URL = base_url
        logger.info(f"API基础URL已设置为: {API_BASE_URL}")
    
    @staticmethod
    def set_image_server_url(image_url: str):
        """设置图片服务器URL
        
        Args:
            image_url: 图片服务器的URL
        """
        global IMAGE_SERVER_URL
        IMAGE_SERVER_URL = image_url
        logger.info(f"图片服务器URL已设置为: {IMAGE_SERVER_URL}")

    @staticmethod
    def get_image_base_url():
        """获取用于图片的基础URL
        
        Returns:
            str: 图片服务器基础URL
        """
        return IMAGE_SERVER_URL if IMAGE_SERVER_URL else API_BASE_URL

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
    def save_image_to_file(image_data: bytes, file_extension: str = '.png', images_dir: str = "static/images") -> Tuple[str, str]:
        """将图片数据保存到文件并返回URL路径

        Args:
            image_data: 图片二进制数据
            file_extension: 图片文件扩展名
            images_dir: 图片保存目录

        Returns:
            Tuple[str, str]: URL路径和文件系统路径
        """
        try:
            # 确保图片目录存在
            os.makedirs(images_dir, exist_ok=True)
            
            # 生成唯一文件名
            image_filename = f"{uuid.uuid4()}{file_extension}"
            image_path = os.path.join(images_dir, image_filename)
            
            # 保存图片文件
            with open(image_path, 'wb') as f:
                f.write(image_data)
            
            # 构建完整URL路径，包含图片服务器地址
            image_base_url = DocumentService.get_image_base_url()
            url_path = f"{image_base_url}/images/{image_filename}"
            
            logger.info(f"图片已保存到: {image_path}, URL: {url_path}")
            return url_path, image_path
        except Exception as e:
            logger.error(f"保存图片失败: {str(e)}", exc_info=True)
            print(f"ERROR: 保存图片失败: {str(e)}")
            print(f"ERROR: 错误详情:\n{traceback.format_exc()}")
            return "", ""


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
            # 创建图片保存目录
            images_dir = os.path.join("static", "images")
            os.makedirs(images_dir, exist_ok=True)
            
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
                                                # 保存图片到文件并获取URL
                                                img_url, img_path = DocumentService.save_image_to_file(
                                                    image_data, 
                                                    file_extension, 
                                                    images_dir
                                                )
                                                # 如果保存失败，则使用空字符串
                                                data_url = img_url if img_url else ""
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
                                        
                                        # 清理宽度和高度值

                                        
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
                                                        # 保存图片到文件并获取URL
                                                        img_url, img_path = DocumentService.save_image_to_file(
                                                            image_data, 
                                                            file_extension, 
                                                            images_dir
                                                        )
                                                        # 如果保存失败，则使用空字符串
                                                        data_url = img_url if img_url else ""
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
    async def export_document(content: Dict[str, Any], format_type: str = "docx", original_file_path: str = None,
                              output_dir: str = "output") -> str:
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
            print(f'original_file_path:{original_file_path}')
            # 生成唯一文件名
            file_id = uuid.uuid4()

            if format_type.lower() == "docx" and original_file_path:
                # 比较并标记差异
                main_content = content.get('main', [])
                marked_content = compare_and_merge_json_for_export(main_content, original_file_path)

                print(f'marked_content:{marked_content}')

                # 创建原始文档的StyleAnalyzer对象副本
                try:
                    # 复制原始docx文件到输出目录
                    output_path = os.path.join(output_dir, f"{file_id}.docx")



                    # 根据marked_content处理docx文档

                    update_document_content(marked_content, original_file_path,output_path)

                    logger.info(f"已创建原始文档StyleAnalyzer副本并准备进行编辑，输出路径: {output_path}")
                except Exception as e:
                    logger.error(f"创建原始文档StyleAnalyzer副本失败: {str(e)}", exc_info=True)
                    print(f"ERROR: 创建原始文档StyleAnalyzer副本失败: {str(e)}")

                # 导出JSON以便调试
                debug_path = os.path.join(output_dir, f"{file_id}_marked.json")
                with open(debug_path, 'w', encoding='utf-8') as f:
                    serializable_content = DocumentService._convert_models_to_dict(content)
                    json.dump(serializable_content, f, ensure_ascii=False, indent=4)

                return output_path
            else:
                # 使用模板文档，如果没有原始文件路径
                main_content = content.get('main', [])
                print(111)
                # 使用static/temp/1.docx作为模板
                template_path = os.path.join("static", "temp", "1.docx")

                # 检查模板是否存在
                if os.path.exists(template_path):
                    try:
                        # 复制模板文件到输出目录
                        output_path = os.path.join(output_dir, f"{file_id}.docx")
                        shutil.copy2(template_path, output_path)

                        # 创建一个空的索引映射，因为没有原始文档
                        dummy_index_mapping = []

                        # 将内容当作全新内容处理
                        for idx, item in enumerate(main_content):






                            # 添加标记，表示这是新增的内容
                            item['__diff_status'] = 'added'
                            item['__original_index'] = idx

                            # 为每个元素生成一个ID并添加到映射
                            if 'id' not in item and 'paragraphId' not in item:
                                item_id = f"auto-{idx}"
                                if item.get('type') == 'paragraph':
                                    item['paragraphId'] = item_id
                                else:
                                    item['id'] = item_id

                            # 记录映射关系
                            dummy_index_mapping.append({
                                'id': item.get('id') or item.get('paragraphId'),
                                'index': idx
                            })

                        # 将映射保存到临时文件
                        temp_mapping_path = os.path.join(output_dir, f"{file_id}_index_mapping.json")
                        with open(temp_mapping_path, 'w', encoding='utf-8') as f:
                            json.dump(dummy_index_mapping, f, ensure_ascii=False, indent=4)

                        # 调用update_document_content处理内容
                        update_document_content(main_content, output_path,output_path)

                        logger.info(f"已使用模板文档创建新文档，输出路径: {output_path}")
                    except Exception as e:
                        logger.error(f"使用模板文档失败: {str(e)}", exc_info=True)
                        print(f"ERROR: 使用模板文档失败: {str(e)}")

                        # 失败时导出为JSON
                        output_path = os.path.join(output_dir, f"{file_id}.json")
                        serializable_content = DocumentService._convert_models_to_dict(content)
                        with open(output_path, 'w', encoding='utf-8') as f:
                            json.dump(serializable_content, f, ensure_ascii=False, indent=4)
                else:
                    logger.error(f"模板文档不存在: {template_path}")
                    # 导出为JSON作为备选方案
                    output_path = os.path.join(output_dir, f"{file_id}.json")
                    serializable_content = DocumentService._convert_models_to_dict(content)
                    with open(output_path, 'w', encoding='utf-8') as f:
                        json.dump(serializable_content, f, ensure_ascii=False, indent=4)
                    
                return output_path
        except Exception as e:
            logger.error(f"导出文档失败: {str(e)}", exc_info=True)
            print(f"ERROR: 导出文档失败: {str(e)}")
            print(f"ERROR: 错误详情:\n{traceback.format_exc()}")
            raise HTTPException(status_code=500, detail=f"导出文档失败: {str(e)}")


def compare_and_merge_json_for_export(edited_content, original_file_path):
    """比较修改后的内容与原始JSON，并标记差异以便导出到DOCX

    Args:
        edited_content: 修改后的内容(main数组中的元素)
        original_file_path: 原始文件路径，用于获取index_mapping和原始JSON

    Returns:
        包含差异标记的内容列表
    """
    # 1. 加载原始JSON文件
    original_json_path = f"{original_file_path[:-5]}.json"
    with open(original_json_path, 'r', encoding='utf-8') as f:
        original_content = json.load(f)

    # 2. 加载索引映射文件
    index_mapping_path = f"{original_file_path[:-5]}_index_mapping.json"
    with open(index_mapping_path, 'r', encoding='utf-8') as f:
        index_mapping = json.load(f)

    # 3. 处理编辑后内容中的空白元素
    for item in edited_content:
        if item.get('type') in ['paragraph', 'title'] and 'valueList' in item and item['valueList']:
            # 检查第一个元素是否为空白元素
            first_elem = item['valueList'][0]
            if first_elem.get('value') and first_elem.get('value').strip() == '':
                # 删除该空白元素
                item['valueList'].pop(0)

    # 4. 创建ID到索引的映射字典，考虑多种ID字段名
    id_to_index = {}
    for item in index_mapping:
        # 检查各种可能的ID字段
        item_id = item.get('id') or item.get('paragraphId') or item.get('titleId')
        if item_id and 'index' in item:
            id_to_index[item_id] = item['index']

    # 5. 创建原始内容ID到元素的映射字典
    original_id_map = {}
    for item in original_content:
        item_id = item.get('id') or item.get('paragraphId') or item.get('titleId')
        if item_id:
            original_id_map[item_id] = item

    # 6. 创建编辑内容ID到元素的映射字典
    edited_id_map = {}
    for item in edited_content:
        item_id = item.get('id') or item.get('paragraphId') or item.get('titleId')
        if item_id:
            edited_id_map[item_id] = item

    # 7. 标记差异
    result = []
    
    # 处理编辑后内容中的元素（新增和修改）
    for idx, item in enumerate(edited_content):
        item_id = item.get('id') or item.get('paragraphId') or item.get('titleId')
        # 跳过没有ID的元素
        if not item_id:
            continue
        index_=id_to_index.get(item_id)
        # 检查元素是否在原始内容中
        if item_id in original_id_map:
            # ID存在，检查内容是否相同
            original_item = original_id_map[item_id]
            if not deep_compare_elements(item, original_item):
                # 内容不同，标记为修改
                item['__diff_status'] = 'modified'
                item['__original_index'] = index_
        else:
            # ID不存在，标记为新增
            item['__diff_status'] = 'added'
            # 尝试找到插入位置

            main_index, sub_index = find_insertion_index(idx, id_to_index, edited_content)

            # 创建复合索引，如 5.1, 5.2, 5.3 等
            # 这样即使多个新元素插入到同一位置，也能保持它们之间的顺序
            item['__original_index'] = main_index + sub_index * 0.001

        result.append(item)
    
    # 处理原始内容中存在但编辑后内容中不存在的元素（删除）
    for item_id, original_item in original_id_map.items():
        if item_id not in edited_id_map:
            # 复制原始元素，避免修改原始数据
            deleted_item = original_item.copy() if isinstance(original_item, dict) else original_item
            # 标记为删除
            deleted_item['__diff_status'] = 'deleted'
            deleted_item['__original_index'] = id_to_index.get(item_id)
            result.append(deleted_item)
    print(result)
    # 7. 按照原始索引排序
    result.sort(key=lambda x: x.get('__original_index', 999999))

    return result


def deep_compare_elements(elem1, elem2):
    """深度比较两个元素是否相同，根据元素类型使用不同的比较策略
    
    Args:
        elem1: 第一个元素
        elem2: 第二个元素
        
    Returns:
        bool: 两个元素是否相同
    """
    # 先检查类型是否相同
    if elem1.get('type') != elem2.get('type'):
        return False
        
    element_type = elem1.get('type')
    
    # 根据元素类型选择不同的比较策略
    if element_type == 'paragraph' or element_type == 'title':
        return compare_paragraph_elements(elem1, elem2)
    elif element_type == 'image':
        return compare_image_elements(elem1, elem2)
    elif element_type == 'table':
        return compare_table_elements(elem1, elem2)
    else:
        # 对于未知类型，使用通用比较
        return compare_generic_elements(elem1, elem2)

def compare_paragraph_elements(elem1, elem2):
    """比较段落元素"""
    # 比较基本属性
    basic_fields = ['value', 'rowFlex', 'indent', 'lineRule', 'line', 'rowMargin']
    for field in basic_fields:
        if field in elem1 or field in elem2:
            # 对数值型字段进行转换后比较
            val1 = elem1.get(field)
            val2 = elem2.get(field)
            
            # 数值类型转换 - 尝试将字符串转为数值进行比较
            if isinstance(val1, (int, float)) or isinstance(val2, (int, float)):
                try:
                    if isinstance(val1, str) and val1.replace('.', '').isdigit():
                        val1 = float(val1)
                    if isinstance(val2, str) and val2.replace('.', '').isdigit():
                        val2 = float(val2)
                except:
                    pass  # 转换失败时保持原值
                
                # 都是数值时尝试浮点数比较
                if isinstance(val1, (int, float)) and isinstance(val2, (int, float)):
                    if abs(float(val1) - float(val2)) > 0.0001:  # 允许小误差
                        return False
                    continue  # 数值相近，视为相同
            
            # 对于非数值字段，如果两个都是None或空字符串或False，视为相同
            if (val1 is None or val1 == '' or val1 is False) and (val2 is None or val2 == '' or val2 is False):
                continue
                
            # 其他情况，标准比较    
            if val1 != val2:
                return False
    
    # 比较valueList
    if 'valueList' in elem1 and 'valueList' in elem2:
        # 长度不同直接返回False
        if len(elem1['valueList']) != len(elem2['valueList']):
            return False
        
        # 逐一比较valueList中的元素
        for i in range(len(elem1['valueList'])):
            val1 = elem1['valueList'][i]
            val2 = elem2['valueList'][i]
            
            # 如果是valueList的最后一个元素，且值为换行符或零宽空格，则完全跳过该元素比较
            if i == len(elem1['valueList']) - 1 and (
                (val1.get('value') == '\n' or val1.get('value') == '\u200B') and
                (val2.get('value') == '\n' or val2.get('value') == '\u200B')
            ):
                continue  # 跳过最后一个元素的所有比较
            
            # 对非最后元素或非换行符/零宽空格的元素进行详细比较
            run_fields = ['value', 'font', 'size', 'bold', 'italic', 'underline', 
                         'strike', 'color', 'highlight', 'rowFlex', 'indent']
            for field in run_fields:
                # 特殊处理换行符和零宽空格的等价比较
                if field == 'value' and (val1.get(field) == '\n' or val1.get(field) == '\u200B') and (val2.get(field) == '\n' or val2.get(field) == '\u200B'):
                    continue  # 视为相同，继续下一个属性比较
                
                # 获取两边的值
                field_val1 = val1.get(field)
                field_val2 = val2.get(field)
                
                # 数值类型转换处理
                if field == 'size' or field == 'indent' or field == 'rowMargin' or field == 'line':
                    try:
                        # 尝试将字符串转为数值
                        if isinstance(field_val1, str) and field_val1.replace('.', '').isdigit():
                            field_val1 = float(field_val1)
                        if isinstance(field_val2, str) and field_val2.replace('.', '').isdigit():
                            field_val2 = float(field_val2)
                    except:
                        pass  # 转换失败时保持原值
                    
                    # 如果都是数值类型，进行浮点数比较
                    if isinstance(field_val1, (int, float)) and isinstance(field_val2, (int, float)):
                        if abs(float(field_val1) - float(field_val2)) > 0.0001:  # 允许小误差
                            return False
                        continue  # 数值相近，视为相同
                
                # 处理缺失字段与False/None的等价性
                # strike, superscript, subscript字段如果一边是缺失(None)一边是False，视为相同
                if field in ['strike', 'superscript', 'subscript']:
                    if (field_val1 is None or field_val1 is False) and (field_val2 is None or field_val2 is False):
                        continue
                
                # 其他情况的标准比较
                if field_val1 != field_val2:
                    return False
    
    return True

def compare_image_elements(elem1, elem2):
    """比较图片元素"""
    image_fields = ['value', 'width', 'height', 'imgDisplay', 'rowFlex', 'rowMargin']
    for field in image_fields:
        if field in elem1 or field in elem2:
            if elem1.get(field) != elem2.get(field):
                return False
    return True

def compare_table_elements(elem1, elem2):
    """比较表格元素"""
    # 比较基本属性
    table_fields = ['borderType', 'borderColor', 'width']
    for field in table_fields:
        if field in elem1 or field in elem2:
            if elem1.get(field) != elem2.get(field):
                return False
    
    # 比较行数
    if len(elem1.get('trList', [])) != len(elem2.get('trList', [])):
        return False
    
    # 逐行比较
    for i in range(len(elem1.get('trList', []))):
        tr1 = elem1['trList'][i]
        tr2 = elem2['trList'][i]
        
        # 比较单元格数量
        if len(tr1.get('tdList', [])) != len(tr2.get('tdList', [])):
            return False
        
        # 逐单元格比较
        for j in range(len(tr1.get('tdList', []))):
            td1 = tr1['tdList'][j]
            td2 = tr2['tdList'][j]
            
            # 比较单元格属性
            if td1.get('colspan') != td2.get('colspan') or td1.get('rowspan') != td2.get('rowspan'):
                return False
                
            if td1.get('verticalAlign') != td2.get('verticalAlign'):
                return False
                
            if td1.get('backgroundColor') != td2.get('backgroundColor'):
                return False
            
            # 比较单元格内容 (递归比较)
            cell_content1 = td1.get('value', [])
            cell_content2 = td2.get('value', [])
            
            if len(cell_content1) != len(cell_content2):
                return False
                
            for k in range(len(cell_content1)):
                if not deep_compare_elements(cell_content1[k], cell_content2[k]):
                    return False
    
    return True

def compare_generic_elements(elem1, elem2):
    """通用元素比较逻辑"""
    # 定义需要比较的关键字段
    key_fields = ['type', 'value', 'width', 'height', 'rowFlex', 'lineRule', 'indent', 'line']

    for field in key_fields:
        if field in elem1 or field in elem2:
            if elem1.get(field) != elem2.get(field):
                return False

    # 特殊处理valueList等数组类型字段
    if 'valueList' in elem1 and 'valueList' in elem2:
        if len(elem1['valueList']) != len(elem2['valueList']):
            return False
        for i in range(len(elem1['valueList'])):
            if not deep_compare_elements(elem1['valueList'][i], elem2['valueList'][i]):
                return False

    return True


def find_insertion_index(idx, id_to_index, edited_content):
    """为新元素找到合适的插入位置

    通过向前查找，直到找到一个在原始内容中存在的元素

    Args:
        idx: 当前元素在edited_content中的索引
        id_to_index: ID到索引的映射字典
        edited_content: 编辑后的内容列表

    Returns:
        合适的插入索引位置和子序号
    """
    # 向前查找，直到找到一个在原始内容中存在的元素
    found_index = None
    sub_index = 0

    # 从当前位置向前遍历
    for i in range(idx - 1, -1, -1):
        if i >= 0:
            prev_item = edited_content[i]
            prev_id = prev_item.get('id') or prev_item.get('paragraphId') or prev_item.get('titleId')

            if prev_id and prev_id in id_to_index:
                # 找到了在原始内容中存在的元素
                found_index = id_to_index[prev_id]

                # 检查是否已有其他新增元素依赖这个原始元素
                for j in range(i + 1, idx):
                    if edited_content[j].get('__diff_status') == 'added':
                        sub_index += 1

                break

    # 如果没有找到，则放在开头
    if found_index is None:
        return 0, sub_index

    # 创建一个复合索引，确保多个新增元素按照正确顺序排列
    # 主索引.子索引格式，例如 5.1, 5.2, 5.3 等
    return found_index, sub_index


def convert_editor_to_docx_paragraph(paragraph_data):
    """将Canvas-Editor格式的段落数据转换为DOCX需要的格式
    
    Args:
        paragraph_data: Canvas-Editor格式的段落数据
        
    Returns:
        dict: DOCX格式的段落样式属性
    """
    style_properties = {}
    
    # 转换对齐方式 - 从Canvas-Editor格式转回DOCX格式
    if 'rowFlex' in paragraph_data:
        reverse_align_map = {
            'left': 'left',
            'right': 'right',
            'center': 'center',
            'alignment': 'both',
            'justify': 'justify'
        }
        alignment = reverse_align_map.get(paragraph_data['rowFlex'])
        if alignment:
            style_properties['alignment'] = alignment
    
    # 转换缩进 - 从Canvas-Editor格式转回DOCX格式
    if 'indent' in paragraph_data and paragraph_data['indent'] is not None:
        # 将Canvas-Editor的缩进值转换回DOCX格式 (乘以转换常数)
        indentation = {
            'firstLine': int(paragraph_data['indent'] * indent_num)  # 使用全局转换常数
        }
        style_properties['indentation'] = indentation
    
    # 转换行距和间距 - 从Canvas-Editor格式转回DOCX格式
    spacing = {}
    
    if 'lineRule' in paragraph_data and paragraph_data['lineRule'] is not None:
        spacing['lineRule'] = paragraph_data['lineRule']
        
    if 'line' in paragraph_data and paragraph_data['line'] is not None:
        spacing['line'] = paragraph_data['line']
        
    if 'rowMargin' in paragraph_data and paragraph_data['rowMargin'] is not None:
        # 将Canvas-Editor的行间距转换回DOCX格式 (乘以转换常数)
        line_value = int(paragraph_data['rowMargin'] * line_spacing_num)  # 使用全局转换常数
        spacing['line'] = line_value
    
    if spacing:
        style_properties['spacing'] = spacing
    
    # 获取差异状态和原始索引
    if 'diff_status' in paragraph_data:
        style_properties['diff_status'] = paragraph_data['diff_status']
        
    if 'original_index' in paragraph_data:
        style_properties['original_index'] = paragraph_data['original_index']
    
    return style_properties

def convert_editor_to_docx_run(run_data):
    """将Canvas-Editor格式的Run数据转换为DOCX需要的格式
    
    Args:
        run_data: Canvas-Editor格式的Run数据
        
    Returns:
        dict: DOCX格式的Run样式属性
    """
    style_properties = {}
    
    # 设置文本内容
    if 'value' in run_data:
        style_properties['text'] = run_data['value']
    
    # 转换字体
    if 'font' in run_data and run_data['font']:
        # 需要将单个字体名称转换为字体字典
        font_name = run_data['font']
        # 反向查找原始字体名称
        original_font = None
        for orig, mapped in FONT_NAME_MAPPING.items():
            if mapped == font_name:
                original_font = orig
                break
        
        if not original_font:
            original_font = font_name  # 如果没找到映射，使用原始名称
        
        # 创建字体字典
        style_properties['fonts'] = {
            'eastAsia': original_font
        }
    
    # 转换字体大小
    if 'size' in run_data and run_data['size'] is not None:
        size_value = run_data['size']
        # 反向查找中文字号
        found_chinese_size = False
        for chinese_size, editor_size in CHINESE_FONT_SIZE_MAPPING.items():
            if abs(editor_size - size_value) < 0.1:  # 允许小误差
                style_properties['size'] = chinese_size
                found_chinese_size = True
                break
        
        if not found_chinese_size:
            # 如果不是标准中文字号，转换为Word磅值的2倍
            style_properties['size'] = str(int(size_value * 2))
    
    # 转换粗体
    if 'bold' in run_data:
        style_properties['bold'] = 'true' if run_data['bold'] else 'false'
    
    # 转换斜体
    if 'italic' in run_data:
        style_properties['italic'] = 'true' if run_data['italic'] else 'false'
    
    # 转换下划线
    if 'underline' in run_data and run_data['underline']:
        style_properties['underline'] = 'single'  # 可以根据需要设置其他下划线类型
    
    # 转换删除线
    if 'strike' in run_data:
        style_properties['strike'] = 'true' if run_data['strike'] else 'false'
    
    # 转换颜色 - 确保颜色值没有#前缀
    if 'color' in run_data and run_data['color']:
        color = run_data['color']
        if color.startswith('#'):
            color = color[1:]  # 移除#前缀
        style_properties['color'] = color
    
    # 转换高亮颜色
    if 'highlight' in run_data and run_data['highlight']:
        highlight = run_data['highlight']
        if highlight.startswith('#'):
            highlight = highlight[1:]  # 移除#前缀
        style_properties['highlight'] = highlight
    
    # 转换上标/下标
    if 'superscript' in run_data and run_data['superscript']:
        style_properties['vert_align'] = 'superscript'
    elif 'subscript' in run_data and run_data['subscript']:
        style_properties['vert_align'] = 'subscript'
    
    return style_properties


def update_document_content(marked_content, original_file_path,output_path):
    """匹配更新项目文档"""
    # 创建原始文档的StyleAnalyzer对象副本
    document_copy = StyleAnalyzer(original_file_path)
    
    # 将元素按照diff_status分类
    deleted_elements = []
    modified_elements = []
    added_elements = []
    print()
    for i in range(len(marked_content)):
        item = marked_content[i]
        diff_status = item.get('__diff_status')
        
        if diff_status == 'deleted':
            deleted_elements.append(item)
        elif diff_status == 'modified':
            modified_elements.append(item)
        elif diff_status == 'added':
            added_elements.append(item)
        else:
            # 默认按照已有元素处理
            modified_elements.append(item)
    
    # 1. 首先处理删除操作 - 按照索引从大到小排序，避免删除操作影响后续元素的索引
    deleted_elements.sort(key=lambda x: x.get('__original_index', 0), reverse=True)
    
    # 记录删除的元素索引，用于后续索引调整
    deleted_indices = set()
    
    for item in deleted_elements:
        original_index = item.get('__original_index')
        if original_index is not None:
            element_type = item.get('type')
            logger.info(f"删除元素: 类型={element_type}, 索引={original_index}")
            
            # 执行删除操作
            if element_type == 'image':
                # 查找图片所在段落和图片索引
                paragraph_index = original_index
                image_index = 0  # 默认为第一个图片
                
                # 如果有额外信息指定图片索引，则使用指定值
                if 'image_index' in item:
                    image_index = item['image_index']
                
                success = document_copy.remove_image_at_paragraph(paragraph_index, image_index)
                if success:
                    deleted_indices.add(original_index)
                    logger.info(f"成功删除图片: 段落索引={paragraph_index}, 图片索引={image_index}")
                else:
                    logger.error(f"删除图片失败: 段落索引={paragraph_index}, 图片索引={image_index}")
            else:
                # 普通元素删除
                success = document_copy.remove_element(original_index)
                if success:
                    deleted_indices.add(original_index)
                    logger.info(f"成功删除元素: 索引={original_index}")
                else:
                    logger.error(f"删除元素失败: 索引={original_index}")
    
    # 2. 处理修改操作
    for item in modified_elements:
        element_type = item.get('type')
        original_index = item.get('__original_index')
        
        # 由于前面的删除操作，需要调整索引
        adjusted_index = original_index
        if original_index is not None:
            # 计算有多少个小于当前索引的元素已被删除
            offset = sum(1 for idx in deleted_indices if idx < original_index)
            adjusted_index = original_index - offset
            
            logger.info(f"更新元素: 类型={element_type}, 原始索引={original_index}, 调整后索引={adjusted_index}")
            
            if element_type == 'paragraph' or element_type == 'title':
                # 处理段落或标题元素
                element_data = {}
                
                # 提取基本属性
                if 'id' in item:
                    element_data['id'] = item['id']
                elif 'paragraphId' in item:
                    element_data['id'] = item['paragraphId']
                elif 'titleId' in item:
                    element_data['id'] = item['titleId']
                    
                # 提取元素类型
                element_data['type'] = item.get('type')
                
                # 提取元素值
                element_data['value'] = item.get('value', '')
                
                # 提取样式属性
                element_data['rowFlex'] = item.get('rowFlex')  # 对齐方式
                element_data['indent'] = item.get('indent')    # 缩进
                element_data['lineRule'] = item.get('lineRule', 'auto')  # 行距规则
                element_data['line'] = item.get('line')        # 行距大小
                element_data['rowMargin'] = item.get('rowMargin')  # 行间距
                
                # 如果是标题，提取标题级别
                if element_type == 'title' and 'level' in item:
                    element_data['level'] = item.get('level')
                
                # 使用调整后的索引
                element_data['original_index'] = adjusted_index
                
                # 处理valueList - 遍历所有子元素
                if 'valueList' in item and item['valueList']:
                    # 根据元素类型选择适当的转换函数
                    if element_type == 'title':
                        style_properties = convert_editor_to_docx_title(element_data)
                    else:
                        style_properties = convert_editor_to_docx_paragraph(element_data)
                    
                    # 首先更新段落样式
                    document_copy.update_paragraph_style_from_xml(document_copy.elements[adjusted_index]['element'], **style_properties)
                    
                    # 标题样式已通过style_id参数设置，不需要额外处理
                    # if element_type == 'title' and 'heading_level' in style_properties:
                    #     document_copy.set_paragraph_as_heading(
                    #         document_copy.elements[adjusted_index]['element'], 
                    #         style_properties['heading_level']
                    #     )
                    
                    # 处理子元素
                    value_list = item['valueList']
                    
                    # 标题类型特殊处理：可能包含段落，段落再包含文本runs
                    if element_type == 'title' and len(value_list) > 0 and value_list[0].get('type') == 'paragraph':
                        # 标题内嵌段落结构处理
                        for para_idx, para in enumerate(value_list):
                            if 'valueList' in para and para['valueList']:
                                para_runs = para['valueList']
                                
                                # 处理段落中的每个run
                                for run_idx, run in enumerate(para_runs):
                                    # 跳过空的Run或最后一个换行符
                                    if not run.get('value') or (run_idx == len(para_runs) - 1 and 
                                                              (run.get('value') == '\n' or run.get('value') == '\u200B')):
                                        continue
                                    
                                    # 转换Run样式
                                    run_style = convert_editor_to_docx_run(run)
                                    
                                    # 应用标题的默认字体大小
                                    if 'default_size' in style_properties and 'size' not in run_style:
                                        run_style['size'] = str(style_properties['default_size'] * 2)  # Word中字号是磅值的2倍
                                    
                                    # 更新Run样式
                                    document_copy.update_run_style_from_xml(document_copy.elements[adjusted_index]['element'], run_idx, **run_style)
                                
                                # 删除多余的runs
                                if para_runs:
                                    document_copy.delete_runs_after_index_from_xml(
                                        document_copy.elements[adjusted_index]['element'], len(para_runs) - 1)
                    else:
                        # 普通段落或简单标题处理
                        for run_idx, run in enumerate(value_list):
                            # 跳过空的Run或最后一个换行符
                            if not run.get('value') or (run_idx == len(value_list) - 1 and 
                                                      (run.get('value') == '\n' or run.get('value') == '\u200B')):
                                continue
                            
                            # 转换Run样式
                            run_style = convert_editor_to_docx_run(run)
                            
                            # 如果是标题且有默认字体大小，使用默认大小
                            if element_type == 'title' and 'default_size' in style_properties and 'size' not in run_style:
                                run_style['size'] = str(style_properties['default_size'] * 2)  # Word中字号是磅值的2倍

                            # 更新Run样式
                            document_copy.update_run_style_from_xml(document_copy.elements[adjusted_index]['element'], run_idx, **run_style)
                        
                        # 删除多余的runs
                        if value_list:
                            document_copy.delete_runs_after_index_from_xml(
                                document_copy.elements[adjusted_index]['element'], len(value_list) - 1)
                else:
                    # 如果没有valueList，仍然更新段落样式
                    if element_type == 'title':
                        style_properties = convert_editor_to_docx_title(element_data)
                    else:
                        style_properties = convert_editor_to_docx_paragraph(element_data)
                    
                    document_copy.update_paragraph_style_from_xml(document_copy.elements[adjusted_index]['element'],
                                                                  **style_properties)
                    
                    # 标题样式已通过style_id参数设置，不需要额外处理
                    # if element_type == 'title' and 'heading_level' in style_properties:
                    #     document_copy.set_paragraph_as_heading(
                    #         document_copy.elements[adjusted_index]['element'], 
                    #         style_properties['heading_level']
                    #     )
                    
                    document_copy.delete_runs_after_index_from_xml(
                        document_copy.elements[adjusted_index]['element'], 0)


            elif element_type == 'table':
                # 使用调整后的索引更新表格
                table_data = item.copy()
                table_data['original_index'] = adjusted_index
                
                # 转换Canvas-Editor表格格式为DOCX格式
                table_style_properties = convert_editor_to_docx_table(table_data)
                
                # 获取表格元素
                table_element = document_copy.elements[adjusted_index]['element']
                
                # 应用表格样式
                document_copy.set_table_style_from_xml(table_element, **table_style_properties)
                
                # 如果表格有行数据，处理每一行
                if 'trList' in table_data and table_data['trList']:
                    for row_idx, row_data in enumerate(table_data['trList']):
                        # 转换行样式
                        row_style_properties = convert_editor_to_docx_table_row(row_data)
                        
                        # 应用行样式
                        document_copy.set_table_row_style_from_xml(table_element, row_idx, **row_style_properties)
                        
                        # 如果行有单元格数据，处理每个单元格
                        if 'tdList' in row_data and row_data['tdList']:
                            for cell_idx, cell_data in enumerate(row_data['tdList']):
                                # 转换单元格样式
                                cell_style_properties = convert_editor_to_docx_table_cell(cell_data)
                                
                                # 应用单元格样式
                                document_copy.set_table_cell_style_from_xml(table_element, row_idx, cell_idx, **cell_style_properties)
                                
                                # 处理单元格内容
                                if 'value' in cell_data and cell_data['value']:
                                    # 获取单元格中的段落元素
                                    cell_paragraphs = document_copy.get_table_cell_paragraphs_from_xml(table_element, row_idx, cell_idx)
                                    
                                    # 处理每个内容段落
                                    for para_idx, para_data in enumerate(cell_data['value']):
                                        if para_idx < len(cell_paragraphs):
                                            # 转换段落样式
                                            para_style = convert_editor_to_docx_paragraph(para_data)
                                            
                                            # 应用段落样式
                                            document_copy.update_paragraph_style_from_xml(cell_paragraphs[para_idx], **para_style)
                                            
                                            # 处理段落中的文本运行
                                            if 'valueList' in para_data and para_data['valueList']:
                                                for run_idx, run_data in enumerate(para_data['valueList']):
                                                    # 跳过空的Run或最后一个换行符
                                                    if not run_data.get('value') or (run_idx == len(para_data['valueList']) - 1 and 
                                                                                  (run_data.get('value') == '\n' or run_data.get('value') == '\u200B')):
                                                        continue
                                                        
                                                    # 转换Run样式
                                                    run_style = convert_editor_to_docx_run(run_data)
                                                    
                                                    # 更新Run样式
                                                    document_copy.update_run_style_from_xml(cell_paragraphs[para_idx], run_idx, **run_style)
                                                    document_copy.delete_runs_after_index_from_xml(
                                                        document_copy.elements[adjusted_index]['element'], run_idx)
                
            elif element_type == 'image':
                # 使用调整后的索引更新图片
                image_data = item.copy()
                image_data['original_index'] = adjusted_index
                
                # 从图片ID中提取关系ID
                image_id = item.get('id', '')
                rel_id = None
                if image_id and '-' in image_id:
                    # 格式如 "img-rId20-13"，提取出rId20部分
                    parts = image_id.split('-')
                    if len(parts) >= 2 and parts[1].startswith('rId'):
                        rel_id = parts[1]
                
                if rel_id:
                    # 获取图片URL和属性
                    image_url = item.get('value', '')
                    width = item.get('width')
                    height = item.get('height')
                    
                    # 如果宽高单位是像素，转换为厘米
                    if width is not None:
                        width = width / px_ch_width
                    if height is not None:
                        height = height / px_ch_width
                    
                    # 获取图片显示方式
                    img_display = item.get('imgDisplay', 'inline')
                    wrap_text_map = {
                        'inline': 'inline',
                        'surround': 'square',
                        'block': 'topAndBottom',
                        'float-bottom': 'behind',
                        'float-top': 'inFront'
                    }
                    wrap_text = wrap_text_map.get(img_display, 'inline')
                    
                    # 如果图片URL是本地服务器上的图片，获取图片路径
                    local_image_path = None
                    if image_url and image_url.startswith((API_BASE_URL, IMAGE_SERVER_URL or API_BASE_URL)):
                        # 从URL提取文件名
                        image_filename = image_url.split('/')[-1]
                        local_image_path = os.path.join("static", "images", image_filename)
                        
                        if not os.path.exists(local_image_path):
                            logger.warning(f"本地图片文件不存在: {local_image_path}")
                            local_image_path = None
                    
                    # 替换图片
                    if local_image_path:
                        document_copy.replace_image(
                            rel_id=rel_id,
                            image_path=local_image_path,
                            width=width,
                            height=height,
                            wrap_text=wrap_text
                        )
                    else:
                        logger.error(f"无法替换图片，未找到有效的图片路径: {image_url}")
                else:
                    logger.error(f"无法从图片ID中提取关系ID: {image_id}")
    
    # 3. 最后处理新增操作

    # 这里需要根据 StyleAnalyzer 提供的 API 来实现新增元素的功能
    
    # 按索引排序处理新增元素，从小到大
    added_elements.sort(key=lambda x: x.get('__original_index', float('inf')))
    
    # 记录已插入元素引起的索引变化
    inserted_count = 0
    print(f'added_elements:{added_elements}')
    for item in added_elements:
        element_type = item.get('type')
        original_index = item.get('__original_index')
        
        if original_index is None:
            # 如果没有指定索引，默认添加到文档末尾
            original_index = len(document_copy.elements)
        
        # 调整索引：减去删除元素造成的偏移，加上前面插入元素的数量
        deleted_before = sum(1 for idx in deleted_indices if idx < original_index)
        adjusted_index = original_index - deleted_before + inserted_count
        
        logger.info(f"新增元素: 类型={element_type}, 目标索引={original_index}, 调整后索引={adjusted_index}")
        
        if element_type == 'paragraph' or element_type == 'title':
            # 构造段落文本和样式
            text = ''
            
            # 获取实际文本内容，处理嵌套结构
            if element_type == 'title' and 'valueList' in item and item['valueList']:
                # 标题可能有嵌套结构：title -> paragraph -> runs
                title_value_list = item['valueList']
                if len(title_value_list) > 0 and title_value_list[0].get('type') == 'paragraph':
                    # 从嵌套的段落中提取文本
                    paragraph = title_value_list[0]
                    if 'valueList' in paragraph and paragraph['valueList']:
                        para_runs = paragraph['valueList']
                        text = ''.join(run.get('value', '') for run in para_runs if run.get('value'))
            elif 'value' in item and item['value']:
                text = item['value']
            elif 'valueList' in item:
                # 从valueList合并文本
                text = ''.join(run.get('value', '') for run in item['valueList'] if run.get('value'))
            
            # 转换样式属性
            if element_type == 'title':
                style_properties = convert_editor_to_docx_title(item)
            else:
                style_properties = convert_editor_to_docx_paragraph(item)
            
            # 插入段落
            position = 'after'
            target_index = adjusted_index - 1 if position == 'after' else 0
            target_index = max(0, min(target_index, len(document_copy.elements) - 1))
            
            logger.info(f"准备插入{element_type}: 文本='{text}', 目标索引={target_index}")
            
            success = document_copy.insert_paragraph(
                element_index=target_index,
                position=position,
                text=text,
                **style_properties,
            )
            
            if success:
                logger.info(f"成功插入{element_type}: 索引={target_index}, 位置={position}")
                inserted_count += 1
                
                # 获取新插入的段落索引
                new_para_index = target_index + (1 if position == 'after' else 0)
                new_para = document_copy.elements[new_para_index]['element']
                
                # 标题样式已通过style_id参数设置，不需要额外处理
                
                # 处理样式
                if element_type == 'title' and 'valueList' in item and item['valueList']:
                    # 标题嵌套结构处理
                    title_value_list = item['valueList']
                    if len(title_value_list) > 0 and title_value_list[0].get('type') == 'paragraph':
                        paragraph = title_value_list[0]
                        if 'valueList' in paragraph and paragraph['valueList']:
                            para_runs = paragraph['valueList']
                            
                            logger.info(f"处理标题嵌套结构: 包含 {len(para_runs)} 个文本运行")
                            
                            # 处理每个run
                            for run_idx, run in enumerate(para_runs):
                                if not run.get('value'):
                                    continue
                                
                                # 转换Run样式
                                run_style = convert_editor_to_docx_run(run)
                                
                                # 应用标题默认字体大小
                                if 'default_size' in style_properties and 'size' not in run_style:
                                    run_style['size'] = str(style_properties['default_size'] * 2)
                                
                                logger.info(f"更新标题中的Run {run_idx}: 文本='{run.get('value')}', 样式={run_style}")
                                
                                # 更新Run样式
                                document_copy.update_run_style_from_xml(new_para, run_idx, **run_style)
                elif 'valueList' in item and item['valueList']:
                    # 普通段落处理
                    for run_idx, run in enumerate(item['valueList']):
                        if not run.get('value'):
                            continue
                        
                        # 转换Run样式
                        run_style = convert_editor_to_docx_run(run)
                        
                        # 如果是标题，应用默认字体大小
                        if element_type == 'title' and 'default_size' in style_properties and 'size' not in run_style:
                            run_style['size'] = str(style_properties['default_size'] * 2)
                        
                        logger.info(f"更新Run {run_idx}: 文本='{run.get('value')}', 样式={run_style}")
                        
                        # 更新Run样式
                        document_copy.update_run_style_from_xml(new_para, run_idx, **run_style)
            else:
                logger.error(f"插入{element_type}失败: 索引={target_index}, 位置={position}")
                
        elif element_type == 'table':
            # 处理表格插入
            # 获取表格数据
            if 'trList' not in item or not item['trList']:
                logger.error("表格数据不完整，缺少行数据")
                continue
                
            # 确定表格维度
            rows_count = len(item['trList'])
            cols_count = 0
            # 查找最大列数
            for row in item['trList']:
                if 'tdList' in row and row['tdList']:
                    cols_count = max(cols_count, len(row['tdList']))
            
            if rows_count == 0 or cols_count == 0:
                logger.error(f"表格维度无效: 行={rows_count}, 列={cols_count}")
                continue
                
            # 准备表格样式
            table_style_properties = convert_editor_to_docx_table(item)
            
            # 构建表格内容 - 一个简单的二维文本数组
            table_content = []
            for row_idx in range(rows_count):
                row_content = []
                row_data = item['trList'][row_idx]
                
                # 如果该行有单元格数据
                if 'tdList' in row_data and row_data['tdList']:
                    for cell_idx in range(cols_count):
                        # 如果单元格存在
                        if cell_idx < len(row_data['tdList']):
                            cell_data = row_data['tdList'][cell_idx]
                            
                            # 提取单元格文本内容
                            cell_text = ""
                            if 'value' in cell_data and cell_data['value']:
                                for para in cell_data['value']:
                                    if 'valueList' in para and para['valueList']:
                                        for run in para['valueList']:
                                            if 'value' in run and run['value'] not in ['\n', '\u200B']:
                                                cell_text += run['value']
                                    elif 'value' in para:
                                        cell_text += para['value']
                            
                            row_content.append(cell_text)
                        else:
                            row_content.append("")  # 空单元格
                else:
                    # 整行为空，填充空单元格
                    row_content = [""] * cols_count
                
                table_content.append(row_content)
            
            # 根据需要计算插入位置
            position = 'after' if adjusted_index > 0 else 'before'
            target_index = adjusted_index - 1 if position == 'after' else 0
            target_index = max(0, min(target_index, len(document_copy.elements) - 1))
            
            # 插入表格
            success = document_copy.insert_table(
                element_index=target_index,
                position=position,
                rows=rows_count,
                cols=cols_count,
                table_content=table_content,
                **table_style_properties
            )
            
            if success:
                logger.info(f"成功插入表格: 索引={target_index}, 位置={position}, 尺寸={rows_count}x{cols_count}")
                inserted_count += 1
                
                # 获取新插入的表格元素索引
                new_table_index = target_index + (1 if position == 'after' else 0)
                # 如果新表格索引有效，获取表格元素
                if new_table_index < len(document_copy.elements):
                    new_table_element = document_copy.elements[new_table_index]['element']
                    
                    # 应用表格样式、行样式和单元格样式
                    # 这部分代码可以重用表格更新时的代码，实现表格、行、单元格样式设置
                    # 这里只展示基本思路，完整实现需要更多的代码
                    for row_idx, row_data in enumerate(item['trList']):
                        if row_idx >= rows_count:
                            break
                            
                        # 设置行样式
                        row_style = convert_editor_to_docx_table_row(row_data)
                        document_copy.set_table_row_style_from_xml(new_table_element, row_idx, **row_style)
                        
                        # 设置单元格样式和内容
                        if 'tdList' in row_data and row_data['tdList']:
                            for cell_idx, cell_data in enumerate(row_data['tdList']):
                                if cell_idx >= cols_count:
                                    break
                                    
                                # 设置单元格样式
                                cell_style = convert_editor_to_docx_table_cell(cell_data)
                                document_copy.set_table_cell_style_from_xml(new_table_element, row_idx, cell_idx, **cell_style)
                                
                                # 处理单元格内容
                                if 'value' in cell_data and cell_data['value']:
                                    # 清除单元格中可能存在的默认段落
                                    existing_paragraphs = document_copy.get_table_cell_paragraphs_from_xml(new_table_element, row_idx, cell_idx)
                                    
                                    # 为每个内容段落创建新段落
                                    for para_idx, para_data in enumerate(cell_data['value']):
                                        # 转换段落样式
                                        para_style = convert_editor_to_docx_paragraph(para_data)
                                        
                                        # 如果是第一个段落且已有段落存在，则更新现有段落
                                        if para_idx == 0 and existing_paragraphs:
                                            # 应用段落样式到现有段落
                                            document_copy.update_paragraph_style_from_xml(existing_paragraphs[0], **para_style)
                                            paragraph = existing_paragraphs[0]
                                        else:
                                            # 否则创建新段落
                                            success, paragraph = document_copy.create_paragraph_in_cell(new_table_element, row_idx, cell_idx)
                                            if not success or not paragraph:
                                                logger.error(f"在单元格({row_idx}, {cell_idx})中创建段落失败")
                                                continue
                                                
                                            # 应用段落样式
                                            document_copy.update_paragraph_style_from_xml(paragraph, **para_style)
                                        
                                        # 处理段落中的文本运行
                                        if 'valueList' in para_data and para_data['valueList']:

                                            
                                            # 处理每个run
                                            for run_idx, run_data in enumerate(para_data['valueList']):
                                                # 跳过空的Run或最后一个换行符
                                                if not run_data.get('value') or (run_idx == len(para_data['valueList']) - 1 and 
                                                                                (run_data.get('value') == '\n' or run_data.get('value') == '\u200B')):
                                                    continue
                                                
                                                # 转换Run样式
                                                run_style = convert_editor_to_docx_run(run_data)
                                                
                                                # 如果是第一个run，更新现有run

                                                document_copy.update_run_style_from_xml(paragraph, run_idx, **run_style)


            else:
                logger.error(f"插入表格失败: 索引={target_index}, 位置={position}")
            
        elif element_type == 'image':
            # 处理图片插入
            # 获取图片数据
            image_path = item.get('value', '')
            
            # 确保图片路径是系统路径
            if not os.path.exists(image_path):
                logger.error(f"图片文件不存在: {image_path}")
                continue
            
            # 获取图片参数
            width = item.get('width')
            height = item.get('height')
            
            # 如果宽高单位是像素或其他单位，转换为厘米
            if width is not None:
                width = width / px_ch_width
            if height is not None:
                height = height / px_ch_width
                
            description = item.get('description', '')
            
            # 获取环绕方式并转换为DOCX格式
            img_display = item.get('imgDisplay', 'inline')
            wrap_text_map = {
                'inline': 'inline',
                'surround': 'square',
                'block': 'topAndBottom',
                'float-bottom': 'behind',
                'float-top': 'inFront'
            }
            wrap_text = wrap_text_map.get(img_display, 'inline')
            
            # 获取行距和行距规则
            line_rule = item.get('lineRule', 'auto')
            line_spacing = item.get('line', 240)
            
            # 根据需要计算插入位置
            position = 'after' if adjusted_index > 0 else 'before'
            target_index = adjusted_index - 1 if position == 'after' else 0
            target_index = max(0, min(target_index, len(document_copy.elements) - 1))
            
            # 插入图片
            success = document_copy.insert_image(
                ele_index=target_index,
                run_index=-1,  # 默认在段落末尾
                position=position,
                image_path=image_path,
                width=width,
                height=height,
                description=description,
                wrap_text=wrap_text,
                line_spacing=line_spacing,
                line_rule=line_rule
            )
            
            if success:
                logger.info(f"成功插入图片: 索引={target_index}, 位置={position}")
                inserted_count += 1
            else:
                logger.error(f"插入图片失败: 索引={target_index}, 位置={position}")
    print(f"elements:{document_copy.elements}")
    document_copy.save(output_path)


    # 将Canvas-Editor格式转换为DOCX格式
def convert_editor_to_docx_title(title_data):
    """将Canvas-Editor格式的标题数据转换为DOCX需要的格式
    
    Args:
        title_data: Canvas-Editor格式的标题数据
        
    Returns:
        dict: DOCX格式的标题样式属性
    """
    style_properties = {}
    
    # 基本属性与段落相同
    paragraph_props = convert_editor_to_docx_paragraph(title_data)
    style_properties.update(paragraph_props)
    
    # 处理标题级别
    if 'level' in title_data:
        # 直接使用Canvas-Editor的标题级别映射到Word样式ID
        level_to_style_id = {
            'first': '2',    # 一级标题
            'second': '3',   # 二级标题
            'third': '4',    # 三级标题
            'fourth': '5',   # 四级标题

        }
        
        level = title_data.get('level')
        
        # 设置标题样式ID
        style_properties['style_id'] = level_to_style_id.get(level, '2')
        
        # 标题通常有特定的字体大小
        # 根据标题级别设置默认字体大小
        default_sizes = {
            'first': 28,
            'second': 24,
            'third': 20,
            'fourth': 18,
            'fifth': 16,
            'sixth': 14
        }
        default_size = default_sizes.get(level, 16)
        style_properties['default_size'] = default_size
    
    return style_properties

def convert_editor_to_docx_table(table_data):
    """将Canvas-Editor格式的表格数据转换为DOCX需要的格式
    
    Args:
        table_data: Canvas-Editor格式的表格数据
        
    Returns:
        dict: DOCX格式的表格样式属性
    """
    style_properties = {}
    
    # 处理表格宽度
    if 'width' in table_data and table_data['width'] is not None:
        # 将Canvas-Editor的宽度值转换回DOCX格式 (乘以转换常数)
        width_value = int(table_data['width'] * width_num)  # 使用全局转换常数
        style_properties['width'] = {'value': width_value, 'type': 'dxa'}
    
    # 处理表格边框类型
    if 'borderType' in table_data:
        border_type = table_data['borderType']
        borders = {}
        
        # 根据边框类型设置不同的边框样式
        if border_type == 'all':
            # 所有边框都设置
            for border_pos in ['top', 'left', 'bottom', 'right', 'inside_h', 'inside_v']:
                borders[border_pos] = {'val': 'single', 'sz': '4', 'color': '000000'}
        elif border_type == 'external':
            # 只设置外边框
            for border_pos in ['top', 'left', 'bottom', 'right']:
                borders[border_pos] = {'val': 'single', 'sz': '4', 'color': '000000'}
            # 内部边框设为无
            borders['inside_h'] = {'val': 'nil'}
            borders['inside_v'] = {'val': 'nil'}
        elif border_type == 'internal':
            # 只设置内边框
            borders['inside_h'] = {'val': 'single', 'sz': '4', 'color': '000000'}
            borders['inside_v'] = {'val': 'single', 'sz': '4', 'color': '000000'}
            # 外部边框设为无
            for border_pos in ['top', 'left', 'bottom', 'right']:
                borders[border_pos] = {'val': 'nil'}
        elif border_type == 'empty' or border_type is None:
            # 所有边框都不显示
            for border_pos in ['top', 'left', 'bottom', 'right', 'inside_h', 'inside_v']:
                borders[border_pos] = {'val': 'nil'}
        elif border_type == 'dash':
            # 虚线边框
            for border_pos in ['top', 'left', 'bottom', 'right', 'inside_h', 'inside_v']:
                borders[border_pos] = {'val': 'dashed', 'sz': '4', 'color': '000000'}
        
        # 如果有边框颜色，覆盖默认颜色
        if 'borderColor' in table_data and table_data['borderColor']:
            color = table_data['borderColor']
            if color.startswith('#'):
                color = color[1:]  # 移除#前缀
            for border_pos in borders:
                if borders[border_pos]['val'] != 'nil':
                    borders[border_pos]['color'] = color
        
        if borders:
            style_properties['borders'] = borders
    
    # 设置表格布局
    style_properties['layout'] = 'fixed'  # 默认固定布局
    
    # 默认单元格边距
    style_properties['cell_margins'] = {
        'top': {'value': '0', 'type': 'dxa'},
        'left': {'value': '108', 'type': 'dxa'},
        'bottom': {'value': '0', 'type': 'dxa'},
        'right': {'value': '108', 'type': 'dxa'}
    }
    
    return style_properties

def convert_editor_to_docx_table_row(row_data):
    """将Canvas-Editor格式的表格行数据转换为DOCX需要的格式
    
    Args:
        row_data: Canvas-Editor格式的表格行数据
        
    Returns:
        dict: DOCX格式的表格行样式属性
    """
    style_properties = {}
    
    # 处理行高
    if 'minHeight' in row_data and row_data['minHeight'] is not None:
        # 转换行高
        style_properties['height'] = {
            'value': int(row_data['minHeight']),
            'rule': 'atLeast'  # 至少为指定高度
        }
    
    # 禁止跨页分割 - 默认允许跨页
    style_properties['cannot_split'] = False
    
    # 是否为表头行 - 默认不是
    style_properties['is_header'] = False
    
    # 如果有行级别的边框设置，可以在这里添加
    
    return style_properties

def convert_editor_to_docx_table_cell(cell_data):
    """将Canvas-Editor格式的表格单元格数据转换为DOCX需要的格式
    
    Args:
        cell_data: Canvas-Editor格式的表格单元格数据
        
    Returns:
        dict: DOCX格式的表格单元格样式属性
    """
    style_properties = {}
    
    # 处理垂直对齐方式
    if 'verticalAlign' in cell_data and cell_data['verticalAlign'] is not None:
        # Canvas-Editor的垂直对齐枚举转换为DOCX格式
        vertical_align_map = {
            'TOP': 'top',
            'CENTER': 'center',
            'BOTTOM': 'bottom'
        }
        
        # 获取值或枚举名
        if hasattr(cell_data['verticalAlign'], 'value'):
            # 如果是枚举对象
            align_value = cell_data['verticalAlign'].value
        else:
            # 如果是字符串
            align_value = cell_data['verticalAlign']
            
        # 转换为大写以匹配枚举
        align_key = align_value.upper() if isinstance(align_value, str) else align_value
        
        # 查找映射
        if align_key in vertical_align_map:
            style_properties['vertical_align'] = vertical_align_map[align_key]
        else:
            # 默认居中
            style_properties['vertical_align'] = 'center'
    
    # 处理背景颜色
    if 'backgroundColor' in cell_data and cell_data['backgroundColor']:
        background_color = cell_data['backgroundColor']
        if background_color.startswith('#'):
            background_color = background_color[1:]  # 移除#前缀
        
        style_properties['shading'] = {
            'val': 'clear',  # 纯色填充
            'color': 'auto',  # 自动文本颜色
            'fill': background_color  # 背景颜色
        }
    
    # 处理单元格边框
    if 'borderTypes' in cell_data and cell_data['borderTypes']:
        borders = {}
        border_sides = cell_data['borderTypes']
        
        # 设置指定的边框
        for border_side in border_sides:
            if border_side in ['top', 'left', 'bottom', 'right']:
                borders[border_side] = {'val': 'single', 'sz': '4', 'color': '000000'}
        
        # 未指定的边框设为无
        for side in ['top', 'left', 'bottom', 'right']:
            if side not in border_sides:
                borders[side] = {'val': 'nil'}
                
        if borders:
            style_properties['borders'] = borders
    
    # 处理合并单元格 - 跨行
    if 'rowspan' in cell_data and cell_data['rowspan'] > 1:
        style_properties['rowspan'] = cell_data['rowspan']
    
    # 处理合并单元格 - 跨列
    if 'colspan' in cell_data and cell_data['colspan'] > 1:
        style_properties['colspan'] = cell_data['colspan']
    
    return style_properties
