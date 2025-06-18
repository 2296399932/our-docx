from pydantic import BaseModel, Field
from typing import List, Dict, Any, Optional, Union
from datetime import datetime
from enum import Enum

class DocumentResponse(BaseModel):
    """文档响应模型"""
    filename: str = Field(..., description="文件名称")
    status: str = Field(..., description="处理状态")
    file_path: str = Field(..., description="文件保存路径")
    upload_time: datetime = Field(default_factory=datetime.now, description="上传时间")
    
    class Config:
        schema_extra = {
            "example": {
                "filename": "example.docx",
                "status": "文件上传成功",
                "file_path": "uploads/example.docx",
                "upload_time": "2023-09-01T12:00:00"
            }
        }

class TextAlign(str, Enum):
    """文本对齐方式"""
    LEFT = "left"
    RIGHT = "right"
    CENTER = "center"
    JUSTIFY = "justify"
    
class VerticalAlign(str, Enum):
    """垂直对齐方式"""
    TOP = "top"
    CENTER = "center"
    BOTTOM = "bottom"

class Border(BaseModel):
    """边框模型"""
    style: Optional[str] = Field(None, description="边框样式(single, double, dash等)")
    width: Optional[float] = Field(None, description="边框宽度(磅)")
    color: Optional[str] = Field(None, description="边框颜色(十六进制)")

class RunModel(BaseModel):
    """文本运行模型(Run)"""
    value: str = Field(..., description="运行文本内容")
    font: Optional[str] = Field(None, description="字体名称")
    size: Optional[float] = Field(None, description="字体大小(磅)")
    bold: bool = Field(False, description="是否粗体")
    italic: bool = Field(False, description="是否斜体")
    underline: bool = Field(False, description="是否下划线")
    strike: bool = Field(False, description="是否删除线")
    color: Optional[str] = Field(None, description="字体颜色(十六进制)")
    highlight: Optional[str] = Field(None, description="高亮颜色(十六进制)")

    
    class Config:
        schema_extra = {
            "example": {
                "text": "示例文本",
                "font": "宋体",
                "size": 12.0,
                "bold": True,
                "color": "#FF0000",
                "highlight": "#FFFF00"
            }
        }

class ParagraphModel(BaseModel):
    """段落模型"""
    id: Optional[str] = Field(None, description="段落ID")
    valueList: List[RunModel] = Field(default_factory=list, description="段落中的文本运行")

    rowFlex: Optional[float] = Field(None, description="段落对齐方式")

    type: Optional[str] = Field(None, description="段落类型")
    value: str = Field(..., description="运行文本内容")

    indent: Optional[float] = Field(None, description="缩进(磅)")

    
    class Config:
        schema_extra = {
            "example": {
                "id": "p1",
                "runs": [
                    {
                        "text": "这是一段示例文本，",
                        "font_name": "宋体",
                        "font_size": 12.0
                    },
                    {
                        "text": "这部分是粗体",
                        "bold": True
                    }
                ],
                "align": "left",
                "indent_first_line": 21.0,
                "spacing_after": 10.0,
                "line_spacing": 1.5
            }
        }

class ImageModel(BaseModel):
    """图片模型"""
    id: Optional[str] = Field(None, description="图片ID")
    file_name: Optional[str] = Field(None, description="图片文件名")
    description: Optional[str] = Field(None, description="图片描述")
    width: Optional[float] = Field(None, description="宽度(厘米)")
    height: Optional[float] = Field(None, description="高度(厘米)")
    rel_id: Optional[str] = Field(None, description="关系ID")
    wrap_type: str = Field("inline", description="环绕方式(inline, square, tight, through, behind)")
    
    class Config:
        schema_extra = {
            "example": {
                "id": "img1",
                "file_name": "image.png",
                "description": "示例图片",
                "width": 10.0,
                "height": 8.0,
                "wrap_type": "inline"
            }
        }

class TableCellModel(BaseModel):
    """表格单元格模型"""
    row_span: int = Field(1, description="行跨度")
    col_span: int = Field(1, description="列跨度")
    vertical_align: VerticalAlign = Field(VerticalAlign.CENTER, description="垂直对齐方式")
    background_color: Optional[str] = Field(None, description="背景颜色(十六进制)")
    borders: Dict[str, Border] = Field(default_factory=dict, description="边框(top, bottom, left, right)")
    paragraphs: List[ParagraphModel] = Field(default_factory=list, description="单元格中的段落")
    
class TableRowModel(BaseModel):
    """表格行模型"""
    height: Optional[float] = Field(None, description="行高(厘米)")
    cells: List[TableCellModel] = Field(..., description="行中的单元格")
    
class TableModel(BaseModel):
    """表格模型"""
    id: Optional[str] = Field(None, description="表格ID")
    rows: List[TableRowModel] = Field(..., description="表格行")
    width: Optional[float] = Field(None, description="表格宽度(厘米)")
    column_widths: List[float] = Field(default_factory=list, description="列宽(厘米)")
    border_type: Optional[str] = Field(None, description="边框类型")
    border_color: Optional[str] = Field(None, description="边框颜色(十六进制)")
    border_width: Optional[float] = Field(None, description="边框宽度(磅)")
    
    class Config:
        schema_extra = {
            "example": {
                "id": "table1",
                "rows": [
                    {
                        "height": 1.0,
                        "cells": [
                            {
                                "row_span": 1,
                                "col_span": 1,
                                "paragraphs": [
                                    {
                                        "runs": [
                                            {
                                                "text": "单元格内容",
                                                "font_size": 12.0
                                            }
                                        ]
                                    }
                                ]
                            }
                        ]
                    }
                ],
                "column_widths": [3.0, 3.0, 3.0]
            }
        }

class SeparatorModel(BaseModel):
    """分隔符模型"""
    id: Optional[str] = Field(None, description="分隔符ID")
    style: str = Field("solid", description="样式(solid, dash, dot等)")
    width: Optional[float] = Field(None, description="宽度(厘米)")
    color: Optional[str] = Field(None, description="颜色(十六进制)")
    
    class Config:
        schema_extra = {
            "example": {
                "id": "sep1",
                "style": "solid",
                "color": "#000000"
            }
        }

class PageBreakModel(BaseModel):
    """分页符模型"""
    id: Optional[str] = Field(None, description="分页符ID")
    
    class Config:
        schema_extra = {
            "example": {
                "id": "pb1"
            }
        }

# 定义文档元素的类型别名，用于文档内容列表
DocElement = Union[ParagraphModel, ImageModel, TableModel, SeparatorModel, PageBreakModel]

class DocumentModel(BaseModel):
    """完整文档模型"""
    title: Optional[str] = Field(None, description="文档标题")
    author: Optional[str] = Field(None, description="作者")
    created_time: Optional[datetime] = Field(None, description="创建时间")
    modified_time: Optional[datetime] = Field(None, description="修改时间")
    subject: Optional[str] = Field(None, description="主题")
    keywords: List[str] = Field(default_factory=list, description="关键词")
    
    # 文档正文内容 - 按顺序包含所有文档元素(段落、图片、表格等)
    content: List[DocElement] = Field(default_factory=list, description="文档内容元素列表")
    
    # 页眉页脚
    headers: Dict[str, List[DocElement]] = Field(default_factory=dict, description="页眉(default、first、even)")
    footers: Dict[str, List[DocElement]] = Field(default_factory=dict, description="页脚(default、first、even)")
    
    # 文档属性
    page_width: float = Field(21.0, description="页面宽度(厘米)")
    page_height: float = Field(29.7, description="页面高度(厘米)")
    margins: Dict[str, float] = Field(
        default_factory=lambda: {"top": 2.54, "right": 2.54, "bottom": 2.54, "left": 2.54},
        description="页边距(厘米)"
    )
    
    class Config:
        schema_extra = {
            "example": {
                "title": "示例文档",
                "author": "用户名",
                "content": [
                    {
                        "id": "p1",
                        "runs": [
                            {
                                "text": "这是文档的第一段",
                                "font_size": 12.0
                            }
                        ]
                    },
                    {
                        "id": "img1",
                        "file_name": "example.png",
                        "width": 10.0,
                        "height": 5.0
                    },
                    {
                        "id": "p2",
                        "runs": [
                            {
                                "text": "这是图片下方的说明文字",
                                "font_size": 10.0,
                                "italic": True
                            }
                        ],
                        "align": "center"
                    }
                ],
                "page_width": 21.0,
                "page_height": 29.7
            }
        }


