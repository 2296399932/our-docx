# Our-Docx

一个功能强大的Python库，用于解析、操作和生成Word文档(DOCX)文件。

## 简介

Our-Docx库提供了一套完整的API，可以轻松地读取、修改和创建Word文档。通过对文档结构的XML解析，提供了丰富的功能来操作文档的各个元素，包括段落、表格、图片、样式等。

## 主要功能

- 📄 **文档结构解析**：读取并解析DOCX文档的XML结构
- 📝 **内容提取**：提取段落、表格、图片和评论等内容
- 🖌️ **样式操作**：读取和修改文本和段落样式
- 📊 **表格处理**：创建、修改和导出表格
- 🖼️ **图片处理**：插入、提取和操作图片
- 📑 **目录生成**：创建和更新目录
- 📄 **页面控制**：添加分页符和页面格式控制

## 安装

```bash
# 安装方式（假设通过pip安装）
pip install our-docx
```
现在根据document =StyleAnalyzer(file_path) document .elements获取原始根据其type类型判断其是为image，table还是段落
## 基本用法

```python
from docx_namespace import DocxElementParser

# 打开一个Word文档
doc = DocxElementParser("example.docx")

# 获取所有段落文本
all_text = doc.get_all_paragraphs_text()
print(all_text)

# 修改段落样式
doc.set_paragraph_alignment(0, "center")

# 插入新段落
doc.insert_paragraph(text="这是新添加的段落", position="after")

# 保存修改后的文档
doc.save("modified_example.docx")
```

## API参考

### 文档解析和内容获取

| 方法名 | 描述 |
|-------|------|
| `get_element()` | 获取文档的根元素 |
| `find_elements_by_tag(tag_name)` | 按标签名查找元素 |
| `get_body_direct_children()` | 获取文档主体的直接子元素 |
| `get_all_paragraphs()` | 获取所有段落元素 |
| `get_all_paragraphs_text()` | 获取所有段落的文本内容 |
| `get_paragraphs_length()` | 获取文档中段落的数量 |
| `get_table_length()` | 获取文档中表格的数量 |
| `get_all_tables()` | 获取所有表格元素 |
| `get_paragraph_by_id(para_id)` | 通过ID获取特定段落 |
| `get_paragraph_text(paragraph)` | 获取指定段落的文本内容 |
| `get_all_text()` | 获取文档中所有文本内容 |
| `get_element_attributes(element)` | 获取元素的所有属性 |
| `get_structured_body_elements()` | 获取结构化的文档主体元素 |
| `get_element_text(num)` | 获取指定元素的文本内容 |
| `print_full_xml()` | 打印完整的XML内容 |

### 表格操作

| 方法名 | 描述 |
|-------|------|
| `extract_table_content(table_element)` | 提取表格内容 |
| `export_table_to_file(table_idx, file_path, format='xlsx')` | 导出表格到文件 |
| `export_all_tables(dir_path, format='xlsx')` | 导出所有表格 |
| `get_table_style(table_index)` | 获取表格样式 |
| `format_table_style(style_info)` | 格式化表格样式信息 |
| `set_table_style(table_index, **style_properties)` | 设置表格样式 |
| `set_table_grid(table_index, column_widths)` | 设置表格网格 |
| `set_table_borders(table_index, **borders)` | 设置表格边框 |
| `set_table_cell_margins(table_index, **margins)` | 设置单元格边距 |
| `set_table_width(table_index, width, width_type='dxa')` | 设置表格宽度 |
| `set_table_row_borders(table_index, row_index, **borders)` | 设置表格行边框 |
| `set_table_cell_borders(table_index, row_index, cell_index, **borders)` | 设置单元格边框 |
| `create_three_line_table(table_index)` | 创建三线表 |
| `insert_table(element_index, position, rows, cols, data, **style_properties)` | 插入表格 |
| `update_table_text_style(table_index, **style_properties)` | 更新表格文本样式 |
| `set_table_text_alignment(table_index, alignment, header_alignment)` | 设置表格文本对齐方式 |
| `insert_table_caption(table_index, chapter_num, caption_text, auto_num, style_id, **style_properties)` | 插入表格标题 |
| `get_table_dimensions(table_index)` | 获取表格尺寸 |
| `get_all_tables_dimensions()` | 获取所有表格尺寸 |
| `get_table_cell_style(table_index, row_idx, col_idx)` | 获取单元格样式 |
| `get_table_cell_paragraphs(table_index, row_idx, col_idx)` | 获取单元格段落 |
| `get_table_cell_text(table_index, row_idx, col_idx)` | 获取单元格文本 |

### 段落样式操作

| 方法名 | 描述 |
|-------|------|
| `extract_paragraph_style(paragraph_element)` | 提取段落样式 |
| `format_paragraph_style(style_info)` | 格式化段落样式信息 |
| `get_paragraph_alignment(num)` | 获取段落对齐方式 |
| `get_paragraph_indentation(num)` | 获取段落缩进 |
| `get_paragraph_spacing(num)` | 获取段落间距 |
| `get_paragraph_borders(num)` | 获取段落边框 |
| `get_paragraph_shading(num)` | 获取段落底纹 |
| `get_paragraph_numbering(num)` | 获取段落编号 |
| `get_paragraph_font(num)` | 获取段落字体 |
| `get_all_paragraph_styles(num)` | 获取所有段落样式 |
| `_get_or_create_pPr(paragraph_element)` | 获取或创建段落属性元素 |
| `set_paragraph_style_id(para_index, style_id)` | 设置段落样式ID |
| `set_paragraph_style_id_from_xml(para_index, style_id)` | 从XML设置段落样式ID |
| `set_paragraph_alignment(para_index, alignment)` | 设置段落对齐方式 |
| `set_paragraph_alignment_from_xml(para_index, alignment)` | 从XML设置段落对齐方式 |
| `set_paragraph_indentation_from_xml(para_index, **indentation)` | 从XML设置段落缩进 |
| `set_paragraph_indentation(para_index, **indentation)` | 设置段落缩进 |
| `set_paragraph_spacing(para_index, **spacing)` | 设置段落间距 |
| `set_paragraph_spacing_from_xml(para_index, **spacing)` | 从XML设置段落间距 |
| `set_paragraph_borders(para_index, **borders)` | 设置段落边框 |
| `set_paragraph_borders_from_xml(para_index, **borders)` | 从XML设置段落边框 |
| `set_paragraph_shading(para_index, val, color, fill)` | 设置段落底纹 |
| `set_paragraph_shading_from_xml(para_index, val, color, fill)` | 从XML设置段落底纹 |
| `set_paragraph_numbering(para_index, num_id, level)` | 设置段落编号 |
| `set_paragraph_font(para_index, **font_properties)` | 设置段落字体 |
| `set_paragraph_font_from_xml(para_index, **font_properties)` | 从XML设置段落字体 |
| `remove_paragraph_property(para_index, property_name)` | 移除段落属性 |
| `update_paragraph_style(para_index, **style_properties)` | 更新段落样式 |
| `update_paragraph_style_from_xml(para_element, **style_properties)` | 从XML更新段落样式 |
| `get_paragraph_style_from_element(paragraph_element)` | 从元素获取段落样式 |

### 文本运行(Run)操作

| 方法名 | 描述 |
|-------|------|
| `get_element_run_text(index)` | 获取元素中运行文本 |
| `get_paragraph_run_text(index)` | 获取段落中运行文本 |
| `get_element_run_content(index)` | 获取元素中运行内容 |
| `get_run_style(para_index, run_index, element_type)` | 获取运行样式 |
| `get_run_style_form_xml(para, run_index, element_type)` | 从XML获取运行样式 |
| `_get_run_style(para_index, run_index)` | 获取运行样式(内部方法) |
| `get_run_font(element_index, run_index, element_type)` | 获取运行字体 |
| `get_run_size(element_index, run_index, element_type)` | 获取运行大小 |
| `get_run_formatting(element_index, run_index, element_type)` | 获取运行格式 |
| `get_run_color(element_index, run_index, element_type)` | 获取运行颜色 |
| `format_run_style(style_info)` | 格式化运行样式信息 |
| `set_paragraph_runs_font(para_index, **font_properties)` | 设置段落中所有运行的字体 |
| `set_runs_bold(para_index, bold)` | 设置段落中所有运行的粗体 |
| `set_runs_italic(para_index, italic)` | 设置段落中所有运行的斜体 |
| `set_runs_underline(para_index, underline_type)` | 设置段落中所有运行的下划线 |
| `set_runs_color(para_index, color)` | 设置段落中所有运行的颜色 |
| `set_runs_size(para_index, size)` | 设置段落中所有运行的大小 |
| `set_runs_highlight(para_index, highlight_color)` | 设置段落中所有运行的高亮 |
| `set_runs_strike(para_index, strike)` | 设置段落中所有运行的删除线 |
| `set_runs_caps(para_index, caps)` | 设置段落中所有运行的大写 |
| `set_runs_vertical_alignment(para_index, alignment)` | 设置段落中所有运行的垂直对齐 |
| `update_runs_style(para_index, **style_properties)` | 更新段落中所有运行的样式 |
| `update_runs_style_from_xml(para_element, **style_properties)` | 从XML更新段落中所有运行的样式 |
| `get_run_element(para_index, run_index)` | 获取运行元素 |
| `_get_run_element(para_index, run_index)` | 获取运行元素(内部方法) |
| `get_run_element_from_xml(para, run_index)` | 从XML获取运行元素 |
| `_get_or_create_rPr(r_element)` | 获取或创建运行属性元素 |
| `get_run_count(para_index)` | 获取段落中运行的数量 |
| `get_run_count_from_xml(para_index)` | 从XML获取段落中运行的数量 |
| `get_run_text(para_index, run_index)` | 获取运行文本 |
| `_get_run_text(para_index, run_index)` | 获取运行文本(内部方法) |
| `get_run_text_from_xml(para, run_index)` | 从XML获取运行文本 |
| `set_run_font(para_index, run_index, **font_properties)` | 设置运行字体 |
| `set_run_size(para_index, run_index, size)` | 设置运行大小 |
| `set_run_bold(para_index, run_index, bold)` | 设置运行粗体 |
| `set_run_italic(para_index, run_index, italic)` | 设置运行斜体 |
| `set_run_underline(para_index, run_index, underline_type)` | 设置运行下划线 |
| `set_run_color(para_index, run_index, color)` | 设置运行颜色 |
| `set_run_highlight(para_index, run_index, highlight_color)` | 设置运行高亮 |
| `set_run_strike(para_index, run_index, strike)` | 设置运行删除线 |
| `set_run_font_from_xml(para, run_index, **font_properties)` | 从XML设置运行字体 |
| `set_run_size_from_xml(para, run_index, size)` | 从XML设置运行大小 |
| `set_run_bold_from_xml(para, run_index, bold)` | 从XML设置运行粗体 |
| `set_run_italic_from_xml(para, run_index, italic)` | 从XML设置运行斜体 |
| `set_run_underline_from_xml(para, run_index, underline_type)` | 从XML设置运行下划线 |
| `set_run_color_from_xml(para, run_index, color)` | 从XML设置运行颜色 |
| `set_run_highlight_from_xml(para, run_index, highlight_color)` | 从XML设置运行高亮 |
| `set_run_strike_from_xml(para, run_index, strike)` | 从XML设置运行删除线 |
| `update_run_style_from_xml(para, run_index, **style_properties)` | 从XML更新运行样式 |
| `update_run_style(para_index, run_index, **style_properties)` | 更新运行样式 |
| `get_runs_from_paragraph(paragraph_element)` | 从段落获取所有运行 |
| `get_run_style_from_element(run_element)` | 从元素获取运行样式 |
| `_extract_run_properties_from_element(rPr)` | 从元素提取运行属性 |

### 图片处理

| 方法名 | 描述 |
|-------|------|
| `extract_images_simple(output_dir)` | 简单提取所有图片 |
| `count_images_simple()` | 计算文档中图片数量 |
| `get_image_by_relation_id(relation_id)` | 通过关系ID获取图片 |
| `save_image_by_relation_id(relation_id, output_path)` | 通过关系ID保存图片 |
| `insert_image(para_index, run_index, position, image_path, width, height, description, wrap_text, new_page)` | 插入图片 |
| `insert_image_with_caption(para_index, image_path, caption_text, chapter_num, width, height, description, wrap_text, new_page, caption_style)` | 插入带标题的图片 |
| `insert_figure_caption(para_index, chapter_num, caption_text, auto_num, style_id, **style_properties)` | 插入图片标题 |
| `get_image_paragraphs_indices()` | 获取包含图片的段落索引 |
| `get_image_details()` | 获取所有图片的详细信息 |
| `remove_image_at_paragraph(paragraph_index, image_index)` | 移除指定段落中的图片 |

### 元素操作

| 方法名 | 描述 |
|-------|------|
| `element_to_dict(element_index, element_type)` | 将元素转换为字典 |
| `_elements_equal(elem1, elem2)` | 比较两个元素是否相等 |
| `insert_paragraph(element_index, position, text, **style_properties)` | 插入段落 |
| `insert_run(para_index, run_index, position, text, **style_properties)` | 插入文本运行 |
| `remove_element(element_index)` | 移除元素 |
| `remove_paragraph(para_index)` | 移除段落 |
| `remove_table(table_index)` | 移除表格 |
| `remove_elements_between(start_index, end_index)` | 移除两个元素之间的所有元素 |
| `remove_content_between_paragraphs(start_para_index, end_para_index)` | 移除两个段落之间的所有内容 |
| `get_element_index_from_paragraph_index(paragraph_index)` | 从段落索引获取元素索引 |
| `get_element_index_from_table_index(table_index)` | 从表格索引获取元素索引 |
| `get_paragraph_index_from_element_index(element_index)` | 从元素索引获取段落索引 |
| `get_table_index_from_element_index(element_index)` | 从元素索引获取表格索引 |
| `get_element_indices_by_type(element_type)` | 获取指定类型的所有元素索引 |
| `get_document_structure()` | 获取文档结构 |

### 目录和标题

| 方法名 | 描述 |
|-------|------|
| `insert_table_of_contents(element_index, position, title, heading_levels, style_id, hyperlinks, show_page_numbers, right_align_page_numbers, leader_char, title_font, title_style, headings, title_style_id)` | 插入目录 |
| `create_table_of_contents_from_headings(element_index, position, max_level, **toc_options)` | 从标题创建目录 |
| `update_toc_field(toc_para_index)` | 更新目录字段 |
| `get_heading_paragraphs()` | 获取所有标题段落 |
| `get_outline_level(para_index)` | 获取段落的大纲级别 |
| `insert_caption(element_index, caption_type, chapter_num, caption_text, position, auto_num, style_id, **style_properties)` | 插入标题 |
| `_insert_seq_field(para_index, seq_name)` | 插入序列字段 |

### 评论操作

| 方法名 | 描述 |
|-------|------|
| `extract_comments()` | 提取文档中的所有评论 |
| `_find_comment_references(comments_map)` | 查找评论引用 |
| `get_comment_at_paragraph(para_index)` | 获取指定段落的评论 |
| `get_comment_by_id(comment_id)` | 通过ID获取评论 |
| `add_comment(element_index, author, comment_text, element_type, run_index)` | 添加评论 |

### 页面控制

| 方法名 | 描述 |
|-------|------|
| `insert_page_break_before_paragraph(para_index)` | 在段落前插入分页符 |
| `insert_page_break(para_index, position)` | 在指定位置插入分页符 |

### 文档操作

| 方法名 | 描述 |
|-------|------|
| `update_document_xml()` | 更新文档XML |
| `save(output_path)` | 保存文档到指定路径 |

## 贡献

欢迎通过提交问题或拉取请求来贡献此项目。


