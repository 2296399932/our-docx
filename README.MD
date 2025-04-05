# Word文档解析与编辑工具 - 项目说明

## 项目概述

这个项目提供了一套全面的工具，用于解析、分析和编辑Word文档(.docx)文件。通过直接操作文档的XML结构，实现了对文档内容、样式和结构的精细控制，超越了现有库的功能限制。

项目分为两个主要类：
1. **DocxFile** - 基础文件操作类，负责docx文件的读取、解压和保存
2. **DocxElementParser** - 核心解析类，提供对文档XML结构的访问和修改功能

## 类结构

### DocxFile 类

处理Word文档的基础类，负责文件的读取、解压缩和保存操作。

**主要功能**：
- 将docx文件解压到内存中
- 解析主要XML部分（document.xml, styles.xml等）
- 处理文档关系和媒体文件
- 保存修改后的文档

### DocxElementParser 类

继承自DocxFile，提供对Word文档内容和结构的详细解析和编辑能力。

**核心属性**：
- `self.tree` - 文档的XML树
- `self.root` - XML树的根元素
- `self.elements` - 所有顶级元素的列表（段落、表格、节等）
- `self.paragraphs` - 所有段落元素的列表
- `self.tables` - 所有表格元素的列表
- `self.sections` - 所有节元素的列表

## 功能分类

### 1. 文档结构解析

| 函数名 | 描述 |
|-------|------|
| `get_structured_body_elements()` | 解析文档body中的所有顶级元素 |
| `get_all_paragraphs()` | 获取文档中所有段落 |
| `get_all_tables()` | 获取文档中所有表格 |
| `find_elements_by_tag()` | 根据XML标签查找元素 |
| `print_full_xml()` | 打印完整的文档XML结构 |

### 2. 文本内容提取

| 函数名 | 描述 |
|-------|------|
| `get_paragraph_text()` | 获取段落文本内容 |
| `get_element_text()` | 获取指定元素的文本内容 |
| `get_all_text()` | 获取文档所有文本内容 |
| `get_element_run_text()` | 获取元素中所有文本运行的文本 |
| `get_run_text()` | 获取特定文本运行的文本内容 |

### 3. 段落样式解析

| 函数名 | 描述 |
|-------|------|
| `extract_paragraph_style()` | 提取段落的所有样式属性 |
| `get_paragraph_alignment()` | 获取段落对齐方式 |
| `get_paragraph_indentation()` | 获取段落缩进信息 |
| `get_paragraph_spacing()` | 获取段落间距信息 |
| `get_paragraph_borders()` | 获取段落边框信息 |
| `get_paragraph_shading()` | 获取段落背景填充信息 |
| `get_paragraph_numbering()` | 获取段落编号信息 |
| `get_paragraph_font()` | 获取段落字体信息 |
| `format_paragraph_style()` | 格式化显示段落样式信息 |

### 4. 文本运行样式解析

| 函数名 | 描述 |
|-------|------|
| `get_run_style()` | 获取文本运行的样式信息 |
| `get_run_font()` | 获取文本运行的字体信息 |
| `get_run_size()` | 获取文本运行的字号信息 |
| `get_run_formatting()` | 获取文本运行的格式信息(粗体、斜体等) |
| `get_run_color()` | 获取文本运行的颜色信息 |
| `format_run_style()` | 格式化显示文本运行样式信息 |

### 5. 表格处理

| 函数名 | 描述 |
|-------|------|
| `extract_table_content()` | 提取表格内容 |
| `export_table_to_file()` | 导出表格到Excel或CSV文件 |
| `export_all_tables()` | 导出所有表格 |
| `get_table_style()` | 获取表格样式信息 |
| `format_table_style()` | 格式化显示表格样式信息 |

### 6. 图片处理

| 函数名 | 描述 |
|-------|------|
| `extract_images_simple()` | 提取文档中的所有图片 |
| `count_images_simple()` | 统计文档中的图片数量 |
| `get_image_by_relation_id()` | 通过关系ID获取图片 |
| `save_image_by_relation_id()` | 保存指定关系ID的图片到文件 |
| `insert_image()` | 在文档中插入图片 |

### 7. 样式修改

#### 7.1 段落样式修改

| 函数名 | 描述 |
|-------|------|
| `set_paragraph_style_id()` | 设置段落样式ID |
| `set_paragraph_alignment()` | 设置段落对齐方式 |
| `set_paragraph_indentation()` | 设置段落缩进 |
| `set_paragraph_spacing()` | 设置段落间距 |
| `set_paragraph_borders()` | 设置段落边框 |
| `set_paragraph_shading()` | 设置段落背景填充 |
| `set_paragraph_numbering()` | 设置段落编号 |
| `set_paragraph_font()` | 设置段落字体属性 |
| `update_paragraph_style()` | 更新段落多个样式属性 |
| `set_paragraph_spacing_preserve_style()` | 设置段落间距并保留样式ID |

#### 7.2 文本运行样式修改（修改段落中所有文本）

| 函数名 | 描述 |
|-------|------|
| `set_paragraph_runs_font()` | 设置段落中所有文本运行的字体 |
| `set_runs_bold()` | 设置段落中所有文本为粗体 |
| `set_runs_italic()` | 设置段落中所有文本为斜体 |
| `set_runs_underline()` | 设置段落中所有文本的下划线 |
| `set_runs_color()` | 设置段落中所有文本的颜色 |
| `set_runs_size()` | 设置段落中所有文本的字号 |
| `set_runs_highlight()` | 设置段落中所有文本的高亮颜色 |
| `update_runs_style()` | 更新段落中所有文本的多个样式属性 |

#### 7.3 单个文本运行样式修改（修改段落中特定文本）

| 函数名 | 描述 |
|-------|------|
| `set_run_font()` | 设置特定文本运行的字体 |
| `set_run_size()` | 设置特定文本运行的字号 |
| `set_run_bold()` | 设置特定文本运行为粗体 |
| `set_run_italic()` | 设置特定文本运行为斜体 |
| `set_run_underline()` | 设置特定文本运行的下划线 |
| `set_run_color()` | 设置特定文本运行的颜色 |
| `set_run_highlight()` | 设置特定文本运行的高亮颜色 |
| `update_run_style()` | 更新特定文本运行的多个样式属性 |

### 8. 文档结构编辑

| 函数名                     | 描述 |
|-------------------------|------|
| `insert_paragraph()`    | 在文档中插入新段落 |
| `insert_paragraph()`    | 在指定位置插入新文本运行 |
| `insert_image()`        | 在文档中插入图片 |
| `update_document_xml()` | 更新文档XML |
| `save()`                | 保存修改后的文档 |

## 使用示例

### 基本解析操作

```python
# 创建解析器对象
parser = DocxElementParser('document.docx')

# 获取文档结构
print(f"文档共有 {len(parser.paragraphs)} 个段落")
print(f"文档共有 {len(parser.tables)} 个表格")

# 提取文本内容
for i, para in enumerate(parser.paragraphs):
    print(f"段落 {i}: {parser.get_paragraph_text(para['element'])}")

# 提取表格内容
for i, table in enumerate(parser.tables):
    table_content = parser.extract_table_content(table['element'])
    print(f"表格 {i} 有 {len(table_content)} 行, {len(table_content[0]) if table_content else 0} 列")
```

### 样式分析

```python
# 分析段落样式
para_index = 5
style_info = parser.extract_paragraph_style(parser.paragraphs[para_index]['element'])
print(parser.format_paragraph_style(style_info))

# 分析文本运行样式
run_style = parser.get_run_style(para_index, 0, element_type="paragraphs")
print(parser.format_run_style(run_style))

# 分析表格样式
table_style = parser.get_table_style(0)
print(parser.format_table_style(table_style))
```

### 内容修改

```python
# 修改段落对齐方式
parser.set_paragraph_alignment(5, "center")

# 修改段落字体
parser.set_paragraph_font(6, 
    eastAsia="黑体", 
    ascii="Times New Roman",
    size=28,  # 14磅
    bold=True,
    color="FF0000"  # 红色
)

# 修改特定文本运行
parser.set_run_bold(10, 2, True)  # 将段落10的第3个文本运行设为粗体
parser.set_run_color(10, 2, "0000FF")  # 将其颜色设为蓝色

# 插入新段落
new_para_index = parser.insert_paragraph(
    element_index=100,
    position='after',
    text="这是新插入的段落",
    style_id="Heading1",
    alignment="center"
)

# 插入图片
parser.insert_image(
    para_index=new_para_index,
    image_path="logo.png",
    width=5,  # 5厘米宽
    height=3,  # 3厘米高
    description="公司标志"
)

# 保存文档
parser.save('output.docx')
```

## 注意事项

1. 所有索引操作均支持负索引（如-1表示最后一个元素）
2. 修改文档后，务必调用`save()`方法保存更改
3. 样式修改会实时应用到XML结构中
4. 插入图片需要安装Pillow库：`pip install Pillow`

## 贡献与扩展

该项目设计为模块化架构，可以方便地扩展新功能。常见的扩展方向包括：

1. 添加对更多文档元素的支持（如页眉、页脚）
2. 增强表格编辑功能
3. 添加样式模板管理
4. 实现更复杂的文档合并和拆分功能

## 环境要求

- Python 3.6+
- 推荐安装的扩展库：
  - Pillow (用于图片处理)
  - pandas (用于表格导出)
  - lxml (可选，用于更快的XML处理)