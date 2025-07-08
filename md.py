from docx_namespace import DocxElementParser
import os

from style_analyzer import StyleAnalyzer


def create_standard_table_example(docx_path, output_path):
    """
    创建一个符合标准要求的三线表格示例

    符合以下规范:
    - 样式：三线表（第一条线1.5磅，第二条线0.5磅，第三条线1.5磅），与上下文各空一行
    - 表格标题：按照章节用阿拉伯数字顺序编号，如"表1-1"
    - 标题字体：五号，中文宋体，外文Times New Roman，居中
    - 标题段落：多倍行距1.25，段前段后0行
    - 空格：表号与表名中间空2个英文半角空格
    - 表格内容字体：五号，中文宋体，外文Times New Roman，文字垂直居中
    - 表格内容段落：多倍行距1.25，段前段后0行

    参数:
        docx_path: 输入的Word文档路径
        output_path: 输出的Word文档路径
    """
    # 初始化文档解析器
    doc = DocxElementParser(docx_path)

    print("开始创建标准三线表格示例...")

    # 在文档末尾插入一个段落作为表格前的空行
    #
    # doc.insert_paragraph(element_index=-1, position='after', text='')
    #
    # # 插入一个3行4列的表格
    # table_data = [
    #     ["项目", "数值1", "数值2", "数值3"],
    #     ["指标A", "89.5", "76.2", "92.1"],
    #     ["指标B", "65.3", "88.7", "77.4"]
    # ]
    #
    # table_index = doc.insert_table(
    #     element_index=-1,
    #     position='after',
    #     rows=3,
    #     cols=4,
    #     data=table_data,
    #     width={'value': 8000, 'type': 'dxa'},  # 修改为字典格式
    #     alignment='center'  # 表格居中对齐
    # )

    # 设置表格样式 - 三线表
    # create_three_line_table函数会创建标准的三线表：
    # 1. 表格顶部线 - 1.5磅粗线
    # 2. 表头底部线 - 0.5磅细线
    # 3. 表格底部线 - 1.5磅粗线
    # 4. 其他所有内部边框线都会被移除
    # doc.create_three_line_table(table_index)

    # 自定义三线表样式（如果create_three_line_table功能不完整或需要更精确控制）
    # 下面的代码展示了如何手动设置三线表样式
    # print( doc.get_table_cell_paragraphs(0,0,0))
    # print(doc.get_paragraph_text(doc.get_table_cell_paragraphs(0,0,0)[0]))
    # 设置表格边框 - 移除所有边框
    table_index = 1
    rows, cols = doc.get_table_dimensions(table_index)

    # 1. 清空所有边框
    doc.set_table_borders(
        table_index,
        top={"val": "none"},
        bottom={"val": "none"},
        left={"val": "none"},
        right={"val": "none"},
        inside_h={"val": "none"},
        inside_v={"val": "none"}
    )
    for row in range(rows):
        doc.set_table_row_borders(
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
            doc.set_table_cell_borders(
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
        doc.set_table_cell_borders(
            table_index, 0, col,
            top={"val": "single", "sz": 12, "color": "000000", "space": "0"}
        )
    # 表头下边框
    for col in range(cols):
        doc.set_table_cell_borders(
            table_index, 0, col,
            bottom={"val": "single", "sz": 4, "color": "000000", "space": "0"}
        )
    # 底线：最下面一行所有单元格加底边框
    for col in range(cols):
        doc.set_table_cell_borders(
            table_index, rows - 1, col,
            bottom={"val": "single", "sz": 12, "color": "000000", "space": "0"}
        )
    # 设置表格文本样式
    # doc.update_table_text_style(
    #     1,
    #     size=24,
    #     font={'ascii': "Times New Roman", 'eastAsia': "宋体"},
    #     alignment="center",
    #     vertical_alignment="center",
    #     spacing={  # 注意这里是一个字典
    #         'line': 420,  # 行间距值（如300=1.5倍行距，240=单倍行距）
    #         'lineRule': 'exact',  # 行间距规则，常用'auto'
    #         'before': 0,  # 段前间距（单位1/20磅，0表示0磅）
    #         'after': 0  # 段后间距（单位1/20磅，0表示0磅）
    #     },
    #
    #     header_row_different=True,
    #     header_style={
    #         'size': 24,
    #         'font': {'ascii': "Arial", 'eastAsia': "黑体"},
    #         'alignment': "center",
    #         'vertical_alignment': "center"
    #     }
    # )
    # # 添加表格标题
    chapter_num = "1"  # 假设是第一章
    table_num = "1"    # 假设是第一个表格
    caption_text = "示例表格标题"

    # # 插入表格标题，并设置两个空格的间隔
    # # 注意：在"表1-1"后面添加两个半角空格，与表名分隔
    # doc.insert_table_caption(
    #     table_index,
    #     chapter_num=chapter_num,
    #     caption_text=f"  {caption_text}",  # 标题前加两个空格
    #     auto_num=True,  # 自动编号
    #     font={'ascii': "Times New Roman", 'eastAsia': "宋体"},
    #     size=21,  # 五号字，10.5磅=21（半磅单位）
    #     alignment="center",
    #     spacing={
    #         'line': 840,  # 1.5倍行距（240=单倍，300=1.25倍，360=1.5倍）
    #         'lineRule': 'exact', # 固定行距
    #         'beforeLines': 0,  # 段前0磅
    #         'afterLines': 0  # 段后0磅
    #     }
    # )
    #
    # # 在表格后插入一个空行
    # element_idx = doc.get_element_index_from_table_index(table_index)
    # doc.insert_paragraph(element_index=element_idx, position='after', text='')
    #
    # # 保存文档
    doc.save(output_path)
    print(f"表格示例已创建并保存到: {output_path}")
    print("注意：此示例创建了一个符合标准的三线表，具有以下特点：")
    print("- 表格上边框：1.5磅粗线")
    print("- 表头下边框：0.5磅细线")
    print("- 表格下边框：1.5磅粗线")
    print("- 表格内部无边框线")
    print("- 表格上下均有空行")
    print("- 表格标题格式：五号字体，宋体/Times New Roman，居中，行距1.25")
    print("- 表号与表名之间有两个半角空格分隔")

def test_set_paragraph_before_line(docx_path, output_path, para_index=0, before_line=400):
    doc = DocxElementParser(docx_path)
    # 设置第para_index段的段前行距（beforeLine，单位twip，1磅=20twip）
    doc.set_paragraph_spacing(para_index, beforeLines=before_line)

    doc.save(output_path)

if __name__ == "__main__":
    # 测试函数
    input_docx = "1.docx"
    doc=StyleAnalyzer(input_docx)
    style_info = doc.get_paragraph_complete_style_info(doc.elements[213]['element'])
    runs = doc.get_runs_from_paragraph(doc.elements[213]['element'])
    run_style_info = doc.get_run_complete_style_info(doc.elements[213]['element'], runs[3],
                                                          )
    direct_style = doc.get_paragraph_style_from_element(doc.elements[213]['element'])
    effective_style = style_info.get('effective_style', {})
    para_props = effective_style.get('paragraph_properties', {})
    print(run_style_info)
    # for element in doc.elements:
    #
    #   print(doc.get_paragraph_text(element['element']),element['index'])
