from docx_namespace import DocxElementParser
from table_style_modifier import extract_and_print_all_content

doc=DocxElementParser('1.docx')
# 获取目标元素
doc_content = extract_and_print_all_content("智算工程学院毕业设计（论文）模板2025届(1).docx")
