from docx_namespace import DocxElementParser

doc = DocxElementParser("sdj-毕业论文(1).docx")

elements = doc.get_element()  # 获取所有元素（段落、表格等）

for idx, element in enumerate(elements):
    text = doc.get_element_text(idx)
    print(f"元素索引: {idx}, 内容: {text}")
