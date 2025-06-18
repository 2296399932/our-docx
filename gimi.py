from style_analyzer import *

doc_path = "1_fixed.docx"
# 创建样式分析器实例获取完整样式信息
style_analyzer = StyleAnalyzer(doc_path)

runs=style_analyzer.get_runs_from_paragraph(style_analyzer.elements[195]['element'])
print(style_analyzer.get_run_complete_style_info(style_analyzer.elements[195]['element'],runs[0]))
