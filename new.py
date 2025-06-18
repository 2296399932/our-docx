from concurrent.futures import ThreadPoolExecutor, as_completed
from auto_fix_style_errors import auto_fix_style_errors
from compare_styles import compare_styles
from gemin2 import docx_first
import json

from table_style_modifier import analyze_document_styles_with_ai



docx_path = "智算工程学院毕业设计（论文）2025届(new)(2).docx"
output_path = "output1.docx"
style_mapping_path = "document_style_mapping.json"

def run_docx_first():
    return docx_first(docx_path)
#
# def run_analyze_document_styles():
#     return analyze_document_styles_with_ai(docx_path_temp)

# 原并发部分删除，改为顺序执行
all_classifications = docx_first(docx_path)
# results = analyze_document_styles_with_ai(docx_path_temp)
# with open("document_classification_results.json", "r", encoding="utf-8") as f:
#     all_classifications = json.load(f)
# 读取API参数JSON文件
with open("智算工程学院毕业设计（论文）模板2025届(1)_api_params.json", "r", encoding="utf-8") as f:
    api_params = json.load(f)

# 等待两个都完成后再继续
# compare_styles(docx_path, all_classifications, style_mapping_path, api_params)
fixed_file = auto_fix_style_errors(
    docx_path,
    all_classifications,
    style_mapping_path,
    api_params,
    output_path,
    interactive=False,
    clean_statistics=True  # 默认进行清理以提高效率
)