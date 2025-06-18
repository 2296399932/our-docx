import os
import json
# from openai import OpenAI
from google import genai
from google.genai import types
from docx_namespace import DocxElementParser
import re

# 配置Google Gemini API
client = genai.Client(api_key="AIzaSyBqmGlZHCAFViUE0Jw-ox_Hi171dCK-XQw")



# 分析文档结构
def analyze_document_structure(doc_path):
    # 打开Word文档
    doc = DocxElementParser(doc_path)

    # 收集文档元素信息
    document_elements = []

    # 遍历所有元素
    for index, item in enumerate(doc.elements):
        element = item.get('element')
        element_info = {"index": index}

        # 检查元素是否是段落类型
        if item.get('type') == 'paragraph':
            # 获取段落文本
            text = doc.get_paragraph_text(element)
            element_info["type"] = "paragraph or heading_level"
            # 限制文本长度为50字符
            if len(text) > 50:
                element_info["text"] = text
            else:
                element_info["text"] = text

            # 检查段落是否有样式信息
            pPr = element.find('.//{*}pPr')
            if pPr is not None:
                pStyle = pPr.find('.//{*}pStyle')
                if pStyle is not None:
                    style_val = pStyle.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val', "")
                    element_info["style_id"] = style_val

            # # 检查空段落是否包含图片
            # if not text.strip():
            #     # 检查是否包含drawing元素(内联图片)
            #     drawings = element.findall('.//{*}drawing')
            #     # 检查是否包含object元素(嵌入对象)
            #     objects = element.findall('.//{*}object')
            #     # 检查是否包含pict元素(VML图形)
            #     picts = element.findall('.//{*}pict')
            #
            #     if drawings or objects or picts:
            #         # 计算图片数量
            #         image_count = len(drawings) + len(objects) + len(picts)
            #         element_info["type"] = "image_paragraph"
            #         element_info["image_count"] = image_count
            #     else:
            #         element_info["type"] = "paragraph or heading_level"
            # else:
            #     element_info["type"] = "paragraph or heading_level"

        # 如果是表格元素
        elif item.get('type') == 'table':
            element_info["type"] = "table"
            # 获取表格行数和列数
            rows = element.findall('.//{*}tr')
            if rows:
                element_info["row_count"] = len(rows)
                cells = rows[0].findall('.//{*}tc')
                element_info["column_count"] = len(cells) if cells else 0

        # 其他元素类型
        else:
            element_type = item.get('type')
            element_info["type"] = element_type

        document_elements.append(element_info)

    return document_elements


def identify_document_sections(document_elements):
    """
    步骤1: 识别文档的主要部分及其索引范围

    Args:
        document_elements: 文档元素信息

    Returns:
        dict: 文档各部分及其索引范围，格式如 {"cover": [start, end], "toc": [start, end], ...}
    """
    # 准备发送给大模型的数据
    elements_json = json.dumps(document_elements, ensure_ascii=False, indent=2)

    # 构建提示词
    system_prompt = """
    你是一个专业的学术论文结构分析专家。请分析提供的Word文档结构信息，并识别文档的主要部分及其索引范围。

    学术论文通常包含以下主要部分：
    1. 封面部分(cover)：包含论文标题、作者信息等
    2. 目录部分(toc)：包含目录标题和目录项
    3. 中文摘要部分(abstract_zh)：中文摘要标题、内容和关键词
    4. 英文摘要部分(abstract_en)：英文摘要标题、内容和关键词
    5. 正文部分(body)：包含论文的主要内容、各级标题和段落
    6. 参考文献部分(references)：包含参考文献列表
    7. 致谢部分(acknowledgements)：感谢他人的贡献
    8. 附录部分(appendix)：补充材料

    请仔细分析文档元素，识别每个部分的起始和结束索引，返回JSON格式的结果。
    只需要识别文档中存在的部分，不需要强制包含所有可能的部分。
    确保所有索引都被包含在某个部分中，不要有遗漏。
    """

    user_prompt = f"""
    请分析以下Word文档的结构信息，并识别文档的主要部分及其索引范围。

    返回格式如下：
    {{
      "cover": [起始索引, 结束索引],  // 封面部分
      "toc": [起始索引, 结束索引],    // 目录部分(包含如第四章等目录项)
      "abstract_zh": [起始索引, 结束索引],  // 中文摘要部分
      "abstract_en": [起始索引, 结束索引],  // 英文摘要部分
      "body": [起始索引, 结束索引],   // 正文部分
      "references": [起始索引, 结束索引],  // 参考文献部分
      "acknowledgements": [起始索引, 结束索引],  // 致谢部分
      "appendix": [起始索引, 结束索引]  // 附录部分
    }}

    注意：
    1. 只需包含文档中实际存在的部分
    2. 各部分的索引范围应该互不重叠，并且覆盖文档的所有有效内容
    3. 每个部分包含该部分的所有内容，包括标题和正文
    4. 确保目录部分包含所有目录项,外文，致谢，附录等都有可能算目录项

    下面是文档元素信息：
    {elements_json}

    请仅返回JSON格式的分析结果，不需要解释或说明。
    """

    try:
        # 使用Google Gemini API调用模型
        combined_prompt = f"{system_prompt}\n\n{user_prompt}"
        response = client.models.generate_content(
            model="gemini-2.5-flash-preview-05-20",
            contents=[combined_prompt],
            config=types.GenerateContentConfig(
                max_output_tokens=30000,
                temperature=0.3
            )
        )
        print(f"1212{response}")
        # 提取响应内容
        response_content = response.text

        # 解析JSON结果
        try:
            # 尝试直接解析
            result = json.loads(response_content)
            print("文档结构分析结果:", result)
            result = filter_indices(result)
            return result
        except json.JSONDecodeError:
            # 如果失败，尝试提取JSON部分
            import re
            json_match = re.search(r'```json\s*([\s\S]*?)\s*```', response_content)
            if json_match:
                try:
                    result = json.loads(json_match.group(1))
                    result = filter_indices(result)
                    return result
                except:
                    pass

            # 如果仍然失败，尝试找到第一个{和最后一个}
            first_brace = response_content.find('{')
            last_brace = response_content.rfind('}')
            if first_brace != -1 and last_brace != -1:
                try:
                    result = json.loads(response_content[first_brace:last_brace + 1])
                    result = filter_indices(result)
                    return result
                except:
                    pass

            print("无法解析文档部分识别结果")
            print("原始返回:", response_content)

            # 后备方案：创建一个基本的文档结构
            max_index = max(e["index"] for e in document_elements)
            fallback_result = {"document": [0, max_index]}
            print("使用后备文档结构:", fallback_result)
            return fallback_result

    except Exception as e:
        print(f"识别文档部分时发生错误: {e}")
        # 后备方案：创建一个基本的文档结构
        max_index = max(e["index"] for e in document_elements)
        fallback_result = {"document": [0, max_index]}
        print("使用后备文档结构:", fallback_result)
        return fallback_result


def classify_document_section(document_elements, section_name, start_index, end_index):
    """
    步骤2: 详细分类文档指定部分内的元素

    Args:
        document_elements: 文档元素信息
        section_name: 部分名称
        start_index: 起始索引
        end_index: 结束索引

    Returns:
        dict: 该部分内各类元素的索引分类
    """
    # 提取该部分的元素
    section_elements = [e for e in document_elements if start_index <= e["index"] <= end_index]
    print(f"正在分析{section_elements}部分...")
    if not section_elements:
        print(f"警告：{section_name}部分没有找到有效元素")
        return {}

    # 准备发送给大模型的数据
    elements_json = json.dumps(section_elements, ensure_ascii=False, indent=2)

    # 根据不同部分构建不同的分类标准和提示词
    classification_standards = {
        "toc": """
        目录部分通常包含以下元素：
        11.目录标题(toc_title)：如"目录"或"CONTENTS"
        2. 一级目录项(toc_level_1)：如"第一章 绪论"、"致谢"、"参考文献"等
        3. 二级目录项(toc_level_2)：如"11.11 研究背景"
        4. 三级目录项(toc_level_3)：如"11.11.11 问题提出"



        特别说明：目录项和页码通常写在同一行，如"第四章 系统的实现"，这里"第四章 系统的实现"是一级目录项，

        """,

        "abstract_zh": """
        中文摘要部分通常包含以下元素：
        1, 摘要标题(abstract_zh_title)：如"摘要"
        2. 摘要内容(abstract_zh_content)
        3. 关键词标签(keywords_zh_label)：如"关键词："
        4. 关键词内容(keywords_zh_content)
        5. 摘要备注(abstract_notes)：如"无页码"等注释
        """,

        "abstract_en": """
        英文摘要部分通常包含以下元素：
        1. 摘要标题(abstract_en_title)：如"Abstract"
        2. 摘要内容(abstract_en_content)
        3. 关键词标签(keywords_en_label)：如"Keywords:"
        4. 关键词内容(keywords_en_content)
        5. 摘要备注(abstract_notes)：如"无页码"等注释
        """,

        "body": """
        【特别注意】：优先根据标题的数字格式进行级别判断：
        - 如果标号格式是\"X.Y\"（只有两级数字，如\"5.11\"），则为二级标题
        - 如果标号格式是\"X.Y.Z\"（有三级数字，如\"5.11.2\"），则为三级标题
        - 只有无法通过数字格式判断时，再参考元素的style_id键（style_id是2的为heading_level_1，3为heading_level_2，4为heading_level_3）
        正文部分通常包含以下元素：
        1. 一级标题(heading_level_1)：如\"第一章 绪论\"、\"第二章 相关技术\"等。一级标题通常以\"第X章\"开头，但也有可能标题被忘记加上编号的情况可能自行判断。
        2. 二级标题(heading_level_2)：如\"11.11 研究背景\"、\"2.3 系统架构\"等。二级标题通常采用\"X.Y+内容\"的格式，但也有可能标题被忘记加上编号的情况可能自行判断。
        3. 三级标题(heading_level_3)：如\"11.11.11 问题提出\"等。三级标题通常采用\"X.Y.Z+内容\"的格式，但也有可能标题被忘记加上编号的情况可能自行判断。
        4. 正文段落(body_paragraph)：普通的正文内容。
        5. 图片(image)：任何图片段落。
        6. 图表标题(figure_caption)：图片的标题或说明文字。
        7. 表格(table)：表格元素。
        8. 表格标题(table_caption)：表格的标题或说明文字。
        9. 公式(equation)：数学公式。
        10. 引用(citation)：引用其他文献的内容。
        11. 代码(code)：如编程代码块。
        

        正确区分二级和三级标题的关键在于标号格式：
        - 如果标号格式是\"X.Y\"（只有两级数字，如\"5.11\"），则为二级标题
        - 如果标号格式是\"X.Y.Z\"（有三级数字，如\"5.11.2\"），则为三级标题
        - 请仔细检查每个标题的格式，确保准确分类
        【再次强调】：只有无法通过数字格式判断时，再参考style_id！
        """,

        "references": """
        参考文献部分通常包含以下元素：
        1. 参考文献标题(references_title)：如"参考文献"或"References"
        2. 参考文献条目(reference_item)
        """,

        "acknowledgements": """
        致谢部分通常包含以下元素：
        1. 致谢标题(acknowledgements_title)：如"致谢"
        2. 致谢内容(acknowledgements_content)
        """,

        "appendix": """
        附录部分通常包含以下元素：
        1.附录标题(appendix_title)：如"附录"或"Appendix"、"外文原文及译文"等
        2. 附录子标题(appendix_subtitle)
        3. 附录内容(appendix_content)
        4. 附录图表(appendix_figure)
        5. 附录表格(appendix_table)
        """
    }

    # 获取当前部分的分类标准，如果没有则使用通用标准
    standard = classification_standards.get(section_name, """
    请根据内容特征将元素分类为合适的类别，如标题、正文、图片等。请确保所有元素都被分类，包括特殊格式的文本和系统元素。
    """)

    # 构建提示词
    system_prompt = f"""
    你是一个专业的学术论文结构分析专家。请详细分析提供的文档"{section_name}"部分的元素，并将它们分类。

    {standard}

    请根据内容特征仔细分析每个元素，返回JSON格式的分类结果。每个类别包含该类型元素的索引列表。
    特别注意：你必须确保对提供的每一个元素进行分类，不遗漏任何索引。
    """

    # 构建用户提示词，为body部分添加特殊处理
    if section_name == "body":
        user_prompt = f"""
        请详细分析以下"body"部分的元素，并将它们分类，特别注意标题级别的分类。
     
        原始索引范围：{start_index} 到 {end_index}

        元素信息：
        {elements_json}

        请将这些元素分类为适当的类别，返回JSON格式的结果。例如：
        {{
          "heading_level_1": [索引列表],
          "heading_level_2": [索引列表],
          "heading_level_3": [索引列表],
          "body_paragraph": [索引列表],
          ...其他类别
        }}

        重要指示：
        1. 每个元素必须被分类，不能遗漏任何索引
        2. 特别注意标题级别的正确分类，这是最重要的任务

        3. 对于特殊元素，如图片、表格、sectPr等，也必须适当分类

        请检查每个可能的标题文本，验证它们的格式是否符合上述规则，然后进行准确分类。
        请仅返回JSON格式的分类结果，不需要解释或说明。
        """
    else:
        user_prompt = f"""
        请详细分析以下"{section_name}"部分的元素，并将它们分类。

        原始索引范围：{start_index} 到 {end_index}

        元素信息：
        {elements_json}

        请将这些元素分类为适当的类别，返回JSON格式的结果。例如：
        {{
          "类别1": [索引列表],
          "类别2": [索引列表],
          ...
        }}

        重要指示：
        1. 每个元素必须被分类，不能遗漏任何索引
        2. 根据元素的内容和格式特征进行分类
        3. 对于特殊元素（如sectPr、空图片段落等），也必须分配适当的类别
        4. 特别注意：确保所有索引都被分类！

        请仅返回JSON格式的分类结果，不需要解释或说明。
        """

    try:
        # 使用Gemini模型
        print(f"正在进行第一次分类: {section_name}部分...")
        combined_prompt = f"{system_prompt}\n\n{user_prompt}"
        response = client.models.generate_content(
            model="gemini-2.5-flash-preview-05-20",
            contents=[combined_prompt],
            config=types.GenerateContentConfig(
                max_output_tokens=30000,
                temperature=0.2
            )
        )

        # 提取响应内容
        original_content = response.text
        print('响应内容为：', original_content)
        import re
        match = re.search(r'```json\s*([\s\S]*?)\s*```', original_content)
        if match:
            response_content = match.group(1)  # 提取匹配的JSON文本
            print("成功从Markdown代码块中提取JSON")
        else:
            # 如果没有找到JSON代码块，尝试直接查找JSON对象
            first_brace = original_content.find('{')
            last_brace = original_content.rfind('}')
            if first_brace != -1 and last_brace != -1:
                response_content = original_content[first_brace:last_brace + 1]
                print("从原始响应中提取JSON对象")
            else:
                response_content = original_content
                print("使用完整的原始响应")

        # 解析JSON结果
        try:
            # 尝试直接解析
            initial_result = json.loads(response_content)
            print(initial_result)
            # 检查是否有遗漏的索引
            section_indices = set(e["index"] for e in section_elements)
            classified_indices = set()
            for indices in initial_result.values():
                for idx in indices:
                    classified_indices.add(idx)

            missing_indices = section_indices - classified_indices
            if missing_indices:
                print(
                    f"警告：{section_name}部分中，有{len(missing_indices)}个索引未被分类: {sorted(list(missing_indices))}")
                # 将未分类的索引添加到"other"类别
                other_category = f"other_{section_name}"
                if other_category not in initial_result:
                    initial_result[other_category] = []
                initial_result[other_category].extend(list(missing_indices))

            # 直接返回大模型一次性分类结果（不再进行二次分类和验证）
            result = initial_result
            result = filter_indices(result)
            return result
        except json.JSONDecodeError:
            print(f"无法解析{section_name}部分的分类结果")
            print("原始返回:", response_content)

            # 创建一个基本分类，确保所有元素都被分类
            basic_result = {f"all_{section_name}_elements": [e["index"] for e in section_elements]}
            return basic_result

    except Exception as e:
        print(f"分类{section_name}部分时发生错误: {e}")
        # 创建一个基本分类，确保所有元素都被分类
        basic_result = {f"all_{section_name}_elements": [e["index"] for e in section_elements]}
        return basic_result


def validate_classification(all_classifications, document_elements):
    """验证最终分类结果，检查是否有缺失的索引"""
    # 收集所有已分类的索引
    classified_indices = set()
    for category, indices in all_classifications.items():
        for idx in indices:
            classified_indices.add(idx)

    # 收集所有应该被分类的索引
    expected_indices = set(element["index"] for element in document_elements)

    # 检查是否有缺失的索引
    missing_indices = expected_indices - classified_indices

    # 检查是否有多余的索引（不在原文档中）
    extra_indices = classified_indices - expected_indices

    return sorted(list(missing_indices)), sorted(list(extra_indices))


def filter_indices(result):
    filtered = {}
    for k, v in result.items():
        filtered[k] = [i for i in v if isinstance(i, int)]
    return filtered


# 主程序
def docx_first(doc_path=""):
    print(f"开始分析文档: {doc_path}")
    document_elements = analyze_document_structure(doc_path)



    print("\n步骤1: 识别文档主要部分及其范围...")
    print(document_elements)
    # 只用过滤后的有效元素识别文档结构
    document_sections = identify_document_sections(document_elements)

    if not document_sections:
        print("识别文档部分失败，无法继续处理。")
        return

    # 打印识别出的文档部分
    print("\n已识别出以下文档部分:")
    for section_name, section_range in document_sections.items():
        print(f"  {section_name}: 索引范围 {section_range[0]} - {section_range[1]}")

    # 存储所有部分的分类结果
    all_classifications = {}

    print("\n步骤2: 详细分类每个部分的内容...")
    for section_name, section_range in document_sections.items():
        # 特殊处理正文部分（如果很长，则分段处理）
        if section_name == "body" and (section_range[1] - section_range[0] > 100):
            print(f"\n正文部分较长，将分为5个小段进行处理，但结果将合并呈现。")

            # 计算每个子部分的大小
            start_idx = section_range[0]
            end_idx = section_range[1]
            total_elements = end_idx - start_idx + 1
            segments = 5  # 将正文分为40段
            part_size = max(1, total_elements // segments)  # 确保每段至少有一个元素

            # 存储每个部分的分类结果
            body_parts_classifications = {}

            # 处理每个子部分
            for i in range(segments):
                part_start = start_idx + i * part_size
                # 最后一段特殊处理，确保包含所有剩余元素
                if i == segments - 1:
                    part_end = end_idx
                else:
                    part_end = min(start_idx + (i + 1) * part_size - 1, end_idx)

                # 如果当前段已经超出范围，则跳出循环
                if part_start > end_idx:
                    break

                print(f"  处理正文子部分 {i + 1}/{segments} (索引 {part_start} - {part_end})...")
                part_classification = classify_document_section(
                    document_elements,
                    "body",  # 使用实际的部分名称作为分类标准
                    part_start,
                    part_end
                )

                # 存储这个部分的分类结果
                for category, indices in part_classification.items():
                    if category not in body_parts_classifications:
                        body_parts_classifications[category] = []
                    body_parts_classifications[category].extend(indices)

            # 合并所有部分的结果
            print(f"\n合并正文部分分类结果...")
            for category, indices in body_parts_classifications.items():
                # 去重并排序
                unique_sorted_indices = sorted(list(set(indices)))
                full_category = f"body_{category}" if category != "body" else category
                all_classifications[full_category] = unique_sorted_indices
                print(f"  - {full_category}: {len(unique_sorted_indices)} 个元素")

        else:
            # 正常处理其他部分
            print(f"\n分析 {section_name} 部分 (索引 {section_range[0]} - {section_range[1]})...")
            section_classification = classify_document_section(
                document_elements,
                section_name,
                section_range[0],
                section_range[1]
            )

            if section_classification:
                # 打印该部分的分类结果摘要
                print(f"  {section_name} 部分分类完成，识别出 {len(section_classification)} 类元素:")
                for category, indices in section_classification.items():
                    indices_str = str(indices)
                    if len(indices_str) > 50:
                        indices_str = indices_str[:47] + "..."
                    print(f"    - {category}: {len(indices)} 个元素 {indices_str}")

                # 将该部分的分类结果添加到总分类结果中
                for category, indices in section_classification.items():
                    full_category = f"{section_name}_{category}" if category != section_name else category
                    all_classifications[full_category] = indices

    print("\n步骤3: 验证分类结果完整性...")
    missing_indices, extra_indices = validate_classification(all_classifications, document_elements)

    if missing_indices:
        print(f"\n警告：分类结果中缺失以下索引: {missing_indices}")
        print("这些索引对应的内容:")
        for idx in missing_indices:
            for element in document_elements:
                if element["index"] == idx:
                    element_type = element.get("type", "unknown")
                    text = element.get("text", "")
                    print(f"  索引 {idx}: [{element_type}] {text}")
                    break

        # 将未分类的索引添加到"other"类别
        if missing_indices:
            if "other_elements" not in all_classifications:
                all_classifications["other_elements"] = []
            all_classifications["other_elements"].extend(missing_indices)
            print(f"已将未分类的索引添加到'other_elements'类别")
    else:
        print("\n验证通过：所有索引都已被分类")

    if extra_indices:
        print(f"\n警告：分类结果中包含无效索引: {extra_indices}")

    # 保存分类结果
    with open("document_classification_results.json", "w", encoding="utf-8") as f:
        json.dump(all_classifications, f, ensure_ascii=False, indent=2)
    print("\n分类结果已保存到 document_classification_results.json")
    return all_classifications

#
if __name__ == '__main__':
    docx_path = "sdj-毕业论文(1).docx"
    docx_first(docx_path)



