#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
清理样式统计结果的工具函数，用于优化标记错误和批注
"""

import json
import copy
from collections import defaultdict

def clean_style_statistics(statistics):
    """
    清理样式统计结果，减少重复的错误标记，确保精确处理run级别错误
    
    完全重新设计的逻辑：
    1. 完全禁止合并不同run的错误
    2. 对于run元素，只保留那些真正有"error"的元素
    3. 对于只有属性缺失的run，进行智能判断
    
    参数:
        statistics: 原始统计结果
        
    返回:
        dict: 清理后的统计结果
    """
    # 深拷贝统计结果，避免修改原始数据
    cleaned = copy.deepcopy(statistics)
    
    if 'elements' not in cleaned:
        return cleaned
    
    # 首先过滤掉没有结果的元素
    cleaned['elements'] = [elem for elem in cleaned['elements'] if elem.get('result')]
    
    # 创建一个新的元素列表
    final_elements = []
    seen_positions = {}  # 用于跟踪已处理的位置
    
    for element in cleaned['elements']:
        element_type = element.get('type')
        index = element.get('index')
        
        # 创建位置键，确保每个元素位置唯一
        position_key = None
        if element_type == 'paragraph':
            position_key = f"paragraph_{index}"
        elif element_type == 'run':
            if isinstance(index, tuple) and len(index) == 2:
                para_idx, run_idx = index
                position_key = f"run_{para_idx}_{run_idx}" 
        elif element_type == 'table':
            position_key = f"table_{index}"
        
        # 跳过已处理的位置
        if position_key in seen_positions:
            continue
        
        # 对run元素进行特殊处理
        if element_type == 'run':
            # 检查是否有真正的错误（带error字段），而不只是缺失
            has_real_error = False
            filtered_results = []
            
            for result in element.get('result', []):
                filtered_result = {}
                for attr, values in result.items():
                    # 仅当有error字段时视为真正的错误
                    if 'error' in values:
                        has_real_error = True
                        filtered_result[attr] = values
                
                if filtered_result:
                    filtered_results.append(filtered_result)
            
            # 如果有真正的错误，更新结果并添加到最终列表
            if has_real_error:
                element['result'] = filtered_results
                seen_positions[position_key] = True
                final_elements.append(element)
            else:
                # 对于没有真正错误的run，检查是否有缺失的关键属性
                # 例如，检查是否缺少size属性和其他可能影响显示的属性
                has_critical_missing = False
                missing_attrs = []
                
                for result in element.get('result', []):
                    for attr, values in result.items():
                        success_value = values.get('success', values.get('scuccess'))
                        if success_value is not None and 'error' not in values:
                            if attr in ['size', 'font_ascii', 'font_chinese', 'bold']:
                                has_critical_missing = True
                                missing_attrs.append({attr: values})
                
                # 如果有关键缺失，添加这些属性并保留元素
                if has_critical_missing:
                    element['result'] = missing_attrs
                    seen_positions[position_key] = True
                    final_elements.append(element)
        else:
            # 非run元素直接添加到最终列表
            seen_positions[position_key] = True
            final_elements.append(element)
    
    # 更新元素列表
    cleaned['elements'] = final_elements
    
    return cleaned

def merge_run_errors(runs):
    """
    合并同一段落中多个run的相同类型错误
    
    参数:
        runs: 同一段落中的run元素列表
        
    返回:
        list: 合并后的错误列表
    """
    error_types = defaultdict(list)
    
    # 收集所有错误类型
    for run in runs:
        for error in run.get('result', []):
            for attr, values in error.items():
                # 使用属性名作为键
                error_types[attr].append(values)
    
    # 合并相同类型的错误
    merged_errors = []
    for attr, values_list in error_types.items():
        # 如果有多个同类型错误，只保留一个代表性的
        if values_list:
            # 检查是否所有错误都有相同的期望值和类似的当前值
            reference = values_list[0]
            expected_value = reference.get('scuccess')
            
            # 检查所有值是否类似
            similar = True
            for vals in values_list[1:]:
                if vals.get('scuccess') != expected_value:
                    similar = False
                    break
            
            # 如果类似，只保留一个代表性错误
            if similar:
                merged_errors.append({attr: reference})
            else:
                # 如果不类似，保留所有不同的错误
                seen = set()
                for vals in values_list:
                    error_key = f"{attr}:{vals.get('scuccess')}:{vals.get('error', '')}"
                    if error_key not in seen:
                        seen.add(error_key)
                        merged_errors.append({attr: vals})
    
    return merged_errors

def is_same_error_type(error1, error2):
    """
    检查两个错误是否是相同类型
    
    参数:
        error1: 第一个错误字典
        error2: 第二个错误字典
        
    返回:
        bool: 如果是相同类型的错误返回True
    """
    # 获取错误属性名
    attr1 = list(error1.keys())[0] if error1 else None
    attr2 = list(error2.keys())[0] if error2 else None
    
    # 如果属性名不同，不是相同类型
    if attr1 != attr2:
        return False
    
    # 如果属性名相同，检查值
    values1 = error1.get(attr1, {})
    values2 = error2.get(attr2, {})
    
    # 检查期望值是否相同
    if values1.get('scuccess') != values2.get('scuccess'):
        return False
    
    # 如果同时都有error值，检查是否相同或类似
    if 'error' in values1 and 'error' in values2:
        # 如果是数字，允许一定的误差
        if isinstance(values1.get('error'), (int, float)) and isinstance(values2.get('error'), (int, float)):
            return abs(values1.get('error') - values2.get('error')) <= 2
        # 否则精确比较
        return values1.get('error') == values2.get('error')
    
    # 如果一个有error值而另一个没有，不是相同类型
    if ('error' in values1) != ('error' in values2):
        return False
    
    # 其他情况下，认为是相同类型
    return True

def get_representative_error_message(errors):
    """
    根据错误生成代表性的错误消息
    
    参数:
        errors: 错误列表
        
    返回:
        str: 代表性的错误消息
    """
    if not errors:
        return "无错误"
    
    messages = []
    for error in errors:
        for attr, values in error.items():
            current = values.get('error', '未设置')
            expected = values.get('scuccess', '未知')
            messages.append(f"{attr}: 当前值={current}, 正确值={expected}")
    
    return "; ".join(messages)

# 演示用例
if __name__ == "__main__":
    # 示例统计信息
    example_stats = {
        "elements": [
            {
                "type": "paragraph",
                "index": 10,
                "result": [
                    {"alignment": {"scuccess": "center", "error": "left"}}
                ]
            },
            {
                "type": "run",
                "index": (10, 0),
                "result": [
                    {"font_chinese": {"scuccess": "黑体", "error": "宋体"}},
                    {"size": {"scuccess": 30, "error": 24}}
                ]
            },
            {
                "type": "run",
                "index": (10, 1),
                "result": [
                    {"font_chinese": {"scuccess": "黑体", "error": "宋体"}},
                    {"size": {"scuccess": 30, "error": 24}}
                ]
            },
            {
                "type": "paragraph",
                "index": 20,
                "result": [
                    {"before": {"scuccess": 120, "error": 0}}
                ]
            },
            {
                "type": "run",
                "index": (20, 0),
                "result": []  # 没有错误
            }
        ]
    }
    
    # 清理统计信息
    cleaned_stats = clean_style_statistics(example_stats)
    
    # 打印清理结果
    print("清理前元素数量:", len(example_stats["elements"]))
    print("清理后元素数量:", len(cleaned_stats["elements"]))
    print("\n清理后的元素:")
    for elem in cleaned_stats["elements"]:
        elem_type = elem["type"]
        if elem_type == "paragraph":
            print(f"段落 {elem['index']} 错误:")
        elif elem_type == "run":
            para_idx, run_idx = elem["index"]
            print(f"段落 {para_idx} 的 Run {run_idx} 错误:")
        elif elem_type == "table":
            print(f"表格 {elem['index']} 错误:")
            
        for error in elem.get("result", []):
            print(f"  {error}") 