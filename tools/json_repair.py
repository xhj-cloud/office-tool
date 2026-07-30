"""
JSON 修复工具
AI 模型生成 JSON 字符串时容易出错（漏逗号、多余逗号等），
此模块在 json.loads 之前自动修复常见问题。
"""

import re
import json


def repair_json(text: str) -> str:
    """
    修复常见 JSON 格式错误，返回修复后的 JSON 字符串。
    如果无法修复则返回原字符串。
    """
    if not text or not text.strip():
        return text

    text = text.strip()

    # 1. 去掉尾部多余逗号: ,}   ,]
    text = re.sub(r",\s*}", "}", text)
    text = re.sub(r",\s*\]", "]", text)

    # 2. 对象 / 数组元素之间缺逗号: "}  {"  "]"  "["
    text = re.sub(r'"\s*\n?\s*"', '",\n"', text)
    text = re.sub(r'"\s*\n?\s*\{', '",\n{', text)
    text = re.sub(r'}\s*\n?\s*{', '},\n{', text)
    text = re.sub(r'}\s*\n?\s*"', '},\n"', text)
    text = re.sub(r'\]\s*\n?\s*"', '],\n"', text)
    text = re.sub(r'\]\s*\n?\s*\{', '],\n{', text)
    text = re.sub(r'"\s*\n?\s*\[', '",\n[', text)
    text = re.sub(r'}\s*\n?\s*\[', '},\n[', text)
    text = re.sub(r'\]\s*\n?\s*\[', '],\n[', text)

    # 3. 两个字符串之间缺逗号（仅限同一行）
    text = re.sub(r'"\s{1,}"', '", "', text)

    return text


def safe_parse_json(text: str):
    """
    安全解析 JSON，先尝试原始文本，失败则修复后重试。
    返回 (parsed_object, error_message) 元组。
    """
    errors = []

    # 尝试原始
    try:
        return json.loads(text), None
    except json.JSONDecodeError as e:
        errors.append(f"原始: {e}")

    # 尝试修复
    repaired = repair_json(text)
    if repaired != text:
        try:
            return json.loads(repaired), None
        except json.JSONDecodeError as e:
            errors.append(f"修复后仍失败: {e}")

    # 都失败，返回详细错误
    error_msg = "; ".join(errors)
    if len(text) > 200:
        snippet = text[:200] + "..."
    else:
        snippet = text
    error_msg += f"\n输入片段: {snippet}"
    error_msg += "\n常见原因: 对象缺少逗号、多了逗号、引号未转义。请检查 JSON 格式。"

    return None, error_msg
