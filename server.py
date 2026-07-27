#!/usr/bin/env python3
"""
Office & Filesystem Tools MCP Server
提供 Word / Excel / PowerPoint 文件读写 + 文件系统基础操作

使用: python server.py   (stdio 模式，CherryStudio / Claude CLI 连接)
"""

from mcp.server.fastmcp import FastMCP

from tools.word_tools import read_docx as _read_docx, write_docx as _write_docx
from tools.excel_tools import read_xlsx as _read_xlsx, write_xlsx as _write_xlsx
from tools.ppt_tools import read_pptx as _read_pptx, write_pptx as _write_pptx
from tools.filesystem_tools import (
    list_directory as _list_directory,
    read_file as _read_file,
    write_file as _write_file,
    edit_file as _edit_file,
    file_info as _file_info,
    create_directory as _create_directory,
    move_file as _move_file,
    delete_file as _delete_file,
)

mcp = FastMCP("office-tools")


# ═══════════════════════════════════════════
# Word
# ═══════════════════════════════════════════

@mcp.tool()
def read_docx(file_path: str, mode: str = "full") -> str:
    """读取 Word 文档内容。mode: full(段落+表格) / paragraphs(仅段落) / tables(仅表格) / structure(含样式)"""
    return _read_docx(file_path, mode)


@mcp.tool()
def write_docx(spec_json: str) -> str:
    """根据 JSON 规格生成 Word 文档。支持标题、段落、表格、分页、签署页。"""
    return _write_docx(spec_json)


# ═══════════════════════════════════════════
# Excel
# ═══════════════════════════════════════════

@mcp.tool()
def read_xlsx(file_path: str, sheet_name: str = "", mode: str = "full") -> str:
    """读取 Excel 文件内容。mode: full(含合并单元格) / values(纯数据) / structure(含样式)"""
    sn = sheet_name if sheet_name else None
    return _read_xlsx(file_path, sn, mode)


@mcp.tool()
def write_xlsx(spec_json: str) -> str:
    """根据 JSON 规格生成 Excel 文件。支持多 Sheet、表头格式、合并单元格、冻结窗格。"""
    return _write_xlsx(spec_json)


# ═══════════════════════════════════════════
# PowerPoint
# ═══════════════════════════════════════════

@mcp.tool()
def read_pptx(file_path: str, mode: str = "full") -> str:
    """读取 PPT 文件内容。mode: full(含形状/表格) / outline(仅大纲文本) / notes(含备注)"""
    return _read_pptx(file_path, mode)


@mcp.tool()
def write_pptx(spec_json: str) -> str:
    """根据 JSON 规格生成 PowerPoint 演示文稿。支持标题、副标题、要点列表、数据表格。"""
    return _write_pptx(spec_json)


# ═══════════════════════════════════════════
# Filesystem
# ═══════════════════════════════════════════

@mcp.tool()
def list_directory(path: str) -> str:
    """列出目录下的所有文件和子目录。"""
    return _list_directory(path)


@mcp.tool()
def read_file(path: str, encoding: str = "utf-8", max_lines: int = 500) -> str:
    """读取任意文本文件的内容。支持 .txt .md .json .py .csv 等纯文本格式。"""
    return _read_file(path, encoding, max_lines)


@mcp.tool()
def write_file(path: str, content: str, encoding: str = "utf-8") -> str:
    """创建或覆盖写入文件（自动创建父目录）。"""
    return _write_file(path, content, encoding)


@mcp.tool()
def edit_file(path: str, old_string: str, new_string: str,
              encoding: str = "utf-8", replace_all: bool = False) -> str:
    """编辑文件内容：查找并替换文本。设置 replace_all=true 可替换所有匹配。"""
    return _edit_file(path, old_string, new_string, encoding, replace_all)


@mcp.tool()
def file_info(path: str) -> str:
    """获取文件或目录的详细信息（大小、修改时间、类型等）。"""
    return _file_info(path)


@mcp.tool()
def create_directory(path: str) -> str:
    """创建目录（自动创建所有父目录）。"""
    return _create_directory(path)


@mcp.tool()
def move_file(source: str, destination: str) -> str:
    """移动或重命名文件/目录。"""
    return _move_file(source, destination)


@mcp.tool()
def delete_file(path: str, recursive: bool = False) -> str:
    """删除文件。设置 recursive=true 可递归删除目录。"""
    return _delete_file(path, recursive)


# ═══════════════════════════════════════════

if __name__ == "__main__":
    mcp.run(transport="stdio")
