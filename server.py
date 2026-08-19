#!/usr/bin/env python3
"""
Office & Filesystem & PDF Tools MCP Server
提供 Word / Excel / PowerPoint 文件读写 + 文件系统基础操作 + PDF 解析与修改

使用: python server.py   (stdio 模式，CherryStudio / Claude CLI 连接)
"""

import functools
import json
import os
from collections import OrderedDict

from mcp.server.fastmcp import FastMCP

from tools.word_tools import read_docx as _read_docx, write_docx as _write_docx, edit_docx as _edit_docx
from tools.excel_tools import read_xlsx as _read_xlsx, write_xlsx as _write_xlsx, edit_xlsx as _edit_xlsx
from tools.ppt_tools import read_pptx as _read_pptx, write_pptx as _write_pptx, edit_pptx as _edit_pptx
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
try:
    from tools.pdf_processor import PDFProcessor
    _pdf_available = True
except ImportError:
    PDFProcessor = None
    _pdf_available = False

mcp = FastMCP("office-tools")

# ── PDF 处理器全局缓存（filepath -> PDFProcessor），LRU 上限防止长驻进程内存无限增长 ──
_PDF_CACHE_MAX = 32
_pdf_processors: "OrderedDict[str, PDFProcessor]" = OrderedDict()


def _get_pdf_proc(filepath: str):
    """获取（或创建）指定 PDF 文件的处理器；超出 LRU 上限时关闭并丢弃最久未用的条目"""
    if not _pdf_available:
        raise RuntimeError("PDF 功能不可用，请安装 PyMuPDF：pip install PyMuPDF>=1.24.0")
    fp = os.path.abspath(filepath)
    proc = _pdf_processors.get(fp)
    if proc is None:
        while len(_pdf_processors) >= _PDF_CACHE_MAX:
            old_fp, old_proc = _pdf_processors.popitem(last=False)
            try:
                old_proc.close()
            except Exception:
                pass
        proc = PDFProcessor(fp)
        _pdf_processors[fp] = proc
    else:
        _pdf_processors.move_to_end(fp)
    return proc


def _pdf_tool(fn):
    """PDF 工具统一错误处理：把异常转为 JSON 错误返回，避免裸 traceback"""
    @functools.wraps(fn)
    def wrapper(*args, **kwargs):
        try:
            return fn(*args, **kwargs)
        except Exception as e:
            return json.dumps({
                "success": False,
                "error": f"{type(e).__name__}: {e}",
            }, ensure_ascii=False)
    return wrapper


# ═══════════════════════════════════════════
# Word
# ═══════════════════════════════════════════

@mcp.tool()
def read_docx(file_path: str, mode: str = "full") -> str:
    """读取 Word 文档内容。mode: full(段落+表格) / paragraphs(仅段落) / tables(仅表格) / structure(含样式)"""
    return _read_docx(file_path, mode)


@mcp.tool()
def write_docx(output: str, content_json: str = "[]", title: str = "",
               body_font: str = "宋体", heading_font: str = "黑体") -> str:
    """生成 Word 文档(.docx)。支持标题、段落、表格、分页、签署页。

    Args:
        output: 输出文件完整路径（必填），如 /Users/xxx/Desktop/文档.docx
        content_json: 正文内容 JSON 数组，每项格式:
            {"type":"heading","text":"第一条 标题"}
            {"type":"paragraph","text":"正文","indent":true}
            {"type":"empty"}
            {"type":"table","headers":["序号","名称"],"rows":[["1","项目A"],["","合计"]],"widths":[2,5]}
            {"type":"page_break"}
            {"type":"signature","left":{"party":"甲方","rep":"______"},"right":{"party":"乙方","rep":"张三"}}
        title: 文档标题（可选）
        body_font: 正文字体（默认宋体）
        heading_font: 标题字体（默认黑体）
    """
    spec = {"output": output, "body_font": body_font, "heading_font": heading_font}
    if title:
        spec["title"] = title
    spec["content"] = json.loads(content_json) if content_json else []
    return _write_docx(json.dumps(spec, ensure_ascii=False))


@mcp.tool()
def edit_docx(file_path: str, spec_json: str) -> str:
    """编辑已有 Word 文档(.docx)。支持全局查找替换、插入/删除段落、修改段落文本、追加段落、修改表格单元格。段落 index 相对于原始文档编号，批量操作互不影响。"""
    return _edit_docx(file_path, spec_json)


# ═══════════════════════════════════════════
# Excel
# ═══════════════════════════════════════════

@mcp.tool()
def read_xlsx(file_path: str, sheet_name: str = "", mode: str = "full") -> str:
    """读取 Excel 文件内容。mode: full(含合并单元格) / values(纯数据) / structure(含样式)"""
    sn = sheet_name if sheet_name else None
    return _read_xlsx(file_path, sn, mode)


@mcp.tool()
def write_xlsx(output: str, sheets_json: str = "[{\"name\":\"Sheet1\",\"headers\":[\"A\"],\"rows\":[[\"\"]]}]") -> str:
    """生成 Excel 文件(.xlsx)。支持多 Sheet、表头格式、合并单元格、冻结窗格。

    Args:
        output: 输出文件完整路径（必填），如 /Users/xxx/Desktop/表格.xlsx
        sheets_json: Sheet 定义 JSON 数组，每项格式:
            {"name":"Sheet1","headers":["序号","名称","金额"],"rows":[["1","交换机",3995]],"col_widths":[8,30,15],"freeze":"A2"}
    """
    spec = {"output": output}
    spec["sheets"] = json.loads(sheets_json) if sheets_json else []
    return _write_xlsx(json.dumps(spec, ensure_ascii=False))


@mcp.tool()
def edit_xlsx(file_path: str, spec_json: str) -> str:
    """编辑已有 Excel 文件(.xlsx/.xlsm)。支持批量/单个单元格写值、追加行、删除行、合并/取消合并单元格、重命名 Sheet。"""
    return _edit_xlsx(file_path, spec_json)


# ═══════════════════════════════════════════
# PowerPoint
# ═══════════════════════════════════════════

@mcp.tool()
def read_pptx(file_path: str, mode: str = "full") -> str:
    """读取 PPT 文件内容。mode: full(含形状/表格) / outline(仅大纲文本) / notes(含备注)"""
    return _read_pptx(file_path, mode)


@mcp.tool()
def write_pptx(output: str, slides_json: str = "[{\"layout\":0,\"title\":\"标题\"}]") -> str:
    """生成 PowerPoint 演示文稿(.pptx)。支持标题、副标题、要点列表、数据表格。

    Args:
        output: 输出文件完整路径（必填），如 /Users/xxx/Desktop/演示.pptx
        slides_json: 幻灯片 JSON 数组，每项格式:
            {"layout":0,"title":"标题页","subtitle":"副标题"}
            {"layout":1,"title":"内容页","bullets":["要点1","要点2"]}
            {"layout":6,"title":"数据页","table":{"headers":["指标","Q1"],"rows":[["营收","100万"]]}}
            layout: 0=标题页 1=标题+内容 6=空白页
    """
    spec = {"output": output}
    spec["slides"] = json.loads(slides_json) if slides_json else []
    return _write_pptx(json.dumps(spec, ensure_ascii=False))


@mcp.tool()
def edit_pptx(file_path: str, spec_json: str) -> str:
    """编辑已有 PowerPoint 文件(.pptx)。支持全局查找替换文本、修改指定页标题、添加/删除幻灯片。"""
    return _edit_pptx(file_path, spec_json)


# ═══════════════════════════════════════════
# Filesystem
# ═══════════════════════════════════════════

@mcp.tool()
def list_directory(path: str) -> str:
    """列出目录下的所有文件和子目录。（路径受白名单限制：OFFICE_TOOLS_ALLOWED_DIRS，默认主目录）"""
    return _list_directory(path)


@mcp.tool()
def read_file(path: str, encoding: str = "utf-8", max_lines: int = 500) -> str:
    """读取任意文本文件的内容。支持 .txt .md .json .py .csv 等纯文本格式。（路径受白名单限制）"""
    return _read_file(path, encoding, max_lines)


@mcp.tool()
def write_file(path: str, content: str, encoding: str = "utf-8") -> str:
    """创建或覆盖写入文件（自动创建父目录）。（路径受白名单限制）"""
    return _write_file(path, content, encoding)


@mcp.tool()
def edit_file(path: str, old_string: str, new_string: str,
              encoding: str = "utf-8", replace_all: bool = False) -> str:
    """编辑文件内容：查找并替换文本。设置 replace_all=true 可替换所有匹配。（路径受白名单限制）"""
    return _edit_file(path, old_string, new_string, encoding, replace_all)


@mcp.tool()
def file_info(path: str) -> str:
    """获取文件或目录的详细信息（大小、修改时间、类型等）。（路径受白名单限制）"""
    return _file_info(path)


@mcp.tool()
def create_directory(path: str) -> str:
    """创建目录（自动创建所有父目录）。（路径受白名单限制）"""
    return _create_directory(path)


@mcp.tool()
def move_file(source: str, destination: str) -> str:
    """移动或重命名文件/目录。（路径受白名单限制）"""
    return _move_file(source, destination)


@mcp.tool()
def delete_file(path: str, recursive: bool = False) -> str:
    """删除文件。设置 recursive=true 可递归删除目录。（路径受白名单限制，禁止递归删除白名单根目录）"""
    return _delete_file(path, recursive)


# ═══════════════════════════════════════════
# PDF（由 pdf-modifier-mcp 合并而来）
# ═══════════════════════════════════════════

@mcp.tool()
@_pdf_tool
def pdf_info(filepath: str) -> str:
    """获取 PDF 文件的基本信息（页数、大小、元数据等）"""
    return json.dumps(_get_pdf_proc(filepath).get_info(), ensure_ascii=False, indent=2)


@mcp.tool()
@_pdf_tool
def pdf_extract_text(filepath: str, page_range: str = "") -> str:
    """提取 PDF 中的文本内容。page_range 如 '1-3'、'2,5'（页码从 1 开始），留空为全部"""
    pr = page_range or None
    return json.dumps(_get_pdf_proc(filepath).extract_text(pr), ensure_ascii=False, indent=2)


@mcp.tool()
@_pdf_tool
def pdf_render_pages(filepath: str, page_range: str = "", dpi: int = 200, output_dir: str = "") -> str:
    """将 PDF 页面渲染为图片文件，返回图片路径供视觉 AI 使用。page_range 如 '1-5'（页码从 1 开始），留空渲染所有页"""
    pr = page_range or None
    od = output_dir or None
    return json.dumps(_get_pdf_proc(filepath).render_pages(pr, dpi, od), ensure_ascii=False, indent=2)


@mcp.tool()
@_pdf_tool
def pdf_apply_modifications(filepath: str, modifications_json: str, page_num: int = 1) -> str:
    """【核心功能】将 AI 生成的修改指令应用到 PDF 页面上（自动执行替换、删除、添加等操作）。page_num 从 1 开始"""
    proc = _get_pdf_proc(filepath)

    mods = json.loads(modifications_json) if isinstance(modifications_json, str) else modifications_json
    page_objects = mods.get("page_objects", mods.get("raw_response", []))
    if isinstance(page_objects, str):
        try:
            parsed = json.loads(page_objects)
            page_objects = parsed.get("page_objects", [])
        except (json.JSONDecodeError, AttributeError):
            return json.dumps({
                "success": False,
                "error": f"无法解析修改指令: {str(page_objects)[:200]}...",
            }, ensure_ascii=False, indent=2)

    page = proc.resolve_page(page_num)
    results = proc.apply_ai_modifications(page_num, page_objects, page.rect.width, page.rect.height)
    return json.dumps(results, ensure_ascii=False, indent=2)


@mcp.tool()
@_pdf_tool
def pdf_save(filepath: str, output_path: str = "") -> str:
    """将修改后的 PDF 保存到文件。output_path 不指定则自动添加 _modified 后缀"""
    op = output_path or None
    out = _get_pdf_proc(filepath).save(op)
    return json.dumps({
        "success": True,
        "output_path": out,
        "message": f"PDF 已保存至: {out}",
    }, ensure_ascii=False, indent=2)


@mcp.tool()
@_pdf_tool
def pdf_manual_replace_text(filepath: str, page_num: int, bbox: list, new_text: str, font_size: float = 12) -> str:
    """手动替换 PDF 中指定区域的文本（不需要 AI 参与）。先遮盖再写入新文本。page_num 从 1 开始；bbox: [x0,y0,x1,y1]（点）"""
    proc = _get_pdf_proc(filepath)
    result = proc.replace_text_in_bbox(page_num, tuple(bbox), new_text, font_size)
    return json.dumps(result, ensure_ascii=False, indent=2)


@mcp.tool()
@_pdf_tool
def pdf_redact(filepath: str, page_num: int, bbox: list, fill_white: bool = False) -> str:
    """遮盖（涂黑/涂白）PDF 中的指定区域。page_num 从 1 开始；bbox: [x0,y0,x1,y1]（点）"""
    proc = _get_pdf_proc(filepath)
    fill_color = (1, 1, 1) if fill_white else (0, 0, 0)
    result = proc.redact_area(page_num, tuple(bbox), fill_color)
    return json.dumps(result, ensure_ascii=False, indent=2)


@mcp.tool()
@_pdf_tool
def pdf_highlight_area(filepath: str, page_num: int, bbox: list) -> str:
    """高亮标注 PDF 中的指定区域。page_num 从 1 开始；bbox: [x0,y0,x1,y1]（点）"""
    proc = _get_pdf_proc(filepath)
    result = proc.add_highlight(page_num, tuple(bbox))
    return json.dumps(result, ensure_ascii=False, indent=2)


@mcp.tool()
@_pdf_tool
def pdf_add_text(filepath: str, page_num: int, text: str, x: float, y: float, font_size: float = 12) -> str:
    """在 PDF 指定位置添加文本（自动支持中文）。page_num 从 1 开始；x、y 为坐标（点）"""
    proc = _get_pdf_proc(filepath)
    result = proc.add_text(page_num, text, x, y, font_size)
    return json.dumps(result, ensure_ascii=False, indent=2)


@mcp.tool()
@_pdf_tool
def pdf_deyellow(filepath: str, page_range: str = "", dpi: int = 200, strength: float = 0.95, threshold: float = 120) -> str:
    """【扫描件去黄】去除扫描件 PDF 的黄底，使纸张变白。page_range 如 '1-3'（页码从 1 开始），留空处理所有页"""
    proc = _get_pdf_proc(filepath)
    pr = page_range or None
    result = proc.deyellow(page_range=pr, dpi=dpi, strength=strength, threshold=threshold)
    return json.dumps(result, ensure_ascii=False, indent=2)


# ═══════════════════════════════════════════

if __name__ == "__main__":
    mcp.run(transport="stdio")
