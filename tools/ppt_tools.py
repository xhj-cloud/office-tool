"""
PowerPoint 演示文稿读写工具
支持 .pptx 格式的读取和生成
"""

import json
from pptx import Presentation
from pptx.util import Inches, Pt, Cm, Emu
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.dml.color import RGBColor


def read_pptx(file_path: str, mode: str = "full") -> str:
    """
    读取 PPT 文件内容

    Args:
        file_path: .pptx 文件路径
        mode: "full"(全部) / "outline"(大纲) / "notes"(含备注)

    Returns:
        JSON 格式的内容
    """
    try:
        prs = Presentation(file_path)
    except Exception as e:
        return json.dumps({"error": f"无法打开文件: {str(e)}"}, ensure_ascii=False)

    output = {
        "file": file_path,
        "slide_count": len(prs.slides),
        "slide_width": prs.slide_width,
        "slide_height": prs.slide_height,
    }

    slides_data = []
    for si, slide in enumerate(prs.slides):
        slide_info = {"slide_num": si + 1, "layout": slide.slide_layout.name}
        shapes = []

        for shape in slide.shapes:
            shape_info = {
                "name": shape.name,
                "type": str(shape.shape_type),
                "left": shape.left,
                "top": shape.top,
                "width": shape.width,
                "height": shape.height,
            }

            if shape.has_text_frame:
                texts = []
                for para in shape.text_frame.paragraphs:
                    para_text = ""
                    for run in para.runs:
                        para_text += run.text
                    if para_text.strip():
                        texts.append(para_text)
                shape_info["texts"] = texts

            if shape.has_table:
                table = shape.table
                tdata = []
                for row in table.rows:
                    tdata.append([cell.text for cell in row.cells])
                shape_info["table"] = {"rows": len(table.rows), "cols": len(table.columns), "data": tdata}

            shapes.append(shape_info)

        slide_info["shapes"] = shapes

        if mode == "notes" and slide.has_notes_slide:
            slide_info["notes"] = slide.notes_slide.notes_text_frame.text

        slides_data.append(slide_info)

    output["slides"] = slides_data
    return json.dumps(output, ensure_ascii=False, indent=2)


def write_pptx(spec_json: str) -> str:
    """
    根据 JSON 规格生成 .pptx 文件

    spec_json 格式:
    {
        "output": "/path/to/output.pptx",
        "slide_width": 13.33,       // 英寸（16:9 默认）
        "slide_height": 7.5,
        "slides": [
            {
                "layout": 0,        // 0=标题, 1=标题+内容, 6=空白（默认）
                "title": "标题文字",
                "subtitle": "副标题",
                "bullets": ["要点1", "要点2", "要点3"],
                "table": {
                    "headers": ["列1", "列2"],
                    "rows": [["a", "b"]],
                    "left": 1.5, "top": 2.0, "width": 7.0, "height": 3.0
                }
            }
        ]
    }
    """
    try:
        spec = json.loads(spec_json)
    except json.JSONDecodeError as e:
        return json.dumps({"error": f"JSON 解析失败: {str(e)}"}, ensure_ascii=False)

    output_path = spec.get("output")
    if not output_path:
        return json.dumps({"error": "必须指定 output 路径"}, ensure_ascii=False)

    prs = Presentation()

    # 页面尺寸（16:9）
    prs.slide_width = Inches(spec.get("slide_width", 13.33))
    prs.slide_height = Inches(spec.get("slide_height", 7.5))

    # 可用的 layouts: 0=Title, 1=Title and Content, 6=Blank
    layout_map = {
        0: 0,   # Title Slide
        1: 1,   # Title and Content
        6: 6,   # Blank
    }

    for slide_spec in spec.get("slides", []):
        layout_idx = slide_spec.get("layout", 6)
        layout_idx = layout_map.get(layout_idx, 6)
        slide_layout = prs.slide_layouts[layout_idx]
        slide = prs.slides.add_slide(slide_layout)

        # 标题
        if slide_spec.get("title") and slide.shapes.title:
            slide.shapes.title.text = slide_spec["title"]

        # 副标题
        if slide_spec.get("subtitle"):
            # 查找副标题占位符
            for ph in slide.placeholders:
                if ph.placeholder_format.idx == 1:
                    ph.text = slide_spec["subtitle"]
                    break

        # 要点
        if slide_spec.get("bullets"):
            body_shape = None
            for shape in slide.shapes:
                if shape.has_text_frame and shape != slide.shapes.title:
                    body_shape = shape
                    break
            if body_shape:
                tf = body_shape.text_frame
                tf.clear()
                for i, bullet in enumerate(slide_spec["bullets"]):
                    if i == 0:
                        tf.paragraphs[0].text = bullet
                    else:
                        p = tf.add_paragraph()
                        p.text = bullet
                        p.level = 0

        # 表格
        if slide_spec.get("table"):
            tspec = slide_spec["table"]
            headers = tspec.get("headers", [])
            rows_data = tspec.get("rows", [])
            n_rows = len(rows_data) + 1
            n_cols = len(headers) or (len(rows_data[0]) if rows_data else 1)

            table_shape = slide.shapes.add_table(
                n_rows, n_cols,
                Inches(tspec.get("left", 1.5)),
                Inches(tspec.get("top", 2.0)),
                Inches(tspec.get("width", 7.0)),
                Inches(tspec.get("height", 3.0)),
            )
            table = table_shape.table

            # 表头
            for ci, h in enumerate(headers):
                cell = table.cell(0, ci)
                cell.text = h
                for p in cell.text_frame.paragraphs:
                    p.font.size = Pt(11)
                    p.font.bold = True
                    p.font.color.rgb = RGBColor(255, 255, 255)
                    p.alignment = PP_ALIGN.CENTER
                cell.fill.solid()
                cell.fill.fore_color.rgb = RGBColor(0x33, 0x33, 0x33)

            # 数据
            for ri, row_data in enumerate(rows_data):
                for ci, val in enumerate(row_data):
                    cell = table.cell(ri + 1, ci)
                    cell.text = str(val)
                    for p in cell.text_frame.paragraphs:
                        p.font.size = Pt(10)
                        p.alignment = PP_ALIGN.CENTER

    prs.save(output_path)
    return json.dumps({"success": True, "output": output_path, "slide_count": len(prs.slides)}, ensure_ascii=False)
