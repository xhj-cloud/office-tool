"""
PowerPoint 演示文稿读写工具
支持 .pptx 格式的读取和生成
"""

import json
import os
from lxml import etree
from pptx import Presentation
from pptx.util import Inches, Pt, Cm, Emu
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE, MSO_SHAPE_TYPE, PP_PLACEHOLDER
from pptx.dml.color import RGBColor
from pptx.oxml import parse_xml
from pptx.oxml.ns import nsdecls, qn
from .json_repair import safe_parse_json


def _iter_shapes(shapes):
    """递归遍历所有形状（包括 group 组合形状内部的子形状）"""
    for shape in shapes:
        yield shape
        try:
            if shape.shape_type == MSO_SHAPE_TYPE.GROUP:
                yield from _iter_shapes(shape.shapes)
        except Exception:
            pass


def _find_body_shape(slide):
    """查找正文内容占位符（BODY/OBJECT 类型）；找不到返回 None。

    注意：layout-0（标题页）的副标题占位符是 SUBTITLE 类型，不会被误认为正文框。
    """
    for shape in slide.shapes:
        if not shape.is_placeholder:
            continue
        try:
            ptype = shape.placeholder_format.type
        except Exception:
            continue
        if ptype in (PP_PLACEHOLDER.BODY, PP_PLACEHOLDER.OBJECT):
            return shape
    return None


def _add_body_textbox(slide, prs):
    """无正文占位符时（如空白 layout-6），添加默认位置的文本框承载内容"""
    left = int(prs.slide_width * 0.067)
    top = int(prs.slide_height * 0.24)
    width = int(prs.slide_width * 0.866)
    height = int(prs.slide_height * 0.64)
    return slide.shapes.add_textbox(left, top, width, height)


def _add_title_textbox(slide, prs):
    """无标题占位符时（如空白 layout-6），在顶部添加标题文本框"""
    left = int(prs.slide_width * 0.067)
    top = int(prs.slide_height * 0.08)
    width = int(prs.slide_width * 0.866)
    height = int(prs.slide_height * 0.16)
    return slide.shapes.add_textbox(left, top, width, height)


def _rgb(color: str):
    """解析十六进制颜色（可带 #），非法值抛 ValueError"""
    c = str(color).strip().lstrip("#")
    if len(c) != 6:
        raise ValueError(f"颜色值非法: {color!r}（应为 #C8102E 形式的十六进制颜色）")
    return RGBColor.from_string(c)


def _apply_font(paragraph, font_name: str = "微软雅黑", size: float = None,
                color: str = None, bold: bool = None):
    """为段落的所有 run 设置字体（含中文 eastAsia 字体）、字号、颜色、加粗"""
    if not paragraph.runs:
        paragraph.add_run()
    for run in paragraph.runs:
        run.font.name = font_name
        rPr = run._r.get_or_add_rPr()
        ea = rPr.find(qn("a:ea"))
        if ea is None:
            ea = etree.SubElement(rPr, qn("a:ea"))
        ea.set("typeface", font_name)
        if size:
            run.font.size = Pt(size)
        if color:
            run.font.color.rgb = _rgb(color)
        if bold is not None:
            run.font.bold = bold


# ── 视觉增强（照片背景 / 渐变 / 文字阴影，对标设计稿级模板）──────────────────

def _set_slide_bg_xml(slide, bgpr_inner: str):
    """设置幻灯片 <p:bg>（bgpr_inner 为 <p:bgPr> 内部 XML，须含填充）"""
    cSld = slide._element.cSld
    old = cSld.find(qn("p:bg"))
    if old is not None:
        cSld.remove(old)
    bg = parse_xml(f'<p:bg {nsdecls("a", "p", "r")}>'
                   f'<p:bgPr>{bgpr_inner}<a:effectLst/></p:bgPr></p:bg>')
    cSld.insert(0, bg)


def _set_background_image(slide, path: str):
    """整页图片背景：cover 裁切（上下 9% 过溢，与参考设计稿同款手法）"""
    if not path or not os.path.exists(path):
        raise ValueError(f"background_image 图片不存在: {path!r}")
    image_part, rId = slide.part.get_or_add_image_part(path)
    inner = (f'<a:blipFill dpi="0" rotWithShape="1">'
             f'<a:blip r:embed="{rId}"/>'
             f'<a:srcRect/>'
             f'<a:stretch><a:fillRect t="-9000" b="-9000"/></a:stretch>'
             f'</a:blipFill>')
    _set_slide_bg_xml(slide, inner)


def _set_gradient_background(slide, colors, angle: float = 90):
    """整页渐变背景。angle: 0=左→右，90=上→下（度）"""
    if not colors or len(colors) < 2:
        raise ValueError("渐变背景需要至少 2 个颜色，"
                         "如 background: {\"colors\": [\"#0d1b2a\", \"#1b3a5c\"], \"angle\": 90}")
    n = len(colors) - 1
    gs = "".join(f'<a:gs pos="{int(100000 * i / n)}">'
                 f'<a:srgbClr val="{str(_rgb(c)).upper()}"/></a:gs>'
                 for i, c in enumerate(colors))
    inner = (f'<a:gradFill rotWithShape="1"><a:gsLst>{gs}</a:gsLst>'
             f'<a:lin ang="{int(angle * 60000)}" scaled="1"/></a:gradFill>')
    _set_slide_bg_xml(slide, inner)


def _replace_shape_fill_with_gradient(shape, gs_xml: str, angle: float = 0):
    """把形状当前的 solidFill 原位替换为渐变填充（保持 spPr 内元素顺序合法）"""
    grad = parse_xml(f'<a:gradFill {nsdecls("a")} rotWithShape="1">'
                     f'<a:gsLst>{gs_xml}</a:gsLst>'
                     f'<a:lin ang="{int(angle * 60000)}" scaled="1"/></a:gradFill>')
    spPr = shape.fill._xPr
    solid = spPr.find(qn("a:solidFill"))
    if solid is None:
        raise ValueError("形状需要先有 solidFill 才能替换为渐变")
    spPr.replace(solid, grad)


def _add_scrim(slide, prs, direction: str = "left"):
    """图片背景上的可读性暗角（黑色渐变渐隐，参考设计稿同款）

    left/right/bottom: 对应方向的黑色渐变遮罩；dark: 全页 35% 黑底
    """
    w, h = prs.slide_width, prs.slide_height
    if direction == "dark":
        shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, w, h)
        shape.line.fill.background()
        shape.fill.solid()
        shape.fill.fore_color.rgb = RGBColor(0, 0, 0)
        clr = shape.fill._xPr.find(qn("a:solidFill")).find(qn("a:srgbClr"))
        etree.SubElement(clr, qn("a:alpha")).set("val", "35000")
        return shape
    if direction == "left":
        rect, ang, dark_first = (0, 0, int(w * 0.62), h), 0, True
    elif direction == "right":
        rect, ang, dark_first = (int(w * 0.38), 0, int(w * 0.62), h), 10800000, True
    elif direction == "bottom":
        rect, ang, dark_first = (0, int(h * 0.45), w, int(h * 0.55)), 5400000, False
    else:
        raise ValueError(f"scrim 只支持 left/right/bottom/dark，收到 {direction!r}")
    shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, *rect)
    shape.line.fill.background()
    shape.fill.solid()
    dark = '<a:gs pos="0"><a:srgbClr val="000000"><a:alpha val="72000"/></a:srgbClr></a:gs>'
    clear = '<a:gs pos="100000"><a:srgbClr val="000000"><a:alpha val="0"/></a:srgbClr></a:gs>'
    gs = (dark + clear) if dark_first else (clear + dark)
    _replace_shape_fill_with_gradient(shape, gs, ang / 60000)
    return shape


def _apply_text_shadow(shape, mode: str = "soft"):
    """给形状内所有 run 加外阴影（参考设计稿白字在照片上可读的关键）

    soft:   轻微漂浮感（blur 3pt / 30% 黑）
    strong: 照片上白字专用（blur 24pt / 57% 黑）
    """
    if mode == "strong":
        attrs, alpha = ('blurRad="241300" dist="38100" dir="5400000" '
                        'algn="tl" rotWithShape="0"'), "57000"
    elif mode == "soft":
        attrs, alpha = ('blurRad="38100" dist="22860" dir="5400000" '
                        'algn="tl" rotWithShape="0"'), "30000"
    else:
        raise ValueError(f"text_shadow 只支持 true/'soft'/'strong'，收到 {mode!r}")
    shadow = (f'<a:effectLst {nsdecls("a")}>'
              f'<a:outerShdw {attrs}>'
              f'<a:srgbClr val="000000"><a:alpha val="{alpha}"/></a:srgbClr>'
              f'</a:outerShdw></a:effectLst>')
    for para in shape.text_frame.paragraphs:
        for run in para.runs:
            rPr = run._r.get_or_add_rPr()
            for old in rPr.findall(qn("a:effectLst")):
                rPr.remove(old)
            el = parse_xml(shadow)
            anchor = None
            for tag in ("a:highlight", "a:uLn", "a:uFill", "a:latin", "a:ea", "a:cs"):
                anchor = rPr.find(qn(tag))
                if anchor is not None:
                    break
            if anchor is not None:
                anchor.addprevious(el)
            else:
                rPr.append(el)


def _set_run_alpha(run, pct: float):
    """设置 run 文字颜色的不透明度（pct: 0~100，如 12 = 12% 可见度）"""
    if run.font.color is None or run.font.color.type is None:
        return
    clr = run.font.color._xFill.find(qn("a:srgbClr"))
    if clr is None:
        return
    for a in clr.findall(qn("a:alpha")):
        clr.remove(a)
    etree.SubElement(clr, qn("a:alpha")).set("val", str(int(pct * 1000)))


def _build_section_slide(slide, prs, slide_spec: dict):
    """章节扉页（PART ONE 风格）：巨号水印字 + 宽字距英文标签 + 强调线 + 大标题"""
    w_in = Emu(prs.slide_width).inches

    # 1) 背景：优先 background_image，其次渐变，最后默认深蓝渐变
    bg_image = slide_spec.get("background_image")
    if bg_image:
        if isinstance(bg_image, dict):
            _set_background_image(slide, bg_image.get("path", ""))
            scrim = bg_image.get("scrim", "left")
        else:
            _set_background_image(slide, bg_image)
            scrim = "left"
        if scrim and scrim != "none":
            _add_scrim(slide, prs, scrim)
    else:
        bg = slide_spec.get("background")
        if isinstance(bg, dict):
            _set_gradient_background(slide, bg.get("colors", ["#0d1b2a", "#1b3a5c"]),
                                     bg.get("angle", 90))
        elif isinstance(bg, str):
            fill = slide.background.fill
            fill.solid()
            fill.fore_color.rgb = _rgb(bg)
        else:
            _set_gradient_background(slide, ["#0d1b2a", "#1b3a5c"], 90)

    shadow = slide_spec.get("text_shadow", True)
    shadow_mode = "soft" if shadow is True else str(shadow)
    left_in = 0.95

    # 2) 巨号水印字（右侧，白色 12% 透明感）
    tb_word = slide.shapes.add_textbox(Inches(w_in * 0.48), Inches(1.0),
                                       Inches(w_in * 0.48), Inches(5.5))
    tb_word.text_frame.word_wrap = True
    p = tb_word.text_frame.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    p.text = str(slide_spec.get("section_word") or "01")
    _apply_font(p, size=150, color="#FFFFFF", bold=True)
    _set_run_alpha(p.runs[0], 12)

    # 3) 英文标签（宽字距）
    tb_label = slide.shapes.add_textbox(Inches(left_in), Inches(2.4),
                                        Inches(7.0), Inches(0.5))
    p2 = tb_label.text_frame.paragraphs[0]
    p2.text = str(slide_spec.get("subtitle") or "PART").upper()
    _apply_font(p2, size=14, color="#FFFFFF")
    _set_run_alpha(p2.runs[0], 75)
    p2.runs[0]._r.get_or_add_rPr().set("spc", "500")  # 5pt 字距

    # 4) 白色强调线
    bar = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE,
                                 Inches(left_in + 0.02), Inches(3.08),
                                 Inches(0.85), Inches(0.05))
    bar.line.fill.background()
    bar.fill.solid()
    bar.fill.fore_color.rgb = RGBColor(0xFF, 0xFF, 0xFF)

    # 5) 中文大标题
    tb_title = slide.shapes.add_textbox(Inches(left_in), Inches(3.4),
                                        Inches(9.8), Inches(1.7))
    p3 = tb_title.text_frame.paragraphs[0]
    p3.text = str(slide_spec.get("title") or "")
    _apply_font(p3, size=44, color="#FFFFFF", bold=True)

    # 6) 可选副说明
    if slide_spec.get("section_note"):
        tb_note = slide.shapes.add_textbox(Inches(left_in), Inches(5.1),
                                           Inches(9.0), Inches(0.6))
        p4 = tb_note.text_frame.paragraphs[0]
        p4.text = str(slide_spec["section_note"])
        _apply_font(p4, size=16, color="#FFFFFF")
        _set_run_alpha(p4.runs[0], 60)

    # 7) 文字阴影（照片背景上白字可读的关键）
    for sh in (tb_word, tb_label, tb_title):
        _apply_text_shadow(sh, shadow_mode)
    return slide


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

        for shape in _iter_shapes(slide.shapes):
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
        "overwrite": false,         // 目标文件已存在时是否覆盖（默认 false，会报错提示）
        "slide_width": 13.33,       // 英寸（16:9 默认）
        "slide_height": 7.5,
        "slides": [
            {
                "layout": 0,        // 0=标题页 1=标题+内容 6=空白（默认）
                                    // "section"=章节扉页（PART ONE 风格：巨号水印字+英文标签+强调线+大标题，
                                    //   配 section_word="01"、subtitle="PART ONE"、section_note="副说明"）
                "title": "标题文字",            // 无标题占位符时自动加文本框
                "subtitle": "副标题",           // 仅标题页（layout 0）有效；section 页=英文标签
                "bullets": ["要点1", "要点2", "要点3"],
                "background": "#C8102E",        // 纯色背景（十六进制，可带 #）
                // 或渐变背景: "background": {"colors": ["#0d1b2a", "#1b3a5c"], "angle": 90}
                //   （angle: 0=左→右，90=上→下；section 页缺省默认深蓝渐变）
                "background_image": "/a/b/c.jpg",  // 整页照片背景（cover 裁切，设计稿级观感的关键）
                //   或对象形式: {"path": "/a/b/c.jpg", "scrim": "left"}
                //   scrim: 照片上文字可读性暗角 left/right/bottom/dark/none（默认 left）
                "text_shadow": true,           // 标题/副标题/要点加外阴影（true=soft，
                //   照片上白字建议 "strong"；参考设计稿同款文字处理）
                "title_color": "#FFD700",       // 标题颜色
                "title_size": 40,               // 标题字号（默认：标题页 40，其他 32）
                "bullet_color": "#333333",      // 要点颜色
                "bullet_size": 18,              // 要点字号（默认 18）
                "image": {"path": "/a/b/c.png", "left": 1.0, "top": 1.0, "width": 4.0},
                "notes": "主持人提示 / 时间控制",  // 演讲者备注
                "table": {
                    "headers": ["列1", "列2"],
                    "rows": [["a", "b"]],
                    "left": 1.5, "top": 2.0, "width": 7.0, "height": 3.0,
                    "col_widths": [3.5, 3.5]    // 列宽（英寸）
                }
            }
        ]
    }
    """
    spec, err = safe_parse_json(spec_json)
    if err:
        return json.dumps({"error": f"JSON 解析失败: {err}"}, ensure_ascii=False)

    output_path = spec.get("output")
    if not output_path:
        return json.dumps({"error": "必须指定 output 路径"}, ensure_ascii=False)

    slides = spec.get("slides")
    if not slides:
        return json.dumps({"error": "slides 不能为空（空数组会生成 0 页的空 PPT）"}, ensure_ascii=False)

    if os.path.exists(output_path) and not spec.get("overwrite"):
        return json.dumps(
            {"error": f"文件已存在: {output_path}。如需覆盖请传 overwrite: true，或换一个输出路径"},
            ensure_ascii=False)

    try:
        return _write_pptx_impl(spec, output_path)
    except ValueError as e:
        return json.dumps({"error": str(e)}, ensure_ascii=False)


def _write_pptx_impl(spec: dict, output_path: str) -> str:
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

    for slide_spec in spec["slides"]:
        if not isinstance(slide_spec, dict):
            raise ValueError(f"slides 每项必须是对象，收到: {slide_spec!r}")

        # 章节扉页（PART ONE 风格）
        layout_key = slide_spec.get("layout", 6)
        if isinstance(layout_key, str) and layout_key in ("section", "part", "divider"):
            slide = prs.slides.add_slide(prs.slide_layouts[6])
            _build_section_slide(slide, prs, slide_spec)
            if slide_spec.get("notes"):
                slide.notes_slide.notes_text_frame.text = str(slide_spec["notes"])
            continue

        if layout_key not in layout_map:
            raise ValueError(f"layout 只支持 0（标题页）/ 1（标题+内容）/ 6（空白）/ "
                             f"'section'（章节扉页），收到 {layout_key!r}")
        layout_idx = layout_map[layout_key]
        slide = prs.slides.add_slide(prs.slide_layouts[layout_idx])

        # 背景：整页图片 > 渐变 > 纯色
        bg_image = slide_spec.get("background_image")
        if bg_image:
            if isinstance(bg_image, dict):
                _set_background_image(slide, bg_image.get("path", ""))
                scrim = bg_image.get("scrim", "left")
            else:
                _set_background_image(slide, bg_image)
                scrim = "left"
            if scrim and scrim != "none":
                _add_scrim(slide, prs, scrim)
        elif slide_spec.get("background"):
            bg = slide_spec["background"]
            if isinstance(bg, dict):
                _set_gradient_background(slide, bg.get("colors"), bg.get("angle", 90))
            else:
                fill = slide.background.fill
                fill.solid()
                fill.fore_color.rgb = _rgb(bg)

        shadow_targets = []  # 需要加文字阴影的形状（text_shadow 时生效）

        # 标题（无标题占位符时自动加文本框，不再静默丢弃）
        if slide_spec.get("title"):
            title_shape = slide.shapes.title or _add_title_textbox(slide, prs)
            shadow_targets.append(title_shape)
            title_shape.text = slide_spec["title"]
            default_title_size = 40 if layout_idx == 0 else 32
            _apply_font(title_shape.text_frame.paragraphs[0],
                        size=slide_spec.get("title_size", default_title_size),
                        color=slide_spec.get("title_color"), bold=True)
            for p in title_shape.text_frame.paragraphs[1:]:
                _apply_font(p, size=slide_spec.get("title_size", default_title_size),
                            color=slide_spec.get("title_color"), bold=True)

        # 副标题（只认 SUBTITLE 类型占位符，避免误用正文占位符）
        if slide_spec.get("subtitle"):
            for ph in slide.placeholders:
                try:
                    if ph.placeholder_format.type == PP_PLACEHOLDER.SUBTITLE:
                        ph.text = slide_spec["subtitle"]
                        for p in ph.text_frame.paragraphs:
                            _apply_font(p, size=20, color=slide_spec.get("title_color"))
                        shadow_targets.append(ph)
                        break
                except Exception:
                    continue

        # 要点（无正文占位符时自动添加文本框，不再静默丢弃）
        if slide_spec.get("bullets"):
            body_shape = _find_body_shape(slide) or _add_body_textbox(slide, prs)
            shadow_targets.append(body_shape)
            tf = body_shape.text_frame
            tf.clear()
            bullet_size = slide_spec.get("bullet_size", 18)
            bullet_color = slide_spec.get("bullet_color")
            for i, bullet in enumerate(slide_spec["bullets"]):
                if i == 0:
                    p = tf.paragraphs[0]
                else:
                    p = tf.add_paragraph()
                p.text = str(bullet)
                p.level = 0
                _apply_font(p, size=bullet_size, color=bullet_color)

        # 文字阴影（照片背景上白字可读的关键；true=soft，也可 'soft'/'strong'）
        shadow = slide_spec.get("text_shadow")
        if shadow:
            shadow_mode = "soft" if shadow is True else str(shadow)
            for sh in shadow_targets:
                _apply_text_shadow(sh, shadow_mode)

        # 图片
        if slide_spec.get("image"):
            ispec = slide_spec["image"]
            ipath = ispec.get("path", "")
            if not ipath or not os.path.exists(ipath):
                raise ValueError(f"图片不存在: {ipath!r}")
            pic_kwargs = {}
            if ispec.get("width"):
                pic_kwargs["width"] = Inches(ispec["width"])
            if ispec.get("height"):
                pic_kwargs["height"] = Inches(ispec["height"])
            if not pic_kwargs:
                pic_kwargs["width"] = Inches(4.0)
            slide.shapes.add_picture(
                ipath,
                Inches(ispec.get("left", 0.5)),
                Inches(ispec.get("top", 0.5)),
                **pic_kwargs,
            )

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

            # 列宽
            for ci, w in enumerate(tspec.get("col_widths", [])):
                if ci < len(table.columns):
                    table.columns[ci].width = Inches(w)

            # 表头
            for ci, h in enumerate(headers):
                cell = table.cell(0, ci)
                cell.text = str(h)
                for p in cell.text_frame.paragraphs:
                    _apply_font(p, size=12, color="#FFFFFF", bold=True)
                    p.alignment = PP_ALIGN.CENTER
                cell.fill.solid()
                cell.fill.fore_color.rgb = RGBColor(0x33, 0x33, 0x33)

            # 数据
            for ri, row_data in enumerate(rows_data):
                for ci, val in enumerate(row_data):
                    cell = table.cell(ri + 1, ci)
                    cell.text = str(val)
                    for p in cell.text_frame.paragraphs:
                        _apply_font(p, size=11)
                        p.alignment = PP_ALIGN.CENTER

        # 演讲者备注
        if slide_spec.get("notes"):
            slide.notes_slide.notes_text_frame.text = str(slide_spec["notes"])

    prs.save(output_path)
    return json.dumps({"success": True, "output": output_path, "slide_count": len(prs.slides)}, ensure_ascii=False)


def _replace_in_paragraph(para, find, replace):
    """在段落内替换所有 find 出现处，支持跨多个 run 的匹配。返回替换次数。

    PPT 文本常被格式差异/拼写检查切成多个 run，旧版只在单个 run 内匹配会静默漏掉。
    替换文本写入首个受影响 run（保留其格式），其余 run 中对应片段删除。
    """
    if not find:
        return 0
    count = 0
    while count < 1000:  # 安全上限：replace 包含 find 时防止无限循环
        full_text = "".join(run.text for run in para.runs)
        idx = full_text.find(find)
        if idx == -1:
            break
        end = idx + len(find)
        pos = 0
        affected = []
        for run in para.runs:
            rlen = len(run.text)
            r_start, r_end = pos, pos + rlen
            if r_end > idx and r_start < end:
                affected.append((run, max(idx - r_start, 0), min(end - r_start, rlen)))
            pos += rlen
        first_run, fs, fe = affected[0]
        first_run.text = first_run.text[:fs] + replace + first_run.text[fe:]
        for run, s, e in affected[1:]:
            run.text = run.text[:s] + run.text[e:]
        count += 1
    return count


def _replace_in_text_frame(tf, find, replace):
    """在文本框内查找替换文本（支持跨 run）。返回替换次数。"""
    total = 0
    for para in tf.paragraphs:
        total += _replace_in_paragraph(para, find, replace)
    return total


def edit_pptx(file_path: str, spec_json: str) -> str:
    """
    编辑已有 PowerPoint 文件（.pptx）

    spec_json 格式:
    {
        "output": "/path/to/output.pptx",   // 可选，不指定则覆盖原文件
        "operations": [
            {"op": "replace_all", "find": "旧", "replace": "新"},       // 全局替换所有形状/表格文本
            {"op": "set_slide_title", "slide": 1, "text": "新标题"},    // 修改指定页标题
            {"op": "add_slide", "layout": 6, "title": "标题", "bullets": ["要点1","要点2"]},  // 添加幻灯片
            {"op": "delete_slide", "slide": 2}                          // 删除指定页
        ]
    }
    """
    spec, err = safe_parse_json(spec_json)
    if err:
        return json.dumps({"error": f"JSON 解析失败: {err}"}, ensure_ascii=False)

    try:
        prs = Presentation(file_path)
    except Exception as e:
        return json.dumps({"error": f"无法打开文件: {str(e)}"}, ensure_ascii=False)

    replacements = 0
    skipped: list[str] = []

    for op in spec.get("operations", []):
        kind = op.get("op")
        try:
            if kind == "replace_all":
                find, replace = op.get("find", ""), op.get("replace", "")
                for slide in prs.slides:
                    for shape in _iter_shapes(slide.shapes):  # 含 group 内部形状
                        if shape.has_text_frame:
                            replacements += _replace_in_text_frame(shape.text_frame, find, replace)
                        if getattr(shape, "has_table", False) and shape.has_table:
                            for row in shape.table.rows:
                                for cell in row.cells:
                                    replacements += _replace_in_text_frame(cell.text_frame, find, replace)

            elif kind == "set_slide_title":
                idx = op.get("slide", 1) - 1
                if not (0 <= idx < len(prs.slides)):
                    skipped.append(f"set_slide_title: 第 {idx + 1} 页超出范围（共 {len(prs.slides)} 页）")
                elif prs.slides[idx].shapes.title is None:
                    skipped.append(f"set_slide_title: 第 {idx + 1} 页没有标题占位符，无法修改")
                else:
                    prs.slides[idx].shapes.title.text = op.get("text", "")

            elif kind == "add_slide":
                layout_idx = op.get("layout", 6)
                if layout_idx not in (0, 1, 6):
                    return json.dumps(
                        {"error": f"add_slide: layout 只支持 0（标题页）/ 1（标题+内容）/ 6（空白），收到 {layout_idx}"},
                        ensure_ascii=False)
                if layout_idx >= len(prs.slide_layouts):
                    return json.dumps(
                        {"error": f"add_slide: 该模板只有 {len(prs.slide_layouts)} 种版式，无法使用 layout {layout_idx}"},
                        ensure_ascii=False)
                slide = prs.slides.add_slide(prs.slide_layouts[layout_idx])
                if op.get("title") and slide.shapes.title:
                    slide.shapes.title.text = op["title"]
                elif op.get("title"):
                    skipped.append(f"add_slide: 新页（layout {layout_idx}）没有标题占位符，标题未写入")
                if op.get("bullets"):
                    body = _find_body_shape(slide) or _add_body_textbox(slide, prs)
                    tf = body.text_frame
                    tf.clear()
                    for i, b in enumerate(op["bullets"]):
                        if i == 0:
                            tf.paragraphs[0].text = b
                        else:
                            p = tf.add_paragraph()
                            p.text = b

            elif kind == "delete_slide":
                idx = op.get("slide", 1) - 1
                xml_slides = prs.slides._sldIdLst
                slides = list(xml_slides)
                if 0 <= idx < len(slides):
                    xml_slides.remove(slides[idx])
                else:
                    skipped.append(f"delete_slide: 第 {idx + 1} 页超出范围（共 {len(slides)} 页）")

            else:
                return json.dumps({"error": f"未知操作: {kind}"}, ensure_ascii=False)

        except Exception as e:
            return json.dumps({"error": f"操作 '{kind}' 执行失败: {str(e)}"}, ensure_ascii=False)

    output_path = spec.get("output", file_path)
    prs.save(output_path)
    result = {"success": True, "output": output_path, "replacements": replacements}
    if skipped:
        result["skipped"] = skipped
    return json.dumps(result, ensure_ascii=False)
