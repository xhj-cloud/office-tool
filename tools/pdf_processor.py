"""
PDF 处理器
基于 PyMuPDF (fitz) 的 PDF 读取、渲染和修改功能
支持中文字体（自动检测 macOS/Windows/Linux 系统字体）
"""

import glob
import io
import json
import os
import platform
from pathlib import Path
from typing import Optional

import fitz  # PyMuPDF


# ── CJK 字体自动检测 ────────────────────────────────────
def _find_cjk_font() -> Optional[str]:
    """自动查找系统中的 CJK 字体文件"""
    system = platform.system()

    if system == "Darwin":  # macOS
        candidates = [
            "/System/Library/Fonts/STHeiti Medium.ttc",
            "/System/Library/Fonts/PingFang.ttc",
            "/System/Library/Fonts/Hiragino Sans GB.ttc",
        ]
        # 也搜索 AssetsV2 中的 PingFang
        pf = glob.glob(
            "/System/Library/AssetsV2/**/PingFang.ttc", recursive=True
        )
        candidates = pf + candidates

    elif system == "Windows":
        candidates = [
            "C:/Windows/Fonts/msyh.ttc",   # 微软雅黑
            "C:/Windows/Fonts/msyhbd.ttc",
            "C:/Windows/Fonts/simsun.ttc",  # 宋体
            "C:/Windows/Fonts/simhei.ttf",  # 黑体
        ]
    else:  # Linux
        candidates = [
            "/usr/share/fonts/truetype/droid/DroidSansFallbackFull.ttf",
            "/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc",
            "/usr/share/fonts/noto-cjk/NotoSansCJK-Regular.ttc",
        ]

    for path in candidates:
        if os.path.exists(path):
            return path
    return None


# 全局 CJK 字体路径
_CJK_FONT_PATH = _find_cjk_font()

# PyMuPDF base-14 内置字体名：insert_text 的 fontname 若取这些值，
# 会直接使用内置字体并忽略 fontfile（中文将渲染成点阵且无法提取）
_BASE14_NAMES = {"helv", "he", "tiro", "cour", "symb", "zadb"}

# 嵌入 CJK 字体文件时使用的非内置别名
_CJK_FONT_NAME = "cjk"


class PDFProcessor:
    """PDF 文件读取、渲染、修改的统一处理器"""

    def __init__(self, filepath: str, cjk_font: Optional[str] = None):
        """
        Args:
            filepath: PDF 文件路径
            cjk_font: CJK 字体文件路径，默认自动检测
        """
        self.filepath = filepath
        self._doc: Optional[fitz.Document] = None
        self.cjk_font = cjk_font or _CJK_FONT_PATH

    @property
    def doc(self) -> fitz.Document:
        if self._doc is None:
            if not os.path.exists(self.filepath):
                raise FileNotFoundError(f"PDF 文件不存在: {self.filepath}")
            self._doc = fitz.open(self.filepath)
        return self._doc

    @staticmethod
    def _has_cjk(text: str) -> bool:
        """检测文本是否包含 CJK 字符"""
        for ch in text:
            cp = ord(ch)
            if (0x4E00 <= cp <= 0x9FFF or  # CJK 统一汉字
                0x3400 <= cp <= 0x4DBF or  # CJK 扩展 A
                0x3000 <= cp <= 0x303F or  # CJK 标点
                0xFF00 <= cp <= 0xFFEF or  # 全角字符
                0x3040 <= cp <= 0x309F or  # 平假名
                0x30A0 <= cp <= 0x30FF):   # 片假名
                return True
        return False

    def close(self):
        if self._doc:
            self._doc.close()
            self._doc = None

    # ── 读取 & 信息 ──────────────────────────────────────

    @property
    def page_count(self) -> int:
        return len(self.doc)

    def resolve_page(self, page_num: int):
        """校验 1-based 页码并返回对应 Page（越界/非法值抛 ValueError）"""
        if isinstance(page_num, bool) or not isinstance(page_num, int):
            raise ValueError(f"page_num 必须是整数，收到: {page_num!r}")
        if not (1 <= page_num <= self.page_count):
            raise ValueError(
                f"页码 {page_num} 超出范围（文档共 {self.page_count} 页，有效值 1-{self.page_count}）"
            )
        return self.doc[page_num - 1]

    def get_info(self) -> dict:
        """获取 PDF 元信息"""
        meta = self.doc.metadata
        return {
            "filepath": self.filepath,
            "page_count": self.page_count,
            "title": meta.get("title", ""),
            "author": meta.get("author", ""),
            "subject": meta.get("subject", ""),
            "format": meta.get("format", "PDF"),
            "file_size_mb": round(os.path.getsize(self.filepath) / (1024 * 1024), 2),
        }

    def extract_text(self, page_range: Optional[str] = None) -> list[dict]:
        """
        提取文本

        Args:
            page_range: 页码范围（1-based），如 "1-3" 或 "2,5" 或 None（全部）

        Returns:
            [{"page": 1, "text": "..."}, ...]
        """
        pages = self._parse_page_range(page_range)
        results = []
        for pn in pages:
            page = self.doc[pn - 1]
            text = page.get_text()
            results.append({"page": pn, "text": text})
        return results

    # ── 渲染为图片（供视觉 AI 使用） ──────────────────────

    def render_pages(
        self,
        page_range: Optional[str] = None,
        dpi: int = 200,
        output_dir: Optional[str] = None,
    ) -> list[dict]:
        """
        将 PDF 页面渲染为 PNG 图片

        Args:
            page_range: 页码范围（1-based），如 "1-5"
            dpi: 渲染分辨率
            output_dir: 如果指定，将图片保存到此目录

        Returns:
            [{"page": 1, "image_path": "...", "width": ..., "height": ...}, ...]
        """
        from PIL import Image

        pages = self._parse_page_range(page_range)
        zoom = dpi / 72  # PDF 默认 72 DPI
        mat = fitz.Matrix(zoom, zoom)

        results = []
        for pn in pages:
            page = self.doc[pn - 1]
            pix = page.get_pixmap(matrix=mat)
            img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)

            if output_dir:
                os.makedirs(output_dir, exist_ok=True)
                img_path = os.path.join(output_dir, f"page_{pn:04d}.png")
                img.save(img_path)
            else:
                # 保存到临时文件
                import tempfile
                tmp = tempfile.NamedTemporaryFile(suffix=".png", delete=False)
                img_path = tmp.name
                img.save(img_path)
                tmp.close()

            results.append({
                "page": pn,
                "image_path": img_path,
                "width": pix.width,
                "height": pix.height,
            })

        return results

    # ── PDF 修改 ─────────────────────────────────────────

    def add_text(
        self,
        page_num: int,
        text: str,
        x: float,
        y: float,
        font_size: float = 12,
        color: tuple = (0, 0, 0),
        font_name: str = "helv",
        font_file: Optional[str] = None,
    ) -> dict:
        """
        在指定位置添加文本（自动使用 CJK 字体处理中文）

        Args:
            page_num: 页码（从 1 开始）
            text: 文本内容
            x, y: 坐标（点）
            font_size: 字号
            color: RGB 颜色元组 (0-1)
            font_name: 字体名（内置字体）
            font_file: 字体文件路径，默认自动使用系统 CJK 字体

        Returns:
            {"success": bool, ...}
        """
        page = self.resolve_page(page_num)
        kwargs = {
            "point": fitz.Point(x, y),
            "text": text,
            "fontsize": font_size,
            "color": color,
            "fontname": font_name,
        }
        # 检测是否包含中文字符，自动使用 CJK 字体
        if self._has_cjk(text):
            detected_font = font_file or self.cjk_font
            if detected_font and os.path.exists(detected_font):
                kwargs["fontfile"] = detected_font
                # fontname 为 base-14 内置名时 PyMuPDF 会忽略 fontfile，
                # 必须改用非内置别名，CJK 字体文件才会真正嵌入
                if kwargs["fontname"] in _BASE14_NAMES:
                    kwargs["fontname"] = _CJK_FONT_NAME

        page.insert_text(**kwargs)
        return {
            "success": True,
            "page": page_num,
            "action": "add_text",
            "position": (x, y),
            "text": text,
        }

    def redact_area(
        self,
        page_num: int,
        rect: tuple,  # (x0, y0, x1, y1)
        fill_color: tuple = (0, 0, 0),
    ) -> dict:
        """
        遮盖（涂黑）指定区域

        Args:
            page_num: 页码（从 1 开始）
            rect: 矩形区域 (x0, y0, x1, y1)，单位：点
            fill_color: 填充颜色，默认黑色

        Returns:
            {"success": bool, ...}
        """
        page = self.resolve_page(page_num)
        r = fitz.Rect(*rect)
        annot = page.add_redact_annot(r, fill=fill_color)
        page.apply_redactions(images=2)  # images=2 彻底移除与遮盖区域重叠的图像
        return {
            "success": True,
            "page": page_num,
            "action": "redact",
            "rect": rect,
        }

    def add_highlight(
        self,
        page_num: int,
        rect: tuple,
        color: tuple = (1, 1, 0),  # 黄色
        opacity: float = 0.3,
    ) -> dict:
        """
        高亮标注区域

        Args:
            page_num: 页码（从 1 开始）
            rect: 矩形区域
            color: RGB 颜色（0-1 范围）
            opacity: 透明度
        """
        page = self.resolve_page(page_num)
        annot = page.add_highlight_annot(fitz.Rect(*rect))
        annot.set_colors(stroke=color)
        annot.set_opacity(opacity)
        annot.update()
        return {
            "success": True,
            "page": page_num,
            "action": "highlight",
            "rect": rect,
        }

    def add_rect_annot(
        self,
        page_num: int,
        rect: tuple,
        text: str = "",
        color: tuple = (1, 0, 0),
        width: float = 1.5,
    ) -> dict:
        """
        添加矩形标注框

        Args:
            page_num: 页码（从 1 开始）
            rect: 矩形区域
            text: 标注文本
            color: RGB（0-1）
            width: 线宽
        """
        page = self.resolve_page(page_num)
        annot = page.add_rect_annot(fitz.Rect(*rect))
        annot.set_colors(stroke=color)
        annot.set_border(width=width)
        if text:
            annot.set_info(content=text)
        annot.update()
        return {
            "success": True,
            "page": page_num,
            "action": "rect_annot",
            "rect": rect,
        }

    def replace_text_in_bbox(
        self,
        page_num: int,
        bbox: tuple,
        new_text: str,
        font_size: float = 12,
        color: tuple = (0, 0, 0),
        font_file: Optional[str] = None,
    ) -> dict:
        """
        替换指定区域内的文本（遮盖旧文本 + 写入新文本）

        Args:
            page_num: 页码（从 1 开始）
            bbox: 目标矩形区域 (x0, y0, x1, y1)
            new_text: 新文本
            font_size: 字号
            color: RGB 颜色 (0-1)
            font_file: 字体文件，默认自动使用系统 CJK 字体
        """
        page = self.resolve_page(page_num)
        r = fitz.Rect(*bbox)

        # 1. 遮盖旧文本
        page.add_redact_annot(r, fill=(1, 1, 1))  # 白色遮盖
        page.apply_redactions(images=2)  # images=2 彻底移除与遮盖区域重叠的图像

        # 2. 写入新文本
        x_center = (bbox[0] + bbox[2]) / 2
        y_center = (bbox[1] + bbox[3]) / 2

        cjk = font_file or self.cjk_font
        use_cjk = self._has_cjk(new_text) and cjk and os.path.exists(cjk)
        # get_text_length 不支持 fontfile，CJK 文字宽度假定为 font_size
        if use_cjk:
            tw = font_size * len(new_text)
        else:
            tw = fitz.get_text_length(new_text, fontsize=font_size)
        x_start = x_center - tw / 2

        ins = {
            "point": fitz.Point(x_start, y_center + font_size * 0.35),
            "text": new_text,
            "fontsize": font_size,
            "color": color,
        }
        if use_cjk:
            ins["fontfile"] = cjk
            # 必须指定非 base-14 的字体名，否则 PyMuPDF 用内置 Helvetica、忽略 fontfile
            ins["fontname"] = _CJK_FONT_NAME
        page.insert_text(**ins)
        return {
            "success": True,
            "page": page_num,
            "action": "replace_text",
            "bbox": bbox,
            "new_text": new_text,
        }

    def apply_ai_modifications(
        self,
        page_num: int,
        modifications: list[dict],
        page_width: float,
        page_height: float,
    ) -> list[dict]:
        """
        应用 AI 返回的修改指令

        Args:
            page_num: 页码（从 1 开始）
            modifications: AI 返回的修改指令列表
            page_width, page_height: 页面尺寸（点）

        Returns:
            执行结果列表
        """
        results = []
        for mod in modifications:
            mod_type = mod.get("type", "")
            action = mod.get("action", "")

            # 将百分比坐标转为实际坐标
            bbox_pct = mod.get("bbox")
            if bbox_pct and bbox_pct != [None, None, None, None]:
                bbox_pts = (
                    bbox_pct[0] / 100 * page_width,
                    bbox_pct[1] / 100 * page_height,
                    bbox_pct[2] / 100 * page_width,
                    bbox_pct[3] / 100 * page_height,
                )
            else:
                bbox_pts = None

            try:
                if mod_type == "text_box":
                    if action == "replace" and bbox_pts:
                        r = self.replace_text_in_bbox(
                            page_num,
                            bbox_pts,
                            mod.get("new_text", ""),
                            mod.get("font_size", 12),
                        )
                    elif action == "delete" and bbox_pts:
                        r = self.redact_area(page_num, bbox_pts, fill_color=(1, 1, 1))
                    elif action == "add" and bbox_pts:
                        r = self.add_text(
                            page_num,
                            mod.get("new_text", ""),
                            bbox_pts[0],
                            bbox_pts[1],
                            mod.get("font_size", 12),
                        )
                    else:
                        r = {"success": False, "error": f"未知操作: {action}"}
                else:
                    r = {"success": False, "error": f"未知类型: {mod_type}"}

                r["reason"] = mod.get("reason", "")
                results.append(r)

            except Exception as e:
                results.append({
                    "success": False,
                    "error": str(e),
                    "modification": mod,
                })

        return results

    # ── 去黄处理（扫描件专用）──────────────────────────

    def deyellow(
        self,
        page_range: Optional[str] = None,
        dpi: int = 200,
        strength: float = 0.95,
        threshold: float = 120.0,
    ) -> dict:
        """
        去除扫描件 PDF 的黄底，使纸张变白。

        原理：逐像素分析亮度，对亮区（纸张）提升蓝色通道至与红绿均等，
        暗区（文字/印章）不受影响。

        Args:
            page_range: 页码范围（1-based），如 '1-3'，留空处理所有页
            dpi: 渲染分辨率（默认 200）
            strength: 去黄强度 0-1（默认 0.95）
            threshold: 亮度阈值，高于此值视为纸张（默认 120）

        Returns:
            {"success": bool, "pages_processed": int, ...}
        """
        from PIL import Image
        import io
        import numpy as np

        pages = self._parse_page_range(page_range)
        results = []

        for pn in pages:
            page = self.doc[pn - 1]
            page_rect = page.rect
            zoom = dpi / 72
            mat = fitz.Matrix(zoom, zoom)
            pix = page.get_pixmap(matrix=mat)

            # 转为 numpy 数组
            arr = np.frombuffer(pix.samples, dtype=np.uint8).reshape(
                pix.height, pix.width, 3
            ).astype(np.float64)

            h, w = arr.shape[:2]

            # 底部扫描边缘清理
            edge_px = max(3, int(h * 0.003))
            arr[h - edge_px:, :, :] = 255

            # 亮度权重：亮区（纸张）完全校正，暗区（文字）忽略
            brightness = arr.mean(axis=2)
            bright_weight = np.clip(brightness / threshold, 0, 1)

            # 逐像素蓝通道校正：使 B 与 R、G 均等
            r, g, b = arr[:, :, 0].copy(), arr[:, :, 1].copy(), arr[:, :, 2].copy()
            target_b = (r + g) / 2.0
            b_deficit = target_b - b
            b_new = np.clip(b + b_deficit * strength * bright_weight, 0, 255)
            arr[:, :, 2] = b_new

            # 轻微绿通道校正
            g_new = np.clip(g + (r - g) * 0.3 * bright_weight, 0, 255)
            arr[:, :, 1] = g_new

            # 插入校正后图像（覆盖在原始内容之上）
            img = Image.fromarray(arr.astype(np.uint8))
            img_bytes = io.BytesIO()
            img.save(img_bytes, format="JPEG", quality=94, optimize=True)
            img_bytes.seek(0)

            # 先用白色矩形遮盖整个页面
            page.draw_rect(page_rect, color=None, fill=(1, 1, 1))
            # 再插入校正后的图像
            page.insert_image(page_rect, stream=img_bytes.read())

            results.append({
                "page": pn,
                "original_rgb": f"({r.mean():.0f},{g.mean():.0f},{b.mean():.0f})",
                "corrected_rgb": f"({arr[:,:,0].mean():.0f},{arr[:,:,1].mean():.0f},{arr[:,:,2].mean():.0f})",
            })

        return {
            "success": True,
            "pages_processed": len(results),
            "dpi": dpi,
            "strength": strength,
            "pages": results,
        }

    # ── 保存 & 工具方法 ──────────────────────────────────

    def save(self, output_path: Optional[str] = None, incremental: bool = False) -> str:
        """
        保存修改后的 PDF

        Args:
            output_path: 输出路径
            incremental: 是否增量保存

        Returns:
            保存的文件路径
        """
        if output_path is None:
            base, ext = os.path.splitext(self.filepath)
            output_path = f"{base}_modified{ext}"

        if os.path.abspath(output_path) == os.path.abspath(self.filepath):
            raise ValueError(
                "不能保存回源文件（PyMuPDF 不允许覆盖自己打开的文件）。"
                "请指定其他输出路径，或先关闭该 PDF 再覆盖。")

        if incremental:
            self.doc.save(output_path, incremental=True, encryption=0)
        else:
            self.doc.save(output_path)

        return output_path

    def _parse_page_range(self, page_range: Optional[str]) -> list[int]:
        """解析页码范围字符串（1-based，如 "1-3"、"2,5"）

        越界或格式非法时抛 ValueError（不再静默丢弃）。
        """
        if page_range is None or not str(page_range).strip():
            return list(range(1, self.page_count + 1))

        total = self.page_count
        pages: set[int] = set()
        for part in str(page_range).split(","):
            part = part.strip()
            if not part:
                continue
            try:
                if "-" in part:
                    start_s, end_s = part.split("-", 1)
                    start, end = int(start_s), int(end_s)
                    if start > end:
                        raise ValueError(f"范围起始大于结束: {part!r}")
                    pages.update(range(start, end + 1))
                else:
                    pages.add(int(part))
            except ValueError as e:
                raise ValueError(
                    f"无法解析页码片段 {part!r}: {e}（格式如 '1-3'、'2,5'，页码从 1 开始）"
                ) from None

        bad = sorted(p for p in pages if not (1 <= p <= total))
        if bad:
            raise ValueError(
                f"页码超出范围: {bad}（文档共 {total} 页，有效值 1-{total}）"
            )
        return sorted(pages)
