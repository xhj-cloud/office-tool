"""
文件系统工具
提供文件的读写、目录浏览、文件移动等基础操作

安全限制：所有路径必须位于白名单目录内（先 resolve 再校验，符号链接指向外部同样被拒绝）。
环境变量 OFFICE_TOOLS_ALLOWED_DIRS 可配置多个允许的根目录（以 os.pathsep 分隔）；
未设置时默认仅允许用户主目录。
"""

import os
import json
import shutil
import mimetypes
from pathlib import Path


# ── 路径白名单（安全限制）──────────────────────────────
def _allowed_roots() -> list:
    """读取允许操作的根目录列表"""
    raw = os.environ.get("OFFICE_TOOLS_ALLOWED_DIRS", "").strip()
    if raw:
        roots = [Path(d).expanduser().resolve() for d in raw.split(os.pathsep) if d.strip()]
        if roots:
            return roots
    return [Path.home().resolve()]


def _check_allowed(p):
    """检查路径是否在白名单内；允许返回 None，否则返回错误 JSON 字符串"""
    for root in _allowed_roots():
        if p == root or root in p.parents:
            return None
    return json.dumps({
        "error": f"路径不在允许范围内: {p}（可通过环境变量 OFFICE_TOOLS_ALLOWED_DIRS 配置允许的目录）",
    }, ensure_ascii=False)


def list_directory(path: str) -> str:
    """列出目录内容"""
    try:
        p = Path(path).resolve()
        err = _check_allowed(p)
        if err:
            return err
        if not p.exists():
            return json.dumps({"error": f"路径不存在: {path}"}, ensure_ascii=False)
        if not p.is_dir():
            return json.dumps({"error": f"不是目录: {path}"}, ensure_ascii=False)

        items = []
        for entry in sorted(p.iterdir()):
            try:
                stat = entry.stat()
                items.append({
                    "name": entry.name,
                    "type": "dir" if entry.is_dir() else "file",
                    "size": stat.st_size,
                    "modified": stat.st_mtime,
                })
            except PermissionError:
                items.append({
                    "name": entry.name,
                    "type": "unknown",
                    "size": 0,
                    "modified": 0,
                })

        return json.dumps({
            "path": str(p),
            "items": items,
            "count": len(items),
        }, ensure_ascii=False, indent=2)
    except Exception as e:
        return json.dumps({"error": str(e)}, ensure_ascii=False)


def read_file(path: str, encoding: str = "utf-8", max_lines: int = 500) -> str:
    """读取文本文件内容"""
    try:
        p = Path(path).resolve()
        err = _check_allowed(p)
        if err:
            return err
        if not p.exists():
            return json.dumps({"error": f"文件不存在: {path}"}, ensure_ascii=False)
        if p.is_dir():
            return json.dumps({"error": f"是目录不是文件: {path}"}, ensure_ascii=False)

        # 检查文件大小
        size = p.stat().st_size
        if size > 5 * 1024 * 1024:  # 5MB 限制
            return json.dumps({
                "error": f"文件过大 ({size / 1024 / 1024:.1f}MB)，超过 5MB 限制",
                "path": str(p),
                "size": size,
            }, ensure_ascii=False)

        with open(p, "r", encoding=encoding, errors="replace") as f:
            lines = f.readlines()

        total = len(lines)
        truncated = total > max_lines
        if truncated:
            lines = lines[:max_lines]

        return json.dumps({
            "path": str(p),
            "size": size,
            "total_lines": total,
            "lines_shown": len(lines),
            "truncated": truncated,
            "content": "".join(lines),
        }, ensure_ascii=False)
    except UnicodeDecodeError:
        # 二进制文件
        return json.dumps({
            "path": str(p),
            "size": size,
            "type": "binary",
            "hint": "这是二进制文件，无法以文本方式读取",
        }, ensure_ascii=False)
    except Exception as e:
        return json.dumps({"error": str(e)}, ensure_ascii=False)


def write_file(path: str, content: str, encoding: str = "utf-8") -> str:
    """写入内容到文件（覆盖模式）"""
    try:
        p = Path(path).resolve()
        err = _check_allowed(p)
        if err:
            return err
        p.parent.mkdir(parents=True, exist_ok=True)

        with open(p, "w", encoding=encoding) as f:
            f.write(content)

        return json.dumps({
            "success": True,
            "path": str(p),
            "size": p.stat().st_size,
        }, ensure_ascii=False)
    except Exception as e:
        return json.dumps({"error": str(e)}, ensure_ascii=False)


def edit_file(path: str, old_string: str, new_string: str,
              encoding: str = "utf-8", replace_all: bool = False) -> str:
    """编辑文件内容（查找替换）"""
    try:
        p = Path(path).resolve()
        err = _check_allowed(p)
        if err:
            return err
        if not p.exists():
            return json.dumps({"error": f"文件不存在: {path}"}, ensure_ascii=False)

        with open(p, "r", encoding=encoding) as f:
            content = f.read()

        if replace_all:
            occurrences = content.count(old_string)
            if occurrences == 0:
                return json.dumps({"error": "未找到匹配文本"}, ensure_ascii=False)
            new_content = content.replace(old_string, new_string)
        else:
            occurrences = content.count(old_string)
            if occurrences == 0:
                return json.dumps({"error": "未找到匹配文本"}, ensure_ascii=False)
            if occurrences > 1:
                return json.dumps({
                    "error": f"找到 {occurrences} 处匹配，请提供更精确的上下文或设置 replace_all=true",
                }, ensure_ascii=False)
            new_content = content.replace(old_string, new_string, 1)

        with open(p, "w", encoding=encoding) as f:
            f.write(new_content)

        return json.dumps({
            "success": True,
            "path": str(p),
            "occurrences": occurrences if replace_all else 1,
        }, ensure_ascii=False)
    except Exception as e:
        return json.dumps({"error": str(e)}, ensure_ascii=False)


def file_info(path: str) -> str:
    """获取文件/目录信息"""
    try:
        p = Path(path).resolve()
        err = _check_allowed(p)
        if err:
            return err
        if not p.exists():
            return json.dumps({"error": f"路径不存在: {path}"}, ensure_ascii=False)

        stat = p.stat()
        info = {
            "path": str(p),
            "name": p.name,
            "type": "dir" if p.is_dir() else "file",
            "size": stat.st_size,
            "created": stat.st_ctime,
            "modified": stat.st_mtime,
            "accessed": stat.st_atime,
        }

        if p.is_file():
            mime, _ = mimetypes.guess_type(str(p))
            info["mime_type"] = mime
            info["extension"] = p.suffix

        return json.dumps(info, ensure_ascii=False, indent=2)
    except Exception as e:
        return json.dumps({"error": str(e)}, ensure_ascii=False)


def create_directory(path: str) -> str:
    """创建目录（含父目录）"""
    try:
        p = Path(path).resolve()
        err = _check_allowed(p)
        if err:
            return err
        p.mkdir(parents=True, exist_ok=True)
        return json.dumps({"success": True, "path": str(p)}, ensure_ascii=False)
    except Exception as e:
        return json.dumps({"error": str(e)}, ensure_ascii=False)


def move_file(source: str, destination: str) -> str:
    """移动/重命名文件或目录"""
    try:
        src = Path(source).resolve()
        dst = Path(destination).resolve()
        err = _check_allowed(src) or _check_allowed(dst)
        if err:
            return err
        if not src.exists():
            return json.dumps({"error": f"源路径不存在: {source}"}, ensure_ascii=False)

        dst.parent.mkdir(parents=True, exist_ok=True)
        shutil.move(str(src), str(dst))

        return json.dumps({
            "success": True,
            "from": str(src),
            "to": str(dst),
        }, ensure_ascii=False)
    except Exception as e:
        return json.dumps({"error": str(e)}, ensure_ascii=False)


def delete_file(path: str, recursive: bool = False) -> str:
    """删除文件或目录"""
    try:
        p = Path(path).resolve()
        err = _check_allowed(p)
        if err:
            return err
        if not p.exists():
            return json.dumps({"error": f"路径不存在: {path}"}, ensure_ascii=False)

        # 禁止递归删除白名单根目录本身（如整个主目录）
        if recursive and any(p == root for root in _allowed_roots()):
            return json.dumps({
                "error": f"拒绝递归删除允许范围的根目录: {p}",
            }, ensure_ascii=False)

        if p.is_dir():
            if recursive:
                shutil.rmtree(str(p))
            else:
                p.rmdir()  # 只删空目录
        else:
            p.unlink()

        return json.dumps({"success": True, "path": str(p)}, ensure_ascii=False)
    except Exception as e:
        return json.dumps({"error": str(e)}, ensure_ascii=False)
