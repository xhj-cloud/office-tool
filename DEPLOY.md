# Office & Filesystem Tools MCP — 部署文档

## 简介

集成了 **Office 文件读写编辑** + **PDF 解析修改** + **文件系统管理** 的 MCP 服务，共 27 个工具。部署到 CherryStudio 或 Claude CLI 后，AI 可以直接操作你电脑上的 Word、Excel、PPT、PDF 以及任意文件。

---

## 快速部署（3 步）

### 1. 复制项目

```bash
git clone https://github.com/xhj-cloud/office-tool.git
cd office-tool
```

> ⚠️ 不要复制 `venv/` 目录，它包含本机绝对路径，到新电脑上必须重建。

### 2. 安装依赖

```bash
cd office-tools-mcp

# 创建虚拟环境
python3 -m venv venv

# 安装依赖
./venv/bin/pip install -r requirements.txt
```

### 3. 配置到你的 AI 工具

根据你使用的工具，二选一：

#### 选项 A：CherryStudio

打开 **设置 → MCP 服务器 → 添加**：

| 字段 | 值 |
|------|-----|
| 名称 | `office-tools` |
| 传输类型 | `stdio` |
| 命令 | `<项目路径>/venv/bin/python` |
| 参数 | `<项目路径>/server.py` |

> 把 `<项目路径>` 替换成实际路径，比如 `/Users/xxx/projects/office-tools-mcp`

#### 选项 B：Claude CLI

编辑 `~/.claude/mcp.json`（没有则新建），写入：

```json
{
  "mcpServers": {
    "office-tools": {
      "type": "stdio",
      "command": "<项目路径>/venv/bin/python",
      "args": ["<项目路径>/server.py"]
    }
  }
}
```

同时同步到 `~/.claude.json` 的 `mcpServers` 字段。

---

## 验证安装

配置完成后重启 AI 工具，在新对话中说：

> "列出我桌面的文件"

如果返回了桌面文件列表，说明部署成功。

也可以手动验证：

```bash
cd <项目路径>
echo '{"jsonrpc":"2.0","id":1,"method":"tools/list","params":{}}' | ./venv/bin/python server.py
```

能输出 14 个工具名即正常。

---

## 工具清单（27 个）

| 类别 | 工具 | 功能 |
|------|------|------|
| Word | `read_docx` | 读取 .docx（段落/表格/样式） |
| | `write_docx` | 生成 .docx（标题/段落/表格/签署页） |
| | `edit_docx` | 编辑已有 .docx（替换/增删段落/改表格） |
| Excel | `read_xlsx` | 读取 .xlsx（多 Sheet/合并单元格） |
| | `write_xlsx` | 生成 .xlsx（多 Sheet/冻结/合并） |
| | `edit_xlsx` | 编辑已有 .xlsx（写值/增删行/合并/改名） |
| PPT | `read_pptx` | 读取 .pptx（文本/表格/大纲/备注） |
| | `write_pptx` | 生成 .pptx（标题/要点/表格） |
| | `edit_pptx` | 编辑已有 .pptx（替换/改标题/增删页面） |
| PDF | `pdf_info` | 获取 PDF 信息（页数/大小） |
| | `pdf_extract_text` | 提取文本内容 |
| | `pdf_render_pages` | 渲染页面为图片 |
| | `pdf_apply_modifications` | AI 驱动修改（替换/删除/添加） |
| | `pdf_save` | 保存修改后的 PDF |
| | `pdf_manual_replace_text` | 手动替换指定区域文本 |
| | `pdf_redact` | 涂黑/涂白遮盖 |
| | `pdf_highlight_area` | 高亮标注 |
| | `pdf_deyellow` | 扫描件去黄底 |
| | `pdf_add_text` | 指定坐标添加文本 |
| 文件 | `list_directory` | 列出目录内容 |
| | `read_file` | 读取文本文件 |
| | `write_file` | 创建或覆盖文件 |
| | `edit_file` | 查找替换编辑 |
| | `file_info` | 文件/目录详细信息 |
| | `create_directory` | 创建目录（含父目录） |
| | `move_file` | 移动或重命名 |
| | `delete_file` | 删除文件/目录 |

---

## 环境要求

| 依赖 | 最低版本 | 说明 |
|------|----------|------|
| Python | 3.10+ | 旧版请先升级 |
| pip | 自动随 Python 安装 |
| 操作系统 | macOS / Linux / Windows | 均支持 |

Python 包（`pip install -r requirements.txt` 自动安装）：

```
mcp>=1.0.0
python-docx>=1.0.0
openpyxl>=3.0.0
python-pptx>=1.0.0
PyMuPDF>=1.24.0
Pillow>=10.0.0
numpy>=1.24.0
```

---

## 目录结构

```
office-tool/
├── server.py               # MCP 服务入口（27 个工具）
├── start.sh                # 启动脚本
├── requirements.txt        # Python 依赖
├── tools/
│   ├── __init__.py
│   ├── word_tools.py        # Word 读写编辑
│   ├── excel_tools.py       # Excel 读写编辑
│   ├── ppt_tools.py         # PPT 读写编辑
│   ├── pdf_processor.py     # PDF 引擎（PyMuPDF）
│   ├── filesystem_tools.py  # 文件系统操作
│   └── json_repair.py       # JSON 自动修复
├── scripts/
│   ├── gen_contract.py      # 合同生成示例
│   └── github-mcp.sh        # GitHub MCP 启动脚本
└── venv/                    # 虚拟环境（部署时重建）
```

---

## Windows 部署指南

### 前提准备

1. **安装 Python 3.10+** — [python.org](https://www.python.org/downloads/)，安装时勾选 `Add Python to PATH`
2. **安装 Git** — [git-scm.com](https://git-scm.com/download/win)

### 部署步骤

打开 **PowerShell**，逐条执行：

```powershell
git clone https://github.com/xhj-cloud/office-tool.git
cd office-tool
python -m venv venv
venv\Scripts\pip install -r requirements.txt
```

> PyMuPDF 安装失败？先装 [VC_redist.x64.exe](https://aka.ms/vs/17/release/vc_redist.x64.exe) 后重试。

### 配置 CherryStudio

设置 → MCP 服务器 → 添加：

| 字段 | 值 |
|------|-----|
| 名称 | `office-tools` |
| 传输类型 | `stdio` |
| 命令 | `C:\Users\你的用户名\office-tool\venv\Scripts\python.exe` |
| 参数 | `C:\Users\你的用户名\office-tool\server.py` |

> 获取完整路径：项目目录下运行 `(Get-Item .\venv\Scripts\python.exe).FullName`

### 配置 Claude CLI

编辑 `%USERPROFILE%\.claude\mcp.json`：

```json
{
  "mcpServers": {
    "office-tools": {
      "type": "stdio",
      "command": "C:\\Users\\用户名\\office-tool\\venv\\Scripts\\python.exe",
      "args": ["C:\\Users\\用户名\\office-tool\\server.py"]
    }
  }
}
```

### 路径对照

| macOS | Windows |
|-------|---------|
| `/Users/xxx/Desktop` | `C:\Users\xxx\Desktop` |
| `python3` | `python` |
| `venv/bin/python` | `venv\Scripts\python.exe` |
| `~/.claude/mcp.json` | `%USERPROFILE%\.claude\mcp.json` |

---

## 常见问题

**Q: 启动报 `connection closed`？**
- 检查 Python 版本是否 ≥ 3.10
- 检查 `venv/` 是否在当前电脑上重建（不能从别的电脑复制）
- 确认 `requirements.txt` 全部安装成功

**Q: AI 说无法访问桌面文件？**
- MCP 本身无路径限制，这是 AI 模型过于谨慎
- 在对话中提供完整绝对路径，如 `/Users/xxx/Desktop/文件.docx`
- 或在 Agent 的系统提示词中说明："可以通过 MCP 工具访问本机任意路径"

**Q: 想新增工具？**
- 在 `tools/` 下新建 `.py` 文件，编写工具函数
- 在 `server.py` 中用 `@mcp.tool()` 注册
- 在 `tools/__init__.py` 中导出

**Q: PDF 中文乱码？** 程序自动检测系统 CJK 字体，如仍乱码请安装中文字体包。

**Q: PyMuPDF 安装失败？** macOS 可尝试 `brew install mupdf` 后重装；Linux 需 `apt install libmupdf-dev`。

**Q: Windows 怎么配置路径？**
- 命令: `C:\Users\xxx\office-tool\venv\Scripts\python.exe`
- 参数: `C:\Users\xxx\office-tool\server.py`
