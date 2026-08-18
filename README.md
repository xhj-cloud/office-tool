# Office & Filesystem & PDF MCP 工具包

## 这是什么？

一个 MCP（Model Context Protocol）服务，装到 CherryStudio 或 Claude CLI 后，你的 AI 助手就能直接操作电脑上的文件——Word/Excel/PPT 三大 Office 格式的读写编辑 + PDF 解析修改 + 文件系统管理。

**本质上就是给 AI 装了一双能操作你电脑文件的手。**

---

## 能做什么？

装好之后，直接用自然语言对 AI 说：

### 📄 Word 文档
| 你对 AI 说 | AI 做的事 |
|------------|-----------|
| "读一下桌面的合同.docx，提取甲乙方信息和金额" | 读取 Word，解析内容 |
| "生成一份采购合同，甲方填XX公司，金额5000" | 生成带格式的合同文档 |
| "把合同.docx 第五条金额改成 5000" | 编辑已有文档 |

### 📊 Excel 表格
| 你对 AI 说 | AI 做的事 |
|------------|-----------|
| "分析销售数据.xlsx 的月度趋势" | 读取 Excel 数据 |
| "把合同报价单提取出来，生成 Excel" | Word → Excel 转换 |
| "在报价单.xlsx 第 5 行后面插入一条新数据" | 编辑已有表格 |

### 🎞 PPT 演示文稿
| 你对 AI 说 | AI 做的事 |
|------------|-----------|
| "根据这份 Word 报告生成一个 5 页的 PPT" | 报告 → 演示文稿 |
| "把第三页的标题改成 Q4 总结" | 编辑已有 PPT |
| "给 PPT 最后加一页总结页" | 增删幻灯片 |

### 📑 PDF 文档
| 你对 AI 说 | AI 做的事 |
|------------|-----------|
| "扫描件去黄底，让纸张变白" | 智能去黄 |
| "把 PDF 第 2 页的公司名改成 XXX" | 替换文本 |
| "把 PDF 渲染成图片看看" | 渲染为图片供分析 |
| "遮盖合同里的身份证号" | 涂黑敏感区域 |

### 📁 文件管理
| 你对 AI 说 | AI 做的事 |
|------------|-----------|
| "列出桌面上所有的文档" | 浏览目录 |
| "把下载文件夹里所有 PDF 归档到新目录" | 批量移动 |
| "读取 app.py 第 20 到 50 行" | 读取文件 |
| "搜索所有 .log 文件" | 递归搜索 |

### 🔒 安全限制与编号约定

**路径白名单**：文件管理工具的所有路径都受白名单限制，只允许操作允许的目录内。

- 默认仅允许用户主目录；
- 用环境变量 `OFFICE_TOOLS_ALLOWED_DIRS` 配置多个允许的根目录（macOS/Linux 以 `:` 分隔）：

```bash
export OFFICE_TOOLS_ALLOWED_DIRS="$HOME/Documents:$HOME/Downloads"
```

路径在符号链接解析后校验，无法通过软链接逃逸；禁止递归删除白名单根目录本身。

**编号约定**：

- PDF 所有页码（`page_num`、`page_range`）统一从 **1** 开始，越界会明确报错；
- `edit_docx` 的段落 `index` 相对于文档初始状态（与 `read_docx` 返回的编号一致），批量操作互不影响。

---

## 全部 27 个工具

| 类别 | 工具 | 说明 |
|------|------|------|
| **Word** | `read_docx` | 读取文档（段落/表格/样式） |
| | `write_docx` | 创建文档（标题/正文/表格/签署页） |
| | `edit_docx` | 编辑已有文档（替换/增删段落/改表格） |
| **Excel** | `read_xlsx` | 读取表格（多 Sheet/合并单元格） |
| | `write_xlsx` | 创建表格（表头/冻结/样式） |
| | `edit_xlsx` | 编辑已有表格（写值/增删行/合并/改名） |
| **PPT** | `read_pptx` | 读取幻灯片（文本/表格/大纲/备注） |
| | `write_pptx` | 创建演示文稿（标题/要点/表格） |
| | `edit_pptx` | 编辑已有 PPT（替换文本/改标题/增删页面） |
| **PDF** | `pdf_info` | 获取 PDF 信息（页数/大小） |
| | `pdf_extract_text` | 提取文本内容 |
| | `pdf_render_pages` | 渲染页面为图片（供 AI 视觉分析） |
| | `pdf_apply_modifications` | AI 驱动修改（自动替换/删除/添加） |
| | `pdf_save` | 保存修改后的 PDF |
| | `pdf_manual_replace_text` | 手动替换指定区域文本 |
| | `pdf_redact` | 涂黑/涂白遮盖敏感区域 |
| | `pdf_highlight_area` | 高亮标注 |
| | `pdf_deyellow` | 扫描件去黄底 |
| | `pdf_add_text` | 指定坐标添加文本 |
| **文件** | `list_directory` | 浏览文件夹内容 |
| | `read_file` | 读取文本文件 |
| | `write_file` | 创建/覆盖写入文件 |
| | `edit_file` | 查找替换编辑 |
| | `file_info` | 文件/目录详细信息 |
| | `create_directory` | 创建目录（含父目录） |
| | `move_file` | 移动/重命名 |
| | `delete_file` | 删除文件/目录 |

---

## 怎么装？

### 前提条件

- Python 3.10+
- 已装 CherryStudio 或 Claude CLI
- 如需 PDF 功能，系统需有中文字体（macOS/Windows 自带，Linux 可能需安装）

### 3 步搞定

**第 1 步：克隆项目**

```bash
git clone https://github.com/xhj-cloud/office-tool.git
cd office-tool
```

**第 2 步：安装依赖**

```bash
python3 -m venv venv
./venv/bin/pip install -r requirements.txt
```

> Windows: `venv\Scripts\pip install -r requirements.txt`

**第 3 步：配置到 AI 工具**

<details>
<summary>🔹 CherryStudio</summary>

设置 → MCP 服务器 → 添加：

| 配置项 | 值 |
|--------|-----|
| 名称 | `office-tools` |
| 传输类型 | `stdio` |
| 命令 | 项目目录下的 `venv/bin/python` |
| 参数 | `server.py` |

</details>

<details>
<summary>🔹 Claude CLI</summary>

```json
{ "mcpServers": { "office-tools": {
    "type": "stdio",
    "command": "实际路径/venv/bin/python",
    "args": ["实际路径/server.py"]
}}}
```

</details>

### 验证

> "列出我桌面的文件"

返回文件列表即成功。

---

## 项目结构

```
office-tool/
├── README.md               ← 本文档
├── DEPLOY.md               ← 详细部署文档
├── server.py               ← MCP 主程序（27 个工具）
├── requirements.txt        ← Python 依赖清单
├── tools/
│   ├── word_tools.py        # Word 读写编辑
│   ├── excel_tools.py       # Excel 读写编辑
│   ├── ppt_tools.py         # PPT 读写编辑
│   ├── pdf_processor.py     # PDF 引擎（PyMuPDF）
│   ├── filesystem_tools.py  # 文件系统操作
│   └── json_repair.py       # JSON 自动修复
└── scripts/
    ├── gen_contract.py      # 合同生成示例
    └── github-mcp.sh       # GitHub MCP 启动
```

---

## 常见问题

**Q: AI 说"无法访问桌面"？** 给完整绝对路径：`/Users/xxx/Desktop/文件.docx`

**Q: 启动报 connection closed？** 检查 Python ≥ 3.10，确保 `pip install -r requirements.txt` 无报错。

**Q: PDF 中文乱码？** 程序自动检测系统 CJK 字体，如仍乱码请安装中文字体包。

**Q: PyMuPDF 安装失败？** macOS 可尝试 `brew install mupdf` 后重装，Linux 需 `apt install libmupdf-dev`。
