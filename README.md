# Office & Filesystem MCP 工具包

## 这是什么？

一个 MCP（Model Context Protocol）服务，装到 CherryStudio 或 Claude CLI 后，你的 AI 助手就能直接操作电脑上的文件——读取、创建、编辑 Word/Excel/PPT，以及管理文件系统。

**本质上就是给 AI 装了一双能操作你电脑文件的手。**

---

## 能做什么？

装好之后，你可以直接用自然语言对 AI 说：

### 📄 Word 文档
| 你对 AI 说 | AI 做的事 |
|------------|-----------|
| "帮我读一下桌面的合同.docx，提取甲乙方信息和金额" | 读取 Word，解析内容 |
| "生成一份采购合同，甲方填XX公司，金额5000，保存到桌面" | 生成带格式的合同文档 |
| "把这份文档的第五条改成..." | 编辑替换指定内容 |

### 📊 Excel 表格
| 你对 AI 说 | AI 做的事 |
|------------|-----------|
| "分析桌面 销售数据.xlsx 的月度趋势" | 读取 Excel 数据 |
| "把合同里的报价单提取出来，生成一个 Excel" | Word → Excel 转换 |
| "汇总这三个 Excel 的总金额，生成汇总表" | 跨文件汇总 |

### 🎞 PPT 演示文稿
| 你对 AI 说 | AI 做的事 |
|------------|-----------|
| "根据这份 Word 报告生成一个 5 页的 PPT" | 报告 → 演示文稿 |
| "把这段数据做成表格放到 PPT 里" | 数据可视化 |
| "读取这个 PPT 的大纲" | 提取 PPT 结构 |

### 📁 文件管理
| 你对 AI 说 | AI 做的事 |
|------------|-----------|
| "列出桌面上所有的 .docx 文件" | 浏览目录 |
| "把下载文件夹里所有 PDF 移动到一个新目录" | 批量移动 |
| "读取这个 .py 文件，帮我改第 20 行" | 编辑代码文件 |
| "帮我创建一个项目文件夹结构" | 创建目录 |

---

## 你的 AI 现在多了 14 项能力

| 类别 | 工具 | 干什么用 |
|------|------|----------|
| **Word** | `read_docx` | 读取文档（段落、表格、样式） |
| | `write_docx` | 创建文档（标题、正文、表格、签署页） |
| **Excel** | `read_xlsx` | 读取表格（多 Sheet、合并单元格） |
| | `write_xlsx` | 创建表格（表头、冻结、样式） |
| **PPT** | `read_pptx` | 读取幻灯片（文本、表格、备注） |
| | `write_pptx` | 创建演示文稿（标题、要点、表格） |
| **文件** | `list_directory` | 浏览文件夹 |
| | `read_file` | 读取任意文本文件 |
| | `write_file` | 创建/覆盖文件 |
| | `edit_file` | 查找替换编辑 |
| | `file_info` | 查看文件信息 |
| | `create_directory` | 新建文件夹 |
| | `move_file` | 移动/重命名 |
| | `delete_file` | 删除文件/文件夹 |

---

## 怎么装？

### 前提条件

- Python 3.10 或更高版本（macOS 自带，Windows 需安装）
- 已装 CherryStudio 或 Claude CLI

### 3 步搞定

**第 1 步：放到任意目录**

```bash
# 把整个 office-tools-mcp 文件夹拷贝到电脑上
# 比如放到 /Users/你的用户名/projects/ 下面
```

**第 2 步：安装依赖**

```bash
cd office-tools-mcp              # 进入目录
python3 -m venv venv             # 创建虚拟环境
./venv/bin/pip install -r requirements.txt   # 安装依赖（约 30 秒）
```

> Windows 用户把 `./venv/bin/pip` 换成 `venv\Scripts\pip`

**第 3 步：配置到 AI 工具**

<details>
<summary>🔹 如果你用 CherryStudio</summary>

打开 CherryStudio → 左下角 **设置** → **MCP 服务器** → **添加**：

| 配置项 | 填什么 |
|--------|--------|
| 名称 | `office-tools` |
| 传输类型 | `stdio` |
| 命令 | 选 `venv/bin/python`（在你的项目目录里） |
| 参数 | `server.py`（在你的项目目录里） |

保存，打开开关，重启 CherryStudio。
</details>

<details>
<summary>🔹 如果你用 Claude CLI</summary>

编辑 `~/.claude/mcp.json`，写入：

```json
{
  "mcpServers": {
    "office-tools": {
      "type": "stdio",
      "command": "实际路径/venv/bin/python",
      "args": ["实际路径/server.py"]
    }
  }
}
```

重启 Claude CLI。
</details>

### 验证是否装好

在 AI 对话框里说：

> "列出我桌面的文件"

如果能返回桌面文件列表，就成功了。

---

## 项目结构

```
office-tools-mcp/
├── README.md               ← 你正在看的文件
├── DEPLOY.md               ← 详细部署文档（含 Windows 说明、常见问题）
├── server.py               ← MCP 服务主程序
├── requirements.txt        ← Python 依赖清单
├── start.sh                ← 快捷启动脚本
├── tools/                  ← 工具实现
│   ├── word_tools.py        # Word 读写
│   ├── excel_tools.py       # Excel 读写
│   ├── ppt_tools.py         # PPT 读写
│   └── filesystem_tools.py  # 文件系统操作
└── scripts/                ← 辅助脚本
    ├── gen_contract.py      # 合同自动生成
    └── github-mcp.sh       # GitHub MCP 启动
```

---

## 常见问题

**Q: AI 说"无法访问桌面"？**

不是技术问题，是 AI 太谨慎了。给完整绝对路径就行：
> ❌ "读桌面的合同"
> ✅ "读 /Users/xxx/Desktop/合同.docx"

**Q: 启动报 connection closed？**

检查 Python 版本（需 ≥ 3.10），确认 `pip install -r requirements.txt` 全部装完没报错。

**Q: 想自己加新功能？**

在 `tools/` 下新建 `.py`，在 `server.py` 里用 `@mcp.tool()` 注册。
