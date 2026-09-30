# course-notes-docx

把教材章节 / 课件（PPT、PDF）/ 已有笔记，压缩成**结构化课程笔记（考点提纲）**的 Word `.docx`，
并把 LaTeX 数学公式转成 **Word 原生公式对象（OMML）**。

这是给 AI Agent 用的 skill（Claude Skill / WorkBuddy Skill 格式），不是普通 Python 库 ——
核心是 `SKILL.md` 里那套经过实测校准的排版规范，`scripts/` 只是执行它的工具。

## 它解决什么问题

用 Word 手写带大量公式的笔记时，最烦的几件事：

- 公式用图片或截图 → 不可编辑、不可搜索、放大糊
- 行内公式和正文**字号不一致**，行高忽大忽小
- 纯公式行被 Word 当"数学段落"默认居中，跟正文左对齐打架
- 括号不随内容高度伸缩（`(a/b)` 里的括号还是小括号）
- 函数名 `det` `tr` `rank` 被渲染成斜体（数学排版里应该是正体）
- 分页把"Definition 2.4"的标题留在上一页、正文甩到下一页

这个 skill 把这些逐条修掉了，做法是**先量样本、再定规范**：仓库里的每个数值
（字号、行距、间距、缩进、颜色）都来自对 7 份手写样本的实测，而不是拍脑袋。

## 功能

- Markdown + LaTeX → `.docx`，公式转为**原生 OMML**（可在 Word 里直接编辑）
- 自动生成**目录**（Word 原生 TOC 域，章 + 节两层，点线引导 + 页码右贴边）
- 自动加**页脚页码**（居中 `X / Y`）
- 条目级"**整块同页**"：不会出现标题与正文被分页切开
- 项目符号按列表体量自动选档（重档 `●◆■▲` / 轻档 `·–—−`），并校验字形是否存在
- 配套**回读校验脚本**，把"排版没坏"变成可执行的硬指标
- 输出语言与素材一致（英文讲义 → 全英文笔记，不翻译、不中英混排）

## 环境要求

- **Windows + 本机安装的 Microsoft Word** —— 导出 PDF、刷新目录/页码域都靠 Word COM，
  这是"所见即所得"的前提（第三方在线预览对 OMML 的 `m:nor` 支持不完整）
- Python 3.10+
- 依赖：`lxml`、`python-docx`、`latex2mathml`、`pymupdf`、`pillow`

```bash
pip install lxml python-docx latex2mathml pymupdf pillow
```

## 安装

把本目录放到 Agent 的 skill 目录下即可：

```
# WorkBuddy / Claude Code 用户级 skill
~/.workbuddy-ai/skills/course-notes-docx/
```

## 用法

```bash
# 1) 生成 docx（输入是 Markdown + LaTeX 中间稿）
python scripts/build_note.py 输入.md 输出.docx

# 2) 回读校验（硬指标，必须全过）
python scripts/verify_docx.py 输出.docx
python scripts/selftest_formulas.py

# 3) 用 Word 刷新目录/页码域 → 回写 docx → 导出 PDF
powershell -NoProfile -ExecutionPolicy Bypass -File scripts/finalize_docx.ps1 \
  -Src "输出.docx" -Pdf "输出.pdf"

# 4) 补回 Word 保存时删掉的公式正体提示
python scripts/restore_upright.py 输出.docx
```

> 排版要调，改 `scripts/build_note.py` 顶部的 `CFG` —— 字号、行距、间距、列表符号、
> 目录样式、页码格式都集中在那里。

## 文件结构

```
SKILL.md                      # 权威规范：排版铁律、字号间距实测值、历史踩坑（先读这个）
references/
  template-spec.md            # 完整模板规范（页面/字号/层级/列表/目录页码的实测值）
  latex-omml.md               # 公式链路与已知限制
scripts/
  build_note.py               # Markdown + LaTeX → docx（顶部 CFG 集中调参）
  verify_docx.py              # 回读硬指标校验
  selftest_formulas.py        # 公式/定界符/正体/行宽估算自检
  finalize_docx.ps1           # Word 刷新目录页码域 + 导出 PDF
  restore_upright.py          # 补回 Word 删掉的 <w:i w:val="0"/>
  docx_to_pdf.ps1             # 纯导出 PDF
  render_pdf_preview.py       # PDF → PNG 目视核对 / dump 行坐标
  analyze_outline.py          # 排版参数分析器（规范只应来自实测）
  probe_bracket_stretch.py    # 括号伸缩度量测（连通域法）
  check_onedrive.py           # OneDrive 占位文件检查
```

## 校验硬指标

`verify_docx.py` 会报这些，前四项必须达标：

| 指标 | 要求 |
|---|---|
| 裸行内公式段落 | = 0 |
| 正文 `w:t` 含 NBSP | = 0 |
| 公式失败残留 | = 0 |
| `w:i=0`（公式内正文正体） | > 0 |
| 目录域 | 已刷新（非占位符） |
| 页脚 | `PAGE` / `NUMPAGES` 域各 1 |

## 已知限制

- **强依赖本机 Word**：PDF 导出与目录/页码刷新走 Word COM；没有 Word 就只能拿到
  "目录未刷新"的 docx
- 脚本里带若干**本机绝对路径**（Python venv、样本目录），换机器需要按需改
- 样本规范针对的是**曼彻斯特大学数学系**的讲义格式（A4、Cambria Math、
  章 15pt / 节 14pt / 条目 12pt / 正文 11pt）；换学校/换排版风格要重跑
  `analyze_outline.py` 重新校准
- 中文 `m:nor`（公式里的正文）在部分在线预览器里仍会被渲染成斜体 —— 这是渲染器的问题，
  交付 PDF 可绕开

## 许可

[MIT](LICENSE)
