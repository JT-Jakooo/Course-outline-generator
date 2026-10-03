---
name: course-notes-docx
description: 根据教材章节、课件 PPT/PDF 或已有笔记，生成符合既定模板规范的 Word（.docx）课程学习笔记，并把 LaTeX 数学公式转为 Word 原生公式对象。This skill should be used when the user asks to produce, rewrite, extend, or reformat university course notes as a .docx file, or when lecture slides, textbook chapters, or existing notes must be condensed into a structured Word document that contains mathematics.
agent_created: true
---

# 课程学习笔记生成（Word / .docx）

## 目的

把教材章节、课件或已有笔记**提炼成一份考点提纲**（Word .docx），保持既定的排版、
层级与编号规范，并让数学公式成为 Word 原生可编辑公式对象（OMML）。

**定位 = 考点提炼，不是"浓缩教材"，更不是复习资料。**
读者拿它当"目录 + 公式索引"用；**要看讲解就去看讲义**。

同一套模板已用于多门科目，可直接复用；换科目只换素材与输出路径，模板规范不用重做。

## 何时使用

- 用户要求"根据这份教材/课件写笔记"、"把这一章整理成 Word 笔记"
- 用户要求按既有模板生成、扩写或重排课程笔记
- 需要输出含大量数学公式的 .docx

## 取舍规则（用户 2026-09-30 明确，优先级高于"完整性"）

> 用户原话：**「你在这版根据 skills 出的提纲里面加了太多根本不需要的冗余的内容，
> 一大段长文字，为什么我不去直接看讲义呢？提纲的重点就是提炼考点。」**

**保留**（手打提纲实测密度：Probability I Chapter 1 = 3 个条目；Chapter 2 整章 = 15 段）
- **编号条目**：`Definition` / `Theorem` / `Lemma` / `Proposition` / `Corollary` /
  `Remark` / `Note` / `Axiom` —— **编号照抄，缺号照缺**
- 条目的**条件与结论**；**关键公式**（尤其含 ∑ ∏ ∫ 的求和/积分式）
- **极短**的连接语（`Let Ω be discrete.`）与作者本人的 `*` 批注
- 章 / 节 / 次级小标题

**删除**
- **例子**（`Example`）—— 除非用户点名要
- **练习**（`Exercise` / `Problem`）
- **动机、直觉、背景**（`An informal view`、`Intuitively, …`、`You may know …`）
- **证明的推导过程**（结论留，过程删；素材只给结论就不写证明）
- **大段解释性段落**（`The advantage of …`、`Because of this, …`、
  `Mathematicians often …`、`We shall denote …` 这类过渡/铺陈）
- **参考性表格**（如 `Table 1: Different sample spaces`）—— 除非用户点名要

> 判断标准一句话：**这一行是不是"要背 / 要会用"的东西？** 不是就删。
> 宁可删多、被用户要求补回，也不要删少、被要求重做。

## 内容边界（硬约束，不可协商）

> 用户原话：**「不要胡编乱造，要实事求是」「一切要严格按照教材的内容来」。**
> 笔记 = **考点提纲**，只做「**筛选 + 压缩 + 重排**」，**不做任何增补**。

1. **绝不编造内容**。定义、定理、命题、引理、推论、例子、编号、公式、符号，
   必须**逐条源自用户提供的素材**。素材里没有的，一律不写。
   素材缺失或读不出来（公式是图片、扫描件模糊、只有目录没有正文）→
   **明确告知用户缺什么**，不要用通用数学知识填空。
2. **只许删减，不许增补**。解释性文字**直接删除**（"压短"不够，要删干净）；
   **不允许**新增定义、新增定理、新增例子、新增证明步骤、新增"直观理解"
   或"背景介绍"。自检标准：**写下的每一句都要能在素材里指到出处**。
3. **不补全推导**。证明里省略的步骤**不要自行补**（哪怕数学上显然）；
   素材只给结论就不写证明。确有必要补，**先问用户**。
4. **编号严格对齐教材**。`Definition 1.1`、`Theorem 2.7` 这类编号照抄素材；
   教材缺号（例如跳过 2.2）**照缺不补**；合并编号（`Definition 2.4. ~ 2.6.`）照原样。
5. **术语与表述照搬素材**。同一概念在素材中说法不一致时**保留原样**，不擅自统一；
   不要用更"规范"或更"通俗"的措辞替换原文。
6. **不做复习型加工**。不添加"易错点 / 自测题 / 章末小结 / 考点提示 / 记忆口诀"；
   也不主动添加目录页、索引、符号表、公式速查表，除非用户明确要求。
   （**"提炼考点" = 删掉非考点，不是标注考点** —— 不要加"考点""重点""易错"这类标签。）
7. **交叉引用与段落标记只写素材里有的**。`<Theorem 10.17>` 这类引用，
   以及 `Etc.` / `Recall` / `Note!!!` / `*注释` 等段落级标记，
   素材里没有就不造。（样本里的 `*` 开头句子是**作者本人**的批注，不是模板要求。）
8. **不确定就问**。任何"要不要补一句 / 加个例子 / 合并两节 / 略去一段"的判断，
   先问用户，不要自行决定。
9. **语言与素材一致（全英文）**。笔记正文、条目名（`Definition 2.4`）、术语与连接语
   一律使用**素材的原语言**；本课程素材为英文讲义 → **输出全英文**，
   **不翻译成中文**，不做中英混排（数学符号与公式除外）。
   **注意**：写源 md 时最容易漏的是**正文第一行的 Scope/说明行** —— 它不在标题、不在条目里，
   手写时极易混进中文（2026-10-03 PDE 提纲就是这里混入了"考点"）。
   写完必须跑 `verify_docx.py` 看"正文含汉字"是否为 0（见铁律三·B）。

> **与下文排版规则的关系**：后面所有排版规则（字号、对齐、公式修复……）
> **只改变呈现形式，不允许改变内容**。两者冲突时，**以内容边界为准**。
> 完成后用 `present_files` 打开预览。

## 三条决定观感的排版铁律

### 铁律一：含公式的行，文字与公式必须同字号，且要能自然折行

样本作者的做法是**把整行文字都写进公式对象，再用引号把文字变回普通文本**
（样本 `<m:nor>` 出现 1102 次）。

但**这个做法有个副作用**：整行变成一个数学段落（`m:oMathPara`）后，
Word 对数学段落的分行会强加缩进 —— 实测正文行会**折行过早**（第 1 行右侧空 186pt）
且**续行缩进 72pt**（看起来像居中）。

因此本脚本对**含正文文字的行**改用等价的"混排段落"：
文字段走普通 `w:r`、数学段走行内 `m:oMath`，段落设 `w:jc=left`。

| 承载方式 | 渲染字号 | 折行表现 |
|---|---|---|
| 整行并入 `m:oMathPara` | 正文 11.04pt / 下标 8.04pt | ✗ 折行过早 + 续行缩进 72pt |
| **混排段落（当前）** | **正文 11.04pt / 下标 8.04pt（完全一致）** | ✓ 占满整行 + 续行齐左 |

两者字号实测完全相同（pymupdf span size），差别只在折行。
开关：`CFG['prose_carrier']`（`'mixed'` 默认 / `'inline_formula'` 回到样本写法）。

**校验时看 `verify_docx.py` 的"正文行(文字+行内公式)"，应 > 0。**

### 铁律二：纯公式段落必须用 `m:oMathPara`，并显式写 `m:jc`

**裸行内 `m:oMath` 且段落里没有文字是坑**：Word 会把这种段落当数学段落处理，
按默认对齐（居中）排，`w:jc` 被忽略 —— 于是同一份文档里有的行靠左、
有的行莫名居中、缩进乱七八糟。实测已复现并确认。

样本自身的对齐约定（按段落分类统计 `oMathPara` 的 `m:jc`）：

| 段落类型 | left | center | 说明 |
|---|---|---|---|
| 含正文文字的段落 | 138 | 48 | → **左对齐** |
| 长正文段落 | 92 | 22 | → **左对齐** |
| 纯公式段落 | 20 | 83 | → **居中** |

因此脚本规则：

- 含正文文字的段落 → `left`（读起来像正常行文，能自然折行）
- 独占一行的纯公式（`$$...$$`）→ `center`
- 多行分步推导（`array` / `aligned`，按 `=` 对齐）→ `left`

**校验时看 `verify_docx.py` 的"裸行内公式段落"，必须为 0。**

### 铁律三：正文里的空格必须是普通空格，不能是 NBSP

`latex2mathml` 会把 `\text{}` 里的普通空格转成 U+00A0（不换行空格）。
**样本 1101 个正文 run 里 NBSP 数量为 0** —— 作者用的全是普通空格。
NBSP 会让 Word 无法在词间断行，整行只能挤在数学运算符处折断，排版塌成一团。

脚本自动把 NBSP 换回普通空格；`\quad` / `\qquad` 这类**有意**的间距另用标记保留为 NBSP。
**校验时看 `verify_docx.py` 的"正文 run 含 NBSP"，应为 0（仅 `\quad` 处允许）。**

### 铁律三·B：生成的提纲/作业文档严禁出现汉字（2026-10-03 PDE 事故）

跨科目约定：**生成的提纲/作业文档一律全英文**，正文（含 Scope 行、标题、目录条目）
严禁混入任何中文。2026-10-03 PDE 提纲的 Scope 行混入"考点"二字，PDF 目录页直接可见，
被用户指出。

`verify_docx.py` 已加硬校验 **"正文含汉字"**（扫 prose `w:t`，含 U+4E00–9FFF 与
扩展 A 区）——**必须为 0，否则退出码 1**。确要出中文文档时加 `--allow-cjk` 豁免。
中文只允许出现在：与用户的对话回复、memory 记录；绝不允许进入生成文档的源 md。

### 铁律四：括号必须随内容高度伸缩

LaTeX 里括号会随内容自动放大，OMML 里只有 **`m:d`（delimiter）对象**才会被 Word
纵向拉伸；普通字符 run 渲染出来是**固定大小**。

有两类写法会让 `latex2mathml` 把定界符退化掉，必须手工修复：

| 类别 | 现象 | 原因 |
|---|---|---|
| (a) 尖括号 / 取整括号 | `⟨u,v⟩`、`⌊x⌋`、`⌈x⌉` 只有默认大小 | `\langle` `\lfloor` `\lceil` 被归成 `<mi>`（标识符）而非 `<mo fence>`，XSL 把整段合并成**一个** `m:t` |
| (b) **函数名紧邻定界符** | `\det\begin{pmatrix}…\end{pmatrix}` 的括号**完全不伸缩**（实测比值 0.56，正常应 ≥0.9） | `latex2mathml` 把函数名和后面的 `(` 合并进同一个 `<m:t>`，例如 `"det("`，定界符不是独立元素 |

脚本自动修复（`_stretch_delims`，两步）：

1. `_split_delim_runs`：把混在普通 run 里的定界符字符拆成独立 run。
   拆分集 = `( ) [ ] { } ⟨ ⟩ ⌊ ⌋ ⌈ ⌉ ⟮ ⟯`。
   `m:nor`（`\text{}` 里的正文）不动 —— 所以 `\text{The matrix (1,0)}` 不会被误包。
2. `_wrap_stretch_delims`：同层里**取距离最近的一对**配对，连同中间内容包进
   `m:d`（`begChr` / `endChr`）。取最近对是为了嵌套时先裹最内层（`⟨⟨a⟩⟩`、`f(g(x))`）。

**`|` 与 `‖` 刻意不加入拆分集**：它们本来就由 `latex2mathml` 正常产出 `m:d`；
而 `|x|+|y|` 这类式子里存在中间的 `|` run，加入后可能被误配对。
（`\mid` 产出的是 U+2223 `∣`，与 U+007C `|` 不是同一字符，互不影响。）

非成对写法**不得**被误配：`(a,b]`、`[a,b)` 的 `(`/`[` 找不到对应右括号，保持原样。

**校验**：`selftest_formulas.py` 的"定界符伸缩自检"39 例（含 3 例非成对）必须全过；
渲染级用 `scripts/probe_bracket_stretch.py` 量测（括号墨迹高度 / 行内墨迹高度，应 ≥ 0.85）。
注意：`--brackets` 只报字体行框，**不能**判伸缩；带 ±3pt 水平 padding 的裁剪窗会把
右侧单元格数字计入墨迹宽度，制造"周期性接缝"的**假阳性** —— 必须用连通域隔离。

### 铁律五：函数名必须正体（`m:sty val="p"`）

样本里 `det` / `dim` 的 run 都带 `<m:sty m:val="p"/>`（正体）。
`latex2mathml` 对**多数**函数名输出 `<mi>`，XSL 会带上 `sty=p`；
但下面这几个输出的是 `<mo>`，会被渲染成**斜体**：

```
det  lim  limsup  liminf  sup  inf  min  max  gcd  Pr
```

`preprocess_tex` 把它们（以及 `\operatorname{…}` / `\operatorname*{…}`）统一改写成
`\mathrm{…}`，实测即可拿到 `sty=p`。正则带 `(?![A-Za-z])`，不会误伤
`\supset` / `\infty` / `\bigcup` 等。
`\lim_{…}` 在本链路本来就是 `m:sSub`（不是 `m:limLow`），换写法不改变下标位置。

**校验**：`selftest_formulas.py` 的"函数名正体自检"（10 个已修正 / 14 个原本正体未破坏 /
9 个易误伤写法未改写 / 3 个 `\operatorname` 已改写）。

## 公式大小规则：什么时候用小公式，什么时候用大公式

> 用户 2026-09-30 原话：**「公式什么时候要用小公式，什么时候要用大公式。
> 就比如原稿中的 1.5 (b) 中的 summation 用的就是小公式，不用分行就可以写完，
> 这个看起来就很整洁，但是可能别的地方就要用大公式。」**

**判据不是"有没有 ∑"，而是"这一行放不放得下"。**

| | 小公式（行内 `$...$`） | 大公式（独占一行 `$$...$$`） |
|---|---|---|
| 触发 | 公式**跟文字同行且整行放得下** | 作者要它独立成行；或行内放不下被迫拆出 |
| 版式 | 跟文字同段，左对齐、随文字自然折行 | `m:oMathPara`，**居中**（多行推导左对齐） |
| 典型 | `(b) Sum up to one, i.e. Σ_{ω∈Ω} ℙ({ω}) = 1`（整行 176pt ≪ 列宽 415pt）<br>`ℙ({ω}) ≥ 0`、`A ⊆ Ω`、`n ∈ ℕ` | 编号公式 `ℙ(E) = Σ_{ω∈E} ℙ({ω})`<br>矩阵、分段函数、多行按 `=` 对齐的推导 |

**自动拆分（安全网）**：行内出现**带上下限的大运算符**
（`\sum` `\prod` `\int` `\lim` `\bigcup` …）**且整行估算宽度 > 列宽 × 0.97** 时，
才把该行拆成三段：

```
正文段（左对齐） → 独立公式（居中） → 正文段（左对齐）
```

原因：**Word 遇到 n 元运算符会强制把折行的续行缩进到运算符位置** ——
实测**行内 `\sum` 也一样**（续行从 333pt 起，正常应 90pt）。
但**只有真的折行才会出问题**：样本的 `(b) … Σ_{ω∈Ω} ℙ({ω}) = 1` 短到不折行，
作者就保持行内小公式，看着很干净。所以脚本按宽度判定，不搞"见 ∑ 必拆"。

**宽度怎么估**（`_line_width_pt`，用本机 Cambria Math 真实字形度量，偏差 < 5%）：
- 文字段：PIL `ImageFont.truetype('C:/Windows/Fonts/cambria.ttc', size, index=1)` 逐段 `getlength`
- 公式段：LaTeX → 近似 Unicode 后度量，二元运算符两侧各补 0.24em、大运算符放大 1.30×
- 上下标宽度并入基数（不缩），故估算**偏宽 = 保守**：宁可误拆，不可误判"放得下"

开关：`CFG['split_big_ops']`（默认 `True`）、`CFG['split_ratio']`（默认 `0.97`）。
**注意**：写 md 时用 `$...$` 表达"作者认为该用小公式"，用 `$$...$$` 表达"该用大公式"，
脚本会尊重；只有放不下时才违背 `$...$` 的意愿去拆。

## 字号 / 颜色 / 间距（7 份样本交叉实测）

| 用途 | 字号 | 字重 | 颜色 |
|---|---|---|---|
| **文档大标题** `<课程全名> Outline` | **16pt（sz=32）** | **加粗** | 黑（左对齐、Cambria Math） |
| **作者行** `By JasonTan` | **14pt（sz=28）** | 常规 | 黑（字体 **Algerian**、**右对齐**） |
| 章 `Chapter N` | 15pt（sz=30） | 加粗 | **#0F4761 深蓝** |
| 节 `N.M` | 14pt（sz=28） | 加粗 | **#0F4761 深蓝** |
| 条目标题 `Definition 1.6. 名称` | 12pt（sz=24） | **编号加粗、名称常规** | 黑 |
| 无编号次级小标题（`####`） | 12pt（sz=24） | 加粗 | **黑** |
| 三级编号标题 `N.M.K`（`#####`） | 16pt（sz=32） | 加粗 | **#0F4761 深蓝**（仅 ITVC 用 heading 3，6 处） |
| 正文 / 公式 | 11pt（sz=22） | 常规 | 黑 |
| 证明正文 | 11pt（可调 10） | 常规 | 黑 |

| 间距 | 值 |
|---|---|
| `docDefaults` 段后 | 8pt（`after=160`） |
| 正文 / 列表 行距 | `278/auto` ≈1.158（**列表与正文同**） |
| **文档大标题 段前** | **6.25pt**（v16 补偿值，见下文 v16 说明） |
| **作者行 段前** | **17.6pt**（v16 补偿值 = 8 + 9.6，见下文 v16 说明） |
| 章 段前 / 段后 | 24pt / 4pt |
| 节 段前 / 段后 | 8pt / 4pt |
| **条目标题 段前 / 段后** | **36pt / 8pt**（视觉间距 ≈ 53pt） |
| **次级小标题 段前 / 段后** | **36pt / 4pt**（视觉间距 ≈ 51pt） |
| 一级列表 | 缩进 `440 twips ≈ 0.306"` + 悬挂 440（符号 x0=90、文字 x0=112） |
| 二级列表 | 缩进 `880 twips ≈ 0.611"` + 悬挂 440（符号 x0=112、文字 x0=134）、**字号 10pt**、符号 `-` |
| **列表块首项 段前** | **6pt**（带 `(a)/(1)/(i)` 标号的分点块 **10pt**） |
| **列表项 段后** | **4pt**（带标号的分点项 **12pt**） |
| **列表块末项 段后** | **18pt**（后接正文时；后接标题类时降到 4pt） |
| **章标题后的第一个块 段前** | **10pt**（不是 36pt！章标题已自带 24pt 段前） |

> **列表块要有上下边界（用户 2026-09-30：「小内容和分点的内容中间没有隔开来，
> 就体现不出分点的优势」）**
> 手打提纲 Prob I Ch1 逐行实测的间距是四档，不是一律 8pt：

> | 位置 | 提纲实测 | 本脚本的实现 |
> |---|---|---|
> | 条目 / 小标题 → 列表块首项 | 25.1pt | `before_list=6` |
> | 列表项之间（一级↔一级 / 一级↔二级） | 17.5~18.6pt | `after_list_item=4` |
> | `(a)` → `(b)` 并列分点之间 | **26.4pt** | `after_label_item=12` |
> | 正文 → 分点块首项 | 26.6pt | `before_label_item=10` |
> | 列表块 → 下一个块 | 51.7pt | 由下一块的 `before`(36) 提供 |

> 关键：**分点块自成两档** —— 并列分点（`(a)(b)`）的块上边界 26.6、项间 26.4，
> 都比普通列表（25.1 / 18.0）大一档，这样并列项彼此显得独立。
> 列表块末项只在**后接普通正文**时才加 18pt 下边界
> （后接条目标题 / 小标题时，它们自带 36pt 段前，块边界已经成立）；
> 这一"看下一行是什么再决定"的判定靠 `_list_flags` 预扫实现。

> **章标题后的第一个块不能用 36pt 段前** —— 章标题自身已带 24pt 段前 + 4pt 段后，
> 提纲实测"章标题 → 第一条目"只有 **29.8pt**，而"正文 → 条目"是 **51.7pt**。
> 一律套 36pt 会让每章开头多空 26pt、整页发飘。见 `_after_h1`。

> **块与块之间要留白（用户 2026-09-30：「看着太乱了，一砣字」）**
> 样本在**次级小标题**与**条目标题**之前**各插一个 12pt 空段落**做间隔：
> 实测视觉间距 **≈ 52pt**，而普通相邻两行只有 **≈ 17pt** —— 差了 3 倍，
> 这就是样本"有呼吸感"的来源。本脚本用等价但更干净的 `space_before` 实现
> （`CFG['before_sub']` / `CFG['before_entry']` = 36pt），不写空段落污染结构。
> **只按样式里的 12pt 是不对的** —— 那样 `Discrete Probability` 会紧贴上一行，
> 和普通列表项分不出来，整块糊成一坨。

> **两级列表**：样本是**一级 `*`（11pt）/ 二级 `-`（10pt）**。
> python-docx 默认模板每个 `abstractNum` 只有 `ilvl=0`，且 `multiLevelType=singleLevel`
> —— 必须三处都改，否则 Word 静默忽略二级定义：
> ① `multiLevelType` → `hybridMultilevel`；② 注入 `ilvl=1`（符号 `-`、ind 880/440）；
> ③ `ilvl=0` 的 `w:ind` 与 **`w:tabs` 的 num tab** 都改成 440
> （**文字位置其实由制表位决定**：默认 360 → 文字只到 108pt，样本是 112pt）。
> 写法：`- 一级项`；`  - 二级项`（缩进 2 空格）。见 `scripts/build_note.py::_setup_list_numbering`。

> **项目符号不必只用实心点（用户 2026-09-30：「bullet point 的样式其实在 word 里
> 有很多，不一定只用点，可以随机在库里挑」）**
> 7 份样本的一级符号实测有**三种**：
> `●`（Wingdings `f06c`：ITVC / LA / Stats I / ODEs）、`◆`（Wingdings `f0b2`：MFA / RA）、
> `-`（Cambria Math：Prob I / ODEs）。二级一律 `-`。
> 本脚本 `CFG['bullet_char']='auto'` 时**先按列表体量定档，再按源文件名确定性挑**
> —— 同一份 md 每次结果一致（可复现），不同科目自然挑到不同符号。
> 想固定就写具体符号，如 `bullet_char='◆'`。
> **换符号必须同时改字体**：默认模板是 `\uf0b7` + `Symbol`，
> 只改 `w:lvlText` 不改 `w:rFonts`，`●` 在 Symbol/Wingdings 下会渲染成别的字形。
>
> **符号要撑得住内容量（用户：「有些时候内容太多符号太小」）**
> 池子按"墨迹高度 / 正文 x-height"分档（PIL 实测，`_bullet_ink`）：
>
> | 档 | 符号 | /x高 | 填充率 |
> |---|---|---|---|
> | **重**（体量 ≥6 用） | `●` 1.51、`◆`/`■`/`▲` 1.46 | >1.4 | 0.5~1.0 |
> | 中 | `▸` 0.74、`*` 0.80、`•` 0.63、`◦` 0.47、`▪` 0.45 | | |
> | **轻**（短列表用） | `·` 0.27、`–`/`—`/`−` **0.14** | <0.3 | |
>
> 判据：`_list_volume(lines)` = **列表项数 + 平均项长/40**，≥ `bullet_volume_heavy`(6)
> 就从 `bullet_pool_heavy`（只放实心）挑；否则用 `bullet_pool_light`。
> 挑完还会**回报一次**并自动升级（空心/过小的会被换掉）：
> ```
> 项目符号：●  墨迹 1.51× x-height、填充率 0.79 | 列表体量 11.8（需"重"档）-> 合适
> ```
> ⚠ **重档只放实心**：`◇`(0.24) / `□`(0.46) / `○`(0.32) 虽然高大但填充率低，
> 满页文字时压不住 —— 空心符号归到中/轻档。
> ⚠ **已实测剔除的豆腐块**（Cambria Math 无字形，渲染成同一个 .notdef 方框）：
> `⬤`(U+2B24)、`‣`(U+2023)、`⁃`(U+2043)。
> **池子想再加符号，先用 PIL 量一次墨迹确认有字形，别盲加。**

> **三条曾经的错误结论（已修正）**
> ① **条目 13pt 是错的** —— 那是量到了 Linear Algebra 里唯一的 sz=26 异常段
>    （`Definition 1.1 Fields`）。全量：sz=24 共 200 段、sz=26 仅 1 段；6 份样本
>    PDF 实测条目一律 **12.0pt**。
> ② **列表 1.5 倍行距是错的** —— 6 份样本列表段均以 `278/auto` 为主
>    （LA 93:55、MFA 90:19、Prob I 74:9、RA 58:3、Stats 53:18）；渲染侧行基线
>    间距众数 12.8–13.0pt（= 11pt × 1.158），不存在 1.5× 的簇。
> ③ **列表左缩进 0.5" 是错的** —— 那是 python-docx 内置 `List Paragraph` 样式的
>    默认值（`w:ind left=720`）。**样本根本没用该样式**：Prob I / LA 的
>    `ListParagraph` 段落数都是 **0**；样本用的是**段落级 `numPr` + `w:ind`**，
>    实测 `left=440`（5 份样本一致，Stats I 的 numbering 定义里也是 440）。
>    渲染实测：符号在左边距 x0=90、文字 x0=112（22pt）。
> ④ **"次级小标题 = 12pt 深蓝" 是把两回事混成了一回事**：
>    · **无编号**次级小标题（Prob I 的 `Discrete Probability`、LA 的同类）→ **12pt 黑**。
>      6 份样本里 **sz=24 且加粗**的 run 颜色只有 `inherit`（= docDefaults 黑）/`000000`，
>      另有个别 `EE0000`（红）/`C6C6C6`（灰，作者手标），**没有一个 `0F4761`**。
>    · **三级编号**标题 `N.M.K`（只有 ITVC 有，6 处：`2.2.1 Gradient` …）→
>      **16pt（sz=32）深蓝**，走的是 heading 3 样式。**不是 12pt。**
>
> **判定加粗别用 PDF 的 flags**：Cambria Math 没有 Bold 字面，`<w:b/>` 由 Word
> 走**合成加粗**（描边）。导出 PDF 里不会嵌入粗体字体，pymupdf 按字体名或
> `flags & 16` 检测会一律报"非粗体"（实测 LA 56 页 PDF 共 0 个 bold span）。
> **这是正常现象**，目视是加粗的即为正确。

## 工作流

### 步骤 1：确认素材与输出路径

确认源素材路径（教材 PDF / 课件 PPT / 已有笔记）、输出 docx 路径、覆盖范围。

**输出位置（用户 2026-09-30 定）**：笔记**放回该科目自己的 OneDrive 文件夹**，
和讲义并排 —— 不集中丢在工作区。命名沿用 Year 1 样本的写法 `<科目名> Outline.docx` / `.pdf`：

```
C:/Users/Admin/OneDrive - The University of Manchester/University/<科目名>/
    ├── Probability 2 lecture notes-4.pdf      ← 讲义（只读）
    └── Probability & Statistics 2 Outline.docx / .pdf   ← 我们的产出（新建）
```

⚠ **只新建，绝不修改 / 覆盖 OneDrive 里的任何既有文件**（讲义、past paper、手打 Outline
一律只读）。该目录**可以新建文件**（实测已确认），但**已有的文件不要动**。
中间稿 `.md`、对比图、探针脚本仍留在工作区 `F:/AI/Workbuddy AI/Study-Year 2/`，
只有最终 `.docx` / `.pdf` 放进科目文件夹。

素材常位于 OneDrive 同步目录。云端占位文件首次读取会触发自动下载（实测 0.2–0.4 秒）：

```bash
PY=C:/Users/Admin/.workbuddy-ai/binaries/python/versions/3.13.12/python.exe
"$PY" scripts/check_onedrive.py "<目录路径>"
```

### 步骤 2：读取素材

- PDF / PPTX：提取文字与公式；PDF 公式常为图片或 Unicode 线性文本，需转 LaTeX。
- 已有 .docx：零依赖读取（zipfile + ElementTree，遍历 `w:t` 与 `m:t`）。
- 遇到无法识别的图片公式，**明确告知用户**，不要猜。

### 步骤 3：撰写 Markdown 中间稿

**动手前先按上面的「取舍规则」过一遍素材**：把 `Example` / `Exercise` / 动机 /
直觉 / 解释段 / 参考表格划掉，只留**编号条目 + 条件结论 + 关键公式**。
写完自检一遍：**这份东西能替读者省下翻讲义的时间吗？** 不能就是删得不够。

> **文档大标题不写在 md 里**（v16）。它由**输出文件名**自动生成：`doc_title=None` →
> 取 `dst` 的文件名主干。所以**输出文件名必须写成 `<科目名> Outline.docx`**，
> 否则大标题会跟着错（样本就是"大标题 == 文件名"）。
> 要显式指定就在 CFG 里设 `doc_title='…'`；作者行默认 `By JasonTan`，`doc_author=''` 可关掉。

| 写法 | 结果 |
|---|---|
| `# Chapter N 章名` | Heading 1（15pt 加粗） |
| `## N.M 小节名` | Heading 2（14pt 加粗，深蓝） |
| `##### N.M.K 三级标题` | Heading 3（16pt 加粗，深蓝） |
| `#### 次级小标题`（无编号） | 12pt 加粗正文（黑） |
| `### Definition 1.1 名称` | 条目标题（12pt；编号加粗、名称常规） |
| 普通段落 | 正文 |
| `- 列表项` | 一级列表（段落级 `numPr`、`ilvl=0`；符号 x0=90、文字 x0=112、11pt） |
| `  - 子项`（缩进 2 空格） | **二级列表**（`ilvl=1`；符号 `-`、文字 x0=134、**10pt**） |
| `- (a) …` / `- (1) …` | 以 `(a)` `(1)` `(i)` 开头的项视为**自带标号** → **不加符号、也不缩进**（样本里 `(a)/(b)` 齐左于 x0=90） |
| `**粗体**` | 加粗（公式行内也支持） |
| `~~删除线~~` | 删除线（内部可含 `$...$`，公式也会被划掉） |
| `*注释`（星号紧跟文字） | 正文段落，星号保留 |
| `$...$`（行内，句子里有文字） | **混排段落**：文字 `w:r` + 行内公式，左对齐、自然折行 |
| 独占一行的 `$$...$$` | 纯公式行（`m:oMathPara`），居中 |
| `$$` 多行块 | 每行各成一条独立公式（可用来手动分行） |
| `\| a \| b \|` + `\|---\|---\|` | Word 表格（固定列宽、booktabs 框线；单元格内支持 `$...$`，**整格只有一个公式时自动用 `oMathPara` 承载**） |

**分行规则：一个逻辑单元 = 一个段落。** 条目标题独占一段；每条陈述句、每个公式、
每个列表项、证明的每个推导步骤各占一段；长句不拆行。

**写公式时必须注意**（细节见 `references/latex-omml.md`）：

- 换行符写成**两个反斜杠** `\\`；用文件写入工具写 md，不要用 `echo`/heredoc。
- 多行推导用 `\begin{array}{rl}…\end{array}`，**不要用 `align` / `aligned`**。
- `\text{}` 内**不要**给 `% _ # $` 加反斜杠转义；`& < > { }` 脚本会自动处理。
- 括号**不用**写 `\left…\right`：`(` `[` `{` `|` `⟨` `⌊` `⌈` 脚本都会包成可伸缩的 `m:d`。
- **不要写 `\big` / `\Big` / `\bigg`**：本链路里 `minsize`/`maxsize` 被 `MML2OMML.XSL` 忽略，
  它们**完全无效**；而 `\big\langle` 更会被 `latex2mathml` 输出成**字面文本** `\langle`（实测）。
  脚本的 `preprocess_tex` 会把这些命令直接删掉（`\bigcup` `\bigcap` `\bigoplus` 等不受影响）。

### 步骤 4：生成 docx

```bash
PY=C:/Users/Admin/.workbuddy-ai/binaries/python/envs/default/Scripts/python.exe
"$PY" scripts/build_note.py 输入.md 输出.docx [left|center|right]
```

第 3 个参数**只覆盖"纯公式段落"的对齐**，默认 `center`。
**调整排版**：改 `build_note.py` 顶部 `CFG` —— 字号、行距、间距、三种对齐、`split_big_ops`、
`split_ratio`、列表四档间距（`before_list` / `after_list_item` / `before_label_item` /
`after_label_item` / `after_list_block`）、两级列表缩进、项目符号
（`bullet_char` / `bullet_pool`）、目录与页码（`toc*` / `page_number*`，见 v15）。

**目录与页码默认开启**（`CFG['toc']` / `CFG['page_numbers']` 均为 `True`）：
章 + 节两层、页脚居中 `X / Y`。域要在 Word 里刷一次才有真实页码 —— 见步骤 6。

依赖 `lxml`、`python-docx`、`latex2mathml`。venv 不存在时先建：

```bash
"C:/Users/Admin/.workbuddy-ai/binaries/python/versions/3.13.12/python.exe" -m venv \
  C:/Users/Admin/.workbuddy-ai/binaries/python/envs/default
```

### 步骤 5：回读校验（必做）

```bash
"$PY" scripts/verify_docx.py 输出.docx
"$PY" scripts/selftest_formulas.py
"$PY" scripts/analyze_outline.py 输出.docx          # 与样本逐项比对排版参数
```

必须确认：

- **`裸行内公式段落` = 0**（否则会被 Word 默认居中）
- **`正文行(文字+行内公式)` > 0**（这是含公式正文的正确状态）
- **`正文 w:t 含 NBSP` = 0**（仅 `\quad` 处允许；否则无法断行）
- **`w:i=0(强制正体)` > 0**（v15：Word 刷新后需用 `restore_upright.py` 补回）
- **`TOC 域` 显示"已含真实条目"**（若显示"仍是占位符"，说明还没跑 `finalize_docx.ps1`）
- **`页脚` PAGE 域 = 1、NUMPAGES 域 = 1**
- `公式失败残留` = 0
- `selftest_formulas.py` 的**定界符伸缩自检 39 例全过** + **函数名正体自检全过**
- Heading 字号 30 / 28（15 / 14pt）、加粗、颜色 `0F4761`（与样本一致）
- 条目标题 run 字号 24（12pt）、`Definition X.Y` 加粗而名称常规

### 步骤 6：刷新目录/页码并导出 PDF（交付默认带上）

在线预览（腾讯文档 / WPS 在线等）对 OMML 的 `m:nor` 支持不完整，
会把公式里的正文也渲染成斜体 —— **那是渲染器的问题，不是文档的问题**。
导出 PDF 可绕开该差异，用户看到的就是 Word 的真实排版。

**笔记带目录/页码时**（v15 起为默认），用 `finalize_docx.ps1` 一次完成
"刷新域 → 回写 docx → 导出 PDF"，再用 `restore_upright.py` 补回 Word 删掉的
`w:i=0`：

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File scripts/finalize_docx.ps1 `
  -Src "输出.docx" -OutDocx "暂存.docx" -Pdf "输出.pdf"
```

```bash
"$PY" scripts/restore_upright.py 暂存.docx 输出.docx
```

`-OutDocx` 可省略（原地保存，仅对本会话新建的文件可靠）；`-NoSave` 则只导出 PDF 不改 docx。

**笔记不带目录/页码时**，用原来的导出脚本即可：

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File scripts/docx_to_pdf.ps1 `
  -Src "输出.docx" -Dst "输出.pdf"
```

两者都用本机 Word 引擎渲染，需要安装 Microsoft Word。

**目视核对**（排版问题只看 XML 判断不出来，必须看渲染结果）：

```bash
"$PY" scripts/render_pdf_preview.py 输出.pdf ./pages 300    # 逐页转 PNG，肉眼检查
"$PY" scripts/render_pdf_preview.py --lines 输出.pdf        # dump 每行 bbox + flags
"$PY" scripts/render_pdf_preview.py --brackets 输出.pdf     # 每个定界符的字形高度
"$PY" scripts/probe_bracket_stretch.py 输出.pdf             # 括号伸缩度（连通域法，>=0.85 合格）
```

`--lines` 模式可用坐标客观判断对齐（居中 / 左对齐 / 缩进），
用 `flags` 的 bit 1 判断是否斜体。依赖 `pymupdf`。

> **注意**：`--brackets` 报的是 pymupdf 的**字体行框**高度，拉伸后的括号仍报同一数值
> （实测 `(a/b)` 与 `(x+y)` 都是 11.1pt），**不能**据此判断伸缩。
> 判伸缩要靠高 dpi 裁剪渲染后目视，或用**连通域**量测
> （括号墨迹高度 / 行内墨迹高度）。两种常见假阳性：
> ① 裁剪窗横向留 ±3pt 会把右侧单元格数字计入墨迹宽度；
> ② 纵向留 ±26pt 会把相邻段落余墨计入。
> 两者都会制造"周期性接缝"的**假象** —— 必须用连通域标记隔离本体。

### 步骤 7：交付

用 `present_files` 同时打开 **docx 与 pdf**，说明覆盖范围与任何未完成、存疑之处。

最终 `.docx` / `.pdf` **落到该科目的 OneDrive 文件夹**（见步骤 1），文件名
`<科目名> Outline.docx` / `.pdf`。

## 常见排版问题与成因（历史踩坑）

| 症状 | 成因 | 处理 |
|---|---|---|
| 有公式的那行文字与公式大小不一 | 段落里文字 run 与行内公式分属两套排版 | **混排段落**：文字 `w:r` + 行内 `m:oMath`，实测字号完全一致 |
| **正文行折行过早（右侧大片空白）+ 续行缩进像居中** | 整行是一个 `m:oMathPara`，Word 对数学段落的分行强加缩进 | **改用混排段落**（`prose_carrier='mixed'`） |
| 公式里正文显示成斜体 | ① 缺 `m:nor`；② 在线预览不支持 `m:nor` | ① 自动加 `m:nor` + `w:i=0`；② **改交 PDF** |
| 有的行靠左、有的行莫名居中、缩进乱 | 裸行内 `m:oMath`（段落里没有文字）被 Word 当数学段落按默认居中排 | 纯公式段包 `m:oMathPara` 并显式写 `m:jc` |
| 正文行被挤断、字距忽大忽小 | `\text{}` 里的空格被转成 NBSP，无法断行 | 自动换回普通空格（`\quad` 除外） |
| 公式错位、与上下行挤压 | 块级公式被当成行内 `m:oMath` 塞进段落 | 改用 `m:oMathPara` |
| 正文中间冒出超大 `∑` 把句子撑断 / 续行缩进到运算符位置 | Word 遇到 n 元运算符会强制缩进折行的续行（**行内 `\sum` 也一样**） | 自动拆行：正文 → 独立公式 → 正文 |
| 公式里空格时有时无 | `\,` `\quad` 等被 MML2OMML 丢弃 | 自动替换为 NBSP |
| **尖括号 / 取整括号只有默认大小，不随内容放大** | `\langle` `\lfloor` `\lceil` 被 latex2mathml 转成 `<mi>`，XSL 把整段合并成一个 run，定界符不是独立元素 | `_split_delim_runs` + `_wrap_stretch_delims` 拆 run 后包成 `m:d` |
| **`\det` / `\operatorname{tr}` 后面的矩阵括号完全不伸缩** | latex2mathml 把函数名和 `(` 合并成**一个** `<m:t>`（如 `"det("`），定界符不是独立元素，包不成 `m:d` | 拆分集加入 `( ) [ ] { }`，同上两步修复 |
| **`det` `tr` `rank` `lim` `max` `min` `gcd` `sup` 显示成斜体** | 这些命令 latex2mathml 输出 `<mo>`（不是 `<mi>`），XSL 不给 `sty=p`；样本是正体 | `preprocess_tex` 改写成 `\mathrm{…}` |
| 括号跑到下标位置 | 带下标的裸括号组被 latex2mathml 拆错 | 自动加 `\left…\right` |
| 多出 `(1)` `(2)` 编号 | 用了 `align` / `aligned` | 自动改 `array{rl}` |
| 标题层级看不出主次 | 加粗被覆盖、字号撞车 | 按样本 15/14/12/11 并强制加粗 |
| 标题颜色和样本不一样 | python-docx 内置模板的 Heading 是另一种蓝（`4F81BD`） | 样本实测为 **#0F4761**，已由 `CFG['heading_color']` 强制；要纯黑改 `'000000'` |
| **分点块和上下文糊在一起、看不出分点的优势** | 列表项一律用 `after_body`(8pt)，块上下都没有边界 | `_list_flags` 预扫出块首/末项，按四档给间距（见上表） |
| **每章开头空一大截、整页发飘** | 章标题后的第一条目也套了 `before_entry=36` | `_after_h1` 标出章标题后的行，改用 `before_entry_first=10` |
| **列表没有项目符号** | 只设了缩进、没挂编号 | 加段落级 `w:numPr`（`numId=1`，bullet）；带 `(a)/(1)` 标号的项跳过 |
| **换了项目符号却渲染成怪字** | 只改了 `w:lvlText`，字体还是默认的 `Symbol` | 同时把 `w:rFonts` 的 ascii/hAnsi/cs 改成 Cambria Math |
| **整块挤成一坨、次级小标题和列表项分不出来** | 段前间距只用了样式的 8/12pt；样本实际在标题前**另插一个 12pt 空段落**，视觉间距 ≈52pt | `CFG['before_sub']` / `CFG['before_entry']` = **36pt** |
| **短公式（如 `(b) … Σ_{ω∈Ω}ℙ({ω})=1`）被强行拆成两行** | 旧规则"见 ∑ 必拆"，不看行宽 | 改为**按行宽判定**：整行估算宽度 ≤ 列宽×0.97 就保持行内小公式 |
| **二级列表符号退化回一级符号（`-` 变 `•`）** | 默认 `abstractNum` 的 `multiLevelType=singleLevel`，Word 忽略 `ilvl>=1` | 改 `hybridMultilevel` 并注入 `ilvl=1` 定义 |
| **列表文字比样本靠左 4pt（108 vs 112）** | 文字位置由 numbering 的 `w:tabs` num tab 决定，默认 360 | 把 `w:tabs` 的 num tab 与 `w:ind` 一起改成 **440** |
| **表格列宽崩坏（短列压成一字宽、长公式撑出页面）** | Word 默认自动布局按内容分配，长公式列抢走全部宽度 | `tbl.autofit=False` + `w:tblLayout type="fixed"` + 逐列 `tcW`/`gridCol` 显式宽度 |
| **表格单元格里超长公式被裁切** | OMML 公式不会自动折行 | 在单元格里用 `\begin{array}{l}…\\…\end{array}` 手工分行（`\{` 与 `\}` 分处不同 `m:mr`，不会被误配成 `m:d`） |
| 导出 PDF 后检测不到粗体 | Cambria Math 无 Bold 字面，Word 走合成加粗，PDF 不嵌粗体字体 | **不是缺陷**，目视为准；勿用 pymupdf flags 判粗体 |

## v10 变更（ODE w1 复盘驱动，2026-09-30）

### D1：`**...**` 包公式时 bold 失效（已修）
- **症状**：markdown `$f$ is **differentiable at $x$** (...)` 生成后正文中
  残留 literal `**`（`' is **differentiable at '`、`'** (otherwise ...'`），bold 未生效。
- **根因**：原 `_INLINE_TOKEN = re.compile(r'(\$[^$]+\$|~~[^~]+~~)')` 先按 `$...$` 切，
  bold 跨公式被拦腰斩断；`_add_text_runs` 的 `\*\*[^*]+\*\*` 匹配不上非对称 `**`。
- **修复**：
  1. 扩展 regex：`_INLINE_TOKEN = re.compile(r'(\$[^$]+\$|~~[^~]+~~|\*\*[^*]+\*\*)')`
     —— 一次性切出 `$...$` / `~~...~~` / `**...**` 三种整 token。
  2. 新增递归函数 `_emit_inline(p, text, size, bold=False, strike=False)`：
     遇 `**...**` 去壳递归传 `bold=True`、`~~...~~` 传 `strike=True`，
     内部 `$...$` 走 `_append_math`（math 不吃 prose-bold）、prose 走
     `_add_text_runs(..., bold=bold)`。两处旧 `for tok in _INLINE_TOKEN.split` 循环均替换。
- **验证**（ODE v1 重生成）：
  - w:t 含 literal `**`：2 → **0**
  - `differentiable at` bold：否 → **是**
  - PDF p1 视觉：星号消失、`differentiable at x` 干净加粗、$x$ 仍为正体数学
- **副作用**：`~~strike~~` 跨公式的递归 strike 行为不变；行宽估算器
  `re.sub(r'\*\*|~~', '', text)` 也不变。

### NBSP 判据修正（ODE v1）
- 原 `verify_docx.py` 用 `<m:r><m:rPr><m:nor/></m:rPr>...<m:t>` 把公式内
  `\text{}\,` `\quad` `\qquad` 还原出的 NBSP 也当"正文 NBSP"算违规。
  实为 Word 公式区**不可断行的合法间距**。
- 改：w:t 与 m:nor 两个计数器分开报：
  - `正文 w:t 含 NBSP`：**必须为 0**（真隐患：词间断不开）
  - `公式内 m:nor 含 NBSP`：**预期行为**（thin/quad 显式间距），不计违规
- ODE v1：w:t NBSP=0 ✓，m:nor NBSP=21（17 来自 `\,`，3 来自 `\qquad`，1 来自 `\quad`）。

### v11：留白节律对齐手打 Outline（ODE w1 间距复盘，2026-09-30）

**问题**（用户看 v1 交付件："分段的多少和标题的位置还是很违和"）：
逐行量测手打 `ODEs Outline.pdf` 与 v1 的 gap（上段底→本段顶，pt），发现两处系统性偏离：

| 过渡类型 | Outline 实测 | v1 实测 | 偏差 |
|---|---|---|---|
| 节/章标题 → 紧接条目 | **≈19** | **40.8 / 40.7** | 标题被 40pt 大空隙"孤立" |
| 节/章标题 → 紧接正文/分点 | **≈19** | **8.4 / 9.1** | 标题挤进正文 |
| 条目 → 自身正文 | ≈14 | 10.2 | 略紧 |
| 正文块末 → 下一条目 | ≈40 | ≈40 | 正确（保留） |

**根因**：`before_entry=36` 被套到**每一个**条目（含"紧跟标题"的条目），且正文/分点紧跟标题时 `space_before=0`。
而手打 Outline 的节律是：**大留白（≈40）只出现在"真正的话题切换"（正文块末→下一条目）；标题与其领起的首段之间只≈19。**

**修复**（`build_note.py` 主循环）：
1. 新增 `prev_kind` 跟踪"上一段块类型"（h1/h2/h3/entry/sub/body/disp/table）。
2. 条目/次级小标题/正文首段/分点首项，**按 `prev_kind` 决定段前**：
   - 上一段是标题（或文档开头 `None`）→ 用 `*_after_heading` 系列（≈14~15pt，总间距≈19）；
   - 否则（正文块末）→ 用 `before_entry`/`before_sub`（36pt，总间距≈40）。
3. `after_entry` 8→**12**（条目→自身正文 10.2→14，对齐 Outline）。
4. 新增 CFG：`before_entry_after_heading=14`、`before_sub_after_heading=14`、
   `before_body_after_heading=15`、`before_list_after_heading=12`、`before_label_item_after_heading=14`；
   旧的 `before_entry_first` / `_after_h1()` 删除（被通用逻辑取代）。
5. Word 对相邻段落取 `max(space_after, space_before)`，故下方数值即"主导间距"。

**验证**（ODE v1 内容用 v11 重生成，导出 PDF 逐行量测）：
- 节标题 → 条目：40.8/40.7 → **18.8 / 18.7** ✓
- 节标题 → 正文首段：8.4/9.1 → **17.8 / 15.1** ✓
- 条目 → 自身正文：10.2 → **14.3** ✓
- 正文块末 → 条目：≈40（保留）✓
- 硬指标无回归：公式失败=0、裸行内公式=0、正文 w:t NBSP=0。
- 交付：`_blindtest/ODE_w1_v2.docx` / `.pdf` + `ODE_v2_vs_Outline.png`（左 v2、右手打 Outline 并排）。

> 注：`analyze_outline.py` 是 docx 版分析器；本复盘用的 PDF 行级量测脚本是临时探针
> `_outline_gaps.py` / `measure_gaps.py`（纯 PDF `get_text('dict')` 取每行 bbox/字号算 gap）。

### v12：条目级"整块同页" —— 消除孤行标题（ODE 复盘之三，2026-09-30）

**问题**：手打 Outline 时长期存在"条目标题在上一页末 / 主体在下一页头"的孤行分页
（如 Definition 2.4）。用户要求生成器保证每条小板块内容在一起 —— 放不下就整块移到下一页。

**根因**：之前 Heading 1/2/3 样式带 `keep_with_next`，但**章/节标题本身不会孤行**
（其下必有正文/条目）。真正会孤行的是 `###` 条目标题和 `####` 次级小标题 —— 它们不走
Heading 样式，是用 `_par` / `_entry_title` 直接创建的 Normal 段，从未带 `keep_with_next`。

**修复**（`build_note.py` 主循环）：
1. **条目级整块同页**：主循环维护 `entry_paras` 缓冲累积当前 `###` 条目的所有段落。
   遇到边界（章 `#` / 节 `##` / 新条目 `###`）时收尾上一条目：除最后一段外，
   **全部**段落设 `keep_with_next=True`，构成"keep-together 链"。
   - 放得下：整条链在同一页。
   - 放不下：整条链整体移到下一页（用户在上一页末尾自然结束，不会留孤行标题）。
   - 跨页（条目 > 一页）：Word 在链内断开，但标题必与首段同行，不会孤立。
2. **标题自身不跨页拆分**：Heading 1/2/3 样式增加 `keep_together=True`；
   `_entry_title` 与 `####` 次级小标题分支增加 `p.paragraph_format.keep_together=True`。
3. **捕获机制**：每行开头记录 `nb = len(doc.paragraphs)`，本行产出结束后
   `entry_paras.extend(doc.paragraphs[nb:])`（章/节标题 `h1/h2` 除外，其 `keep_with_next`
   已在 Heading 样式中处理）。
4. **收尾**：循环结束后对最后一条 `entry_paras` 做同样收尾（除末段外全部
   `keep_with_next`）。

**确定性验证**（`/f/tmp/deterministic_orphan.py`，filler 调在"标题能放下、主体放不下"的
窄窗口）：

| filler | NOFIX（关闭 v12 链） | FIXED（v12） |
|---|---|---|
| 52 行 | 标题+正文同页（未孤行） | 整块一起移到 p3 |
| 53 行 | **标题 p2 末尾 / 正文 p3 头部 → ORPHAN** | 整块一起在 p3 |
| 54 行 | **标题 p2 末尾 / 正文 p3 头部 → ORPHAN** | 整块一起在 p3 |

证据图 `/f/tmp/orphan_vs_fixed.png`（左 NOFIX p2：标题孤立；右 FIXED p3：标题+正文同页）。

**回归验证**（`ODE_w1_v1.md` → v3）：
- `verify_docx.py`：公式失败=0、裸行内公式=0、正文 w:t NBSP=0、Heading 字号/颜色正确；
  段后间距分布、行距分布无变化。
- 页数：v2 4 页 → v3 5 页（一个条目原本被分割，现整块移到下一页 —— 用户期望的
  "把分开的板块放到下一页"行为）。
- 交付：`_blindtest/ODE_w1_v3.docx` / `.pdf`。

**已知限制 / 待办**：
- 当前**对所有条目（含 Proof）一视同仁地用整块链**。如果某门课的长证明（>半页）
  经常刚好填不到当前页剩余，整链会带着证明一起跳到下一页，在上一页留较大空白。
  应对方法（Y2 Prob 2 有 54 个 Proof 时再评估）：在 `Proof` 起手处把当前条目 flush、
  关闭 `entry_paras` 累积，证明内段落仅在首段设 `keep_with_next` 让 "Proof." 不孤行，
  证明主体允许自然跨页。本轮 ODE 验证无需此处理，留待做 Prob 2 时再改。

### v13：正文上方标题留"一格"空（2026-09-30，用户第四轮反馈）

**问题**：用户原话 —— **「小标题和大标题上面如果是小字正文的话最好还是空一格，
就按标题本身的字号和间距空就行」**。即：`##` 小标题 / `#` 大标题若正上方是小字正文，
应留一格空档，大小 = 标题自身字号 × 行距。

**实测根因**：Heading 2/3 样式的 `space_before` 只有 **8pt**，所以正文→节标题渲染出的
gap 仅 **12.1~12.4pt**（标题像挤在正文里）。而手打 ODEs Outline 的"正文→节标题"
实测 **20.3 / 20.4 / 20.5pt**。用 `14pt × 1.158 ≈ 16.2pt` 段前 → 渲染 gap ≈ **20.4pt**，
与 Outline 逐位吻合。

**修复**（`build_note.py` 主循环，`#`/`##`/`#####` 三个 Heading 分支）：
当 `prev_kind == 'body'`（正上方是小字正文/列表）时，
`space_before = max(样式原段前, 标题字号 × CFG['line_body'])` —— **只补不削**：
- `##`(14pt)：8 → 16.2pt（渲染 gap 12.2 → **20.5**）✓ 对齐 Outline
- `#`(15pt)：max(24, 17.4)=24，不变（本就够宽）
- `#####`(16pt)：8 → 18.5pt
- `####`(before_sub=36)、`###`(before_entry=36) 不动（本就够宽，max 后不变）
新增 CFG：`before_sub3=8`（Heading 3 段前，与样式统一）。
**为何用 max（而非直接覆盖）**：用户说"最好还是空一格"= 补上缺失的空档，
不是把已够宽的标题压紧；max 保证只补不削。

**验证**（`ODE_w1_v1.md` → v4）：
- 正文→`##`：12.2/12.1/12.4/12.4 → **20.5/20.4/20.6/20.6**（Outline = 20.3~20.5）✓
- 正文→条目、正文→`####` 全部不变（38~46pt）✓
- 硬指标无回归：公式失败=0、裸行内公式=0、正文 w:t NBSP=0、Heading 字号/颜色正确。
- 证据图 `_blindtest/heading_blank_v3_vs_v4.png`（左 v3 挤、右 v4 有空档）。
- 交付：`_blindtest/ODE_w1_v4.docx` / `.pdf`。

### v14：空底数脚本的"幻影宽度"修复（撇号/点号，2026-09-30，用户 `/pdf` 提问）

**问题**：用户问 —— **「2.4 的位置 prime 和后面那个撇为什么隔这么远」**。
实测 `Lagrange: prime ` 结束于 x=191.64，而撇号 `′` 起始于 x=201.77 → **gap 10.1pt**
（正常空格仅 2.5pt）。`Convention: ` 同理（10.1pt）。用户接着追问
**「这个撇是不是应该更加右上角一点」** —— 说明不仅要消空隙，撇还要**保持右上角（上标）位置**。

**根因**：latex2mathml 把**孤立的** `'`（前面没有底数）转成
`<msup><mi/><mi>′</mi></msup>` —— **空底数**。MML2OMML 产出带空 `<m:e>` 的 `m:sSup`，
Word 给空底数槽留出 ~7.6pt 幻影宽度，撇号因此被推远。
（`$\dot{}$` 空点号同源：`<m:limUpp><m:e/>`。）

**两次尝试（记录踩坑）**：
1. ❌ 先把孤立撇号换成 `\prime`（→ `<mi>′</mi>` 纯 run）：空隙没了，**但撇掉到基线**，
   不再上标 —— 用户随即指出"应该更右上角"。**已回退**。
2. ✅ 正确做法：**保留 `m:sSup` 结构**（撇号仍是 8.04pt 上标），只把**空底数**换成
   **零宽空格(U+200B)** —— 宽度归零，上标位置/字号完全不变。

**修复**（`build_note.py`，`latex_to_omml` 里新增 `_fix_empty_script_bases`）：
遍历 `m:sSup/sSub/sSubSup/limUpp/limLow/mover/munder`，若其底数 `<m:e>` 无可见文字，
清空并塞入一个含 ZWSP 的 run。一次修复同时覆盖**孤立撇号**与**空点号 `\dot{}`**。
```python
_SCRIPT_BASE = {'sSup','sSub','sSubSup','limUpp','limLow','mover','munder'}
def _fix_empty_script_bases(root):
    for el in root.iter():
        if etree.QName(el).localname not in _SCRIPT_BASE: continue
        e = el.find('{%s}e' % M)
        if e is None or _has_visible_text(e): continue
        for c in list(e): e.remove(c)
        r = etree.SubElement(e, '{%s}r' % M)
        etree.SubElement(r, '{%s}t' % M).text = '\u200b'   # ZWSP
```

**验证**（`ODE_w1_v1.md` → v6）：
- `Lagrange: prime ′`：撇号 x0=191.57（正文末 191.64）→ **gap ≈ 0**，
  y0=568.52 vs 正文 570.26 → **上移 1.74pt（上标位，8.04pt）** ✓
- `Convention: ′ denotes…` 同上 ✓
- 附带修复：`$\dot{}$` 空点号两侧 ~4pt 空隙 → ≈0 ✓
- `h'(x) = f'(g(x))g'(x)` 的上标撇不变 ✓
- 硬指标无回归：公式失败=0、裸行内公式=0、正文 w:t NBSP=0。
- 证据图 `_blindtest/prime_three_way.png`（上 v4 上标+空隙 / 中 v5 基线无空隙 /
  下 v6 上标且无空隙）。
- 交付：`_blindtest/ODE_w1_v6.docx` / `.pdf`。

> **通用性**：任何"空底数"脚本结构（孤立撇号、空 `\dot{}`/`\ddot{}`/`\bar{}` 等）都适用。

### v15：目录 + 页码（2026-09-30，用户要求"对教材分析并加入目录、页码"）

**用户要求**：给笔记加**目录**与**页码**。层级定 **章 + 节（1-2）**，页码定
**页脚居中 `X / Y`**。用户明确：目录**看教材自己的目录**，教材没有就**按各章写出来的大标题**
建；"暂时还不需要做什么分析的工作"（即不做讲义 PDF 的章节结构抽取）。

**样本实测**（7 份 Outline 中 3 份带目录）：

| 样本 | TOC 域 | 目录层级 | 页脚 |
|---|---|---|---|
| MFA / Probability I | ` TOC \o "1-2" \h \z \u ` | TOC1 章 + TOC2 节 | 居中 `PAGE` / `NUMPAGES` → `1 / 46` |
| Linear Algebra | ` TOC \o "1-3" \h \z \u ` | 收到条目层（目录很长） | 无 |
| 其余 4 份（ITVC / ODEs / RA / Stats I） | 无目录 | — | 无 |

其它实测结论：
- **目录后另起一页**：Prob I 目录 p1–p2、正文 p3；LA 目录 p1–p3、正文 p4。
- **页码全文连续编号**，目录页也算第 1 页；目录里的页码就是这个编号（Prob I 目录写
  "Chapter 1 … 3"，正文确实在 p3）。
- TOC1 基于 Normal（**11pt，无额外格式**）；TOC2 只多一个 `ind left=420` twips。
- 条目末尾是**点线引导 + 右对齐页码**（Word 生成，`w:leader="dot"`）。

**实现**（`build_note.py` 新增 `_ensure_toc_styles` / `_insert_toc` / `_add_page_numbers`
/ `_set_update_fields` / `_finalize_toc_and_pages`，由 `build()` 收尾调用）：

1. **TOC1/TOC2 样式**：直接注入 `styles.xml`，`w:styleId` 与**内建名**
   （`w:name w:val="toc 1"`）都与样本一致 —— Word 的 TOC 域是**按内建名**给各级条目套格式的，
   自造名字会套不上。样式里带一个右对齐 + `w:leader="dot"` 的制表位，
   落点 `8306` twips = 页宽 11906 − 左右边距 1800×2，页码才会右贴边。
2. **TOC 域**：在正文最前面插 `[目录标题][TOC 域]` 两段；域是**原生域**
   （`TOC \o "1-2" \h \z \u`），页码由 Word 生成，不是硬编码 —— 增删内容后按 F9 即可刷新。
   域内先放一行占位结果，Word 一更新就被真实条目替换。
3. **页码**：`sections[0].footer` 居中放 `PAGE` + `" / "` + `NUMPAGES` 三段域
   （begin / instrText / separate / 占位 / end），页脚距 992 twips（= 样本 `pgMar footer`）。
4. **另起一页**：给第一个 `Heading 1` 设 `page_break_before`（样本实测行为）。
5. **`updateFields`**：`settings.xml` 写 `<w:updateFields w:val="true"/>`，Word 打开即刷新。
   ⚠ `CT_Settings` 是**有序**类型，必须插在 `compat` / `rsids` 等**之前**，否则 schema 非法。

**刷新与导出**（新增 `scripts/finalize_docx.ps1`）：

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File scripts/finalize_docx.ps1 `
  -Src "输出.docx" -OutDocx "暂存.docx" -Pdf "输出.pdf"
```

域在 Word 构建前没有缓存结果，所以**必须**跑一次 Word 才有真实目录与页码。要点：

- `$doc.Fields.Update()` **只覆盖正文 story**，页脚的 `PAGE`/`NUMPAGES` 拿不到 ——
  必须遍历 `$doc.StoryRanges`（含 `NextStoryRange` 链）才能刷到页眉页脚文本框。
  只写 `Fields.Update()` 时实测页脚停在占位值 `1 / 1`（页数其实是 6），改成 StoryRanges
  遍历后立刻变成 `1 / 6`。
- 顺序：`Fields.Update()` → `TablesOfContents.Update()` → `Repaginate()` →
  StoryRanges 遍历 → `Repaginate()` → 保存/导出。
- 诊断行 `pages=N` / `footer[1]=...` 会写进日志，便于确认域真的刷新了。

**Word 往返的副作用**：Word 自己保存 docx 时会把 `<w:i w:val="0"/>`（公式内正文强制正体，
见"铁律五"）当作冗余删掉 —— 实测 **21 → 0**。这会影响第三方在线预览。
故新增 `scripts/restore_upright.py`，在 Word 刷新后把它补回来：

```bash
"$PY" scripts/restore_upright.py 暂存.docx 输出.docx     # 补 w:i=0，其余字节原样透传
```

只改 `word/document.xml`，其余部件逐字节复制。补回后 `w:i=0` 计数恢复 21。

**验证**（`ODE_w1_v1.md` → v8）：

- 目录页：`Contents` 标题 + 9 条（TOC1 3 / TOC2 7），点线引导、页码右贴边、页码 2/2/2/2/3/4/4/5/6
- 页脚 `1 / 6` … `6 / 6`，`pages=6`
- Word 往返前后**公式渲染逐像素一致**（v6 p2 ↔ v8 p3、v6 p4 ↔ v8 p5 并排比对）
- 硬指标无回归：公式失败=0、裸行内公式=0、正文 w:t NBSP=0、`w:i=0`=21、Heading 字号/颜色正确
- 交付 `_blindtest/ODE_w1_v8.docx` / `.pdf`；证据 `v15_toc_page.png`、`v15_wordroundtrip_check.png`

**目录美化验证**（`ODE_w1_v1.md` → v9，同一份内容只改目录样式）：
章条目 12pt 加粗深蓝、节条目 11pt 黑缩进 21pt；组内行距 ≈16.9pt、组间 ≈25.2pt；
页数仍 6 页、页脚 `1 / 6`；硬指标与 v8 一致。证据 `v15_toc_plain_vs_styled.png`（左素 / 右美化）。
交付 `_blindtest/ODE_w1_v9.docx` / `.pdf`。

**目录条目的外观**（v15 第二轮，用户"可以美观一点吗"）：样本的目录很朴素
（TOC1 继承 Normal = 11pt 不加粗、TOC2 只多一个缩进）。默认做得更"立得住"：

| 层级 | 字号 | 字重 | 颜色 | 缩进 | 段前 / 段后 |
|---|---|---|---|---|---|
| `Contents` 标题 | 16pt | 加粗 | `#0F4761` | 0 | 0 / 18pt |
| TOC1（章） | **12pt** | **加粗** | **`#0F4761`** | 0 | **10pt** / 2pt |
| TOC2（节） | 11pt | 常规 | 黑 | 21pt（420twips） | 0 / 2pt |

效果：每章成为**视觉分组**（章条目加粗深蓝 + 组前 10pt 留白），节条目缩进小字；
点线引导 + 页码右贴边不变。实测组内行距 ≈16.9pt、组间 ≈25.2pt。

**CFG 开关**：

- 目录：`toc` / `toc_levels` / `toc_break`
- 标题：`toc_title`（置 `None` 则不加标题行）/ `toc_title_size` / `toc_title_after`
- 条目：`toc_entry`（节）/ `toc_entry1`（章）/ `toc_bold1` / `toc_color1`（`None` → 黑）/
  `toc_before1` / `toc_after1` / `toc_before2` / `toc_after2` / `toc_indent2`
- 页码：`page_numbers` / `page_number_format`（`'both'` → `X / Y`，`'page'` → 仅 `X`）/
  `page_number_size` / `page_number_gap`

> 想回到样本的素净样子：`toc_bold1=False`、`toc_color1=None`、`toc_entry1=11`、
> `toc_before1=0`、`toc_after1=0`。

### v16：文档大标题 + 作者行（2026-10-03，用户：「是不是没有按照之前的模板加大标题？」）

**问题**：用户发现生成的提纲**没有文档大标题**。实测确认属实 —— `build_note.py` 只有
`toc_title`（= `Contents`），**没有任何"文档标题"概念**；`analyze_outline.py` 也没有 `title` 分类。
`template-spec.md` 里那句"我们的笔记没有文档标题"是**生成侧的自我描述**（循环论证），不是样本事实。

**样本实测**（7 份 Year 1 Outline 的段落 0 / 1，OOXML + PDF 双向量测）：

| 项 | 实测值 | 样本一致性 |
|---|---|---|
| 段落 0 文字 | **`<课程全名> Outline`** | 6/6 可读样本都有（LA 那份 docx 损坏读不出） |
| 段落 0 格式 | Cambria Math、sz=32(16pt)、加粗、左对齐、**无 `w:jc` / 无 `w:spacing` / 无 color** | 5/6 加粗（Statistics I 未加粗，PDF 里还退化成 DengXian，属其笔误） |
| 段落 1 文字 | `By JasonTan`（ITVC 作 `BY`） | **3/6** 有；ODEs / Real Analysis 是空段，Statistics I 没有 |
| 段落 1 格式 | Algerian、sz=28(14pt)、常规、**`w:jc=right`**、无 spacing | 同上 |
| 渲染（Prob I / MFA PDF） | 标题行框顶 **y0=81.07pt**；标题→作者 Δ=**23.47pt**；作者→目录首行 Δ=**17.55pt** | 两份**逐位相同** |

**关键坑一：样本的 `sectPr` 带行网格** —— `<w:docGrid w:type="lines" w:linePitch="312"/>`
（Word 东亚版式默认），本脚本输出的是 python-docx 默认（无网格）。**单变量实验**：
把样本的 `docGrid` 塞进我们的 docx 后，标题 y0 74.83 → **81.07**、Δ 13.87 → **23.47**，
与样本**完全一致** —— 差异 100% 来自行网格。

但行网格是**文档级**属性，会改变全篇行距（实测 **6 页 → 7 页，+17%**），
而 v11/v13 的间距是在**无网格**渲染下校准的，启用网格会把它们全部推翻。
**决定（用户 2026-10-03 拍板）：保持无网格，用显式段前把这两行补回样本的渲染值。**

**关键坑二：Word 段间间距取 `max(上一段 after, 本段 before)`，不是相加。**
标题继承来的 `after=8pt` 会盖住作者行较小的 `before`。所以作者行 `before` 必须写成
**17.6pt（= 8 + 9.6）**；只写 9.6pt 只净增 1.6pt（实测 13.87 → 15.43，正是此因）。

**实现**（`build_note.py`）：
- 新增 CFG：`doc_title` / `doc_title_size` / `doc_title_bold` /
  `doc_author` / `doc_author_size` / `doc_author_font` / `doc_title_before` / `doc_author_before`
- `doc_title=None` → **自动取输出文件名主干**（样本"大标题 == 文件名"：
  `Probability I Outline.docx` 的标题就是 `Probability I Outline`）；置 `''` 则不加
- 插入顺序（`body.insert(0, …)` 倒序执行）：**大标题 → 作者行 → `Contents` → TOC 域**
- 新增 `_xml_esc()`：标题里的 `&`（`Mathematical Foundation & Analysis Outline`）
  直接拼进 `w:t` 会产出非法 XML，必须转义。**注意它与 `_esc_prose()` 的私用区占位符是两回事**
- 若将来决定启用行网格，把 `doc_title_before` / `doc_author_before` 改成 `0` 即可

**验证**（Word 导出 PDF 后逐行量 bbox，非读 XML 猜）：

| 指标 | 样本 Prob I / MFA | 本脚本 v16 |
|---|---|---|
| 标题 y0 | 81.07 | **81.07** ✓ |
| 标题→作者 Δ | 23.47 | **23.47** ✓ |
| 标题字号 | 15.96 | **15.96** ✓ |
| 作者字号 | 14.04 | **14.04** ✓ |
| 作者右边缘 x1 | 508.93 | **509.05** ✓（页宽差 0.1pt） |
| 字体 | CambriaMath / Algerian | **同** ✓ |
| 硬指标（`verify_docx.py`） | — | 裸行内公式 0 / 正文 NBSP 0 / 汉字 0 / 公式失败 0 ✓ |

**顺带修了 `analyze_outline.py`**：新增 `title` / `author` 两个分类。
之前样本的大标题和作者行被算进 `body`，把 16pt/14pt 混进正文字号统计、把
`Algerian` 混进字体统计、把 `right` 混进对齐统计 —— 直接污染"规范只来自实测"。
修后样本 body 从 182 → 180 段，字体统计里的 `Algerian` 消失。

> ⚠ **已知限制**：`analyze_outline.py` 报告的 `after` / `before` / `line` 是**段落直接 pPr**
> 的值，**不沿样式链解析**（只有字号走了 `eff_sz`）。所以本脚本产物的 `title` 会显示
> `after=200 line=276`（docDefaults），而实际生效的是 Normal 样式里的 `160 / 278`。
> 拿它下间距结论时要注意。

## 参考

- `references/template-spec.md` — 页面、字体、层级、编号、列表、分行规则的完整实测规范，
  含 7 份样本的跨科目校准结论
- `references/latex-omml.md` — 公式链路、整行并入公式的实现细节、自动修复项与已知限制

脚本一览：

| 脚本 | 作用 |
|---|---|
| `scripts/build_note.py` | Markdown + LaTeX → docx。顶部 `CFG` 集中调参；自带目录/页码（v15） |
| `scripts/verify_docx.py` | 回读硬指标：裸行内公式=0、正文 NBSP=0、公式失败=0、**正文含汉字=0**（`--allow-cjk` 豁免）、`w:i=0`、目录/页码域状态 |
| `scripts/selftest_formulas.py` | 公式 / 定界符 / 正体 / 行宽估算自检（不依赖 docx） |
| `scripts/finalize_docx.ps1` | **Word 刷新目录与页码域** → 回写 docx → 导出 PDF（v15 必跑） |
| `scripts/restore_upright.py` | 补回 Word 保存时删掉的 `<w:i w:val="0"/>` 正体提示（v15） |
| `scripts/docx_to_pdf.ps1` | 纯导出 PDF（无目录/页码时用） |
| `scripts/render_pdf_preview.py` | PDF → PNG 目视核对；`--lines` 行坐标、`--brackets` 定界符高度 |
| `scripts/analyze_outline.py` | **Outline 排版参数分析器**。**规范只应来自实测**，复核或新增科目时用它重跑：`analyze_outline.py <样本.docx>`（单份画像）/ `--compare`（跨样本对照表）/ `--json out.json`（存盘比对）。按**段落类别**（chapter/section/entry/body/list/toc）聚合真实生效的字号、加粗、颜色、间距、行距，并解析 `basedOn` 样式链 —— 样本主要靠直接格式化，只看 `styles.xml` 会得出错误结论。 |

## 环境备注

- `C:\Program Files\Microsoft Office\root\Office16\MML2OMML.XSL` 已确认存在。
- 隔离 venv：`C:/Users/Admin/.workbuddy-ai/binaries/python/envs/default/Scripts/python.exe`
- 已安装：lxml 6.1.3、python-docx 1.2.0、latex2mathml 3.81.1、pymupdf 1.28.2、pillow
- **pillow 用于行宽估算**（`_line_width_pt` 决定公式用小还是大）。缺了也能跑
  （退化为"字符数 × 平均宽度"的粗估），但精度会下降，建议装上。
  字体路径 `C:\Windows\Fonts\cambria.ttc` 的 **index=1 才是 Cambria Math**（index=0 是 Cambria）。
- 样本真身（7 份 Outline，均在 OneDrive）：
  `C:\Users\Admin\OneDrive - The University of Manchester\University\Year 1\`
  - Semester 1：`Introduction to Vector Calculus/ITVC Outline.docx`、`MFA/MFA Outline.docx`、
    `Probability I/Probability I Outline.docx`
  - Semester 2：`Linear Algebra/Linear Algebra Outline.docx`、`ODEs/ODEs Outline.docx`、
    `Real Analysis/Real Analysis Outline.docx`、`Statistics I/Statistics I Outline.docx`
  - 其中 6 份有配套 `.pdf`（Real Analysis 没有），可作渲染对照。
- 从 bash 调 `powershell` 会被安全策略拦截，导出 PDF 要用 PowerShell 工具执行。
- `docx_to_pdf.ps1` **必须保持纯 ASCII**：PowerShell 5.1 按 ANSI(GBK) 解码无 BOM 文件，
  中文字节会注入杂散大括号并破坏 try/catch 解析（报 `catch: CommandNotFound`）。
  本机执行策略为 Restricted，必须以 `-ExecutionPolicy Bypass -File` 调用，否则**静默不执行**。
