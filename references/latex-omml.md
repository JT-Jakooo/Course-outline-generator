# LaTeX 公式 → Word 原生公式（OMML）

## 链路

```
LaTeX --(latex2mathml)--> MathML --(Office MML2OMML.XSL)--> OMML --(python-docx)--> .docx
```

| 环节 | 依赖 | 位置 |
|---|---|---|
| LaTeX → MathML | `latex2mathml` 3.81.1 | pip |
| MathML → OMML | `MML2OMML.XSL` | `C:\Program Files\Microsoft Office\root\Office16\MML2OMML.XSL` |
| MathML → OMML 执行 | `lxml` 6.1.3（XSLT 1.0） | pip |
| 写入 docx | `python-docx` 1.2.0 | pip |

安装（隔离 venv，勿污染系统）：

```bash
PY=C:/Users/Admin/.workbuddy-ai/binaries/python/envs/default/Scripts/python.exe
"$PY" -m pip install lxml python-docx latex2mathml
```

---

## 一、含正文的行：混排段落（解决"字号不一"且不破坏折行）

**样本证据**：`Linear Algebra Outline.docx` 中 `<m:nor>`（Word 公式内用引号括起来的
"普通文本"）出现 **1102 次**；1395 段里纯公式段 626、纯文字段 434、混排仅 38。
即作者把**整行文字都写进公式对象**，再用引号把文字变回普通文本。

**但这个做法有个副作用**：整行变成一个数学段落（`m:oMathPara`）后，
**Word 对数学段落的分行会强加缩进**。实测同一段文字：

| 承载方式 | 第 1 行 | 第 2 行 | 字号 |
|---|---|---|---|
| 整行 `m:oMathPara` `jc=left` | 90 → 318.8（**右侧空 186pt**） | 162 起（**缩进 72pt**） | 正文 11.04 / 下标 8.04 |
| **混排段落 `w:jc=left`** | 90 → **524.5（占满）** | **90 起（齐左）** | 正文 11.04 / 下标 8.04 |

字号实测完全一致（pymupdf span size），差别只在折行。

**实现**：`_emit_line()` 判断行里除公式外是否还有正文（`_has_prose`）：

- **有** → 按 `$...$` 切成 `_split_math_segments()`，文字段走 `_add_text_runs()`（普通 `w:r`），
  数学段走 `_append_math(display=False)`（行内 `m:oMath`），段落设 `w:jc` = `prose_jc`。
- **没有**（纯公式行）→ 整行转成一个 `m:oMathPara`。

开关 `CFG['prose_carrier']`：`'mixed'`（默认）/ `'inline_formula'`（回到样本写法）。

`m:nor` 仍会在**纯公式**里出现 —— 公式体内部的 `\text{}` 依然映射为 `<m:nor>`。

配套细节：

| 需求 | 处理 |
|---|---|
| `**加粗**` 在公式内 | `\text{}` 内插私用区哨兵 `\uE000/\uE001`，转换后按哨兵拆 run 并写 `<w:b/>`（`\text{}` 相邻段会被合并，靠位置索引不可靠） |
| `&` `<` `>` | latex2mathml 的 `\text{}` **不做转义**，这些字符会产出非法 XML → 先换成 `\uE002/\uE003/\uE004`，回写时还原 |
| `{` `}` | 会被当作 TeX 分组吃掉 → 换成 `\uE005/\uE006` |
| `%` `_` `#` `$` `^` `~` `\` `|` | 在 `\text{}` 内原样透传，**不要**加反斜杠转义（加了反而会显示成 `\%`） |
| 正体（非斜体） | 除 `m:nor` 外再补 `<w:i w:val="0"/>`（`_force_upright`）。Word 只靠 `m:nor` 即可，但在线预览只认 `w:rPr`，会误渲染成斜体 |
| **NBSP** | latex2mathml 会把 `\text{}` 内的普通空格转成 U+00A0，导致 Word 无法断行 → 一律换回普通空格；`\quad` 等**有意**间距用私用区标记 `\uE007` 保留为 NBSP |
| `w:rPr` 子元素顺序 | 手工插入 `w:b`/`w:i` 后必须按 OOXML 规定顺序重排（`_normalize_rpr_order`），否则严格校验器报错 |

---

## 二、纯公式段落必须用 `m:oMathPara` + 显式 `m:jc`

**实测坑**：段落里只有裸行内 `m:oMath`（没有文字）时，
Word 会当"数学段落"按默认对齐（**居中**）处理，段落自身的 `w:jc="left"` 被忽略。
对照实验：

| 方案 | 首行 x0（正文列起点 90pt） | 结论 |
|---|---|---|
| 裸 `oMath` + 段落 `w:jc=left` | 134.7 | **仍居中**（`w:jc` 无效） |
| 裸 `oMath` + 前置零宽空格 run | 90.0 | 左对齐 |
| `oMathPara` + `m:jc=left` | 90.0 | 左对齐 |

后果就是同一份文档里"有的行靠左、有的行莫名居中、缩进忽大忽小"。

```xml
<w:p><w:pPr>…</w:pPr>
  <m:oMathPara>
    <m:oMathParaPr><m:jc m:val="center"/></m:oMathParaPr>
    <m:oMath>…</m:oMath>
  </m:oMathPara>
</w:p>
```

**对齐按段落类型分配**（样本分类实测：含文字段落 left 138:48、纯公式段落 center 83:20）：

| 段落 | 对齐 | CFG 键 |
|---|---|---|
| 含正文文字 | 混排段落 + `w:jc=left` | `prose_jc` |
| 独占一行的纯公式 | `center` | `display_jc` |
| 多行分步推导（`array` / `aligned`） | `left` | `derive_jc` |

判定"多行推导"：`preprocess_tex` 后公式体含 `\begin{array}` 即视为推导。
矩阵 / 分段函数用的是 `pmatrix` / `cases`，不会误判。

---

## 二·补、大运算符自动分行

正文行里出现**带上下限**的大运算符时，自动拆成
`正文段（left） → 独立公式（center/left） → 正文段（left）`。

原因：**Word 遇到 n 元运算符会强制把折行的续行缩进到运算符位置**。
实测行内 `\sum` 也一样 —— `Proof. The (i,l)-entry of both sides equals $\sum_{j,k}…$. Hence…`
的续行从 333.4pt 起（正常应 90pt）。样本作者从不在正文行内使用 `∑`，一律另起独立公式。

触发条件（保守，避免误伤）：

- 数学段匹配 `\\(sum|prod|int|lim|bigcup|…)`，**且**含 `_` 或 `^`
- 前后正文合计 ≥ 8 个非公式字符

> 注意正则边界：**不能用 `\b` 收尾** —— 正则里 `_` 也算单词字符，`\sum_{…}` 处 `\b` 不成立，
> 会导致该规则静默失效（踩过）。用 `(?![A-Za-z])`。

开关：`CFG['split_big_ops']`，默认 `True`。

---

## 二·补二、定界符必须包成 `m:d` 才会伸缩

LaTeX 里括号随内容高度自动放大（`\left…\right` 只是显式声明，TeX 默认就会伸缩）。
OMML 里对应的是 **`m:d`（delimiter）对象** —— Word 只对 `m:d` 里的括号做纵向拉伸。
如果括号只是普通字符 run，渲染出来就是**固定大小**，跟内容一样高。

### 实测：哪些写法能自动产出 `m:d`

| 写法 | MathML | OMML | 结果 |
|---|---|---|---|
| `(\frac{a}{b})`、`[\frac{a}{b}]`、`\{\frac{a}{b}\}` | `<mo fence>` | `m:d` | ✅ 伸缩 |
| `\left(\frac{a}{b}\right)` | `<mo fence>` | `m:d`（**不带** `begChr`/`endChr`，默认圆括号） | ✅ 伸缩 |
| `|\frac{a}{b}|`、`\|\frac{a}{b}\|`、`\lvert`、`\lVert` | `<mo fence>` | `m:d` | ✅ 伸缩 |
| `pmatrix` / `bmatrix` / `vmatrix` / `cases` | — | `m:d` | ✅ 伸缩 |
| `\langle…\rangle`、`\lfloor…\rfloor`、`\lceil…\rceil` | `<mi>`（**标识符，不是 fence**） | 普通 `m:r` | ❌ **固定大小** |
| `\left\langle…\right\rangle`、`\left\lfloor…\right\rfloor` | `<mo fence>` | 仍被合并成普通 `m:r` | ❌ **固定大小** |
| **`\det\begin{pmatrix}…\end{pmatrix}`** | `<mo>det</mo>` + `<mo fence>` | 函数名与 `(` 合并成**一个** `m:r` | ❌ **完全不伸缩** |
| **`\operatorname{tr}\begin{pmatrix}…\end{pmatrix}`** | 同上 | 同上 | ❌ **完全不伸缩** |
| **`\sin(x)`、`f(x)`、`f(g(x))`、`\log_{2}(x+1)`** | 同上 | 同上（`"f(x)"` 一个 run） | ❌ 固定大小（内容矮时看不出） |

### 根因（两类）

**（a）尖括号 / 取整括号：**

`latex2mathml` 把 `\langle` 归成 `<mi>`（identifier）而不是 `<mo>`。
于是 `\langle u,v\rangle` 转出的是

```xml
<mrow><mi>⟨</mi><mi>u</mi><mo>,</mo><mi>v</mi><mi>⟩</mi></mrow>
```

`MML2OMML.XSL` 会把**连续的简单 token 合并成一个 `m:r`**：

```xml
<m:oMath><m:r><m:t>⟨u,v⟩</m:t></m:r></m:oMath>
```

定界符根本不是独立元素，无法配对成 `m:d`。
（`\left\langle\frac{a}{b}\right\rangle` 之所以能成，是因为中间的 `m:f` 打断了合并 ——
这也解释了为什么"括号里有分式"时看起来是对的，而 `⟨u,v⟩` 是坏的。）

**（b）函数名紧邻定界符：**

`\det` / `\operatorname{tr}` / `\sin` / `f` 这类函数名后面紧跟 `(` 时，
`latex2mathml` 输出的 `<mo>`/`<mi>` 与 `<mo fence>` 会被 XSL 合并进**同一个** `m:r`：

```xml
<m:oMath>
  <m:r><m:t>det(</m:t></m:r>     <!-- 函数名和 ( 粘在一起 -->
  <m:m>…矩阵…</m:m>
  <m:r><m:t>)</m:t></m:r>
</m:oMath>
```

`(` 不是独立元素 → 包不成 `m:d` → 矩阵括号完全不伸缩。
实测 `\det\begin{pmatrix}a&b\\c&d\end{pmatrix}` 的括号墨迹高度只有行内墨迹的 **0.56 倍**
（正常应 ≥ 0.9）。内容矮（如 `f(x)`）时肉眼看不出来，**内容一高就露馅**。


### 修复（`_stretch_delims`）

两步，在 `_append_math` 里转换后立即执行：

1. **`_split_delim_runs`** —— 遍历所有非 `m:nor` 的 `m:r`，
   若 `m:t` 含拆分集中任一字符，按"每个定界符字符独立成 run"重新切分
   （`⟨u,v⟩=0` → `⟨` / `u,v` / `⟩` / `=0`；`det(` → `det` / `(`）。

   拆分集 = `( ) [ ] { } ⟨ ⟩ ⌊ ⌋ ⌈ ⌉ ⟮ ⟯`。
   其中 `( ) [ ] { }` 是为了兜住上面 (b) 类的合并情形；
   `⟨ ⟩ ⌊ ⌋ ⌈ ⌉ ⟮ ⟯` 是 (a) 类。
2. **`_wrap_stretch_delims`** —— 在同一父节点下找**成对**的定界符 run，
   连同中间内容包进 `m:d`：

```xml
<m:d>
  <m:dPr><m:begChr m:val="⟨"/><m:endChr m:val="⟩"/></m:dPr>
  <m:e>…中间内容…</m:e>
</m:d>
```

配对策略是**取距离最近的一对**，这样嵌套时总是先裹最内层：

```
⟨⟨a⟩⟩  →  ⟨ m:d(⟨a⟩) ⟩  →  m:d(⟨ m:d(⟨a⟩) ⟩)
f(g(x)) →  f( g m:d(x) )  →  f m:d( g m:d(x) )
```

（若按"从左往右第一个匹配"，外层 `⟨` 会先跟内层的 `a` 裹在一起，留下一只固定大小的 `⟩`。）

细节：

- `m:nor`（公式内正文文本）里的字符**不动** —— 那是 `\text{The matrix (1,0)}`，
  本就该是普通文字。实测 `\text{}` 在 `_stretch_delims` 阶段**已带 `m:nor`**
  （`<mtext>` 由 XSL 直接映射），所以保护是有效的。
- **`|` 与 `‖` 刻意不加入拆分集**。它们本来就由 `latex2mathml` 正常产出 `m:d`；
  而 `|x|+|y|` 这类式子里存在中间的 `|` run，加入后可能被误配对成一个 `m:d`。
  （`\mid` 产出的是 U+2223 `∣`，与 U+007C `|` 不是同一字符，互不影响。）
- **非成对写法不得被误配**：`(a,b]`、`[a,b)` 的 `(`/`[` 找不到对应右括号，
  保持原样（配对只认 `_STRETCH_PAIRS` 里登记的同型右括号）。
- `m:d` 省略 `begChr`/`endChr` 时 OMML 默认就是圆括号，所以 `\left(…\right)` 与
  `pmatrix` 不写这两项；但 `_wrap_stretch_delims` 新包出来的 `m:d` **会显式写**。
- 空的定界符对（`⟨⟩`）直接删掉两端，避免死循环。
- 未配对的单个定界符保持原样，渲染为小号字符 —— 与 LaTeX 行为一致。

### 验证

渲染级验证：用 Word 导出 PDF，量测括号**墨迹高度**与**所在行墨迹高度**之比。

```bash
"$PY" scripts/render_pdf_preview.py --brackets out.pdf   # 逐行列出每个定界符的字形高度
```

> 注意：pymupdf 的 `char['bbox']` 报的是**字体行框**，拉伸后的括号仍报同一高度
> （实测 `(a/b)` 与 `(x+y)` 都是 11.1pt），**不能**用它判断伸缩。

**可靠的量测法：连通域标记。**

1. 以行 bbox 为基准开一个上下各留 30pt 的大窗；
2. 二值化后做 8-邻接连通域标记；
3. 只保留 bbox 与本行 y 区间相交的连通域（相邻行的墨迹是独立连通域，天然排除）；
4. 括号 = 包含括号字符 bbox 中心的那个连通域；
5. 比值 = 括号墨迹高度 / 本行全部墨迹高度。**≥ 0.85 为合格。**

两种常见**假阳性**（本会话踩过，一度误判为"括号有周期性接缝"）：

- 裁剪窗横向留 ±3pt → 把右侧矩阵单元格的数字计入每行墨迹宽度，
  于是波动周期恰好等于矩阵行数，看起来像"拼接括号的接缝"；
- 裁剪窗纵向留 ±26pt → 把相邻段落的余墨计入，中间出现空白段，
  把 `max(mid) - min(mid)` 撑大，也会被读成"接缝"。

**用连通域隔离本体后，实测我的 3×3 矩阵括号与样本的括号在高度、质心偏移、
墨宽、笔画宽度上完全一致** —— 当时以为的"接缝"纯属量测伪影。
真正的缺陷是上面 (b) 类：括号根本没被包成 `m:d`。

---

## 三、脚本已内置的自动修复（`preprocess_tex`）

### 1. 带下标的裸括号组会错位

`(A\times B)_{i,k}` 单独出现正常；**后面还有内容**时（`(A\times B)_{i,k}=\sum…`），
latex2mathml 会把右括号塞进下标基数。
自动修复：括号组后紧跟 `_`/`^` 的裸括号改写成 `\left(…\right)`（跳过 `\begin{}`/`\text{}` 等参数内括号）。

### 2. 间距命令被整段丢弃

`MML2OMML.XSL` 没有 `<mspace>` 的映射，`\,` `\:` `\;` `\quad` `\qquad` 在 OMML 里**完全消失**。
自动修复：换成 `\text{…}` 里的私用区标记 `\uE007`（还原时变 NBSP），宽度 1/1/2/4/8 个；`\!` 删除。
用标记而非直接写空格，是为了让"有意的间距"与"普通词间空格"在还原阶段能被区分处理。

### 3. `align` / `aligned` 会凭空插入公式编号

实测转出的 OMML 是 3 列矩阵，第 3 列是 `(1)`、`(2)` 编号文本；`aligned` 还可能产出非法 MathML。
自动修复：统一改写为 `array{rl}` —— 保留 `&` 对齐、不编号。

### 4. 不支持的命令

| 写法 | 处理 |
|---|---|
| `\big` `\Big` `\bigg` `\Bigg`（含 `l`/`r`/`m` 变体） | **空操作**：`<mo minsize/maxsize>` 被 `MML2OMML.XSL` 忽略，尺寸只由 `m:d` 内容高度决定。且 `\big\langle` 会被 `latex2mathml` 输出成**字面文本** `\langle`（实测）→ 预处理直接删除 |
| 自定义宏 / `\newcommand` | 不支持，先展开 |
| 依赖宏包的命令（`\bm`、`\nicefrac` …） | 不支持，换等价基础写法 |
| TikZ / 图形 | 不支持，另存图片插入 |

> 删除 `\big` 系列用 `\\(?:big|Big|bigg|Bigg)(?:l|r|m)?(?![A-Za-z])` ——
> `(?![A-Za-z])` 是必需的，否则会误伤 `\bigcup` `\bigcap` `\bigoplus` `\bigotimes`
> `\bigvee` `\bigwedge` `\bigsqcup` 等大运算符命令（自检里有 7 例守这条线）。

### 5. 函数名必须正体（`m:sty val="p"`）

样本里 `det` / `dim` 的 run 都带 `<m:sty m:val="p"/>`（正体，`m:nor` 为 False）。
`latex2mathml` 对**多数**函数名输出 `<mi>`，XSL 会给 `sty=p`；
但下面这几个输出的是 `<mo>`，XSL **不给** `sty`，于是渲染成**斜体**：

```
det  lim  limsup  liminf  sup  inf  min  max  gcd  Pr
```

已实测**本来正体**（不需要动）：`sin` `cos` `tan` `cot` `sec` `arcsin` `arccos`
`arctan` `sinh` `cosh` `tanh` `coth` `log` `ln` `lg` `exp` `dim` `ker` `hom` `arg` `deg`。

自动修复：`preprocess_tex` 把它们改写成 `\mathrm{…}`（实测 `\mathrm{det}` → `sty=p`）。
`\operatorname{tr}` / `\operatorname*{…}` 同样输出 `<mo>`，一并改写成 `\mathrm{…}`。

```python
_OPNAME = re.compile(r'\\(det|limsup|liminf|lim|sup|inf|min|max|gcd|Pr)(?![A-Za-z])')
_OPERATORNAME = re.compile(r'\\operatorname\*?\{([^{}]*)\}')
```

- `(?![A-Za-z])` 必需：否则 `\sup` 会咬住 `\supset`、`\inf` 会咬住 `\infty`。
- 交替顺序 `limsup|liminf|lim` 必需：正则交替是"先匹配先赢"。
- `\lim_{…}` 在本链路本来就是 `m:sSub`（不是 `m:limLow`），
  换成 `\mathrm{lim}` **不改变下标位置**，安全。
- `\mathrm{…}` 不会打断后面的 `(` 合并（`\mathrm{det}(` 仍是一个 run），
  所以 (b) 类的 `_split_delim_runs` 依然生效。

**副作用与校验**：`selftest_formulas.py` 的"函数名正体自检"覆盖
10 个已修正 / 14 个原本正体未破坏 / 9 个易误伤写法未改写 / 3 个 `\operatorname` 已改写。

---

## 四、已实测通过的写法

分式、上下标、求和 / 连乘 / 积分（含 ±∞ 限）、根号、希腊字母全集、
`\mathbb` `\mathcal` `\mathrm` `\operatorname`、`\text{}`、极限、箭头
（`\Rightarrow` `\Leftrightarrow` `\mapsto` `\to`）、`\overline` `\hat` `\tilde`、
`\langle\rangle`（自动包 `m:d` 伸缩）、`\|`、偏导、`\pmod`、集合 `\{\}`、`\emptyset`、
`pmatrix` / `bmatrix` / `vmatrix` / `cases` / `array`、嵌套分式与矩阵、
`\left…\right`、行变换记号 `r_i \to r_i + \lambda r_j`。

**多行推导**：用 `\begin{array}{rl}…\end{array}`（左右两列，对齐在 `=`），
或 `\begin{array}{c}…\end{array}`（转成 Word 的 `m:eqArr`）。不要用 `align` / `aligned`。

---

## 五、写作约定

- 行内公式 `$...$`，独立公式独占一行 `$$...$$`。跨行的 `$$` 块（单独一行 `$$` 起止）也支持。
- **换行符必须写成两个反斜杠** `\\`。用文件写入工具写 md；
  若通过 shell 传参，bash 会吃掉一层转义，矩阵会塌成一行。
- 公式失败时退化为 `[公式失败: …]` 文本并在 stderr 打印，**不会静默丢内容**。

## 六、校验

```bash
PY=C:/Users/Admin/.workbuddy-ai/binaries/python/envs/default/Scripts/python.exe
"$PY" scripts/verify_docx.py 输出.docx      # 文档级：公式数、段落分类、NBSP、对齐、字号、间距
"$PY" scripts/selftest_formulas.py          # 链路级：57 公式 + 7 整行 + 5 分行 + 22 定界符 + 空格/正体/混排
```

`verify_docx.py` 关键指标：

- **`裸行内公式段落` 必须为 0** —— 段落里只有公式没有文字时会被 Word 默认居中
- **`正文行(文字+行内公式)` 应 > 0** —— 这是含公式正文的正确状态
- **`正文 run 含 NBSP` 必须为 0** —— 仅 `\quad` 处允许；否则 Word 无法断行
- `纯公式行对齐` 应同时出现 `center`（独立公式）与 `left`（推导）
- Heading 字号 30 / 28（15 / 14pt）、加粗、颜色 `000000`

## 七、导出 PDF

Word 的真实排版可由本机 Word 引擎导出为 PDF（第三方在线预览对 `m:nor` 支持不完整）：

```powershell
powershell -ExecutionPolicy Bypass -File scripts/docx_to_pdf.ps1 -Src "note.docx"
```

> **踩坑**：Word COM 会把正斜杠当 URL 分隔符，`F:/AI/x.docx` 会变成 `F:\//AI/x.docx`
> 而报"找不到您的文件"。脚本内部已把 `/` 统一换成 `\`；手工内联调用时务必自己转。

想反向核对渲染效果，可把 PDF 逐页渲染成图再目视检查：

```bash
"$PY" scripts/render_pdf_preview.py note.pdf pages 300        # 逐页 PNG（第 3 参数是 dpi）
"$PY" scripts/render_pdf_preview.py --lines note.pdf          # 每行 bbox + 字体 flags（判对齐 / 斜体）
"$PY" scripts/render_pdf_preview.py --brackets note.pdf       # 每个定界符的字形高度（判伸缩，需目视复核）
```
