"""Markdown(+LaTeX) -> 符合模板规范的 DOCX 课程笔记（浓缩教材版）。

用法:
    python build_note.py <input.md> <output.docx> [display_jc]
    display_jc: left(默认，与样本一致) | center | right

Markdown 约定:
    # Chapter N 标题                  -> Heading 1（章，15pt 加粗）
    ## N.M 小节                       -> Heading 2（节，14pt 加粗，深蓝）
    ##### N.M.K 三级标题               -> Heading 3（16pt 加粗，深蓝；仅 ITVC 样本有）
    #### 次级小标题（无编号）           -> 12pt 加粗正文（黑）
    ### Definition 1.1 名称            -> 条目标题（12pt；编号加粗、名称常规）
    普通段落
    - 列表项                           -> 项目符号列表（左缩进 0.306"、符号悬挂；行距同正文）
                                          以 (a)/(1)/(i) 开头的项自带标号，不再另加符号
      - 缩进 2 空格的列表项             -> 二级列表（符号 '-'、10pt、缩进 44pt）
    ~~删除线~~                        -> 删除线（公式也支持）
    *注释（星号紧跟文字，无空格）        -> 补充说明段落（星号保留）
    **粗体**                          -> 行内加粗
    $...$                             -> 行内公式
    独占一行的 $$...$$                 -> 独立公式（oMathPara 块级公式）
    | a | b | + |---|---| 表格         -> Word 表格（单元格内支持 $...$）
    Proof. ...                        -> 证明段（字号由 CFG['proof'] 控制）

核心排版决策（全部据样本 Linear Algebra Outline.docx 实测）:

1. **含公式的行，文字与公式同字号**。样本作者的做法是把整行文字都写进公式对象、
   再用引号把文字变回普通文本（`<m:nor>` 出现 1102 次）。本脚本对**含正文文字的行**
   改为等价的"混排段落"：文字走普通 `w:r`、公式走行内 `m:oMath`。
   实测两种写法的渲染字号完全一致（正文 11.04pt / 下标 8.04pt），
   但 Word 对数学段落（`m:oMathPara`）的分行会强加缩进 —— 会把正文行折得
   过早（右侧空 186pt）且续行缩进 72pt；混排段落的折行与普通文字完全一致。
   （开关 `CFG['prose_carrier']`，`'mixed'` 默认；`'inline_formula'` 回到样本写法。）

2. **纯公式段落必须用 m:oMathPara 承载，并显式写 m:jc**。
   裸行内 m:oMath 在 Word 里会按"数学段落"默认对齐（center）处理，
   导致同一份文档里有的行靠左、有的行莫名居中、缩进乱七八糟 —— 实测踩坑。
   样本实测（按段落分类统计 oMathPara 的 m:jc）:
       含正文文字的段落: left 138 / center 48 / 无 107   -> 左对齐
       长正文段落      : left  92 / center 22 / 无  36   -> 左对齐
       纯公式段落      : center 83 / left  20 / right 3  -> 居中
   故本脚本规则：**含文字的段落 left，纯公式段落 center，多行推导 left**。

3. **正文里的空格必须是普通空格，不能是 NBSP**。
   latex2mathml 会把 \\text{} 内的普通空格转成 U+00A0（不换行空格），
   样本 1101 个正文 run 里 NBSP 数量为 0 —— 全是普通空格。
   NBSP 会让 Word 无法在词间断行，整行只能挤在数学运算符处折断，排版塌成一团。
   本脚本自动把 NBSP 换回普通空格；\\quad 等显式间距另用标记保留为 NBSP。

4. **字号 / 颜色 / 间距（7 份样本实测，取多数值）**:
   章 15pt(sz=30,heading1) / 节 14pt(sz=28,heading2) / 条目标题 12pt(sz=24) /
   正文 11pt(sz=22,docDefaults) / 证明 11pt。标题一律加粗。
   章、节标题颜色 = **#0F4761 深蓝**（6 份样本 PDF 实测 5 份如此，仅 ODEs 为黑）；
   条目标题与正文为黑。
   间距：段后 8pt(docDefaults)；章 before 24pt / after 4pt，节 before 8pt / after 4pt，
   条目 after 8pt。行距 278/auto ≈1.158（正文与列表同）。

   注意（易误判为缺陷）：**Cambria Math 没有 Bold 字面**，`<w:b/>` 由 Word 走
   "合成加粗"（描边加粗）。导出的 PDF 里因此**不会出现独立粗体字体**，用 pymupdf
   按字体名或 flags 位检测会一律报"非粗体"。目视是加粗的即为正确。

5. **分行**：一个逻辑单元 = 一个段落。条目标题独占一段，每条陈述 / 公式 / 列表项 /
   证明步骤各占一段。公式需要分行时另起一段独立公式，不要让 Word 自动折行。

6. **公式大小（行内 vs 独立）按"整行放不放得下"决定，而不是按"有没有 ∑"**。
   行内公式（`$...$`）默认保持行内；只有当**整行估算宽度超过正文列宽**时，
   才把行内的大运算符（∑ ∏ ∫ lim …）拆成独立公式。
   样本实证：`(b) Sum up to one, i.e. Σ_{ω∈Ω} ℙ({ω}) = 1` 整行仅 176pt（列宽 415pt），
   **作者就是保持行内小公式**，看着干净；只有行放不下时才另起大公式。
   估算用 PIL + 本机 Cambria Math 真实字形度量（实测与 Word 渲染偏差 < 5%）：
      文字段 —— 逐字符 getlength；公式段 —— LaTeX 转近似 Unicode 后度量，
      并对二元运算符补两侧空白、对大运算符放大。宁大勿小（避免误判"放得下"）。

7. **块与块之间要留白，但留白位置要对**。手打 Outline（ODEs Outline.pdf 逐行实测）
   的留白节律：标题(章/节) → 紧接首段 ≈ 19pt（标题与"它领起的内容"紧凑）；
   条目 → 自身正文 ≈ 14pt；正文→正文 ≈ 13pt；**只有"正文块末 → 下一条目"才 ≈ 40pt**
   （大留白只出现在真正的话题切换处）。
   本脚本用 `space_before` 实现，并按"上一段块类型"决定本段段前
   （`before_entry`/`before_sub` 用于正文块末，`*_after_heading` 用于紧跟标题），
   避免空段落污染文档结构，也避免"节标题被 40pt 大空隙孤立"或"标题挤进正文"。
"""
import copy
import re
import sys
from pathlib import Path

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls, qn
from docx.shared import Inches, Pt, RGBColor
from lxml import etree
from latex2mathml.converter import convert

# ---------------------------------------------------------------- 配置区
# 全部取自样本实测值。想调层级观感只改这里。
CFG = dict(
    font='Cambria Math',
    h1=15,          # 章标题 pt（样本 heading1 sz=30）
    h2=14,          # 节标题 pt（样本 heading2 sz=28）
    entry=12,       # 条目标题 pt（样本 sz=24）
    sub=12,         # 无编号次级小标题 pt（样本 sz=24，黑色）
    sub3=16,        # 三级编号标题 N.M.K pt（仅 ITVC 用 heading 3，sz=32，深蓝）
    body=11,        # 正文/公式 pt（样本 docDefaults sz=22）
    proof=11,       # 证明正文 pt（样本多为 11；个别 10pt -> 改 10 即可）
    # ---- 对齐（样本实测：含文字的段落 left，纯公式段落 center）----
    prose_jc='left',      # 含正文文字的公式段落（"Let K be a field, ..." 这类）
    display_jc='center',  # 独占一行的纯公式
    derive_jc='left',     # 多行推导（array / aligned，按 = 对齐的步骤）
    # 含公式的段落一律用 m:oMathPara 承载并显式写 m:jc。
    # 'inline' 会退化成裸 m:oMath —— Word 会按数学段落默认居中，排版不可控，仅作兼容开关。
    math_line_mode='para',
    # 含正文文字的行怎么承载：
    #   'mixed'        -> 文字段用普通 w:r、数学段用行内 m:oMath（推荐）
    #   'inline_formula' -> 整行并入一个 m:oMathPara（样本作者的做法）
    # 实测：'mixed' 与 'inline_formula' 渲染出的字号完全一致（正文 11.04pt / 下标 8.04pt），
    # 但 Word 对数学段落的分行会强加缩进 —— 用 'inline_formula' 时正文行会"折行过早 +
    # 续行缩进 72pt"（看起来像居中），而 'mixed' 的折行与普通文字完全相同。
    prose_carrier='mixed',
    # ---- 行距（7 份样本 docDefaults 实测一致：after=160twips, line=278/auto）----
    # 278/240 = 1.1583，Word 写出 w:line="278"，与样本逐位一致（写 1.15 会得到 276）。
    line_body=278 / 240,
    # 样本 5/5 份的列表段落以 278/auto 为主（并非 1.5 倍）；列表与正文同行距。
    line_list=278 / 240,
    after_body=8,
    # ---- 标题间距（7 份样本 heading1/2 样式实测一致）----
    # heading1: before=480twips(24pt) after=80twips(4pt)
    before_h1=24, after_h1=4,
    # heading2: before=160twips(8pt) after=80twips(4pt)
    before_h2=8, after_h2=4,
    # heading3（三级编号标题）：样本 before=8pt
    before_sub3=8,
    # === v13：正文上方标题留白 ===
    # 用户要求：小标题/大标题若正上方是小字正文，应留"一格"空，大小按标题自身
    # 字号 × 行距（即一行标题高）。实现为 max(样式原段前, 标题字号 × line_body)，
    # 只补不削 —— 已够宽的标题不受影响。手打 Outline 实测 body→节标题 ≈20.3pt，
    # 而 14pt×1.158≈16.2pt 段前 → 渲染 gap ≈20.4pt，完全吻合。
    # 条目标题 after：样本 8pt，但手打 Outline 实测"条目 → 自身正文"≈14pt，
    # 故提到 12pt（正文段前=0，max(12,0)+行距≈14，对齐 Outline）。
    before_entry=36, after_entry=12,
    # 无编号次级小标题（#### ）段前：同样对齐样本的 ≈ 52pt 视觉间距（正文块末→小标题）。
    before_sub=36, after_sub=4,
    # === v10：标题后首段间距（对齐手打 ODEs Outline 逐行实测，见 _outline_gaps.txt）===
    # 手打 Outline 的留白节律（gap = 上段底→本段顶）：
    #   标题(章/节) → 紧接首段(条目/正文/分点) ≈ 19pt  —— 标题与"它领起的那段"紧凑；
    #   条目 → 自身正文 ≈ 14pt；正文→正文 ≈ 13pt；
    #   正文块末 → 下一条目 ≈ 40pt  —— 大留白只出现在"真正的话题切换"处。
    # v1 把 before_entry=36 套到"节标题紧接的条目"上，节标题与条目间出现 40pt 大空隙
    # （标题像被孤立）；而"节标题 → 正文首段"又只有 ~8pt（标题挤进正文）。两处都违和。
    # 故按"上一段类型"决定本段段前：上一段是标题时收紧到 ≈19，是正文块末时用 40。
    # Word 对相邻段落取 max(space_after, space_before)，故下方数值即为"主导间距"。
    before_entry_after_heading=14,   # 条目紧跟章/节/三级标题：≈19 总间距
    before_sub_after_heading=14,     # 次级小标题紧跟标题
    before_body_after_heading=15,    # 正文首段紧跟标题：≈19 总间距
    before_list_after_heading=12,    # 分点块首项（带符号）紧跟标题
    before_label_item_after_heading=14,  # 带 (a)/(1)/(i) 分点首段紧跟标题
    # （deprecated）旧逻辑只区分"章标题后"，现由上面的 after_heading 系列取代：
    before_entry_first=10,
    # ---- 公式大小（行内小公式 vs 独立大公式）----
    # 行内公式默认保持行内；只有当"整行估算宽度 > 列宽 × split_ratio"时才拆成独立公式。
    # 这样 `(b) Sum up to one, i.e. Σ_{ω∈Ω} ℙ({ω}) = 1`（全行 176pt ≪ 列宽 415pt）
    # 保持行内小公式（与样本一致），只有真的放不下才另起大公式。
    split_ratio=0.97,
    # ---- 二级列表（样本：一级 * / 二级 -，二级字号 10pt）----
    sub_bullet=10,          # 二级列表项字号 pt（样本 sz=20）
    indent_list2=880 / 1440,    # 二级缩进 44pt（样本文字 x0=134）
    hanging_list2=440 / 1440,
    # 章/节标题颜色：6 份样本 PDF 实测，5 份为 #0F4761（深蓝），仅 ODEs 为纯黑。
    # 改 '000000' 即回到纯黑。
    heading_color='0F4761',
    # 列表缩进：5 份样本实测 —— 文字左缩进 440twips(0.306")、项目符号悬挂 440twips
    # （符号贴在左边距，文字缩进 22pt）。此前写的 0.5" 是错的：那是 python-docx
    # 内置 "List Paragraph" 样式的默认值，而样本根本没用该样式
    # （Prob I / LA 的 ListParagraph 段落数都是 0），用的是段落级 numPr + ind left=440。
    indent_list=440 / 1440,
    hanging_list=440 / 1440,
    upright_in_text=True,   # 给公式内正文 run 补 <w:i w:val="0"/>，增强第三方渲染器兼容
    split_big_ops=True,     # 正文行内出现 ∑/∏/∫/lim 等大运算符时，自动拆成独立公式行
    # ---- 列表块间距（v9，按手打提纲 Prob I Ch1 逐行实测校准）----
    # 提纲实测：条目/小标题 → 列表块首项 25.1pt；列表项之间 16.1~17.4pt；
    # (a)/(b) 并列分点之间 26.4pt（普通项的 1.6 倍）；列表块 → 下一块 51.7pt。
    # 生成侧原来一律用 after_body(8)，导致"分点块"上下都糊在正文里、看不出分点。
    before_list=6,          # 列表块首项额外段前（给"引导句 → 分点"留出边界）
    after_list_item=4,      # 列表项之间的段后（收紧，让分点块内部紧凑）
    # (a)/(1)/(i) 并列分点自成"两档"：提纲实测块上边界 26.6pt、项间 26.4pt，
    # 都比普通列表（25.1 / 18.0）大一档 —— 并列分点要显得彼此独立。
    before_label_item=10,   # 带标号分点块首项的段前
    after_label_item=12,    # (a)/(1)/(i) 并列分点之间的段后（比普通项大一档）
    after_list_block=18,    # 列表块末项段后（块下边界；后接标题类时自动降为 after_list_item）
    # ---- 项目符号（可在库里挑，不必只用实心点）----
    # 样本实测一级符号有三种：● (ITVC / LA / Stats I / ODEs)、◆ (MFA / RA)、
    # - (Prob I / ODEs)。二级一律 '-'（10pt，见 sub_bullet）。
    #
    # 'auto' = 先按**列表体量**定档（长列表用视觉重的符号），再按源文件名确定性挑一个
    #   （同一份 md 每次一致，不同科目自然不同）。想固定就直接写符号，如 '◆'。
    #
    # 池子按"墨迹高度 / 正文 x-height"分档（PIL 实测，见 _bullet_ink_ratio）：
    #   重  1.2~1.5×：● ◆ ■ ▲ ◇ □ ○   —— 项多/项长时用，压得住
    #   中  0.45~0.8×：▸ * • ◦ ▪
    #   轻  0.14~0.27×：· ∙ – — −      —— 只适合短列表（样本 Prob I 用的就是这类）
    #
    # ⚠ 已实测剔除的**豆腐块** —— Cambria Math 没有这三个字形，PIL 渲染出的
    #   bbox 完全相同（0.480×0.665、填充率 0.44）= 同一个 .notdef 方框：
    #       ⬤(U+2B24)  ‣(U+2023)  ⁃(U+2043)
    # ⚠ ▬(U+25AC) 宽 0.82em、高 0.565em —— 过宽，视觉突兀，不用。
    #   ⇒ 池子想再加符号，先跑一次墨迹实测确认有字形，别盲加。
    bullet_char='auto',
    # 重档只放**实心**符号：空心（◇ □ ○）虽然高大，但填充率低（0.24~0.46），
    #   满页文字时视觉重量不够，压不住。
    bullet_pool_heavy=['●', '◆', '■', '▲'],
    bullet_pool_light=['▸', '*', '•', '◦', '▪', '□', '○', '◇', '·', '–'],
    # 列表体量 ≥ 此值时用重档符号（"内容太多符号太小"就用不上轻符号）
    bullet_volume_heavy=6.0,
    # === v15：目录 + 页码（对齐手打 Outline 中带目录的 3 份样本实测）===
    # 7 份样本里 3 份有目录，实测：
    #   MFA / Prob I : TOC 域 ` TOC \o "1-2" \h \z \u ` + TOC1/TOC2 样式
    #                  + 页脚居中 PAGE / NUMPAGES（渲染成 " 1 / 46 "）
    #   LA           : `\o "1-3"`（收到条目层，目录很长）
    #   目录后另起一页：Prob I 目录 p1–p2、正文 p3；LA 目录 p1–p3、正文 p4
    #   页码全文连续编号（目录页也算第 1 页，目录里的页码即该编号）
    #   TOC1 基于 Normal（11pt，无额外格式）；TOC2 再加 ind left=420twips
    toc=True,                   # 生成目录（Word 原生 TOC 域，可按 F9 重新生成）
    toc_levels='1-2',           # 收录层级：'1-2' = 章 + 节（样本 MFA / Prob I）
    toc_title='Contents',       # 目录标题；置 None 或 '' 则不加标题行
    toc_title_size=16,          # 目录标题字号（LA 的 "TOC Heading" 实测 sz=32）
    toc_title_after=18,         # 目录标题段后 pt（把标题与首条拉开）
    toc_break=True,             # 目录后另起一页（样本实测）
    # --- 目录条目外观 ---
    # 样本的目录很朴素（TOC1 继承 Normal = 11pt 不加粗、TOC2 只多一个缩进）。
    # 这里默认做得更"立得住"：章条目加粗 + 章标题同色深蓝 + 稍大字号 +
    # 组前留白，让每一章成为视觉分组；节条目保持黑色小字。
    # 想回到样本的素净样子：toc_bold1=False、toc_color1=None、toc_entry1=11、
    # toc_before1=0、toc_after1=0。
    toc_entry=11,               # TOC2（节）字号 pt
    toc_entry1=12,              # TOC1（章）字号 pt
    toc_bold1=True,             # TOC1 加粗
    toc_color1='0F4761',        # TOC1 颜色（章标题同色深蓝）；None -> 正文黑
    toc_before1=10,             # TOC1 段前 pt（章组之间的留白）
    toc_after1=2,               # TOC1 段后 pt
    toc_before2=0,              # TOC2 段前 pt
    toc_after2=2,               # TOC2 段后 pt
    toc_indent2=420,            # TOC2 左缩进 twips（样本 TOC2 ind left=420）
    page_numbers=True,          # 页脚页码
    page_number_format='both',  # 'both' -> "X / Y"（样本）；'page' -> 仅 "X"
    page_number_size=10,        # 页脚字号
    page_number_gap=992,        # 页脚距下边距 twips（样本 pgMar footer=992）
)

XSL_PATH = r'C:/Program Files/Microsoft Office/root/Office16/MML2OMML.XSL'
MML_NS = 'http://www.w3.org/1998/Math/MathML'
W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
M = 'http://schemas.openxmlformats.org/officeDocument/2006/math'
XMLNS = 'http://www.w3.org/XML/1998/namespace'
FONT = CFG['font']

# 占位符：latex2mathml 的 \text{} 不做转义，& < > 会产出非法 XML，{ } 会被当分组。
# 先用私用区字符替换，转换回写时再还原。
_PH = {'&': '\uE002', '<': '\uE003', '>': '\uE004', '{': '\uE005', '}': '\uE006'}
_PH_BACK = {v: k for k, v in _PH.items()}
B_ON, B_OFF = '\uE000', '\uE001'    # 公式内加粗标记（\text{} 会原样透传，可跨 run 合并后仍可定位）
_NB = '\uE007'                      # 显式间距标记：还原成 NBSP（真正不可断行的空格）
_NB_TMP = '\uE008'                  # 归一过程中的中转字符

_XSLT = None
_ERR = 0


def _xslt():
    global _XSLT
    if _XSLT is None:
        _XSLT = etree.XSLT(etree.parse(XSL_PATH))
    return _XSLT


# ---------------------------------------------------------------- 行宽估算
# 用途：判断"行内公式能不能跟文字同行放下" —— 放得下就保持行内小公式，
# 放不下才把大运算符拆成独立公式。用本机 Cambria Math 的真实字形度量，
# 实测与 Word 渲染偏差 < 5%（文字段 0.98~1.05，公式段 0.98~1.08）。
_FONT_FILE = r'C:\Windows\Fonts\cambria.ttc'   # index=1 = Cambria Math
_FONT_INDEX = 1
_FONT_CACHE = {}
_PIL_OK = None


def _font(size_pt):
    """取 Cambria Math 的 PIL 字体对象；不可用时返回 None。"""
    global _PIL_OK
    if _PIL_OK is False:
        return None
    key = round(float(size_pt), 2)
    if key in _FONT_CACHE:
        return _FONT_CACHE[key]
    try:
        from PIL import ImageFont
        ft = ImageFont.truetype(_FONT_FILE, key, index=_FONT_INDEX)
        _PIL_OK = True
    except Exception:
        _PIL_OK = False
        return None
    _FONT_CACHE[key] = ft
    return ft


# LaTeX -> 近似 Unicode（用于宽度估算）。丢掉上下标标记 => 估算偏宽（保守）。
_GREEK = {'alpha': 'α', 'beta': 'β', 'gamma': 'γ', 'delta': 'δ', 'epsilon': 'ε',
          'varepsilon': 'ε', 'zeta': 'ζ', 'eta': 'η', 'theta': 'θ', 'vartheta': 'ϑ',
          'iota': 'ι', 'kappa': 'κ', 'lambda': 'λ', 'mu': 'μ', 'nu': 'ν', 'xi': 'ξ',
          'pi': 'π', 'rho': 'ρ', 'sigma': 'σ', 'tau': 'τ', 'upsilon': 'υ',
          'phi': 'φ', 'varphi': 'φ', 'chi': 'χ', 'psi': 'ψ', 'omega': 'ω',
          'Gamma': 'Γ', 'Delta': 'Δ', 'Theta': 'Θ', 'Lambda': 'Λ', 'Xi': 'Ξ',
          'Pi': 'Π', 'Sigma': 'Σ', 'Upsilon': 'Υ', 'Phi': 'Φ', 'Psi': 'Ψ',
          'Omega': 'Ω'}
_MATH_SYMS = {'sum': '∑', 'prod': '∏', 'int': '∫', 'oint': '∮', 'infty': '∞',
              'partial': '∂', 'in': '∈', 'notin': '∉', 'subset': '⊂',
              'subseteq': '⊆', 'supset': '⊃', 'supseteq': '⊇', 'cup': '∪',
              'cap': '∩', 'emptyset': '∅', 'varnothing': '∅', 'setminus': '\\',
              'to': '→', 'rightarrow': '→', 'leftarrow': '←', 'mapsto': '↦',
              'Rightarrow': '⇒', 'leq': '≤', 'geq': '≥', 'le': '≤', 'ge': '≥',
              'neq': '≠', 'approx': '≈', 'equiv': '≡', 'times': '×',
              'cdot': '⋅', 'div': '÷', 'pm': '±', 'mp': '∓',
              'ldots': '…', 'cdots': '⋯', 'dots': '…', 'vdots': '⋮', 'ddots': '⋱',
              'forall': '∀', 'exists': '∃', 'neg': '¬', 'land': '∧', 'lor': '∨',
              'circ': '∘', 'langle': '⟨', 'rangle': '⟩',
              'lfloor': '⌊', 'rfloor': '⌋', 'lceil': '⌈', 'rceil': '⌉'}
_BB = {'P': 'ℙ', 'R': 'ℝ', 'N': 'ℕ', 'Z': 'ℤ', 'Q': 'ℚ', 'C': 'ℂ', 'E': '𝔼',
       'F': '𝔽', '1': '𝟙'}
_CAL = {'F': 'ℱ', 'L': 'ℒ', 'A': '𝒜', 'B': 'ℬ', 'D': '𝒟', 'P': '𝒫', 'S': '𝒮',
        'H': 'ℋ', 'M': 'ℳ', 'E': 'ℰ'}
_WRAP_CMDS = ('mathbb', 'mathcal', 'mathrm', 'mathbf', 'mathit', 'operatorname',
              'text', 'textrm', 'textbf', 'textit', 'mbox', 'boldsymbol')
# 两侧会带空白的二元/关系运算符
_WIDE_OPS = set('=+-<>≈≠≤≥∈∉⊆⊂⊃∪∩→←↦⇒×⋅±')


def _math_unicode(tex):
    out, i, n = [], 0, len(tex)
    while i < n:
        c = tex[i]
        if c == '\\':
            m = re.match(r'\\([A-Za-z]+)', tex[i:])
            if m:
                cmd = m.group(1)
                i += m.end()
                if cmd in _WRAP_CMDS:
                    if i < n and tex[i] == '{':
                        j = _skip_group(tex, i)
                        inner = tex[i + 1:j - 1]
                        i = j
                        if cmd == 'mathbb':
                            out.append(''.join(_BB.get(x, x) for x in inner))
                        elif cmd == 'mathcal':
                            out.append(''.join(_CAL.get(x, x) for x in inner))
                        else:
                            out.append(inner)
                    continue
                out.append(_GREEK.get(cmd) or _MATH_SYMS.get(cmd) or 'M')
                continue
            out.append(' ')          # \, \; \: \! 等间距命令
            i += 2
            continue
        if c in '{}':
            i += 1
            continue
        if c in ' \t':
            i += 1
            continue
        if c in '_^':
            i += 1                  # 上下标：宽度并入基数（保守），只消费标记
            continue
        out.append(c)
        i += 1
    return ''.join(out)


def _est_math_width(tex, size, op_pad=0.24, bigop=1.30):
    ft = _font(size)
    if ft is None:
        return len(_math_unicode(tex)) * 0.62 * size
    w = 0.0
    for ch in _math_unicode(tex):
        cw = ft.getlength(ch)
        if ch in _WIDE_OPS:
            cw += op_pad * 2 * size
        if ch in '∑∏∐∫∮':
            cw *= bigop
        w += cw
    return w


def _est_text_width(text, size, scale=1.04):
    t = re.sub(r'\*\*|~~', '', text)
    ft = _font(size)
    if ft is None:
        return len(t) * 0.58 * size
    return ft.getlength(t) * scale


def _line_width_pt(text, size):
    """估算"文字 + 行内公式"整行宽度（pt）。"""
    w = 0.0
    for i, seg in enumerate(re.split(r'(\$[^$]+\$)', text)):
        if not seg:
            continue
        w += _est_math_width(seg[1:-1], size) if i % 2 == 1 \
            else _est_text_width(seg, size)
    return w


def _column_width_pt(doc):
    sec = doc.sections[0]
    return float(sec.page_width - sec.left_margin - sec.right_margin) / 12700.0


# ---------------------------------------------------------------- LaTeX 预处理
_LDELIMS = ('\\left', '\\bigl', '\\Bigl', '\\biggl', '\\Biggl',
            '\\big', '\\Big', '\\bigg', '\\Bigg')
_OPAQUE = ('\\begin', '\\end', '\\text', '\\textrm', '\\mathrm', '\\mathbf',
           '\\mathit', '\\mathbb', '\\mathcal', '\\operatorname', '\\mbox',
           '\\textbf', '\\textit')


def _skip_group(tex, i):
    depth, n = 0, len(tex)
    while i < n:
        if tex[i] == '\\':
            i += 2
            continue
        if tex[i] == '{':
            depth += 1
        elif tex[i] == '}':
            depth -= 1
            if depth == 0:
                return i + 1
        i += 1
    return n


def wrap_subscript_parens(tex):
    """把"带下/上标的裸括号组"改成 \\left(...\\right)。

    latex2mathml 对 `(A\\times B)_{i,k}` 这种写法，只有当它是整个表达式时才对；
    一旦后面还有内容（如 `(A\\times B)_{i,k}=\\sum...`），产出的 MathML 会把右括号
    塞进下标基数里 —— 括号跑到下标位置，就是"公式错位"。
    """
    stack, marks = [], {}
    i, n = 0, len(tex)
    while i < n:
        ch = tex[i]
        if ch == '\\':
            cmd = next((c for c in _LDELIMS if tex.startswith(c, i)), None)
            if cmd:
                j = i + len(cmd)
                if j < n and tex[j] in '()[]|.':
                    j += 1
                (stack.append(('managed', i)) if cmd == '\\left'
                 else (stack.pop() if stack else None))
                i = j
                continue
            op = next((c for c in _OPAQUE if tex.startswith(c, i)), None)
            if op:
                j = i + len(op)
                while j < n and tex[j].isalpha():
                    j += 1
                if j < n and tex[j] == '{':
                    j = _skip_group(tex, j)
                i = j
                continue
            i += 2
            continue
        if ch == '(':
            stack.append(('bare', i))
        elif ch == ')':
            if stack:
                kind, idx = stack.pop()
                j = i + 1
                while j < n and tex[j] == ' ':
                    j += 1
                if kind == 'bare' and j < n and tex[j] in '_^':
                    marks[idx] = 'L'
                    marks[i] = 'R'
        i += 1
    if not marks:
        return tex
    out = []
    for k, c in enumerate(tex):
        if k in marks:
            out.append('\\left' if marks[k] == 'L' else '\\right')
        out.append(c)
    return ''.join(out)


# LaTeX 间距命令 -> 显式不换行空格。MML2OMML.XSL 没有 <mspace> 的映射，
# \, \; \: \quad \qquad 会被整段丢弃（实测）。用 _NB 标记，还原时再变回 NBSP。
_SPACE_MAP = [
    (r'\qquad', '\\text{%s}' % (_NB * 8)),
    (r'\quad', '\\text{%s}' % (_NB * 4)),
    (r'\;', '\\text{%s}' % (_NB * 2)),
    (r'\:', '\\text{%s}' % _NB),
    (r'\,', '\\text{%s}' % _NB),
    (r'\!', ''),
]


def _fix_spaces(t):
    for pat, rep in _SPACE_MAP:
        t = re.sub(r'(?<!\\)' + re.escape(pat), lambda m, r=rep: r, t)
    return t


# \big \Big \bigg \Bigg 及 l/r/m 变体，在本链路里是**空操作**：
# MML2OMML.XSL 忽略 <mo minsize/maxsize>，括号尺寸只由 m:d 的内容高度决定。
# 更糟的是 \big\langle 会被 latex2mathml 原样输出成**字面文本** "\langle"（实测）。
# 故一律删除，让括号回到标准写法，交给 _stretch_delims 处理。
# 注意 (?![A-Za-z])：不能误伤 \bigcup \bigcap \bigoplus \bigotimes \bigvee \bigwedge …
_BIGFAM = re.compile(r'\\(?:big|Big|bigg|Bigg)(?:l|r|m)?(?![A-Za-z])')

# 函数名正体修复。
# latex2mathml 对多数函数名输出 <mi>（sin log exp ln dim ker arg deg …），
# MML2OMML.XSL 会带上 <m:sty m:val="p"/>（正体）；但下面这几个输出的是 <mo>，
# 于是渲染成**斜体** —— 实测样本里 det / dim 都带 sty=p，必须修正。
# 转成 \mathrm{} 即可拿到 sty=p（实测 \mathrm{det} -> sty=p）。
# 注意：\lim_{...} 在本链路本来就是 m:sSub（不是 m:limLow），
#       换成 \mathrm{lim} 不改变下标位置，安全。
# (?![A-Za-z]) 防误伤：\supset / \infty / \limsup …
_OPNAME = re.compile(r'\\(det|limsup|liminf|lim|sup|inf|min|max|gcd|Pr)'
                     r'(?![A-Za-z])')
# \operatorname{tr} / \operatorname*{tr} 同样输出 <mo>，一并转 \mathrm
_OPERATORNAME = re.compile(r'\\operatorname\*?\{([^{}]*)\}')


def preprocess_tex(tex):
    """修掉已知会让转换失败、错位或丢空格的写法。"""
    t = tex.strip()
    # align / aligned 会被当成"带编号"环境，OMML 里凭空多出 (1)(2)；aligned 还可能产出非法 MathML
    for env in ('align', 'aligned'):
        t = re.sub(r'\\begin\{%s\*?\}' % env, r'\\begin{array}{rl}', t)
        t = re.sub(r'\\end\{%s\*?\}' % env, r'\\end{array}', t)
    t = _BIGFAM.sub('', t)
    t = _OPERATORNAME.sub(r'\\mathrm{\1}', t)
    t = _OPNAME.sub(r'\\mathrm{\1}', t)
    t = wrap_subscript_parens(t)
    t = _fix_spaces(t)
    return t


def latex_to_omml(tex):
    mml = convert(preprocess_tex(tex))
    if 'xmlns' not in mml:
        mml = mml.replace('<math', f'<math xmlns="{MML_NS}"', 1)
    dom = etree.fromstring(mml.encode('utf-8'))
    omml = _xslt()(dom).getroot()
    _fix_empty_script_bases(omml)
    return omml


# 空底数的上/下标、极限、重音结构（如孤立的 `$'$` → m:sSup 空底数、`$\dot{}$` → m:limUpp 空底数）
# 会让 Word 给底数槽留出 ~7.6pt 的幻影宽度：表现为"prime 和 ′ 之间隔一大截"
# （实测 gap 10.1pt，正常空格仅 2.5pt）。
# 修复：往空底数里塞一个**零宽空格**(U+200B) —— 宽度归零，脚本的字号与右上角位置不变。
# （实测：sSup 空底数 → 撇号被推远 10.1pt；塞 ZWSP 后 gap ≈ 0，撇号仍在 8.04pt/上标位。）
_SCRIPT_BASE = {'sSup', 'sSub', 'sSubSup', 'limUpp', 'limLow', 'mover', 'munder'}


def _has_visible_text(el):
    for t in el.iter('{%s}t' % M):
        if t.text and t.text.strip():
            return True
    return False


def _fix_empty_script_bases(root):
    for el in root.iter():
        if etree.QName(el).localname not in _SCRIPT_BASE:
            continue
        e = el.find('{%s}e' % M)
        if e is None or _has_visible_text(e):
            continue
        for c in list(e):
            e.remove(c)
        r = etree.SubElement(e, '{%s}r' % M)
        etree.SubElement(r, '{%s}t' % M).text = '\u200b'   # ZWSP


# ---------------------------------------------------------------- 公式 run 修饰
def _style_math(root, half_pt):
    """给公式树里所有 run 补齐 rFonts + sz + szCs，并保留 m:t 空格。"""
    for r in root.iter('{%s}r' % M):
        rpr = r.find('{%s}rPr' % W)
        if rpr is None:
            rpr = etree.Element('{%s}rPr' % W)
            mpr = r.find('{%s}rPr' % M)
            r.insert(1 if mpr is not None else 0, rpr)
        _fill_rpr(rpr, half_pt)
    for rpr in root.iter('{%s}rPr' % W):
        _fill_rpr(rpr, half_pt)
    for t in root.iter('{%s}t' % M):
        t.set('{%s}space' % XMLNS, 'preserve')


def _fill_rpr(rpr, half_pt):
    rf = rpr.find('{%s}rFonts' % W)
    if rf is None:
        rf = etree.Element('{%s}rFonts' % W)
        rpr.insert(0, rf)
    for a in ('ascii', 'hAnsi', 'eastAsia', 'cs'):
        rf.set('{%s}%s' % (W, a), FONT)
    for tag in ('sz', 'szCs'):
        el = rpr.find('{%s}%s' % (W, tag))
        if el is None:
            el = etree.SubElement(rpr, '{%s}%s' % (W, tag))
        el.set('{%s}val' % W, str(half_pt))


def _mark_bold(run_el):
    rpr = run_el.find('{%s}rPr' % W)
    if rpr is None:
        rpr = etree.Element('{%s}rPr' % W)
        run_el.insert(1, rpr)
    if rpr.find('{%s}b' % W) is None:
        b = etree.Element('{%s}b' % W)
        rf = rpr.find('{%s}rFonts' % W)
        rpr.insert(1 if rf is not None else 0, b)


def _set_strike(run_el, on=True):
    """给 w:r 或 m:r 加删除线（w:strike）；子元素顺序由 _normalize_rpr_order 兜底。"""
    rpr = run_el.find('{%s}rPr' % W)
    if rpr is None:
        rpr = etree.Element('{%s}rPr' % W)
        mrpr = run_el.find('{%s}rPr' % M)
        if mrpr is not None:            # m:r 里的 w:rPr 必须排在 m:rPr 之后
            mrpr.addnext(rpr)
        else:
            run_el.insert(0, rpr)
    if on and rpr.find('{%s}strike' % W) is None:
        etree.SubElement(rpr, '{%s}strike' % W)
    return rpr


# w:rPr 子元素必须按 OOXML 规定顺序排列，否则严格校验器会报错。
_RPR_ORDER = ['rStyle', 'rFonts', 'b', 'bCs', 'i', 'iCs', 'caps', 'smallCaps',
              'strike', 'dstrike', 'outline', 'shadow', 'emboss', 'imprint',
              'noProof', 'snapToGrid', 'vanish', 'webHidden', 'color',
              'spacing', 'w', 'kern', 'position', 'sz', 'szCs', 'highlight',
              'u', 'effect', 'bdr', 'shd', 'fitText', 'vertAlign', 'rtl', 'cs']


def _normalize_rpr_order(root):
    """把公式树里所有 w:rPr 的子元素按 OOXML 规定顺序重排。"""
    idx = {n: i for i, n in enumerate(_RPR_ORDER)}
    for rpr in root.iter('{%s}rPr' % W):
        kids = list(rpr)
        kids.sort(key=lambda e: idx.get(etree.QName(e).localname, 999))
        for k in kids:
            rpr.append(k)


# 这些定界符 latex2mathml 会退化成普通文字 run（固定大小），必须手工包成 m:d 才能伸缩。
# 实测分两类：
#   (a) 尖括号/取整括号：\langle \rangle \lfloor \rfloor \lceil \rceil
#       —— latex2mathml 把 \langle 转成 <mi>（标识符）而不是 <mo fence>，
#          MML2OMML.XSL 于是把 "⟨u,v⟩" 合并成**一个** m:r / m:t。
#   (b) 函数名紧邻定界符：\det(...) \operatorname{tr}(...) \sin(x) f(x)
#       —— latex2mathml 把函数名和后面的 ( 合并进同一个 <m:t>，例如 "det("。
#          这时 ( 不是独立元素，配对会失败 —— 实测 det(pmatrix) 的括号完全不伸缩。
# 两类都必须先 _split_delim_runs 把字符拆出来，再 _wrap_stretch_delims 配对。
#
# 注意： | 与 ‖ 刻意**不**加入。它们本来就由 latex2mathml 正常产出 m:d；
#        而 |x|+|y| 这类式子里存在中间的 | run，加入后可能被误配对成 m:d。
#        \mid 产出的是 U+2223（∣），与 U+007C（|）不是同一字符，互不影响。
#        (...) [...] \{...\} pmatrix vmatrix cases 在常规写法下会正常产出 m:d，
#        这里加入 ( ) [ ] { } 只是为了兜住上面 (b) 类的合并情形。
_STRETCH_PAIRS = {'(': ')', '[': ']', '{': '}',
                  '⟨': '⟩', '⌊': '⌋', '⌈': '⌉', '⟮': '⟯'}
_STRETCH_CHARS = set(_STRETCH_PAIRS) | set(_STRETCH_PAIRS.values())
# 字符集里含正则元字符 ( ) [ ] { }，必须 re.escape
_STRETCH_CLS = re.escape(''.join(sorted(_STRETCH_CHARS)))
_STRETCH_RE = re.compile('[%s]|[^%s]+' % (_STRETCH_CLS, _STRETCH_CLS))


def _split_delim_runs(root):
    """把混在普通 run 里的可伸缩定界符字符拆成独立 run。

    latex2mathml 对 `\\langle u,v\\rangle` 产出单个 `<mi>⟨u,v⟩</mi>`，
    MML2OMML.XSL 再合并成一个 m:r —— 定界符不是独立元素，无法包成 m:d。
    这里按字符切分，让后续配对能工作。m:nor（正文文本）里的字符不动。
    """
    for r in list(root.iter('{%s}r' % M)):
        mpr = r.find('{%s}rPr' % M)
        if mpr is not None and mpr.find('{%s}nor' % M) is not None:
            continue
        t = r.find('{%s}t' % M)
        if t is None:
            continue
        txt = t.text or ''
        if len(txt) < 2 or not any(c in _STRETCH_CHARS for c in txt):
            continue
        parent = r.getparent()
        if parent is None:
            continue
        pos = list(parent).index(r)
        for k, part in enumerate(_STRETCH_RE.findall(txt)):
            nr = copy.deepcopy(r)
            nr.find('{%s}t' % M).text = part
            parent.insert(pos + k, nr)
        parent.remove(r)


def _is_plain_delim_run(el):
    """判断元素是否为"整个 run 就是一个待伸缩定界符字符"。"""
    if el.tag != '{%s}r' % M:
        return None
    mpr = el.find('{%s}rPr' % M)
    if mpr is not None and mpr.find('{%s}nor' % M) is not None:
        return None                      # m:nor 是正文文本，不该当数学定界符
    t = el.find('{%s}t' % M)
    if t is None:
        return None
    txt = (t.text or '').strip()
    if len(txt) == 1 and txt in _STRETCH_CHARS:
        return txt
    return None


def _wrap_stretch_delims(parent):
    """在同一层里把成对的定界符 run 连同中间内容包成 m:d（Word 才会按内容高度伸缩）。

    配对取"距离最近的一对"，这样嵌套时总是先裹最内层：
        ⟨⟨a⟩⟩  ->  ⟨ m:d(⟨a⟩) ⟩  ->  m:d(⟨ m:d(⟨a⟩) ⟩)
    而不是把外层 ⟨ 和内层 a 先裹在一起。
    """
    while True:
        kids = list(parent)
        marks = [(i, c) for i, el in enumerate(kids)
                 for c in [_is_plain_delim_run(el)] if c]
        hit = None
        for x in range(len(marks)):
            for y in range(x + 1, len(marks)):
                if _STRETCH_PAIRS.get(marks[x][1]) != marks[y][1]:
                    continue
                dist = marks[y][0] - marks[x][0]
                if hit is None or dist < hit[0]:
                    hit = (dist, x, y)
        if hit is None:
            return
        idx_a, ca = marks[hit[1]]
        idx_b, cb = marks[hit[2]]
        el_a, el_b = kids[idx_a], kids[idx_b]
        inner = kids[idx_a + 1:idx_b]
        if not inner:                    # 空的定界符对，删掉以免死循环
            parent.remove(el_a)
            parent.remove(el_b)
            continue
        d = etree.Element('{%s}d' % M)
        dpr = etree.SubElement(d, '{%s}dPr' % M)
        etree.SubElement(dpr, '{%s}begChr' % M).set('{%s}val' % M, ca)
        etree.SubElement(dpr, '{%s}endChr' % M).set('{%s}val' % M, cb)
        e = etree.SubElement(d, '{%s}e' % M)
        parent.insert(idx_a, d)
        for el in inner:
            e.append(el)                 # append 会自动从 parent 移走
        parent.remove(el_a)
        parent.remove(el_b)


def _stretch_delims(root):
    """全文修复：把退化成普通 run 的定界符改成可伸缩的 m:d。"""
    _split_delim_runs(root)
    for parent in list(root.iter()):
        _wrap_stretch_delims(parent)


def _force_upright(root):
    """给公式内"普通文本" run（m:nor）补 <w:i w:val="0"/>。

    Word 靠 m:nor 就能显示正体；但部分第三方渲染器（在线预览等）只认 w:rPr，
    会把这些文字也渲染成斜体。显式写 w:i=0 可兼容这类渲染器，对 Word 无副作用。
    """
    if not CFG.get('upright_in_text'):
        return
    for r in root.iter('{%s}r' % M):
        mpr = r.find('{%s}rPr' % M)
        if mpr is None or mpr.find('{%s}nor' % M) is None:
            continue
        rpr = r.find('{%s}rPr' % W)
        if rpr is None:
            rpr = etree.Element('{%s}rPr' % W)
            r.insert(1, rpr)
        i = rpr.find('{%s}i' % W)
        if i is None:
            i = etree.SubElement(rpr, '{%s}i' % W)
        i.set('{%s}val' % W, '0')


def _restore_text_runs(root):
    """还原占位符 / 空格，并按加粗哨兵把 m:nor run 拆成加粗 / 不加粗两段。"""
    for t in root.iter('{%s}t' % M):
        txt = t.text or ''
        if not txt:
            continue
        # 1) 显式间距标记 -> NBSP（先转成中转字符，避免下一步被误伤）
        if _NB in txt:
            txt = txt.replace(_NB, _NB_TMP)
        # 2) latex2mathml 把 \text{} 里的普通空格转成了 NBSP；样本正文 run 里
        #    NBSP 数量为 0 —— 必须换回普通空格，否则 Word 无法在词间断行。
        if '\xa0' in txt:
            txt = txt.replace('\xa0', ' ')
        if _NB_TMP in txt:
            txt = txt.replace(_NB_TMP, '\xa0')
        # 3) 特殊字符占位符还原
        if any(ph in txt for ph in _PH_BACK):
            for k, v in _PH_BACK.items():
                txt = txt.replace(k, v)
        if txt != (t.text or ''):
            t.text = txt

    for t in list(root.iter('{%s}t' % M)):
        txt = t.text or ''
        if B_ON not in txt and B_OFF not in txt:
            continue
        r = t.getparent()
        parent = r.getparent()
        pos = list(parent).index(r)
        bold, new = False, []
        for seg in re.split('([%s%s])' % (B_ON, B_OFF), txt):
            if seg == B_ON:
                bold = True
                continue
            if seg == B_OFF:
                bold = False
                continue
            if not seg:
                continue
            nr = copy.deepcopy(r)
            nt = nr.find('{%s}t' % M)
            nt.text = seg
            if bold:
                _mark_bold(nr)
            new.append(nr)
        for k, nr in enumerate(new):
            parent.insert(pos + k, nr)
        parent.remove(r)


# ---------------------------------------------------------------- 行 -> 公式
def _esc_prose(s):
    for ch, ph in _PH.items():
        s = s.replace(ch, ph)
    return s


def _text_seg(s, bold=False):
    s = _esc_prose(s)
    if bold:
        return '\\text{%s%s%s}' % (B_ON, s, B_OFF)
    return '\\text{%s}' % s


def line_to_tex(text, force_bold=False):
    """整行 -> 单个 LaTeX 表达式。

    文字段包成 \\text{...}（OMML 里就是 <m:nor>），数学段原样保留；
    加粗用哨兵 B_ON/B_OFF 标在 \\text{} 内部，转换后再拆 run 落成 <w:b/>。
    """
    out = []
    for i, seg in enumerate(re.split(r'\$([^$]+)\$', text)):
        if not seg:
            continue
        if i % 2 == 1:
            out.append(seg)
            continue
        for sub in re.split(r'(\*\*.+?\*\*)', seg):
            if not sub:
                continue
            if sub.startswith('**') and sub.endswith('**') and len(sub) > 4:
                out.append(_text_seg(sub[2:-2], True))
            else:
                out.append(_text_seg(sub, force_bold))
    return ''.join(out)


def _append_math(par, tex, size_pt, display=False, jc='center', strike=False):
    global _ERR
    try:
        omml = latex_to_omml(tex)
        _stretch_delims(omml)
        _restore_text_runs(omml)
        _style_math(omml, int(round(size_pt * 2)))
        _force_upright(omml)
        if strike:                       # 公式整体加删除线（如原文被划掉的 ℙ(ω)）
            for mr in omml.iter('{%s}r' % M):
                _set_strike(mr, True)
        _normalize_rpr_order(omml)
        if display:
            holder = parse_xml(
                '<m:oMathPara %s><m:oMathParaPr><m:jc m:val="%s"/></m:oMathParaPr>'
                '</m:oMathPara>' % (nsdecls('m'), jc))
            holder.append(omml)
            par._p.append(holder)
        else:
            par._p.append(omml)
        return True
    except Exception as e:
        _ERR += 1
        r = par.add_run(f'[公式失败: {tex}]')
        _set_run(r, size_pt)
        print(f'  ! 公式失败: {tex[:60]} -> {e}', file=sys.stderr)
        return False


# ---------------------------------------------------------------- run / 段落
def _set_run(run, size=None, bold=None):
    run.font.name = FONT
    if size:
        run.font.size = Pt(size)
    if bold is not None:
        run.font.bold = bold
    rpr = run._element.get_or_add_rPr()
    rf = rpr.find('{%s}rFonts' % W)
    if rf is None:
        rf = etree.Element('{%s}rFonts' % W)
        rpr.insert(0, rf)
    for a in ('ascii', 'hAnsi', 'eastAsia'):
        rf.set('{%s}%s' % (W, a), FONT)


def _set_para_mark(par, size, bold=False):
    """段落标记 rPr —— 公式 run 缺 rPr 时继承它，是公式字号的兜底。"""
    ppr = par._p.get_or_add_pPr()
    rpr = ppr.find('{%s}rPr' % W)
    if rpr is None:
        rpr = etree.Element('{%s}rPr' % W)
        ppr.append(rpr)
    _fill_rpr(rpr, int(round(size * 2)))
    b = rpr.find('{%s}b' % W)
    if bold and b is None:
        etree.SubElement(rpr, '{%s}b' % W)
    elif not bold and b is not None:
        rpr.remove(b)


def _norm(text):
    return re.sub(r'[ \t\u3000]{2,}', ' ', text).replace('\u3000', ' ')


# w:pPr 子元素顺序（只列用得到的部分，其余忽略）
_PPR_ORDER = ['pStyle', 'keepNext', 'keepLines', 'pageBreakBefore', 'framePr',
              'widowControl', 'numPr', 'suppressLineNumbers', 'pBdr', 'shd',
              'tabs', 'suppressAutoHyphens', 'kinsoku', 'wordWrap',
              'overflowPunct', 'topLinePunct', 'autoSpaceDE', 'autoSpaceDN',
              'bidi', 'adjustRightInd', 'snapToGrid', 'spacing', 'ind',
              'contextualSpacing', 'mirrorIndents', 'suppressOverlap', 'jc',
              'textDirection', 'textAlignment', 'textboxTightWrap',
              'outlineLvl', 'divId', 'cnfStyle', 'rPr', 'sectPr']


def _set_ind(par, left=None, hanging=None, first_line=None):
    """直接写 w:ind。python-docx 的 first_line_indent 不接受负值，悬挂缩进须手写。"""
    ppr = par._p.get_or_add_pPr()
    ind = ppr.find('{%s}ind' % W)
    if ind is None:
        ind = etree.Element('{%s}ind' % W)
        _insert_ordered(ppr, ind, _PPR_ORDER)
    for tag, val in (('left', left), ('hanging', hanging),
                     ('firstLine', first_line)):
        if val is not None:
            ind.set('{%s}%s' % (W, tag), str(int(round(val * 1440))))


def _add_numpr(par, num_id=1, ilvl=0):
    """挂项目符号（用文档自带 numbering.xml 的 numId=1：bullet \\uf0b7）。"""
    ppr = par._p.get_or_add_pPr()
    if ppr.find('{%s}numPr' % W) is not None:
        return
    _insert_ordered(ppr, parse_xml(
        '<w:numPr %s><w:ilvl w:val="%d"/><w:numId w:val="%d"/></w:numPr>'
        % (nsdecls('w'), ilvl, num_id)), _PPR_ORDER)


_BULLET_CHAR = None


_BULLET_RATIO = {}


def _bullet_ink(ch):
    """量符号的墨迹（PIL 实测，带缓存），返回 (高度/x-height, 填充率)。

    高度/x-height >1 表示比小写 x 还高、视觉够重；<0.3 说明是个小点，
    只撑得住短列表。填充率用来区分实心/空心 —— 空心的 ◇ □ ○ 虽然高大，
    但填充率只有 0.24~0.46，满页文字时压不住，所以重档只放实心符号。
    字形缺失（豆腐块）时返回 (0.0, 0.0)。
    """
    if ch in _BULLET_RATIO:
        return _BULLET_RATIO[ch]
    ft = _font(200)
    if ft is None:
        _BULLET_RATIO[ch] = (None, None)
        return _BULLET_RATIO[ch]
    try:
        from PIL import Image, ImageDraw
    except Exception:
        _BULLET_RATIO[ch] = (None, None)
        return _BULLET_RATIO[ch]

    def _ink(c):
        img = Image.new('L', (600, 600), 0)
        ImageDraw.Draw(img).text((200, 200), c, font=ft, fill=255)
        bb = img.getbbox()
        if bb is None:
            return (0.0, 0.0)
        # 用 tobytes 而不是 getdata：Pillow 14 会移除后者，且前者快得多
        # （'L' 模式每像素 1 字节）
        px = sum(1 for p in img.crop(bb).tobytes() if p > 60)
        area = max(1, (bb[2] - bb[0]) * (bb[3] - bb[1]))
        return ((bb[3] - bb[1]) / 200.0, px / area)

    xh, _ = _ink('x')
    h, fill = _ink(ch)
    _BULLET_RATIO[ch] = ((h / xh if xh else 0.0), fill)
    return _BULLET_RATIO[ch]


def _bullet_ink_ratio(ch):
    """只要"高度/x-height"一项。"""
    return _bullet_ink(ch)[0]


def _list_volume(lines):
    """列表体量 = 项数 + 平均项长/40。用来决定符号该多"重"。

    用户：「有些时候内容太多符号太小」—— 项多或项长的列表需要视觉更重的符号，
    否则满页文字配一个小点，压不住、也失去了分点的引导作用。
    """
    items = []
    for raw in lines:
        s = _norm(raw.strip())
        if s.startswith('- '):
            items.append(re.sub(r'\$[^$]*\$', '', s[2:]))
    if not items:
        return 0.0
    avg = sum(len(it) for it in items) / len(items)
    return len(items) + avg / 40.0


def _report_bullet(ch, vol):
    """生成后回报符号是否撑得住当前内容量（用户要求"使用之后检查一下"）。"""
    r, fill = _bullet_ink(ch)
    need = '重' if vol >= CFG['bullet_volume_heavy'] else '轻'
    if r is None:
        print(f'项目符号：{ch}（无法实测墨迹尺寸，请目视确认字形是否存在）')
        return
    if r == 0.0 and fill == 0.0:
        ok = '⚠ 字形缺失（豆腐块）-> 换 ● / ◆ / ■'
    elif vol >= CFG['bullet_volume_heavy']:
        if r < 0.30 or fill < 0.35:
            ok = '⚠ 偏小或空心，内容多时压不住 -> 建议 ● / ◆ / ■'
        elif r < 0.80:
            ok = '偏轻，长列表建议换更重的符号'
        else:
            ok = '合适'
    else:
        ok = '合适' if r >= 0.14 else '⚠ 墨迹异常'
    print(f'项目符号：{ch}  墨迹 {r:.2f}× x-height、填充率 {fill:.2f}'
          f' | 列表体量 {vol:.1f}（需"{need}"档）-> {ok}')


def _pick_bullet(src, lines=()):
    """挑一级项目符号：先按**列表体量**定档，再按源文件名确定性挑。

    样本实测一级符号有三种：● (ITVC / LA / Stats I / ODEs)、◆ (MFA / RA)、
    - (Prob I / ODEs)。
    CFG['bullet_char'] 写 'auto' 时：
      ① 体量 ≥ bullet_volume_heavy -> 从 bullet_pool_heavy 挑（重档）
      ② 否则从 bullet_pool_light 挑
      ③ 挑中后仍撑不住（轻档遇到大体量）-> 自动升级到重档
    按源文件名做确定性选择，同一份 md 每次结果一致（可复现），不同科目自然不同。
    想固定就直接把 CFG['bullet_char'] 写成具体符号。
    """
    global _BULLET_CHAR
    ch = CFG.get('bullet_char') or 'auto'
    vol = _list_volume(lines)
    if ch != 'auto':
        _BULLET_CHAR = ch
        _report_bullet(ch, vol)
        return _BULLET_CHAR
    heavy = CFG.get('bullet_pool_heavy') or ['●']
    light = CFG.get('bullet_pool_light') or heavy
    h = 0
    for c in Path(src).stem:
        h = (h * 31 + ord(c)) % 1000003
    # ① 按体量定档
    ch = (heavy if vol >= CFG['bullet_volume_heavy'] else light)[h % len(
        heavy if vol >= CFG['bullet_volume_heavy'] else light)]
    # ③ 自动升级：挑中的撑不住当前内容量时换重档
    r = _bullet_ink_ratio(ch)
    if vol >= CFG['bullet_volume_heavy'] and r is not None and r < 0.80:
        ch = heavy[h % len(heavy)]
    _BULLET_CHAR = ch
    _report_bullet(ch, vol)
    return _BULLET_CHAR


def _setup_list_numbering(doc, num_id=1):
    """把 numId 对应的 abstractNum 调成样本的两级列表。

    要做四件事：
    1. 注入 ilvl=1（符号 '-'、缩进 880/440）—— python-docx 默认模板每个
       abstractNum 只有 ilvl=0 一级；样本（Prob I / LA）是一级 `*` / 二级 `-`。
    2. 把 ilvl=0 的 ind 从默认 360/360 改成 440/440 —— **Word 以 numbering 里的
       ind 为准**（实测：段落自己写 w:ind left=440 也会被 numbering 的 360 盖掉，
       文字只到 108pt，而样本是 112pt）。
    3. 把 ilvl=0 的符号从默认 `\\uf0b7`/Symbol 换成 `_BULLET_CHAR`/Cambria Math
       —— 样本一级符号有 ● / ◆ / - 三种，可在 CFG['bullet_pool'] 里挑。
    4. 把 multiLevelType 从 singleLevel 改成 hybridMultilevel，
       否则 Word 忽略 ilvl>=1，二级列表退化回一级符号。
    """
    done = getattr(doc, '_wb_num_done', None)
    if done is None:
        done = set()
        doc._wb_num_done = done
    if num_id in done:
        return
    done.add(num_id)
    try:
        numbering = doc.part.numbering_part.element
    except Exception:
        return
    abstract_id = None
    for num in numbering.findall(qn('w:num')):
        if num.get(qn('w:numId')) == str(num_id):
            a = num.find(qn('w:abstractNumId'))
            abstract_id = a.get(qn('w:val')) if a is not None else None
    if abstract_id is None:
        return
    for ab in numbering.findall(qn('w:abstractNum')):
        if ab.get(qn('w:abstractNumId')) != abstract_id:
            continue
        # --- ilvl=0 缩进对齐样本 ---
        l0 = ab.find(qn('w:lvl'))
        if l0 is not None:
            # 符号：默认是 \uf0b7 + Symbol 字体（渲染成实心圆点）。
            # 换成挑中的符号，字体必须一并改成 Cambria Math —— 否则 ● 在
            # Symbol/Wingdings 下会渲染成别的字形（样本里 '-' 就是 Cambria Math）。
            if _BULLET_CHAR:
                lt = l0.find(qn('w:lvlText'))
                if lt is not None:
                    lt.set(qn('w:val'), _BULLET_CHAR)
                rpr = l0.find(qn('w:rPr'))
                if rpr is None:
                    rpr = etree.SubElement(l0, qn('w:rPr'))
                rf = rpr.find(qn('w:rFonts'))
                if rf is None:
                    rf = etree.SubElement(rpr, qn('w:rFonts'))
                for attr in ('ascii', 'hAnsi', 'cs'):
                    rf.set(qn('w:' + attr), FONT)
                rf.set(qn('w:hint'), 'default')
            # 文字位置其实由"编号后的制表位"决定（w:tabs 的 num tab），
            # 默认模板是 360twips -> 文字只到 108pt；样本是 440 -> 112pt。
            for tb in l0.findall('.//' + qn('w:tab')):
                if tb.get(qn('w:val')) == 'num':
                    tb.set(qn('w:pos'), '440')
            ind0 = l0.find('.//' + qn('w:ind'))
            if ind0 is not None:
                ind0.set(qn('w:left'), '440')
                ind0.set(qn('w:hanging'), '440')
        if any(l.get(qn('w:ilvl')) == '1' for l in ab.findall(qn('w:lvl'))):
            return
        mlt = ab.find(qn('w:multiLevelType'))
        if mlt is not None:
            mlt.set(qn('w:val'), 'hybridMultilevel')
        lvl = parse_xml(
            '<w:lvl %s w:ilvl="1">'
            '<w:start w:val="1"/>'
            '<w:numFmt w:val="bullet"/>'
            '<w:lvlText w:val="-"/>'
            '<w:lvlJc w:val="left"/>'
            '<w:pPr><w:ind w:left="880" w:hanging="440"/></w:pPr>'
            '<w:rPr><w:rFonts w:ascii="Cambria Math" w:hAnsi="Cambria Math" '
            'w:hint="default"/></w:rPr>'
            '</w:lvl>' % nsdecls('w'))
        ab.append(lvl)
        return


# 已自带标号的列表项（(a) (1) (i) …）—— 标号就在正文里，不再加项目符号
_LABEL_ITEM_RE = re.compile(r'^\((?:\d+|[A-Za-z]|[ivxlcdmIVXLCDM]+)\)\s')


def _list_flags(lines):
    """预扫全文，标出每个列表项在"列表块"里的位置，用于加边界间距。

    为什么要预扫：间距是**块级**属性（首项的上边界、末项的下边界），
    而主循环是流式的，读到末项时还不知道下一行是什么。

    返回 {行号: (is_list, is_start, is_end, tail_gap)}
      is_list  —— 该行是 `- ` 列表项
      is_start —— 列表块首项（上一非空行不是列表项）
      is_end   —— 列表块末项（下一非空行不是列表项）
      tail_gap —— 末项 **且** 下一非空行是"普通正文"（不是 # 标题类）时才 True。
                  下一行若是条目标题/次级小标题，它们自带 36pt 段前，
                  块边界已经成立，末项就不必再加下边界间距了。
    """
    info = []
    for raw in lines:
        s = _norm(raw.strip())
        info.append(s or None)
    n = len(info)
    out = {}
    for k in range(n):
        s = info[k]
        if not s or not s.startswith('- '):
            continue
        prev = next((info[j] for j in range(k - 1, -1, -1) if info[j]), None)
        nxt = next((info[j] for j in range(k + 1, n) if info[j]), None)
        is_start = not (prev or '').startswith('- ')
        is_end = not (nxt or '').startswith('- ')
        tail_gap = bool(is_end and nxt and not nxt.startswith('#'))
        out[k] = (True, is_start, is_end, tail_gap)
    return out


_TEXT_TOKEN = re.compile(r'(\*\*[^*]+\*\*)')
# 行内混排的分词器：$...$（行内公式）、~~...~~（删除线，内部可含公式）、
# **...**（加粗，内部可含公式）。三个一起识别才能处理"$f$ is **diff at $x$**"
# 这种加粗跨公式的情形 —— 之前若先按 $...$ 切，** 就会被拦腰斩断、漏出来。
_INLINE_TOKEN = re.compile(r'(\$[^$]+\$|~~[^~]+~~|\*\*[^*]+\*\*)')


def _emit_inline(p, text, size, bold=False, strike=False):
    """行内混排发射：按 _INLINE_TOKEN 一次切完，对 $...$ / ~~...~~ / **...**
    三种整 token 递归处理（内部允许再含公式），其余正文走 _add_text_runs。
    bold/strike 由调用方在递归时传入，用来跨公式保留外层格式。
    """
    for tok in _INLINE_TOKEN.split(text):
        if not tok:
            continue
        # 整 token 是 $...$ -> 行内公式
        if tok.startswith('$') and tok.endswith('$') and len(tok) > 2:
            _append_math(p, tok[1:-1], size, display=False, strike=strike)
            continue
        # 整 token 是 ~~...~~ -> 删除线包内部（内部可含公式，递归）
        if tok.startswith('~~') and tok.endswith('~~') and len(tok) > 4:
            _emit_inline(p, tok[2:-2], size, bold=bold, strike=True)
            continue
        # 整 token 是 **...** -> 加粗包内部（内部可含公式，递归 —— D1 修复点）
        if tok.startswith('**') and tok.endswith('**') and len(tok) > 4:
            _emit_inline(p, tok[2:-2], size, bold=True, strike=strike)
            continue
        # 普通正文
        _add_text_runs(p, tok, size, bold=bold, strike=strike)


def _add_text_runs(par, text, size, bold=None, strike=False):
    for seg in _TEXT_TOKEN.split(_norm(text)):
        if not seg:
            continue
        if seg.startswith('**') and seg.endswith('**') and len(seg) > 4:
            r = par.add_run(seg[2:-2])
            _set_run(r, size, True)
        else:
            r = par.add_run(seg)
            _set_run(r, size, bold)
        if strike:
            r.font.strike = True


def _par(doc, size=None, bold=False, after=None, before=None,
         line=None, align=None, indent=None, style=None,
         hanging=None, bullet=False, ilvl=0):
    p = doc.add_paragraph(style=style) if style else doc.add_paragraph()
    pf = p.paragraph_format
    pf.space_after = Pt(CFG['after_body'] if after is None else after)
    if before is not None:
        pf.space_before = Pt(before)
    pf.line_spacing = CFG['line_body'] if line is None else line
    if align is not None:
        p.alignment = align
    if indent is not None or hanging is not None:
        _set_ind(p, left=indent, hanging=hanging)
    if bullet:
        _setup_list_numbering(doc)
        _add_numpr(p, ilvl=ilvl)
    _set_para_mark(p, size if size else CFG['body'], bold)
    return p


def _split_math_segments(text):
    """把一行按 $...$ 切成 [(是否数学段, 片段), ...]。"""
    out = []
    for i, seg in enumerate(re.split(r'(\$[^$]+\$)', text)):
        if not seg:
            continue
        out.append((i % 2 == 1, seg[1:-1] if i % 2 == 1 else seg))
    return out


def _has_prose(text):
    """去掉公式段后是否还剩正文文字。"""
    return bool(re.sub(r'\$[^$]+\$', '', text).strip())


def _emit_line(doc, text, size, style=None, line=None, indent=None,
               before=None, after=None, hanging=None, bullet=False, ilvl=0):
    """渲染一行（不再做分行判断）。"""
    p = _par(doc, size, after=after, before=before, line=line,
             indent=indent, style=style, hanging=hanging, bullet=bullet,
             ilvl=ilvl)
    if '$' not in text:
        _add_text_runs(p, text, size)
        return p

    # 含正文文字 -> 混排段落：文字走 w:r、公式走行内 m:oMath。
    # 这样折行与普通文字完全一致（Word 对 m:oMathPara 的分行会强加缩进）。
    if _has_prose(text) and CFG['prose_carrier'] == 'mixed':
        if CFG['prose_jc'] == 'left':
            p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        elif CFG['prose_jc'] == 'center':
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        else:
            p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        _emit_inline(p, text, size)
        return p

    # 纯公式行，或 prose_carrier='inline_formula' 的兼容路径：整行并入一个公式对象。
    tex = line_to_tex(text)
    if CFG['math_line_mode'] != 'para':
        _append_math(p, tex, size, display=False)
    else:
        # 列表项里的纯公式保持左对齐（跟着缩进走），其余按纯公式规则居中
        jc = CFG['prose_jc'] if (style == 'List Paragraph' or indent) \
            else _display_jc(tex)
        _append_math(p, tex, size, display=True, jc=jc)
    return p


def _content_line(doc, text, size, style=None, line=None, indent=None,
                  before=None, after=None, hanging=None, bullet=False, ilvl=0):
    """一行内容：含 $...$ 就整行并入公式，否则普通文字 run。

    含文字的段落用 prose_jc（左对齐），保证读起来像正常行文。
    若行内出现"带上下限的大运算符"（∑ ∏ ∫ lim 等）**且整行放不下**，自动拆行：
        正文段 -> 独立公式 -> 正文段
    判据是"整行估算宽度是否超过正文列宽"，不是"有没有 ∑"：
    短行（如 `(b) Sum up to one, i.e. Σ_{ω∈Ω} ℙ({ω}) = 1`，176pt ≪ 415pt）
    保持行内小公式 —— 与样本一致；只有真的放不下才另起独立大公式。
    理由：Word 遇到 n 元运算符会强制把折行的续行缩进到运算符位置，
    所以"会折行"的行内大运算符必须拆出来；不折行的则不必拆。
    """
    if CFG.get('split_big_ops') and '$' in text:
        col = _column_width_pt(doc)
        if _line_width_pt(text, size) > col * CFG.get('split_ratio', 0.97):
            parts = _split_big_op_line(text)
            if parts:
                first = True
                for kind, seg in parts:
                    # 项目符号 / 悬挂缩进只落在拆行后的第一段上
                    kw = dict(hanging=hanging, bullet=bullet, ilvl=ilvl) if first \
                        else dict()
                    if kind == 'math':
                        p = _par(doc, size, style=style, line=line,
                                 indent=indent, before=before, after=after)
                        # 从列表项里拆出来的公式跟着列表缩进左对齐，避免"居中浮空"
                        _append_math(p, seg, size, display=True,
                                     jc=CFG['prose_jc'] if indent
                                     else _display_jc(seg))
                    else:
                        _emit_line(doc, seg, size, style=style, line=line,
                                   indent=indent, before=before, after=after,
                                   **kw)
                    first = False
                return
    return _emit_line(doc, text, size, style=style, line=line,
                      indent=indent, before=before, after=after,
                      hanging=hanging, bullet=bullet, ilvl=ilvl)


# 会撑破正文行的大运算符
# 注意：不能用 \b 收尾 —— 正则里 '_' 也算单词字符，"\sum_" 处 \b 不成立。
_BIGOP = re.compile(
    r'\\(?:sum|prod|coprod|int|oint|iint|iiint|bigcup|bigcap|bigoplus|'
    r'bigotimes|bigvee|bigwedge|bigsqcup|limsup|liminf|lim|max|min|'
    r'sup|inf|argmax|argmin)(?![A-Za-z])')


def _split_big_op_line(text):
    """把含"带上下限大运算符"的正文行拆成 [(kind, seg), ...]。

    为什么必须拆：Word 遇到 n 元运算符（∑ ∏ ∫ lim）会强制把折行的续行缩进到
    运算符位置 —— 实测行内 `\\sum` 也一样（续行从 333pt 起，正常应 90pt）。
    所以只要正文行里出现大运算符，就把它单独拎成一行独立公式。

    返回 None 表示不拆。
    """
    segs = re.split(r'(\$[^$]+\$)', text)
    hit = None
    for i, seg in enumerate(segs):
        if i % 2 == 1 and _BIGOP.search(seg[1:-1]) and re.search(r'[_^]', seg[1:-1]):
            hit = i
            break
    if hit is None:
        return None
    before = ''.join(segs[:hit]).strip()
    math = segs[hit][1:-1].strip()
    after = ''.join(segs[hit + 1:]).strip()
    # 前后正文合计要有实质内容，否则不值得拆
    prose_len = len(re.sub(r'\$[^$]*\$', '', before + after))
    if prose_len < 8:
        return None
    after = re.sub(r'^[.,;:]\s*', '', after)      # 去掉句读
    out = []
    if before:
        out.append(('text', before))
    out.append(('math', math))
    if after:
        out.append(('text', after))
    return out


def _is_derivation(tex):
    """判断是否为"有步骤那种"多行推导。

    preprocess 会把 align/aligned 统一改写成 array{rl}，故只要出现顶层 array
    就认为是按 = 对齐的分步推导 -> 左对齐。矩阵/分段函数用的是 pmatrix/cases，
    不会误判。
    """
    return '\\begin{array}' in preprocess_tex(tex)


def _display_jc(tex):
    return CFG['derive_jc'] if _is_derivation(tex) else CFG['display_jc']


# ---------------------------------------------------------------- 条目标题
_LABEL_RE = re.compile(
    r'^((?:Definition|Theorem|Proposition|Remark|Corollary|Example|Lemma|Note|'
    r'Recall|Claim|Axiom|Proof)\b[^A-Za-z]*)', re.I)


def _split_label(text):
    m = _LABEL_RE.match(text)
    if not m:
        return '', text
    return m.group(1), text[m.end():]


def _entry_title(doc, text, before=None):
    """条目标题：12pt；"Definition 1.6." 加粗，名称常规（样本实测）。

    名称里的公式走行内 m:oMath，标题过长时按普通文字折行。
    before 由主流程按"上一段类型"传入：正文块末→条目用 before_entry(≈40pt)，
    紧跟标题→条目用 before_entry_after_heading(≈19pt 总间距)。
    """
    p = _par(doc, CFG['entry'],
             before=CFG['before_entry'] if before is None else before,
             after=CFG['after_entry'])
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    label, name = _split_label(text)
    if label:
        r = p.add_run(label)
        _set_run(r, CFG['entry'], True)
    for is_math, seg in _split_math_segments(name):
        if is_math:
            _append_math(p, seg, CFG['entry'], display=False)
        else:
            _add_text_runs(p, seg, CFG['entry'])
    p.paragraph_format.keep_together = True   # v12：条目标题自身不跨页拆分
    return p


# ---------------------------------------------------------------- 样式初始化
def _setup_styles(doc):
    st = doc.styles

    n = st['Normal']
    n.font.name = FONT
    n.font.size = Pt(CFG['body'])
    n.element.rPr.rFonts.set('{%s}eastAsia' % W, FONT)
    n.paragraph_format.space_after = Pt(CFG['after_body'])
    n.paragraph_format.space_before = Pt(0)
    n.paragraph_format.line_spacing = CFG['line_body']

    for lvl, (sz, before, after) in {
            1: (CFG['h1'], CFG['before_h1'], CFG['after_h1']),
            2: (CFG['h2'], CFG['before_h2'], CFG['after_h2']),
            3: (CFG['sub3'], CFG['before_sub3'], 4)}.items():
        s = st['Heading %d' % lvl]
        s.font.name = FONT
        s.font.size = Pt(sz)
        s.font.bold = True
        s.font.italic = False
        # 样本实测：章/节标题为深蓝 #0F4761（6 份样本 PDF 中 5 份如此）。
        # 改 CFG['heading_color'] = '000000' 可回到纯黑。
        s.font.color.rgb = RGBColor.from_string(CFG['heading_color'])
        rpr = s.element.get_or_add_rPr()
        rf = rpr.find('{%s}rFonts' % W)
        if rf is None:
            rf = etree.Element('{%s}rFonts' % W)
            rpr.insert(0, rf)
        for a in ('ascii', 'hAnsi', 'eastAsia', 'cs'):
            rf.set('{%s}%s' % (W, a), FONT)
        s.paragraph_format.space_before = Pt(before)
        s.paragraph_format.space_after = Pt(after)
        s.paragraph_format.line_spacing = CFG['line_body']
        s.paragraph_format.keep_with_next = True
        s.paragraph_format.keep_together = True   # v12：章/节标题自身不跨页拆分

    lp = st['List Paragraph']
    lp.font.name = FONT
    lp.font.size = Pt(CFG['body'])
    lp.paragraph_format.line_spacing = CFG['line_list']


# ---------------------------------------------------------------- 目录 / 页码（v15）
# 正文列宽（twips）= 页宽 11906 - 左右边距 1800×2 —— 目录页码的右对齐制表位落点。
_TOC_TAB = 11906 - 1800 * 2


def _ensure_toc_styles(doc):
    """注入 TOC1 / TOC2 样式，与样本的 styleId 和**内建名**逐字一致。

    Word 的 TOC 域按内建样式名（"toc 1" / "toc 2"）给各级条目套格式，
    所以必须写 `w:name w:val="toc N"` 而不是自造名字，否则更新域后样式对不上。
    样本实测：TOC1 基于 Normal（无额外格式，11pt）；TOC2 仅多一个 ind left=420。
    这里额外给一个右对齐 + 点线引导的制表位，页码才能像样本一样右贴边。

    章条目默认**加粗 + 章标题同色深蓝 + 稍大 + 组前留白**（`toc_bold1` /
    `toc_color1` / `toc_entry1` / `toc_before1`），让每一章成为视觉分组；
    想回到样本的素净样子就把这几个值调回 `False` / `None` / `11` / `0`。
    """
    styles = doc.styles.element
    have = {s.get('{%s}styleId' % W) for s in styles.findall('{%s}style' % W)}
    normal_id = doc.styles['Normal'].style_id
    levels = (
        ('TOC1', 0, CFG['toc_entry1'], CFG['toc_bold1'], CFG['toc_color1'],
         CFG['toc_before1'], CFG['toc_after1']),
        ('TOC2', CFG['toc_indent2'], CFG['toc_entry'], False, None,
         CFG['toc_before2'], CFG['toc_after2']),
    )
    for sid, indent, size, bold, color, before, after in levels:
        if sid in have:
            continue
        # w:rPr 子元素顺序：rFonts -> b/bCs -> color -> sz/szCs（CT_RPr 是序列）
        rpr = ('<w:rFonts w:ascii="%s" w:hAnsi="%s" w:eastAsia="%s" w:cs="%s"/>'
               % (FONT, FONT, FONT, FONT))
        if bold:
            rpr += '<w:b/><w:bCs/>'
        if color:
            rpr += '<w:color w:val="%s"/>' % color
        rpr += ('<w:sz w:val="%d"/><w:szCs w:val="%d"/>'
                % (int(round(size * 2)), int(round(size * 2))))
        xml = (
            '<w:style %s w:type="paragraph" w:styleId="%s">'
            '<w:name w:val="toc %s"/>'
            '<w:basedOn w:val="%s"/><w:next w:val="%s"/>'
            '<w:uiPriority w:val="39"/><w:unhideWhenUsed/><w:qFormat/>'
            '<w:pPr>'
            '<w:tabs><w:tab w:val="right" w:leader="dot" w:pos="%d"/></w:tabs>'
            '<w:spacing w:before="%d" w:after="%d" w:line="%d" w:lineRule="auto"/>'
            '%s'
            '</w:pPr>'
            '<w:rPr>%s</w:rPr>'
            '</w:style>'
        ) % (
            nsdecls('w'), sid, sid[-1], normal_id, normal_id,
            _TOC_TAB,
            int(round(before * 20)), int(round(after * 20)),
            int(round(CFG['line_body'] * 240)),
            ('<w:ind w:left="%d"/>' % indent) if indent else '',
            rpr,
        )
        styles.append(parse_xml(xml))


def _insert_toc(doc):
    """在正文最前面插入 [目录标题][TOC 域]，并让第一个章标题另起一页。

    TOC 域是**原生域**：页码由 Word 生成，不是我们硬编码的，所以增删内容后
    按 F9（或跑 finalize_docx.ps1）即可刷新。域内先放一段占位结果，
    Word 一更新就被真实条目替换；即使不更新，文件也不会是空的。
    """
    body = doc.element.body
    field_p = parse_xml(
        '<w:p %s><w:pPr><w:pStyle w:val="TOC1"/></w:pPr>'
        '<w:r><w:fldChar w:fldCharType="begin"/></w:r>'
        '<w:r><w:instrText xml:space="preserve"> TOC \\o "%s" \\h \\z \\u '
        '</w:instrText></w:r>'
        '<w:r><w:fldChar w:fldCharType="separate"/></w:r>'
        '<w:r><w:rPr><w:rFonts w:ascii="%s" w:hAnsi="%s" w:eastAsia="%s"/>'
        '<w:sz w:val="%d"/></w:rPr>'
        '<w:t xml:space="preserve">Update this field (F9) to build the '
        'table of contents.</w:t></w:r>'
        '<w:r><w:fldChar w:fldCharType="end"/></w:r>'
        '</w:p>' % (nsdecls('w'), CFG['toc_levels'], FONT, FONT, FONT,
                    int(round(CFG['toc_entry'] * 2)))
    )
    body.insert(0, field_p)

    if CFG['toc_title']:
        title_p = parse_xml(
            '<w:p %s><w:pPr><w:keepNext/>'
            '<w:spacing w:before="0" w:after="%d" w:line="%d" '
            'w:lineRule="auto"/></w:pPr>'
            '<w:r><w:rPr><w:rFonts w:ascii="%s" w:hAnsi="%s" w:eastAsia="%s"/>'
            '<w:b/><w:color w:val="%s"/><w:sz w:val="%d"/><w:szCs w:val="%d"/>'
            '</w:rPr><w:t xml:space="preserve">%s</w:t></w:r></w:p>'
            % (nsdecls('w'), int(round(CFG['toc_title_after'] * 20)),
               int(round(CFG['line_body'] * 240)),
               FONT, FONT, FONT, CFG['heading_color'],
               int(round(CFG['toc_title_size'] * 2)),
               int(round(CFG['toc_title_size'] * 2)),
               CFG['toc_title'])
        )
        body.insert(0, title_p)

    # 目录后另起一页（样本实测：目录独占前面的页，正文从新页开始）
    if CFG['toc_break']:
        for p in doc.paragraphs:
            if p.style is not None and p.style.name == 'Heading 1':
                p.paragraph_format.page_break_before = True
                break


def _add_page_numbers(doc):
    """页脚居中页码：样本实测格式 " X / Y "（PAGE 域 / NUMPAGES 域）。"""
    sec = doc.sections[0]
    sec.footer_distance = Pt(CFG['page_number_gap'] / 20.0)
    ftr = sec.footer
    ftr.is_linked_to_previous = False
    p = ftr.paragraphs[0]
    for r in list(p.runs):                 # 清掉默认空 run，避免残留格式
        r._r.getparent().remove(r._r)
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    pf = p.paragraph_format
    pf.space_before = Pt(0)
    pf.space_after = Pt(0)
    pf.line_spacing = 1.0

    def _field(instr, placeholder):
        """一段完整的域：begin / instrText / separate / 占位结果 / end。"""
        runs = []
        r = p.add_run()
        r._r.append(parse_xml('<w:fldChar %s w:fldCharType="begin"/>'
                              % nsdecls('w')))
        runs.append(r)
        r = p.add_run()
        r._r.append(parse_xml(
            '<w:instrText %s xml:space="preserve"> %s </w:instrText>'
            % (nsdecls('w'), instr)))
        runs.append(r)
        r = p.add_run()
        r._r.append(parse_xml('<w:fldChar %s w:fldCharType="separate"/>'
                              % nsdecls('w')))
        runs.append(r)
        runs.append(p.add_run(placeholder))
        r = p.add_run()
        r._r.append(parse_xml('<w:fldChar %s w:fldCharType="end"/>'
                              % nsdecls('w')))
        runs.append(r)
        for rr in runs:
            _set_run(rr, CFG['page_number_size'])

    _field('PAGE', '1')
    if CFG['page_number_format'] == 'both':
        _set_run(p.add_run(' / '), CFG['page_number_size'])
        _field('NUMPAGES', '1')


def _set_update_fields(doc):
    """settings.xml 加 <w:updateFields w:val="true"/> —— 打开文档即刷新域（目录/页码）。"""
    st = doc.settings.element
    if st.find(qn('w:updateFields')) is not None:
        return
    el = parse_xml('<w:updateFields %s w:val="true"/>' % nsdecls('w'))
    # CT_Settings 是有序类型，updateFields 必须排在 compat/rsids 等之前
    for tag in ('w:hdrShapeDefaults', 'w:footnotePr', 'w:endnotePr',
                'w:compat', 'w:docVars', 'w:rsids'):
        anchor = st.find(qn(tag))
        if anchor is not None:
            anchor.addprevious(el)
            return
    st.append(el)


def _finalize_toc_and_pages(doc):
    """按 CFG 装配目录与页码。"""
    if CFG['toc']:
        _ensure_toc_styles(doc)
        _insert_toc(doc)
    if CFG['page_numbers']:
        _add_page_numbers(doc)
    if CFG['toc'] or CFG['page_numbers']:
        _set_update_fields(doc)


# ---------------------------------------------------------------- 表格
_TBLPR_ORDER = ['tblStyle', 'tblpPr', 'tblOverlap', 'bidiVisual',
                'tblStyleRowBandSize', 'tblStyleColBandSize', 'tblW', 'jc',
                'tblCellSpacing', 'tblInd', 'tblBorders', 'shd', 'tblLayout',
                'tblCellMar', 'tblLook']


def _insert_ordered(parent, el, order=_TBLPR_ORDER):
    """按 OOXML schema 顺序把 el 插进 parent，否则 Word 会报文档损坏。"""
    tag = etree.QName(el).localname
    idx = order.index(tag)
    pos = len(parent)
    for j, child in enumerate(parent):
        ct = etree.QName(child).localname
        if ct in order and order.index(ct) > idx:
            pos = j
            break
    parent.insert(pos, el)


def _parse_table_row(line):
    line = line.strip()
    if line.startswith('|'):
        line = line[1:]
    if line.endswith('|'):
        line = line[:-1]
    return [c.strip() for c in line.split('|')]


def _is_sep_row(line):
    return bool(re.match(r'^\|[\s:\-|]+\|$', line.strip()))


def _cell_paragraph(cell, text, size, align=WD_ALIGN_PARAGRAPH.CENTER):
    """单元格内容：文字走 w:r、$...$ 走行内 m:oMath（与正文同一套混排逻辑）。"""
    p = cell.paragraphs[0]
    pf = p.paragraph_format
    pf.space_before = Pt(0)
    pf.space_after = Pt(0)
    pf.line_spacing = CFG['line_body']
    p.alignment = align
    _set_para_mark(p, size)
    text = text.strip()
    if '$' not in text:
        _add_text_runs(p, text, size)
        return p
    # 整格只有一个公式 -> 用 oMathPara 承载（裸 m:oMath 会被 Word 当数学段落另排）
    if re.fullmatch(r'\$[^$]+\$', text):
        _append_math(p, text[1:-1], size, display=True, jc='center')
        return p
    if CFG['prose_carrier'] == 'mixed':
        _emit_inline(p, text, size)
    else:
        _add_text_runs(p, text, size)
    return p


def _emit_table(doc, block, size):
    """Markdown 管道表 -> Word 表格。

    版面（依讲义原表）：表头与内容居中、不加粗；
    框线为"上 / 下 + 表头下横线 + 列间竖线"（booktabs 风格）。
    列宽用固定布局显式指定 —— 否则 Word 的自动布局会把短列压成一字宽、
    长公式列撑出页面（实测踩坑）。
    """
    header = _parse_table_row(block[0])
    body = [_parse_table_row(l) for l in block[2:]]
    ncol = max([len(header)] + [len(r) for r in body]) if body else len(header)
    ncol = max(ncol, 1)

    sec = doc.sections[0]
    avail = int(sec.page_width - sec.left_margin - sec.right_margin)
    frac = [0.30, 0.70] if ncol == 2 else [1.0 / ncol] * ncol
    widths = [int(avail * f) for f in frac]

    tbl = doc.add_table(rows=1, cols=ncol)
    tbl.autofit = False
    _insert_ordered(tbl._tbl.tblPr,
                    parse_xml('<w:tblLayout %s w:type="fixed"/>' % nsdecls('w')))
    tblW = tbl._tbl.tblPr.find('{%s}tblW' % W)
    if tblW is None:
        tblW = parse_xml('<w:tblW %s w:w="0" w:type="auto"/>' % nsdecls('w'))
        _insert_ordered(tbl._tbl.tblPr, tblW)
    tblW.set('{%s}w' % W, str(sum(widths)))
    tblW.set('{%s}type' % W, 'dxa')

    for j in range(ncol):
        _cell_paragraph(tbl.rows[0].cells[j],
                        header[j] if j < len(header) else '', size)
    for row in body:
        cells = tbl.add_row().cells
        for j in range(ncol):
            _cell_paragraph(cells[j], row[j] if j < len(row) else '', size)

    for j in range(ncol):
        tbl.columns[j].width = widths[j]
        for row in tbl.rows:
            row.cells[j].width = widths[j]

    tblPr = tbl._tbl.tblPr
    _insert_ordered(tblPr, parse_xml(
        '<w:tblBorders %s>'
        '<w:top w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
        '<w:left w:val="none" w:sz="0" w:space="0" w:color="auto"/>'
        '<w:bottom w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
        '<w:right w:val="none" w:sz="0" w:space="0" w:color="auto"/>'
        '<w:insideH w:val="none" w:sz="0" w:space="0" w:color="auto"/>'
        '<w:insideV w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
        '</w:tblBorders>' % nsdecls('w')))
    for cell in tbl.rows[0].cells:
        tcPr = cell._tc.get_or_add_tcPr()
        tcPr.append(parse_xml(
            '<w:tcBorders %s>'
            '<w:bottom w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
            '</w:tcBorders>' % nsdecls('w')))
    return tbl


# ---------------------------------------------------------------- 主流程
def build(src, dst, jc=None):
    if jc:
        CFG['display_jc'] = jc
    lines = Path(src).read_text(encoding='utf-8').splitlines()
    flags = _list_flags(lines)
    _pick_bullet(src, lines)     # 依赖 lines：按列表体量定符号档位
    doc = Document()

    sec = doc.sections[0]
    sec.page_width, sec.page_height = Inches(8.27), Inches(11.69)   # A4
    sec.top_margin = sec.bottom_margin = Inches(1.0)
    sec.left_margin = sec.right_margin = Inches(1.25)

    _setup_styles(doc)

    n_disp = 0
    in_proof = False
    i = 0
    prev_kind = None                # 上一段块类型：h1/h2/h3/entry/sub/body/disp/table
    entry_paras = []                # v12：当前条目（###）累积段落，用于"整块同页"
    while i < len(lines):
        cur = i                     # 当前行下标（给 _list_flags 用）
        raw = lines[i].rstrip()
        n_indent = len(raw) - len(raw.lstrip(' \t'))   # 前导空格数（判二级列表）
        s = _norm(raw.strip())
        i += 1
        if not s:
            continue
        nb = len(doc.paragraphs)    # v12：本行产出前的段落数，用于事后捕获新段落
        # 本行块类型（v10：用于按"上一段类型"决定本段段前间距）
        if s.startswith('# '):
            cur_kind = 'h1'
        elif s.startswith('## '):
            cur_kind = 'h2'
        elif s.startswith('##### '):
            cur_kind = 'h3'
        elif s.startswith('#### '):
            cur_kind = 'sub'
        elif s.startswith('### '):
            cur_kind = 'entry'
        elif s.startswith('$$') or s == '$$':
            cur_kind = 'disp'
        elif s.startswith('|'):
            cur_kind = 'table'
        else:
            cur_kind = 'body'

        # v12：条目级"整块同页" —— 遇到章/节/新条目边界，先收尾上一块：
        # 除最后一段外全部 keep_with_next，杜绝"标题在上一页、主体在下一页"的孤行分页。
        if cur_kind in ('h1', 'h2', 'entry'):
            for p in entry_paras[:-1]:
                p.paragraph_format.keep_with_next = True
            entry_paras = []

        # Markdown 管道表：首行 | a | b |，次行 |---|---|
        if s.startswith('|'):
            block = [s]
            while i < len(lines):
                nxt = _norm(lines[i].strip())
                if nxt.startswith('|'):
                    block.append(nxt)
                    i += 1
                else:
                    break
            size = CFG['proof'] if in_proof else CFG['body']
            if len(block) >= 2 and _is_sep_row(block[1]):
                _emit_table(doc, block, size)
                prev_kind = 'table'
            else:                       # 不是表格，按普通正文逐行处理
                for ln in block:
                    _content_line(doc, ln, size)
                prev_kind = 'body'
            entry_paras.extend(doc.paragraphs[nb:])
            continue

        # $$ 独占一行 -> 收集到闭合 $$（可写多行，每行各自成为一条独立公式）
        if s == '$$':
            buf = []
            while i < len(lines) and lines[i].strip() != '$$':
                buf.append(lines[i].strip())
                i += 1
            i += 1
            size = CFG['proof'] if in_proof else CFG['body']
            for k, one in enumerate(buf):
                if not one:
                    continue
                p = _par(doc, size)
                if _append_math(p, one, size, display=True,
                                jc=_display_jc(one)):
                    n_disp += 1
            entry_paras.extend(doc.paragraphs[nb:])
            prev_kind = 'disp'
            continue

        if s.startswith('$$') and s.endswith('$$') and len(s) > 4:
            size = CFG['proof'] if in_proof else CFG['body']
            body = s[2:-2]
            p = _par(doc, size)
            if _append_math(p, body, size, display=True,
                            jc=_display_jc(body)):
                n_disp += 1
        elif s.startswith('##### '):
            in_proof = False
            h = doc.add_heading(s[6:], level=3)
            for r in h.runs:
                _set_run(r, CFG['sub3'], True)
            _set_para_mark(h, CFG['sub3'], True)
            h.paragraph_format.keep_together = True
            if prev_kind == 'body':   # v13：正文上方标题，至少留一行(按标题字号)
                h.paragraph_format.space_before = Pt(
                    max(CFG['before_sub3'], CFG['sub3'] * CFG['line_body']))
        elif s.startswith('#### '):
            in_proof = False
            before = (CFG['before_sub_after_heading']
                      if prev_kind in (None, 'h1', 'h2', 'h3')
                      else CFG['before_sub'])
            p = _par(doc, CFG['sub'], True, before=before, after=CFG['after_sub'])
            _add_text_runs(p, s[5:], CFG['sub'], True)
            p.paragraph_format.keep_together = True
        elif s.startswith('### '):
            in_proof = False
            before = (CFG['before_entry_after_heading']
                      if prev_kind in (None, 'h1', 'h2', 'h3')
                      else CFG['before_entry'])
            _entry_title(doc, s[4:], before=before)
        elif s.startswith('## '):
            in_proof = False
            h = doc.add_heading(s[3:], level=2)
            for r in h.runs:
                _set_run(r, CFG['h2'], True)
            _set_para_mark(h, CFG['h2'], True)
            if prev_kind == 'body':   # v13：正文上方标题，至少留一行(按标题字号)
                h.paragraph_format.space_before = Pt(
                    max(CFG['before_h2'], CFG['h2'] * CFG['line_body']))
        elif s.startswith('# '):
            in_proof = False
            h = doc.add_heading(s[2:], level=1)
            for r in h.runs:
                _set_run(r, CFG['h1'], True)
            _set_para_mark(h, CFG['h1'], True)
            if prev_kind == 'body':   # v13：正文上方标题，至少留一行(按标题字号)
                h.paragraph_format.space_before = Pt(
                    max(CFG['before_h1'], CFG['h1'] * CFG['line_body']))
        else:
            if _LABEL_RE.match(s) and not s.lower().startswith('proof'):
                in_proof = False
            if s.lower().startswith('proof'):
                in_proof = True
            size = CFG['proof'] if in_proof else CFG['body']
            if s.startswith('- '):
                item = s[2:]
                # 已自带 (a)/(1)/(i) 标号的项：标号就是标记 —— 既不加项目符号，
                # 也不缩进（样本里 (a)/(b) 齐左于 x0=90）。
                no_bullet = bool(_LABEL_ITEM_RE.match(item))
                # --- 列表块间距（v9）---
                # 首项加"上边界"、末项加"下边界"，项间收紧，带标号分点单独一档。
                _is, _start, _end, _tail = flags.get(
                    cur, (True, False, False, False))
                if _end:
                    after = CFG['after_list_block'] if _tail \
                        else CFG['after_list_item']
                elif no_bullet:
                    after = CFG['after_label_item']
                else:
                    after = CFG['after_list_item']
                if _start:
                    if no_bullet:
                        before = (CFG['before_label_item_after_heading']
                                  if prev_kind in (None, 'h1', 'h2', 'h3')
                                  else CFG['before_label_item'])
                    else:
                        before = (CFG['before_list_after_heading']
                                  if prev_kind in (None, 'h1', 'h2', 'h3')
                                  else CFG['before_list'])
                else:
                    before = None
                if n_indent >= 2:
                    # 缩进 2 空格 -> 二级列表（样本：一级 * 11pt / 二级 - 10pt）
                    _content_line(doc, item, CFG['sub_bullet'],
                                  line=CFG['line_list'],
                                  before=before, after=after,
                                  indent=None if no_bullet else CFG['indent_list2'],
                                  hanging=None if no_bullet else CFG['hanging_list2'],
                                  bullet=not no_bullet, ilvl=1)
                else:
                    _content_line(doc, item, size,
                                  line=CFG['line_list'],
                                  before=before, after=after,
                                  indent=None if no_bullet else CFG['indent_list'],
                                  hanging=None if no_bullet else CFG['hanging_list'],
                                  bullet=not no_bullet)
            else:
                before = (CFG['before_body_after_heading']
                          if prev_kind in (None, 'h1', 'h2', 'h3') else None)
                _content_line(doc, s, size, before=before)
            if '∎' in s:
                in_proof = False
        # v12：把本行产生的段落归入"当前条目"缓冲（章/节标题除外，已单独处理）
        if cur_kind not in ('h1', 'h2'):
            entry_paras.extend(doc.paragraphs[nb:])
        # 记录本段块类型，供下一段按"上一段类型"决定段前间距（v10）
        prev_kind = cur_kind

    # v12：收尾最后一个条目（除末段外全部 keep_with_next，保证整块同页）
    for p in entry_paras[:-1]:
        p.paragraph_format.keep_with_next = True
    _finalize_toc_and_pages(doc)      # v15：目录 + 页码
    doc.save(dst)
    print(f'已生成 {dst}（块级公式 {n_disp} 个，失败 {_ERR} 个'
          + ('，含目录' if CFG['toc'] else '')
          + ('，含页码' if CFG['page_numbers'] else '') + '）')
    return _ERR


if __name__ == '__main__':
    if len(sys.argv) not in (3, 4):
        print(__doc__)
        sys.exit(1)
    sys.exit(1 if build(*sys.argv[1:]) else 0)
