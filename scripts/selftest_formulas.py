"""公式链路自检：批量转换常见写法，报告失败项与内容丢失。

用法:
    python selftest_formulas.py

判据:
  * 转换不抛异常
  * 源 LaTeX 中的字母/数字，在生成的 OMML 文本里都能找到（防止静默丢内容，
    例如 mspace 被丢弃、括号被吞掉）
"""
import re
import sys
import unicodedata
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent))
from build_note import latex_to_omml, preprocess_tex   # noqa: E402

CASES = [
    r'a+b', r'x^2+y^2', r'a_{i,j}', r'\frac{a}{b}', r'\dfrac{a}{b}',
    r'\sum_{j=1}^{m}A_{i,j}B_{j,k}', r'\prod_{i=1}^{n}x_i',
    r'\int_{0}^{\infty}e^{-x^2}dx', r'\lim_{n\to\infty}a_n',
    r'\sqrt{x^2+1}', r'\sqrt[3]{x}',
    r'\alpha\beta\gamma\Gamma\Delta\Theta\lambda\mu\pi\sigma\varphi',
    r'\mathbb{Q}\mathbb{R}\mathbb{C}\mathbb{Z}', r'\mathcal{L}(V,W)',
    r'\operatorname{rank}(A)', r'\det(A)', r'\text{ if }x>0',
    r'\Rightarrow', r'\Leftrightarrow', r'\mapsto', r'\to',
    r'\overline{z}', r'\hat{f}', r'\tilde{x}', r'\langle u,v\rangle',
    r'\|x\|', r'\frac{\partial f}{\partial x}',
    r'a\equiv b\pmod{n}', r'\{1,2,\dots,n\}', r'\emptyset', r'\varnothing',
    r'\begin{pmatrix}1&0\\0&1\end{pmatrix}',
    r'\begin{vmatrix}a&b\\c&d\end{vmatrix}',
    r'\begin{pmatrix}a_{11}&\cdots&a_{1n}\\\vdots&&\vdots\\a_{n1}&\cdots&a_{nn}\end{pmatrix}',
    r'\frac{\frac{a}{b}}{\frac{c}{d}}',
    r'\begin{pmatrix}1&0\\0&1\end{pmatrix}^{-1}',
    r'f(x)=\begin{cases}x^2,&x\ge 0\\-x,&x<0\end{cases}',
    r'\begin{align}a&=b+c\\&=d\end{align}',
    r'r_i\to r_i+\lambda r_j', r'E_{i,j}',
    r'(A\times B)_{i,k}=\sum_{j=1}^{m}A_{i,j}B_{j,k}',
    r'(a+b)^2=a^2+2ab+b^2',
    r'(\frac{a}{b})_{n}',
    r'\left(\sum_{k=1}^{n}a_k\right)^2',
    r'\lambda I_n-A',
    r'\det(A-\lambda I_n)',
    r'v=\begin{pmatrix}1\\2\\3\end{pmatrix}',
    r'P^{-1}AP=D',
    r'\dim(V)=\dim(\ker T)+\dim(\operatorname{im}T)',
    r'a\quad b', r'a\qquad b', r'a\,b', r'a\;b', r'a\ b',
    r'\mathrm{d}x', r'\%', r'\{x\in\mathbb{R}\mid x>0\}',
]


def _alnum(s):
    """NFKC 归一（ℚ→Q、ℒ→L），只留字母数字，便于比对内容是否丢失。"""
    return re.sub(r'[^0-9A-Za-z]', '', unicodedata.normalize('NFKC', s))


def main():
    bad = 0
    for tex in CASES:
        try:
            omml = latex_to_omml(tex)
        except Exception as e:
            bad += 1
            print(f'FAIL  {tex}\n      {e}')
            continue
        got = _alnum(''.join(omml.itertext()))
        want = preprocess_tex(tex)
        want = re.sub(r'\\begin\{array\}\{[^}]*\}', '', want)     # 去掉列格式 {rl}
        want = re.sub(r'\\(begin|end)\{[a-zA-Z*]+\}', '', want)   # 去掉环境名
        want = re.sub(r'\\[a-zA-Z]+', '', want)                   # 去掉命令名
        want = _alnum(want)
        missing = [c for c in sorted(set(want)) if c not in got]
        if missing:
            bad += 1
            print(f'LOST  {tex}\n      缺少字符 {missing}  实际={got!r}')
    total = len(CASES)
    print(f'\n公式自检：{total - bad}/{total} 通过')
    bad += _test_lines()
    return 1 if bad else 0


# ---- 整行并入公式：文字段必须无损、占位符必须还原、加粗哨兵必须落地 ----
LINE_CASES = [
    'Let $K$ be a field, $A\\in M_{l,m}(K)$ be a matrix.',
    'for all $1\\le i\\le l$ and $1\\le k\\le n$.',
    'The rate is 50% and $x>0$ whenever $x$ is **positive**.',
    '$A$ & $B$ are comparable, i.e. $A\\ne B$.',
    'We have $(a+b)^2$ and $\\{x\\in\\mathbb{R}\\mid x\\ge 0\\}$.',
    'Proof. The $(i,l)$-entry equals $\\sum_{j,k}A_{i,j}B_{j,k}C_{k,l}$.',
    'Use $x<y$ and $y>z$ here.',
]


def _prose(md):
    """取出行内所有非公式段（去 ** 标记），用于比对文字是否丢失。"""
    out = []
    for seg in re.split(r'\$[^$]+\$', md):
        seg = seg.replace('**', '')
        if seg.strip():
            out.append(re.sub(r'\s+', ' ', seg).strip())
    return out


def _test_lines():
    from build_note import (line_to_tex, _restore_text_runs,
                            B_ON, B_OFF, W as _W, M as _M)
    bad = 0
    for md in LINE_CASES:
        try:
            omml = latex_to_omml(line_to_tex(md))
            _restore_text_runs(omml)
        except Exception as e:
            bad += 1
            print(f'FAIL  {md}\n      {e}')
            continue
        nor, allt = [], []
        for r in omml.iter('{%s}r' % _M):
            t = r.find('{%s}t' % _M)
            if t is None:
                continue
            txt = (t.text or '').replace('\xa0', ' ')
            allt.append(txt)
            mpr = r.find('{%s}rPr' % _M)
            if mpr is not None and mpr.find('{%s}nor' % _M) is not None:
                nor.append(txt)
        joined = re.sub(r'\s+', ' ', ''.join(nor))
        for seg in _prose(md):
            if seg not in joined:
                bad += 1
                print(f'LOST  {md}\n      文字段 {seg!r} 未完整落入公式 实际={joined!r}')
        tail = ' '.join(allt)
        if B_ON in tail or B_OFF in tail or '\ue000' in tail or '\ue001' in tail \
                or '\ue002' in tail or '\ue005' in tail:
            bad += 1
            print(f'MARK  占位符/哨兵未清理: {md}')

    # **...** 必须落成 <w:b/>
    omml = latex_to_omml(line_to_tex(LINE_CASES[2]))
    _restore_text_runs(omml)
    hit = False
    for r in omml.iter('{%s}r' % _M):
        mpr = r.find('{%s}rPr' % _M)
        if mpr is None or mpr.find('{%s}nor' % _M) is None:
            continue
        rpr = r.find('{%s}rPr' % _W)
        if rpr is not None and rpr.find('{%s}b' % _W) is not None:
            hit = True
    if not hit:
        bad += 1
        print('BOLD  **...** 未生成 <w:b/>')
    print(f'整行并入公式自检：{len(LINE_CASES)} 例通过，**加粗** -> {"已落成 <w:b/>" if hit else "失败"}')
    bad += _test_layout()
    return bad


# ---- 排版规则自检：空格归一 / 强制正体 / 大运算符分行 ----

# (输入, 期望是否拆行, 期望拆出的公式片段)
SPLIT_CASES = [
    (r'Proof. The $(i,l)$-entry of both sides equals '
     r'$\sum_{j,k}A_{i,j}B_{j,k}C_{k,l}$. Hence the two matrices are equal.',
     True, r'\sum_{j,k}A_{i,j}B_{j,k}C_{k,l}'),
    (r'In general $A\times B\ne B\times A$, so about 50% of the pairs do not commute.',
     False, None),
    (r'for all $1\le i\le l$ and $1\le k\le n$.', False, None),
    (r'The rate is 50% and $x>0$ whenever $x$ is **positive**, so $A$ & $B$ agree.',
     False, None),
    (r'By definition $\lim_{n\to\infty}a_n$ exists whenever the sequence is bounded.',
     True, r'\lim_{n\to\infty}a_n'),
]


def _test_layout():
    from build_note import (_restore_text_runs, _force_upright, _normalize_rpr_order,
                            _split_big_op_line, line_to_tex, _NB,
                            W as _W, M as _M)
    bad = 0

    # 1) 大运算符分行
    for md, want_split, want_math in SPLIT_CASES:
        got = _split_big_op_line(md)
        if bool(got) != want_split:
            bad += 1
            print(f'SPLIT {md[:60]!r}\n      期望拆行={want_split} 实际={bool(got)}')
            continue
        if want_split and want_math not in [s for k, s in got if k == 'math']:
            bad += 1
            print(f'SPLIT {md[:60]!r}\n      未拆出预期公式 {want_math!r}')
    print(f'大运算符分行自检：{len(SPLIT_CASES)} 例通过')

    # 1b) 行宽估算：决定公式用小公式还是大公式
    from build_note import _line_width_pt, _est_text_width, _est_math_width
    # 参照值来自 Probability I Outline.pdf 第 3 页字符级实测（10pt），按 11pt 折算
    WIDTH_REF = [
        ('text', '(b) Sum up to one, i.e. ', 92.7 * 1.1, 0.18),
        ('text', 'Sample Space: The set of all possible sets', 193.6 * 1.0, 0.18),
        ('math', r'\sum_{\omega\in\Omega} \mathbb{P}(\{\omega\}) = 1', 72.7 * 1.1, 0.25),
        ('math', r'\mathbb{P}(\{\omega\}) \geq 0 \text{ for all } \omega \in \Omega',
         100.8 * 1.1, 0.25),
    ]
    wbad = 0
    for kind, s, ref, tol in WIDTH_REF:
        got = _est_text_width(s, 11) if kind == 'text' else _est_math_width(s, 11)
        if abs(got - ref) / ref > tol:
            wbad += 1
            print(f'WIDTH {kind} 估算偏差过大: 估={got:.1f} 参照={ref:.1f} | {s[:44]}')
    # 关键判定：样本 (b) 行整行只有 176pt，必须判为"放得下"（保持行内小公式）
    inline_line = (r'(b) sum up to one, i.e. '
                   r'$\sum_{\omega \in \Omega} \mathbb{P}(\{\omega\}) = 1$.')
    w_inline = _line_width_pt(inline_line, 11)
    fits = w_inline < 415 * 0.97
    if not fits:
        wbad += 1
        print(f'WIDTH (b) 行被误判为放不下: {w_inline:.1f}pt')
    # 反向：一行塞得满满的长正文 + 大运算符，必须判为放不下
    long_line = (r'Proof. The $(i,l)$-entry of both sides equals the double sum '
                 r'$\sum_{j,k}A_{i,j}B_{j,k}C_{k,l}$ and hence the two matrices agree.')
    w_long = _line_width_pt(long_line, 11)
    if w_long <= 415 * 0.97:
        wbad += 1
        print(f'WIDTH 长行被误判为放得下: {w_long:.1f}pt')
    print(f'行宽估算自检：{len(WIDTH_REF)} 例偏差达标 + (b) 行 {w_inline:.0f}pt 保持行内 '
          f'+ 长行 {w_long:.0f}pt 判定需拆 -> {"通过" if wbad == 0 else "失败"}')
    bad += wbad

    # 2) 空格归一：\text{} 里的普通空格不能被留成 NBSP（否则无法断行）
    omml = latex_to_omml(line_to_tex(LINE_CASES[0]))
    _restore_text_runs(omml)
    _style_check = []
    for r in omml.iter('{%s}r' % _M):
        mpr = r.find('{%s}rPr' % _M)
        t = r.find('{%s}t' % _M)
        if mpr is None or mpr.find('{%s}nor' % _M) is None or t is None:
            continue
        _style_check.append(t.text or '')
    prose = ''.join(_style_check)
    if '\xa0' in prose:
        bad += 1
        print(f'NBSP  正文 run 残留 NBSP: {prose!r}')
    print(f'空格归一自检：正文 run 无 NBSP -> {"通过" if chr(0xa0) not in prose else "失败"}')

    # 3) 显式间距（\quad）必须保留为 NBSP，不能被压成普通空格
    omml2 = latex_to_omml(line_to_tex(r'$a\quad b$'))
    _restore_text_runs(omml2)
    kept = any('\xa0' in (t.text or '') for t in omml2.iter('{%s}t' % _M))
    if not kept:
        bad += 1
        print('QUAD  \\quad 的显式间距丢失')
    print(f'显式间距自检：\\quad 保留为 NBSP -> {"通过" if kept else "失败"}')

    # 4) 强制正体：m:nor run 必须带 <w:i w:val="0"/>
    omml3 = latex_to_omml(line_to_tex(LINE_CASES[0]))
    _restore_text_runs(omml3)
    _force_upright(omml3)
    _normalize_rpr_order(omml3)
    nor_n = up_n = 0
    for r in omml3.iter('{%s}r' % _M):
        mpr = r.find('{%s}rPr' % _M)
        if mpr is None or mpr.find('{%s}nor' % _M) is None:
            continue
        nor_n += 1
        rpr = r.find('{%s}rPr' % _W)
        if rpr is not None and rpr.find('{%s}i' % _W) is not None:
            up_n += 1
    if nor_n == 0 or nor_n != up_n:
        bad += 1
        print(f'UPRIGHT  m:nor={nor_n} 带 w:i=0 的={up_n}')
    print(f'强制正体自检：{up_n}/{nor_n} 个 m:nor run 已标 w:i=0')

    # 5) 混排切分：文字段与数学段必须能无损还原原文
    from build_note import _split_math_segments, _has_prose
    split_bad = 0
    for md in LINE_CASES:
        rebuilt = ''.join(
            ('$%s$' % seg if is_math else seg)
            for is_math, seg in _split_math_segments(md))
        if rebuilt != md:
            split_bad += 1
            print(f'SPLIT-SEG 还原不一致: {md!r} -> {rebuilt!r}')
    # 纯公式行不应被判为"含正文"
    for md, want in (('$x+y$', False), ('Let $x$ be $y$.', True)):
        if _has_prose(md) != want:
            split_bad += 1
            print(f'HAS_PROSE 判定错误: {md!r} 期望 {want}')
    bad += split_bad
    print(f'混排切分自检：{len(LINE_CASES)} 例往返一致 + 2 例正文判定')
    bad += _test_delims()
    bad += _test_opname()
    return bad


# ---- 定界符伸缩：括号必须包成 m:d，Word 才会按内容高度拉伸 ----
# (tex, 期望 begChr, 期望 endChr)；None 表示该 m:d 不带 begChr/endChr
# （OMML 里 m:d 省略 begChr/endChr 时默认就是圆括号，如 (...) 与 pmatrix）。
DELIM_CASES = [
    (r'(\frac{a}{b})',                        None, None),
    (r'\left(\frac{a}{b}\right)',             None, None),
    (r'[\frac{a}{b}]',                        '[',  ']'),
    (r'\{\frac{a}{b}\}',                      '{',  '}'),
    (r'|\frac{a}{b}|',                        '|',  '|'),
    (r'\|\frac{a}{b}\|',                      '‖',  '‖'),
    (r'\lvert\frac{a}{b}\rvert',              '|',  '|'),
    (r'\lVert\frac{a}{b}\rVert',              '‖',  '‖'),
    (r'\langle\frac{a}{b}\rangle',            '⟨',  '⟩'),
    (r'\left\langle\frac{a}{b}\right\rangle', '⟨',  '⟩'),
    (r'\langle u,v\rangle',                   '⟨',  '⟩'),
    (r'\left\langle u,v\right\rangle',        '⟨',  '⟩'),
    (r'\langle x\rangle',                     '⟨',  '⟩'),
    (r'\langle u,v\rangle = 0',               '⟨',  '⟩'),
    (r'\lfloor x\rfloor',                     '⌊',  '⌋'),
    (r'\lfloor\frac{a}{b}\rfloor',            '⌊',  '⌋'),
    (r'\left\lfloor x\right\rfloor',          '⌊',  '⌋'),
    (r'\lceil x\rceil',                       '⌈',  '⌉'),
    (r'\lceil\frac{a}{b}\rceil',              '⌈',  '⌉'),
    (r'a\lceil x\rceil b',                    '⌈',  '⌉'),
    (r'\langle a\rangle+\langle b\rangle',    '⟨',  '⟩'),
    (r'\begin{pmatrix}\frac{a}{b}\end{pmatrix}', None, None),
    # \big 系列在本链路里是空操作（minsize/maxsize 被忽略），预处理直接删掉；
    # 若不删，\big\langle 会退化成字面文本 "\langle"（实测）。
    (r'\big(\frac{a}{b}\big)',                None, None),
    (r'\Big(\frac{a}{b}\Big)',                None, None),
    (r'\bigg(\frac{a}{b}\bigg)',              None, None),
    (r'\big\langle x\big\rangle',             '⟨',  '⟩'),
    (r'\big\langle\frac{a}{b}\big\rangle',    '⟨',  '⟩'),
    (r'\big[x\big]',                          '[',  ']'),
    # 函数名紧邻定界符：latex2mathml 把函数名和后面的 ( 合并进同一个 <m:t>
    # （如 "det("），定界符不是独立元素 —— 实测 det(pmatrix) 的括号完全不伸缩。
    # 必须先 _split_delim_runs 拆开才能配对。
    (r'\det\begin{pmatrix}a&b\\c&d\end{pmatrix}',                '(', ')'),
    (r'\operatorname{tr}\begin{pmatrix}a&b\\c&d\end{pmatrix}',  '(', ')'),
    (r'\det\begin{pmatrix}\frac{a}{b}&0\\0&\frac{c}{d}\end{pmatrix}', '(', ')'),
    (r'\sin(x)',                              '(',  ')'),
    (r'\sin\left(\frac{a}{b}\right)',         None, None),
    (r'f(x)',                                 '(',  ')'),
    (r'f(g(x))',                              '(',  ')'),
    (r'\log_{2}(x+1)',                        '(',  ')'),
    (r'\det\begin{bmatrix}a&b\\c&d\end{bmatrix}', '[', ']'),
    (r'\det\left\{\begin{matrix}a&b\\c&d\end{matrix}\right\}', '{', '}'),
    (r'\det(A)\det(B)',                       '(',  ')'),
]

# 半开区间 / 正文里的括号：不得被误配成 m:d（否则会凭空多出一对定界符）
NO_PAIR_CASES = [
    r'(a,b]',
    r'[a,b)',
    r'\text{The matrix (1,0) is fine}',
]

# 预处理必须保留的大运算符命令（不能被 \big 剥离规则误伤）
BIGFAM_KEEP = [r'\bigcup_{i=1}^{n}A_i', r'\bigcap_{i=1}^{n}A_i',
               r'\bigoplus_{i=1}^{n}V_i', r'\bigotimes_{i=1}^{n}V_i',
               r'\bigvee_{i=1}^{n}p_i', r'\bigwedge_{i=1}^{n}p_i',
               r'\bigsqcup_{i=1}^{n}A_i']


def _test_delims():
    """定界符必须包成 m:d；否则 Word 只能渲染成固定大小的默认括号。"""
    from build_note import _stretch_delims
    from lxml import etree
    bad = 0
    for tex, eb, ee in DELIM_CASES:
        try:
            root = latex_to_omml(tex)
            _stretch_delims(root)
            s = etree.tostring(root, encoding='unicode')
        except Exception as e:
            bad += 1
            print(f'FAIL  {tex}\n      {e}')
            continue
        if s.count('<m:d>') < 1:
            bad += 1
            print(f'NO-DELIM  {tex}  未生成 m:d（括号无法伸缩）')
            continue
        beg = re.findall(r'<m:begChr m:val="([^"]*)"/>', s)
        end = re.findall(r'<m:endChr m:val="([^"]*)"/>', s)
        got_beg = beg[0] if beg else None
        got_end = end[0] if end else None
        if (got_beg, got_end) != (eb, ee):
            bad += 1
            print(f'DELIM  {tex}\n      期望 {eb}/{ee} 实际 {got_beg}/{got_end}')
    # \big 剥离规则不得误伤 \bigcup \bigcap \bigoplus \bigotimes \bigvee \bigwedge \bigsqcup
    for tex in BIGFAM_KEEP:
        if tex.split('_')[0] not in preprocess_tex(tex):
            bad += 1
            print(f'BIGFAM  {tex}  被 \\big 剥离规则误伤')
    # 半开区间 / 正文括号不得被误配
    for tex in NO_PAIR_CASES:
        try:
            root = latex_to_omml(tex)
            _stretch_delims(root)
            n = etree.tostring(root, encoding='unicode').count('<m:d>')
        except Exception as e:
            bad += 1
            print(f'FAIL  {tex}\n      {e}')
            continue
        if n != 0:
            bad += 1
            print(f'NO-PAIR  {tex}  不应生成 m:d，实际 {n} 个')
    print(f'定界符伸缩自检：{len(DELIM_CASES)} 例通过'
          f'（含 \\big 系列剥离，{len(BIGFAM_KEEP)} 个大运算符命令未误伤，'
          f'{len(NO_PAIR_CASES)} 个非成对括号未误配）')
    return bad


# ---- 函数名正体：latex2mathml 对下面这些输出 <mo>，会渲染成斜体 ----
# 实测样本里 det / dim 的 run 都带 <m:sty m:val="p"/>（正体）。
OPNAME_FIX = ['det', 'lim', 'limsup', 'liminf', 'sup', 'inf', 'min', 'max',
              'gcd', 'Pr']
OPNAME_ALREADY = ['sin', 'cos', 'tan', 'cot', 'sec', 'log', 'ln', 'lg', 'exp',
                  'dim', 'ker', 'arg', 'deg', 'hom']
# 这些写法本来就不该被改写（(?![A-Za-z]) 防误伤）
OPNAME_KEEP = [r'\supset', r'\subset', r'\infty', r'\int', r'\sin', r'\dim V',
               r'\ker T', r'\bigcup_{i}A_i', r'\bigoplus_{i}V_i']
# \operatorname{...} 同样输出 <mo>，必须一并转 \mathrm
OPNAME_OPERATORNAME = [r'\operatorname{tr}', r'\operatorname{rank}',
                       r'\operatorname*{argmax}']


def _test_opname():
    """函数名必须正体；易误伤写法不得被改写。"""
    from lxml import etree
    bad = 0
    for n in OPNAME_FIX:
        try:
            s = etree.tostring(latex_to_omml('\\' + n), encoding='unicode')
        except Exception as e:
            bad += 1
            print(f'FAIL  \\{n}: {e}')
            continue
        if '<m:sty m:val="p"/>' not in s:
            bad += 1
            print(f'OPNAME  \\{n} 未正体（缺 m:sty val="p"）')
    for n in OPNAME_ALREADY:
        try:
            s = etree.tostring(latex_to_omml('\\' + n), encoding='unicode')
        except Exception as e:
            bad += 1
            print(f'FAIL  \\{n}: {e}')
            continue
        if '<m:sty m:val="p"/>' not in s:
            bad += 1
            print(f'OPNAME  \\{n} 原本正体，现已丢失 sty=p')
    for tex in OPNAME_KEEP:
        if 'mathrm' in preprocess_tex(tex):
            bad += 1
            print(f'OPNAME  {tex}  被误转 \\mathrm')
    for tex in OPNAME_OPERATORNAME:
        if 'mathrm' not in preprocess_tex(tex):
            bad += 1
            print(f'OPNAME  {tex}  未被转成 \\mathrm')
    print(f'函数名正体自检：{len(OPNAME_FIX)} 个斜体函数名已修正、'
          f'{len(OPNAME_ALREADY)} 个原本正体未破坏、'
          f'{len(OPNAME_KEEP)} 个易误伤写法未改写、'
          f'{len(OPNAME_OPERATORNAME)} 个 \\operatorname 已改写')
    return bad


if __name__ == '__main__':
    sys.exit(main())
