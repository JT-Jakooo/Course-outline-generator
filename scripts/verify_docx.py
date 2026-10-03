"""回读校验生成的 docx，输出关键排版指标，用于确认没有静默降级。

用法:
    python verify_docx.py <doc.docx>

检查项:
  * oMath / oMathPara 数量
  * 是否有 [公式失败 ...] 残留
  * 正文文本里是否残留 NBSP（会导致无法断行、排版塌成一团）
  * 公式内正文 run 是否补了 <w:i w:val="0"/>（正体）
  * 目录：TOC 域是否已由 Word 刷新（而非占位符）、TOC1/TOC2 条目与内建样式名
  * 页码：页脚 PAGE / NUMPAGES 域、settings.xml 的 updateFields
  * 段落分类：正文行(混排) / 纯公式行 / 裸行内公式段落（后者必须为 0）
  * 标题与条目字号、是否加粗
  * 正文/公式是否统一 Cambria Math
  * 段落间距与行距是否一致
"""
import re
import sys
import zipfile
from collections import Counter


def main(path):
    z = zipfile.ZipFile(path)
    d = z.read('word/document.xml').decode('utf-8')
    s = z.read('word/styles.xml').decode('utf-8')

    om = len(re.findall(r'<m:oMath[ >]', d))
    omp = len(re.findall(r'<m:oMathPara[ >]', d))
    nor = d.count('<m:nor/>')
    fail = len(re.findall(r'\[公式失败', d))
    print(f'oMath           : {om}')
    print(f'oMathPara(块级) : {omp}')
    print(f'm:nor(公式内正文): {nor}   <- 整行并入公式的证据，应为正文行数量级')
    print(f'公式失败残留     : {fail}')

    # 段落分类
    #   含正文文字 + 行内公式  -> 正文行（混排），期望存在，靠 w:jc 左对齐、自然折行
    #   只有行内公式、没有文字 -> Word 会当数学段落默认居中排，必须为 0
    #   有 m:oMathPara         -> 纯公式行（$$ 块 / 拆出来的大运算符公式）
    prose_par, bare_par, para_par, bare_list = 0, 0, 0, 0
    for pm in re.finditer(r'<w:p\b.*?</w:p>', d, re.S):
        blk = pm.group(0)
        if '<m:oMathPara' in blk:
            para_par += 1
            continue
        if '<m:oMath' not in blk:
            continue
        # 注意：正文 run 常以空格开头（" is a field iff "），不能要求 w:t 首字符非空白
        has_text = re.search(
            r'<w:r\b(?:(?!</w:r>).)*?<w:t[^>]*>[^<]*\S', blk, re.S) is not None
        if has_text:
            prose_par += 1
        elif '<w:numPr>' in blk:
            # 列表项：项目符号在 numbering 里，段落可以没有文字 run；
            # 且 Word 往返会把列表段里的 oMathPara 外壳拆掉（实测 3->1），渲染仍左对齐
            #（2026-10-01 Practical Statistics 实测：编号列表的缩进与对齐不受影响）。
            # 故单独计数，不计入 bare_par。
            bare_list += 1
        else:
            bare_par += 1
    print(f'正文行(文字+行内公式): {prose_par}   <- 期望存在；左对齐 + 自然折行')
    print(f'纯公式行(oMathPara)  : {para_par}   <- $$ 块 + 从正文拆出的大运算符公式')
    print(f'裸行内公式段落       : {bare_par}   <- 必须为 0；>0 会被 Word 默认居中')
    print(f'裸公式列表项(单独计数): {bare_list}   <- 纯公式列表项，Word 往返后属预期，渲染仍左对齐')

    # 正文（w:t，公式之外的 prose run）里的 NBSP：样本实测为 0。
    #   —— 这才是真正的"词间断不开"隐患，必须为 0。
    # 公式内部 m:nor（\text{} 正常文本）里的 NBSP 来自 \, \quad \qquad 等显式间距，
    #   是 Word 公式区里"不可断行的间距"，属预期行为，单独计数仅作信息、不计违规。
    nbsp_body = 0
    for mm in re.finditer(r'<w:r\b(?:(?!</w:r>).)*?<w:t[^>]*>(.*?)</w:t>', d, re.S):
        if '\xa0' in mm.group(1):
            nbsp_body += 1
    nbsp_math = 0
    for mm in re.finditer(r'<m:r><m:rPr><m:nor/></m:rPr>.*?<m:t[^>]*>(.*?)</m:t>',
                          d, re.S):
        if '\xa0' in mm.group(1):
            nbsp_math += 1
    print(f'正文 w:t 含 NBSP     : {nbsp_body}   <- 必须为 0；>0 会让 Word 无法在词间断行')
    print(f'公式内 m:nor 含 NBSP : {nbsp_math}   <- 预期行为（thin/quad 等显式间距，不可断行），不计违规')

    # ---- CJK 汉字检查（2026-10-03 PDE 提纲事故）----
    # 跨科目约定：生成的提纲/作业文档必须全英文，正文严禁混入中文。
    # 只查 prose w:t（标题/段落文本）；公式内 m:t 不查（数学符号正常应无汉字，
    # 但为降低误报风险先不纳入硬失败）。加 --allow-cjk 可豁免（确要出中文文档时）。
    cjk_re = re.compile(r'[\u4e00-\u9fff\u3400-\u4dbf]')
    cjk_hits = []
    for mm in re.finditer(r'<w:r\b(?:(?!</w:r>).)*?<w:t[^>]*>(.*?)</w:t>', d, re.S):
        if cjk_re.search(mm.group(1)):
            cjk_hits.append(mm.group(1).strip()[:60])
    allow_cjk = '--allow-cjk' in sys.argv[2:]
    print(f'正文含汉字         : {len(cjk_hits)}   <- 必须为 0；示例: {cjk_hits[:3] if cjk_hits else "（无）"}'
          + ('   <-- --allow-cjk 已豁免' if allow_cjk and cjk_hits else ''))

    up = len(re.findall(r'<w:i w:val="0"/>', d))
    print(f'w:i=0(强制正体)  : {up}   <- 公式内正文 run 的正体保障')

    # ---- v15：目录 / 页码 ----
    instr = [i.strip() for i in re.findall(
        r'<w:instrText[^>]*>([^<]*)</w:instrText>', d)]
    toc_instr = [i for i in instr if i.startswith('TOC')]
    toc_lv = Counter(re.findall(r'<w:pStyle w:val="(TOC\d)"/>', d))
    placeholder = 'Update this field' in d
    print(f'TOC 域           : {toc_instr if toc_instr else "（无）"}'
          + ('   <- 仍是占位符，未用 Word 刷新（finalize_docx.ps1）'
             if placeholder else '   <- 已含真实条目'))
    print(f'目录条目段落     : {dict(sorted(toc_lv.items())) or "（无）"}'
          '   <- TOC1=章 / TOC2=节')
    toc_style = sorted(set(re.findall(r'<w:name w:val="(toc \d)"/>', s)))
    print(f'TOC 样式(内建名) : {toc_style or "（无）"}   <- 必须是 "toc N"，Word 按此套用')

    ftr = [n for n in z.namelist() if 'footer' in n]
    pg = npg = 0
    for n in ftr:
        for f in re.findall(r'<w:instrText[^>]*>([^<]*)</w:instrText>',
                            z.read(n).decode('utf-8')):
            if f.strip() == 'PAGE':
                pg += 1
            elif f.strip() == 'NUMPAGES':
                npg += 1
    print(f'页脚             : {len(ftr)} 个部件，PAGE 域={pg} NUMPAGES 域={npg}')
    if 'word/settings.xml' in z.namelist():
        se = z.read('word/settings.xml').decode('utf-8')
        print(f'settings.xml     : updateFields={"updateFields" in se}'
              '   <- true 时 Word 打开即刷新目录/页码')

    jc = Counter(re.findall(r'<m:oMathParaPr><m:jc m:val="([a-z]+)"/>', d))
    print(f'纯公式行对齐     : {dict(jc)}   <- 独立公式 center / 推导 left')

    fonts = Counter(re.findall(r'<w:rFonts w:ascii="([^"]+)"', d))
    print(f'显式字体分布     : {dict(fonts.most_common(5))}')
    nofont = len(re.findall(r'<w:r>(?:(?!</w:r>).)*?<w:t', d, re.S))
    print(f'（正文 run 总数参考）: {nofont}')

    # 标题
    for lvl in (1, 2, 3):
        m = re.search(r'<w:style [^>]*w:styleId="Heading%d"[^>]*>.*?</w:style>' % lvl, s, re.S)
        if not m:
            m = re.search(r'<w:name w:val="heading %d"/>.*?</w:style>' % lvl, s, re.S)
        if m:
            sz = re.findall(r'<w:sz w:val="(\d+)"', m.group(0))
            b = 'bold' if '<w:b/>' in m.group(0) else 'regular'
            col = re.findall(r'<w:color w:val="([0-9A-Fa-f]+)"', m.group(0))
            print(f'Heading{lvl}        : sz={sz} {b} color={col}')

    sz_doc = Counter(re.findall(r'<w:sz w:val="(\d+)"/>', d))
    print(f'正文中显式字号   : {dict(sz_doc.most_common(6))}')

    sp_after = Counter(re.findall(r'<w:spacing[^>]*w:after="(\d+)"', d))
    sp_line = Counter(re.findall(r'<w:spacing[^>]*w:line="(\d+)"', d))
    print(f'段后间距分布     : {dict(sp_after.most_common(5))}')
    print(f'行距分布         : {dict(sp_line.most_common(5))}')

    return 1 if (fail or bare_par or (cjk_hits and not allow_cjk)) else 0


if __name__ == '__main__':
    # 第 2 参起仅接受 --allow-cjk（豁免正文中文字符检查）
    extra = [a for a in sys.argv[2:] if a != '--allow-cjk']
    if len(sys.argv) < 2 or extra:
        print(__doc__)
        sys.exit(1)
    sys.exit(main(sys.argv[1]))
