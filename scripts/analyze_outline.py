# -*- coding: utf-8 -*-
"""Outline 排版参数分析器 —— 从样本 .docx 提取可复用的排版画像。

为什么需要它：
    课程笔记模板规范不能靠单一样本推断，也不能只看 styles.xml —— 实测样本
    （Linear Algebra Outline.docx）里 1089/1240 个段落**没有 pStyle**，排版靠
    直接格式化（direct formatting）。只看命名单一样式会得到完全错误的结论。

    本脚本按**段落类别**聚合真实生效的参数，便于跨样本交叉对比，区分：
      * 跨科目稳定不变的约定  -> 写死进生成器默认值
      * 科目间有差异的参数    -> 暴露成 CFG 开关

提取内容：
    1. 页面设置（纸张、页边距、页眉页脚引用）
    2. styles.xml：docDefaults + 各命名样式（含 basedOn 链解析）
    3. 段落分类统计：title / author / chapter / section / entry / body / list / toc / empty
       每类给出真实生效的字号、加粗、字体、颜色、段前后间距、行距、缩进、对齐
    4. 公式对象统计：m:oMath / m:oMathPara / m:jc(对齐) / m:nor / m:d / m:m
    5. 结构统计：Chapter 数、条目标题模式、列表、图表、TOC 域

用法:
    python analyze_outline.py <a.docx> [b.docx ...]
    python analyze_outline.py --json out.json <a.docx> ...
    python analyze_outline.py --compare <a.docx> [b.docx ...]   # 跨样本对照表
"""
import json
import re
import sys
import zipfile
from collections import Counter, OrderedDict

from lxml import etree

W = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
M = '{http://schemas.openxmlformats.org/officeDocument/2006/math}'

ENTRY_RE = re.compile(
    r'^\s*(Definition|Theorem|Proposition|Corollary|Lemma|Remark|Example|'
    r'Note|Recall|Axiom|Claim|Fact|Notation|Exercise)\b[\s.]*\d')
CHAPTER_RE = re.compile(r'^\s*(Chapter|CHAPTER)\s+\d', re.I)
# 条目标题里"加粗前缀"的形态：Definition 1.1 / Theorem 2.7. / Proposition 1.9.
ENTRY_LABEL_RE = re.compile(
    r'^((?:Definition|Theorem|Proposition|Corollary|Lemma|Remark|Example|'
    r'Note|Recall|Axiom|Claim|Fact|Notation|Exercise)\s*[\d.]+\.?)')

CATS = ('title', 'author', 'chapter', 'section', 'entry', 'body', 'list',
        'toc', 'empty')


def _w(el, name):
    return el.get(W + name) if el is not None else None


def _m(el, name):
    return el.get(M + name) if el is not None else None


# ---------------------------------------------------------------- rPr / pPr

def rpr_profile(rpr):
    """从一个 w:rPr 提取关心的属性（b/i 只在显式出现时记录）。"""
    if rpr is None:
        return {}
    out = {}
    rf = rpr.find(W + 'rFonts')
    if rf is not None:
        for k in ('ascii', 'hAnsi', 'eastAsia', 'cs'):
            v = rf.get(W + k)
            if v:
                out.setdefault('fonts', {})[k] = v
        # 主题字体（ITVC 这类样本用 asciiTheme 而非 ascii）
        for k in ('asciiTheme', 'hAnsiTheme'):
            v = rf.get(W + k)
            if v:
                out.setdefault('themeFonts', {})[k] = v
    for tag in ('sz', 'szCs'):
        e = rpr.find(W + tag)
        if e is not None:
            out[tag] = _w(e, 'val')
    for tag in ('b', 'bCs', 'i', 'iCs'):
        e = rpr.find(W + tag)
        if e is not None:
            # <w:b/> 无 val 表示开启；<w:b w:val="0"/> 表示显式关闭
            out[tag] = _w(e, 'val') or 'on'
    c = rpr.find(W + 'color')
    if c is not None:
        out['color'] = _w(c, 'val')
    return out


def ppr_profile(ppr):
    if ppr is None:
        return {}
    out = {}
    sp = ppr.find(W + 'spacing')
    if sp is not None:
        for k in ('before', 'after', 'line', 'lineRule', 'beforeLines',
                  'afterLines'):
            v = sp.get(W + k)
            if v:
                out['spacing_' + k] = v
    ind = ppr.find(W + 'ind')
    if ind is not None:
        for k in ('left', 'start', 'hanging', 'firstLine'):
            v = ind.get(W + k)
            if v:
                out['ind_' + k] = v
    jc = ppr.find(W + 'jc')
    if jc is not None:
        out['jc'] = _w(jc, 'val')
    st = ppr.find(W + 'pStyle')
    if st is not None:
        out['pStyle'] = _w(st, 'val')
    if ppr.find(W + 'numPr') is not None:
        out['numPr'] = True
    # 段落标记自身的 rPr（样本用它在无 run 的段落上定字号）
    mark = rpr_profile(ppr.find(W + 'rPr'))
    if mark:
        out['mark'] = mark
    return out


def load_styles(z):
    """返回 {styleId: {name, type, basedOn, rPr, pPr}} 与 docDefaults。"""
    out = {'styles': {}, 'docDefaults': {}}
    if 'word/styles.xml' not in z.namelist():
        return out
    st = etree.fromstring(z.read('word/styles.xml'))
    dd = st.find(W + 'docDefaults')
    if dd is not None:
        rd = dd.find(W + 'rPrDefault')
        pd = dd.find(W + 'pPrDefault')
        out['docDefaults'] = {
            'rPr': rpr_profile(rd.find(W + 'rPr') if rd is not None else None),
            'pPr': ppr_profile(pd.find(W + 'pPr') if pd is not None else None),
        }
    for s in st.findall(W + 'style'):
        sid = _w(s, 'styleId')
        if not sid:
            continue
        bo = s.find(W + 'basedOn')
        out['styles'][sid] = {
            'type': _w(s, 'type'),
            'name': _w(s.find(W + 'name'), 'val'),
            'basedOn': _w(bo, 'val') if bo is not None else None,
            'rPr': rpr_profile(s.find(W + 'rPr')),
            'pPr': ppr_profile(s.find(W + 'pPr')),
        }
    return out


def resolve(sty, sid, kind='rPr', _seen=None):
    """沿 basedOn 链合并样式属性：子覆盖父，最后叠 docDefaults。"""
    if _seen is None:
        _seen = set()
    acc = {}
    chain = []
    cur = sid
    while cur and cur in sty['styles'] and cur not in _seen:
        _seen.add(cur)
        chain.append(cur)
        acc = {**sty['styles'][cur].get(kind, {}), **acc}
        cur = sty['styles'][cur].get('basedOn')
    dflt = sty.get('docDefaults', {}).get(kind, {})
    acc = {**dflt, **acc}
    return acc, chain


# ---------------------------------------------------------------- 段落分类

def classify(p, pstyle_name, text, has_math):
    if not text and not has_math:
        return 'empty'
    if pstyle_name and pstyle_name.lower().startswith('toc'):
        return 'toc'
    if pstyle_name == 'heading 1' or CHAPTER_RE.match(text):
        return 'chapter'
    if pstyle_name in ('heading 2', 'heading 3', 'heading 4'):
        return 'section'
    if ENTRY_RE.match(text):
        return 'entry'
    if pstyle_name == 'List Paragraph':
        return 'list'
    return 'body'


def eff_sz(cat_runs, pprof, style_rpr, default_sz):
    """段落真实生效字号：run 级 > ¶标记 > 样式链 > docDefaults。"""
    for rp in cat_runs:
        if rp.get('sz'):
            return rp['sz'], 'run'
    if pprof.get('mark', {}).get('sz'):
        return pprof['mark']['sz'], 'mark'
    if style_rpr.get('sz'):
        return style_rpr['sz'], 'style'
    if default_sz:
        return default_sz, 'default'
    return None, 'none'


def analyze(path):
    z = zipfile.ZipFile(path)
    doc = etree.fromstring(z.read('word/document.xml'))
    sty = load_styles(z)
    rep = {'file': path}

    # ---------- 1. 页面设置 ----------
    sect = doc.find('.//' + W + 'sectPr')
    if sect is not None:
        pg = sect.find(W + 'pgSz')
        mar = sect.find(W + 'pgMar')
        rep['page'] = {
            'w': _w(pg, 'w'), 'h': _w(pg, 'h'), 'orient': _w(pg, 'orient'),
            'margin': {k: _w(mar, k) for k in
                       ('top', 'bottom', 'left', 'right', 'header', 'footer')},
            'headers': len(sect.findall(W + 'headerReference')),
            'footers': len(sect.findall(W + 'footerReference')),
        }
    rep['docDefaults'] = sty['docDefaults']
    # 只保留命名样式，便于人工核对
    rep['styles'] = OrderedDict(
        (sid, s) for sid, s in sty['styles'].items()
        if s['name'] and not s['name'].lower().startswith(('toc', 'grid')))

    dflt_sz = sty['docDefaults'].get('rPr', {}).get('sz')
    dflt_ppr = sty['docDefaults'].get('pPr', {})

    # ---------- 3. 段落分类统计 ----------
    body = doc.find(W + 'body')
    paras = body.findall(W + 'p') if body is not None else doc.findall(W + 'p')

    cat_data = {c: {
        'n': 0,
        'sz': Counter(), 'sz_src': Counter(), 'bold': Counter(),
        'font': Counter(), 'color': Counter(),
        'after': Counter(), 'before': Counter(), 'line': Counter(),
        'ind': Counter(), 'jc': Counter(), 'pstyle': Counter(),
        'label_bold': Counter(), 'examples': [],
    } for c in CATS}

    n_list_style = n_numpr = 0
    entry_types = Counter()
    first_text_idx = None          # v16：正文里第一个非空段落 = 文档大标题

    for idx, p in enumerate(paras):
        ppr = p.find(W + 'pPr')
        prof = ppr_profile(ppr)
        sid = prof.get('pStyle')
        sname = sty['styles'].get(sid, {}).get('name') if sid else None
        style_rpr, _ = resolve(sty, sid, 'rPr') if sid else (
            sty['docDefaults'].get('rPr', {}), [])
        style_ppr, _ = resolve(sty, sid, 'pPr') if sid else (
            sty['docDefaults'].get('pPr', {}), [])

        runs = p.findall(W + 'r')
        run_profiles = [rpr_profile(r.find(W + 'rPr')) for r in runs]
        txt_parts = [''.join(t.text or '' for t in r.findall(W + 't'))
                     for r in runs]
        math_txt = []
        math_runs = []
        for mr in p.iter(M + 'r'):
            mt = mr.find(M + 't')
            if mt is not None and mt.text:
                math_txt.append(mt.text)
            math_runs.append(rpr_profile(mr.find(W + 'rPr')))
        text = ''.join(txt_parts).strip()
        full = text or ''.join(math_txt).strip()
        has_math = bool(math_runs)

        cat = classify(p, sname, text, has_math)
        # v16：文档大标题 / 作者行。样本段落 0 = 大标题（16pt 加粗、无样式），
        # 段落 1 = 作者行（jc=right、Algerian）。不单列的话会被算进 body，
        # 把 16pt/14pt、Algerian、right 这些值混进正文统计，污染"规范只来自实测"。
        if cat != 'empty' and first_text_idx is None:
            first_text_idx = idx
        if cat == 'body' and idx == first_text_idx:
            cat = 'title'
        elif (cat == 'body' and first_text_idx is not None
              and idx == first_text_idx + 1 and prof.get('jc') == 'right'):
            cat = 'author'
        d = cat_data[cat]
        d['n'] += 1
        d['pstyle'][sname or '(none)'] += 1

        sz, src = eff_sz(run_profiles, prof, style_rpr, dflt_sz)
        if sz:
            d['sz'][sz] += 1
            d['sz_src'][src] += 1

        # 加粗：段落里**第一个非空文本 run** 的 b（条目标题就是靠它加粗的）
        bold = '(no run)'
        for rp, t in zip(run_profiles, txt_parts):
            if t.strip():
                bold = rp.get('b', 'off')
                break
        d['bold'][bold] += 1

        # 条目标题的"加粗前缀"形态：Definition 1.1 加粗 / 其后标题不加粗
        if cat == 'entry':
            m = ENTRY_LABEL_RE.match(text)
            if m and len(runs) >= 2:
                r0 = run_profiles[0].get('b', 'off')
                r1 = run_profiles[1].get('b', 'off')
                d['label_bold'][f'label={r0}/rest={r1}'] += 1
            entry_types[ENTRY_LABEL_RE.match(text).group(1).split()[0]
                        if m else '?'] += 1

        for rp in run_profiles:
            for v in (rp.get('fonts') or {}).values():
                d['font'][v] += 1
            for v in (rp.get('themeFonts') or {}).values():
                d['font']['theme:' + v] += 1
            if rp.get('color'):
                d['color'][rp['color']] += 1

        after = prof.get('spacing_after', style_ppr.get('spacing_after'))
        before = prof.get('spacing_before', style_ppr.get('spacing_before'))
        line = prof.get('spacing_line', style_ppr.get('spacing_line'))
        lrule = prof.get('spacing_lineRule', style_ppr.get('spacing_lineRule'))
        ind = prof.get('ind_left', style_ppr.get('ind_left'))
        jc = prof.get('jc', style_ppr.get('jc'))
        if after:
            d['after'][after] += 1
        if before:
            d['before'][before] += 1
        if line:
            d['line'][f'{line}/{lrule}'] += 1
        if ind:
            d['ind'][ind] += 1
        if jc:
            d['jc'][jc] += 1
        if prof.get('numPr'):
            n_numpr += 1
        if sname == 'List Paragraph':
            n_list_style += 1
        if len(d['examples']) < 3 and full:
            d['examples'].append(full[:64])

    rep['paragraphs'] = {
        'total': len(paras),
        'categories': {c: {
            'n': cat_data[c]['n'],
            'sz': dict(cat_data[c]['sz'].most_common(6)),
            'sz_src': dict(cat_data[c]['sz_src'].most_common()),
            'bold': dict(cat_data[c]['bold'].most_common(4)),
            'font': dict(cat_data[c]['font'].most_common(5)),
            'color': dict(cat_data[c]['color'].most_common(4)),
            'after': dict(cat_data[c]['after'].most_common(4)),
            'before': dict(cat_data[c]['before'].most_common(3)),
            'line': dict(cat_data[c]['line'].most_common(4)),
            'ind': dict(cat_data[c]['ind'].most_common(4)),
            'jc': dict(cat_data[c]['jc'].most_common(4)),
            'pstyle': dict(cat_data[c]['pstyle'].most_common(5)),
            'label_bold': dict(cat_data[c]['label_bold'].most_common(4)),
            'examples': cat_data[c]['examples'],
        } for c in CATS},
        'list_style_paragraphs': n_list_style,
        'numpr_paragraphs': n_numpr,
    }

    # ---------- 4. 公式对象 ----------
    omath = list(doc.iter(M + 'oMath'))
    opara = list(doc.iter(M + 'oMathPara'))
    jc = Counter()
    for op in opara:
        pr = op.find(M + 'oMathParaPr')
        if pr is None:
            jc['(no oMathParaPr)'] += 1
            continue
        j = pr.find(M + 'jc')
        # 关键：OMML 的 m:jc 用 m:val（math 命名空间），不是 w:val
        jc[(_m(j, 'val') if j is not None else '(no jc)') or '(empty)'] += 1
    rep['math'] = {
        'oMath': len(omath), 'oMathPara': len(opara),
        'm_nor': len(list(doc.iter(M + 'nor'))),
        'oMathPara_jc': dict(jc.most_common()),
        'delimiter_m_d': len(list(doc.iter(M + 'd'))),
        'matrix_m_m': len(list(doc.iter(M + 'm'))),
        'm_sSub': len(list(doc.iter(M + 'sSub'))),
        'm_frac': len(list(doc.iter(M + 'f'))),
    }

    # ---------- 5. 结构 ----------
    n_chapter = cat_data['chapter']['n']
    n_entry = cat_data['entry']['n']
    rep['structure'] = {
        'chapter_headings': n_chapter,
        'section_headings': cat_data['section']['n'],
        'entry_titles': n_entry,
        'entry_types': dict(entry_types.most_common()),
        'tables': len(list(doc.iter(W + 'tbl'))),
        'drawings': (len(list(doc.iter(W + 'drawing')))
                     + len(list(doc.iter(W + 'pict')))),
        'has_toc_field': b'TOC' in z.read('word/document.xml'),
    }
    return rep


# ---------------------------------------------------------------- 输出

SZ = lambda v: f'{int(v)/2:g}pt' if v and str(v).isdigit() else str(v)


def fmt(rep):
    L = []
    a = L.append
    a('=' * 78)
    a('文件: ' + rep['file'].split('/')[-1])
    a('=' * 78)
    pg = rep.get('page', {})
    if pg:
        mm = pg['margin']
        a(f"页面 {pg['w']}x{pg['h']} twips  页边距 上{mm['top']} 下{mm['bottom']} "
          f"左{mm['left']} 右{mm['right']}  页眉{pg['headers']} 页脚{pg['footers']}")
    dd = rep.get('docDefaults', {})
    a(f"docDefaults rPr: {dd.get('rPr')}")
    a(f"docDefaults pPr: {dd.get('pPr')}")
    a('-- 命名样式（含 basedOn 链）--')
    for sid, s in rep.get('styles', {}).items():
        nm = s['name']
        if nm in ('Normal', 'heading 1', 'heading 2', 'heading 3',
                  'heading 4', 'List Paragraph'):
            a(f"   {sid:6} {nm:16} basedOn={s['basedOn']!s:6} "
              f"sz={s['rPr'].get('sz')} b={s['rPr'].get('b')} "
              f"color={s['rPr'].get('color')} pPr={s['pPr']}")
    p = rep.get('paragraphs', {})
    a(f"段落 {p.get('total')}  列表样式 {p.get('list_style_paragraphs')}  "
      f"编号段 {p.get('numpr_paragraphs')}")
    for c in CATS:
        d = p['categories'][c]
        if not d['n']:
            continue
        a(f"  [{c:8}] n={d['n']:<4} sz={d['sz']} (src={d['sz_src']}) "
          f"bold={d['bold']}")
        a(f"            font={d['font']} color={d['color']}")
        a(f"            after={d['after']} before={d['before']} "
          f"line={d['line']}")
        a(f"            jc={d['jc']} ind={d['ind']} pstyle={d['pstyle']}")
        if d['label_bold']:
            a(f"            entry label/rest bold={d['label_bold']}")
        if d['examples']:
            a(f"            e.g. {d['examples'][0]!r}")
    m = rep.get('math', {})
    a(f"公式 oMath={m.get('oMath')} oMathPara={m.get('oMathPara')} "
      f"m:nor={m.get('m_nor')} m:d={m.get('delimiter_m_d')} "
      f"m:m={m.get('matrix_m_m')} m:sSub={m.get('m_sSub')} "
      f"m:f={m.get('m_frac')}")
    a(f"oMathPara 对齐 {m.get('oMathPara_jc')}")
    s = rep.get('structure', {})
    a(f"Chapter {s.get('chapter_headings')}  节 {s.get('section_headings')}  "
      f"条目标题 {s.get('entry_titles')} {s.get('entry_types')}")
    a(f"表格 {s.get('tables')}  图 {s.get('drawings')}  TOC域 {s.get('has_toc_field')}")
    return '\n'.join(L)


def compare(reps):
    """跨样本对照：每个类别只报最高频字号/加粗/行距，便于横向比对。"""
    L = []
    a = L.append
    hdr = f"{'样本':28s} {'Ch':>4} {'Sec':>4} {'条目':>4} {'正文':>5} " \
          f"{'章sz':>7} {'节sz':>7} {'条目sz':>8} {'正文sz':>8} " \
          f"{'条目粗':>7} {'行距':>10}"
    a(hdr)
    a('-' * len(hdr))
    for r in reps:
        s = r.get('structure', {})
        c = r.get('paragraphs', {}).get('categories', {})

        def top(cat, key, dflt='-'):
            dd = c.get(cat, {}).get(key, {})
            return next(iter(dd), dflt) if dd else dflt

        def topsz(cat):
            v = top(cat, 'sz', None)
            return SZ(v) if v else '-'
        name = r['file'].split('/')[-1].replace(' Outline.docx', '')[:27]
        a(f"{name:28s} {s.get('chapter_headings'):>4} "
          f"{s.get('section_headings'):>4} {s.get('entry_titles'):>4} "
          f"{c.get('body', {}).get('n', 0):>5} "
          f"{topsz('chapter'):>7} {topsz('section'):>7} "
          f"{topsz('entry'):>8} {topsz('body'):>8} "
          f"{top('entry', 'bold'):>7} {top('body', 'line'):>10}")
    return '\n'.join(L)


if __name__ == '__main__':
    args = sys.argv[1:]
    js = None
    cmp_mode = False
    if args and args[0] == '--json':
        js, args = args[1], args[2:]
    if args and args[0] == '--compare':
        cmp_mode, args = True, args[1:]
    if not args:
        print(__doc__)
        sys.exit(1)
    reps = [analyze(p) for p in args]
    if cmp_mode:
        print(compare(reps))
    else:
        for r in reps:
            print(fmt(r))
            print()
    if js:
        with open(js, 'w', encoding='utf-8') as f:
            json.dump(reps, f, ensure_ascii=False, indent=1)
        print('JSON ->', js)
