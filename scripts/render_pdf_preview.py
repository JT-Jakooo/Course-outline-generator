"""把 PDF 逐页渲染成 PNG，用于**目视核对排版**。

用法:
    python render_pdf_preview.py <in.pdf> <out_dir> [dpi]

为什么需要它:
    排版问题（字号不一、莫名居中、行被撑断、文字斜体）只看 docx 的 XML 是判断不出来的，
    必须看渲染结果。标准做法是先用 `docx_to_pdf.ps1` 让 Word 导出 PDF，
    再用本脚本转成图片逐页检查。

附带能力:
    `--lines` 模式会 dump 每行文字的 bbox 与字体 flags，
    可用坐标客观判断对齐方式（居中 / 左对齐 / 缩进），
    用 flags 的 bit 1 判断是否斜体。

    `--brackets` 模式会量测每个定界符字形的实际高度，
    用来客观验证"括号是否随内容伸缩"——
    同一种括号在矮内容行与高内容行里的高度差就是伸缩幅度。

依赖: pymupdf
    "$PY" -m pip install pymupdf
"""
import sys
from pathlib import Path

# 可伸缩定界符候选字符（含 OMML 默认圆括号）
_DELIMS = set('()[]{}|‖⟨⟩⌊⌋⌈⌉⟮⟯')


def dump_lines(pdf):
    """打印每行文字的 bbox 与字体信息，用于量测对齐。"""
    import pymupdf
    doc = pymupdf.open(pdf)
    for pno in range(doc.page_count):
        page = doc[pno]
        print(f'===== page {pno + 1}  (width={page.rect.width:.0f}) =====')
        for blk in page.get_text('dict')['blocks']:
            if blk.get('type') != 0:
                continue
            for ln in blk['lines']:
                txt = ''.join(s['text'] for s in ln['spans']).strip()
                if not txt:
                    continue
                x0, y0, x1, y1 = ln['bbox']
                fonts = sorted({s['font'] for s in ln['spans']})
                flags = sorted({s['flags'] for s in ln['spans']})
                italic = any(f & 2 for f in flags)
                print(f'x0={x0:6.1f} x1={x1:6.1f} y={y0:6.1f} '
                      f'{"ITALIC" if italic else "      "} | {txt[:78]}')
                print(f'      fonts={fonts} flags={flags}')


def dump_brackets(pdf):
    """量测每个定界符字形的实际渲染高度（pt），验证括号伸缩。

    原理：同一个字符在矮内容（如 `(x+y)`）与高内容（如 `(a/b)`）里，
    如果渲染高度不同，就说明 Word 真的按内容高度做了伸缩。
    """
    import pymupdf
    doc = pymupdf.open(pdf)
    for pno in range(doc.page_count):
        page = doc[pno]
        print(f'===== page {pno + 1} =====')
        for blk in page.get_text('rawdict')['blocks']:
            if blk.get('type') != 0:
                continue
            for ln in blk['lines']:
                chars = [c for s in ln['spans'] for c in s['chars']]
                txt = ''.join(c['c'] for c in chars).strip()
                if not txt:
                    continue
                hits = [(c['c'], c['bbox'][3] - c['bbox'][1])
                        for c in chars if c['c'] in _DELIMS]
                if not hits:
                    continue
                desc = '  '.join(f'{ch}={h:5.1f}' for ch, h in hits)
                print(f'  {txt[:60]:62s} | {desc}')


def render(pdf, out_dir, dpi=150):
    import pymupdf
    out = Path(out_dir)
    out.mkdir(parents=True, exist_ok=True)
    doc = pymupdf.open(pdf)
    print('pages =', doc.page_count)
    for i, page in enumerate(doc):
        fn = out / f'page{i + 1}.png'
        page.get_pixmap(dpi=dpi).save(str(fn))
        print('saved', fn)


if __name__ == '__main__':
    if len(sys.argv) < 3:
        print(__doc__)
        sys.exit(1)
    if sys.argv[1] == '--lines':
        dump_lines(sys.argv[2])
    elif sys.argv[1] == '--brackets':
        dump_brackets(sys.argv[2])
    else:
        render(sys.argv[1], sys.argv[2],
               int(sys.argv[3]) if len(sys.argv) > 3 else 150)
