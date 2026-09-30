# -*- coding: utf-8 -*-
"""Re-add the `<w:i w:val="0"/>` upright hint to formula-embedded normal text.

Why this exists
---------------
`build_note.py` writes `<w:i w:val="0"/>` into every OMML run that carries
`<m:nor/>` (normal text inside a formula, e.g. "Lagrange notation"). Word itself
renders `m:nor` upright without it, but third-party renderers -- the Tencent Docs
web preview among them -- default those runs to italic, which mangles the body
text inside formulas.

`finalize_docx.ps1` drives Word to build the TOC / page-number fields, and Word's
own save strips the hint as redundant. This script puts it back, so the delivered
.docx keeps a correct cached TOC *and* stays readable in an online preview.

Usage
-----
    python restore_upright.py <in.docx> [out.docx]

With no `out.docx` the file is rewritten in place. The rest of the package is
copied through byte-for-byte; only `word/document.xml` is touched.
"""
import shutil
import sys
import zipfile

from lxml import etree

W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
M = 'http://schemas.openxmlformats.org/officeDocument/2006/math'


def restore(src, dst=None):
    dst = dst or src
    zin = zipfile.ZipFile(src)
    entries = [(i, zin.read(i.filename)) for i in zin.infolist()]
    zin.close()

    doc_name = 'word/document.xml'
    payload = dict((i.filename, b) for i, b in entries)
    if doc_name not in payload:
        raise SystemExit('not a docx (no %s): %s' % (doc_name, src))

    root = etree.fromstring(payload[doc_name])
    fixed = 0
    for run in root.iter('{%s}r' % M):
        # only runs whose rPr marks them as "normal text inside a formula"
        if run.find('{%s}rPr/{%s}nor' % (M, M)) is None:
            continue
        rpr = run.find('{%s}rPr' % W)
        if rpr is None:
            rpr = etree.Element('{%s}rPr' % W)
            mrpr = run.find('{%s}rPr' % M)
            if mrpr is not None:
                mrpr.addnext(rpr)
            else:
                run.insert(0, rpr)
        it = rpr.find('{%s}i' % W)
        if it is None:
            it = etree.Element('{%s}i' % W)
            # CT_RPr order: rStyle, rFonts, b, bCs, i, iCs, ... -> right after rFonts
            rf = rpr.find('{%s}rFonts' % W)
            if rf is not None:
                rf.addnext(it)
            else:
                rpr.insert(0, it)
            fixed += 1
        it.set('{%s}val' % W, '0')

    payload[doc_name] = etree.tostring(root, xml_declaration=True,
                                       encoding='UTF-8', standalone=True)

    with zipfile.ZipFile(dst, 'w', zipfile.ZIP_DEFLATED) as zout:
        for info, data in entries:
            zout.writestr(info, payload.get(info.filename, data))
    return fixed


if __name__ == '__main__':
    if len(sys.argv) not in (2, 3):
        print(__doc__)
        sys.exit(1)
    n = restore(*sys.argv[1:])
    print('restored %d upright hint(s) -> %s' % (n, sys.argv[-1]))
