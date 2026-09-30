# -*- coding: utf-8 -*-
"""括号伸缩度探针（连通域标记法）—— 判断 Word 是否把括号按内容高度拉伸。

为什么需要它：
    pymupdf 的 char['bbox'] 报的是**字体行框**，括号被纵向拉伸后仍报同一高度
    （实测 (a/b) 与 (x+y) 都是 11.1pt），所以 `render_pdf_preview.py --brackets`
    **不能**用来判伸缩。本脚本用像素级连通域隔离括号本体，再与"本行全部墨迹高度"
    相比，得到可靠判据。

原理：
    1. 以「行 bbox」为基准开一个上下各留 PAD 的大窗；
    2. 二值化后做 8-邻接连通域标记；
    3. 只保留 bbox 与「本行 y 区间」相交的连通域 —— 相邻行的墨迹是独立连通域，
       天然被排除（这是本方法相对"矩形裁剪窗"的关键优势）；
    4. 括号 = 包含括号字符 bbox 中心的那个连通域；
    5. 比值 = 括号墨迹高度 / 本行全部墨迹高度。**>= 0.85 判合格。**

两个已知假阳性（用矩形裁剪窗时会踩）：
    * 横向留 ±3pt  → 把右侧矩阵单元格的数字计入每行墨迹宽度，
      波动周期恰好等于矩阵行数，看起来像"拼接括号的周期性接缝"；
    * 纵向留 ±26pt → 把相邻段落的余墨计入，中间出现空白段，
      把 max(mid) - min(mid) 撑大，也会被读成"接缝"。

用法:
    python probe_bracket_stretch.py <pdf> [page] [--min-ratio 0.85]

退出码: 0 = 全部合格；1 = 存在不合格项（可用于 CI / 自检）。
"""
import sys
from collections import deque

import pymupdf

DPI = 600
PAD = 30            # 行 bbox 上下各留多少 pt 开窗
OPEN_CHARS = '([{'  # 只考察左定界符（右定界符形状对称，量一个即可）


def label(pm, thr=128):
    """8-邻接连通域标记，返回 [(x0, y0, x1, y1, area)]（像素坐标）。"""
    W, H, step, buf = pm.width, pm.height, pm.stride, pm.samples
    ink = [[buf[y * step + x] < thr for x in range(W)] for y in range(H)]
    seen = [[False] * W for _ in range(H)]
    comps = []
    for sy in range(H):
        for sx in range(W):
            if not ink[sy][sx] or seen[sy][sx]:
                continue
            q = deque([(sx, sy)])
            seen[sy][sx] = True
            x0 = x1 = sx
            y0 = y1 = sy
            area = 0
            while q:
                x, y = q.popleft()
                area += 1
                if x < x0: x0 = x
                if x > x1: x1 = x
                if y < y0: y0 = y
                if y > y1: y1 = y
                for dy in (-1, 0, 1):
                    for dx in (-1, 0, 1):
                        nx, ny = x + dx, y + dy
                        if 0 <= nx < W and 0 <= ny < H and ink[ny][nx] \
                                and not seen[ny][nx]:
                            seen[ny][nx] = True
                            q.append((nx, ny))
            comps.append((x0, y0, x1, y1, area))
    return comps


def pt(px):
    return px / DPI * 72


def measure(pdf, page=0):
    """返回 [(字符, 括号高pt, 行内墨迹高pt, 比值, 行文本)]。"""
    doc = pymupdf.open(pdf)
    pg = doc[page]
    out = []
    for blk in pg.get_text('rawdict')['blocks']:
        if blk.get('type') != 0:
            continue
        for ln in blk['lines']:
            lb = ln['bbox']
            txt = ''.join(c['c'] for s in ln['spans'] for c in s['chars'])[:46]
            dels = [c for s in ln['spans'] for c in s['chars']
                    if c['c'] in OPEN_CHARS]
            if not dels:
                continue
            clip = pymupdf.Rect(lb[0] - 2, lb[1] - PAD, lb[2] + 2, lb[3] + PAD)
            pm = pg.get_pixmap(dpi=DPI, clip=clip, colorspace=pymupdf.csGRAY)
            comps = label(pm)
            # 本行 y 区间（像素行）；留 6% DPI 的容差吸收基线偏差
            ry0 = (lb[1] - clip.y0) / 72 * DPI
            ry1 = (lb[3] - clip.y0) / 72 * DPI
            tol = 0.06 * DPI
            mine = [c for c in comps if c[3] >= ry0 - tol and c[1] <= ry1 + tol]
            if not mine:
                continue
            fh = max(c[3] for c in mine) - min(c[1] for c in mine) + 1
            c0 = dels[0]
            cx = ((c0['bbox'][0] + c0['bbox'][2]) / 2 - clip.x0) / 72 * DPI
            cy = ((c0['bbox'][1] + c0['bbox'][3]) / 2 - clip.y0) / 72 * DPI
            best, bd = None, 1e18
            for c in mine:
                ddx = 0 if c[0] <= cx <= c[2] else min(abs(c[0] - cx),
                                                       abs(c[2] - cx))
                ddy = 0 if c[1] <= cy <= c[3] else min(abs(c[1] - cy),
                                                       abs(c[3] - cy))
                if ddx * ddx + ddy * ddy < bd:
                    bd = ddx * ddx + ddy * ddy
                    best = c
            if best is None:
                continue
            bh = best[3] - best[1] + 1
            out.append((c0['c'], pt(bh), pt(fh), bh / fh, txt))
    return out


def main():
    args = [a for a in sys.argv[1:] if not a.startswith('--')]
    if not args:
        print(__doc__)
        return 1
    min_ratio = 0.85
    for a in sys.argv[1:]:
        if a.startswith('--min-ratio'):
            min_ratio = float(a.split('=')[1] if '=' in a
                              else sys.argv[sys.argv.index(a) + 1])
    pdf = args[0]
    page = int(args[1]) if len(args) > 1 else 0

    bad = 0
    rows = measure(pdf, page)
    if not rows:
        print(f'（{pdf} 第 {page} 页未找到定界符）')
        return 0
    for ch, bh, fh, ratio, txt in rows:
        flag = 'OK  ' if ratio >= min_ratio else 'SMALL'
        if ratio < min_ratio:
            bad += 1
        print(f'{ch} 括号={bh:5.1f}pt 行内墨迹={fh:5.1f}pt 比={ratio:.2f} {flag} | {txt}')
    print(f'\n合计 {len(rows)} 例，不合格 {bad} 例（阈值 {min_ratio}）')
    return 1 if bad else 0


if __name__ == '__main__':
    sys.exit(main())
