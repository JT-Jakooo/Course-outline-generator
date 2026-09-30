"""检测 OneDrive 文件是否已在本机同步（云端占位符首次读取会触发自动下载）。

用法:  python check_onedrive.py [目录路径]
省略路径时使用默认 OneDrive 大学目录。
"""
import ctypes
import os
import sys

OD = sys.argv[1] if len(sys.argv) > 1 else \
    r'C:\Users\Admin\OneDrive - The University of Manchester\University'
OFFLINE = 0x1000
RECALL_OPEN = 0x40000
RECALL_DATA = 0x400000
SPARSE = 0x200
REPARSE = 0x400


def attrs(a):
    out = []
    if a & RECALL_DATA:
        out.append('云端占位')
    if a & RECALL_OPEN:
        out.append('需下载')
    if a & OFFLINE:
        out.append('离线')
    if a & SPARSE:
        out.append('稀疏')
    if a & REPARSE:
        out.append('重解析')
    return ','.join(out) or '-'


rows = []
for root, dirs, files in os.walk(OD):
    for f in files:
        p = os.path.join(root, f)
        try:
            sz = os.path.getsize(p)
        except OSError:
            sz = -1
        a = ctypes.windll.kernel32.GetFileAttributesW(p)
        cloud = bool(a > 0 and (a & RECALL_DATA or a & RECALL_OPEN))
        rows.append((sz, p, hex(a) if a > 0 else 'ERR', attrs(a), cloud))

rows.sort(key=lambda r: (not r[4], r[1]))
print(f'总文件数: {len(rows)}')
cloud = [r for r in rows if r[4]]
print(f'仅云端(不可直接读): {len(cloud)}')
print(f'已在本机: {len(rows) - len(cloud)}')
print()
print('--- 样例（前 8 个）---')
for r in rows[:8]:
    tag = 'CLOUD' if r[4] else 'LOCAL'
    print(f'{tag} {r[0]:>10d}  {r[2]:<10s} {r[3]:<12s} {r[1][len(OD)+1:]}')
if cloud:
    print()
    print('--- 仅云端样例 ---')
    for r in cloud[:8]:
        print(f'CLOUD {r[0]:>10d}  {r[3]:<12s} {r[1][len(OD)+1:]}')
