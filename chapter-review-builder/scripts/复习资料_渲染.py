# -*- coding: utf-8 -*-
"""复习资料视觉验收配套工具（不启动 PowerShell，Word COM 由 render.ps1 走 PowerShell 工具执行）。

用法：
  python 复习资料_渲染.py --prepare <docx> [<docx2> ...]   # 复制副本到过程目录并写清单
  python 复习资料_渲染.py --render [--dpi 140]             # 逐页字数 + 空页筛查 + 渲染 PNG
  python 复习资料_渲染.py --clean                          # 清理临时副本/清单/PNG

为什么分两段：踩坑清单——Bash 调 PowerShell 被安全策略拦截；且 Word COM 必须由 PowerShell 工具
直接执行。过程目录默认在工作区 复习资料/_render（中文路径已实测可用），--clean 即清场。
"""
import io
import os
import shutil
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import _env

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = _env.find_root()
OUT = _env.render_dir(ROOT)                    # 过程文件目录：docx/pdf/png 的暂存地
LIST = os.path.join(OUT, '.renderlist.txt')
MAP = os.path.join(OUT, '.rendermap.txt')
CACHE = os.path.join(OUT, '_render_cache')
PS1_SRC = os.path.join(HERE, 'render.ps1')
PS1 = os.path.join(OUT, 'render.ps1')


def prepare(paths):
    os.makedirs(OUT, exist_ok=True)
    shutil.copy2(PS1_SRC, PS1)                 # 把 ps1 一起放进过程目录，靠 $PSScriptRoot 定位清单
    lines, mapping = [], []
    for i, raw in enumerate(paths, 1):
        src = os.path.abspath(raw)
        if not os.path.exists(src):
            print('跳过（不存在）:', src)
            continue
        d = os.path.join(OUT, 'p%d.docx' % i)
        p = os.path.join(OUT, 'p%d.pdf' % i)
        shutil.copy2(src, d)
        if os.path.exists(p):
            os.remove(p)
        lines += [d, p]
        mapping.append('p%d = %s' % (i, os.path.basename(src)))
    with open(LIST, 'w', encoding='utf-8') as f:
        f.write('\n'.join(lines))
    with open(MAP, 'w', encoding='utf-8') as f:
        f.write('\n'.join(mapping))
    print('清单:', LIST, '共', len(lines), '行')
    for m in mapping:
        print(' ', m)
    print('下一步：用 PowerShell 工具执行  & "%s"' % PS1)


def render(dpi=140):
    import pymupdf
    if not os.path.exists(LIST):
        print('清单不存在，先 --prepare')
        raise SystemExit(1)
    lines = [x for x in open(LIST, encoding='utf-8').read().splitlines() if x.strip()]
    os.makedirs(CACHE, exist_ok=True)
    for i in range(0, len(lines), 2):
        pdf = lines[i + 1]
        if not os.path.exists(pdf):
            print(os.path.basename(pdf), 'MISSING')
            continue
        d = pymupdf.open(pdf)
        print('===', os.path.basename(pdf), '|', d.page_count, '页 |', os.path.getsize(pdf), 'bytes')
        for pi in range(d.page_count):
            t = d[pi].get_text().strip()
            flag = '  << 疑似空页' if len(t) < 40 else ''
            print('  p%02d: %4d 字%s' % (pi + 1, len(t), flag))
            pix = d[pi].get_pixmap(dpi=dpi)
            png = os.path.join(CACHE, '%s_p%02d.png' % (os.path.splitext(os.path.basename(pdf))[0], pi + 1))
            pix.save(png)
        d.close()
    print('PNG 输出:', CACHE)


def clean():
    for p in (LIST, MAP):
        if os.path.exists(p):
            os.remove(p)
    shutil.rmtree(OUT, ignore_errors=True)
    shutil.rmtree(CACHE, ignore_errors=True)
    print('临时副本/清单/PNG 已清理')


if __name__ == '__main__':
    a = sys.argv[1:]
    if a and a[0] == '--prepare':
        prepare(a[1:])
    elif a and a[0] == '--render':
        dpi = 140
        if '--dpi' in a:
            dpi = int(a[a.index('--dpi') + 1])
        render(dpi)
    elif a and a[0] == '--clean':
        clean()
    else:
        print(__doc__)
