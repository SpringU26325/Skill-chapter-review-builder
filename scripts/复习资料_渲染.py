# -*- coding: utf-8 -*-
"""复习资料视觉验收配套工具（不启动 PowerShell，Word COM 由 render.ps1 走 PowerShell 工具执行）。

用法：
  python 复习资料_渲染.py --prepare <docx> [<docx2> ...]    # 建本章私有过程目录并写清单
  python 复习资料_渲染.py --render [目录名] [--dpi 140]     # 逐页字数 + 空页筛查 + 渲染 PNG
  python 复习资料_渲染.py --clean [目录名]                  # 只清自己那个目录

为什么一个章节一个子目录（11-14 实测教训）：
  并行会话共用 复习资料\\_render\\p1.docx 会互相覆盖——同一轮里前 6 页是自己的资料、
  第 7 页忽然变成另一章。现在按 `_render\\<科目>_<章节>\\` 开子目录（如
  `_render\\程序设计_第1章\\`），各用各的 pN.docx / 清单 / PNG；科目必须在目录名里，
  将来两个科目同时做「第1章」才不会撞。`--clean` 只删指定或唯一的那个目录，不再
  rmtree 整个 _render（那会清掉别人正在用的过程文件）。目录不止一个时必须显式指定，不猜。

  `--prepare` 也不 rmtree 自己的目录：里面的 pN.docx 正被 Word 占着就报错退出，
  而不是删掉别人的清单、留下删不掉的副本。重跑只清自己上一轮的遗留（多余编号、
  旧 .renderlog.txt、旧 _render_cache）。
"""
import glob
import io
import os
import shutil
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import _env

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = _env.find_root()
BASE = _env.render_dir(ROOT)                   # 上级目录：复习资料\_生产\_render 或 CRB_RENDER_DIR
PS1_SRC = os.path.join(HERE, 'render.ps1')
LIST_NAME = '.renderlist.txt'
MAP_NAME = '.rendermap.txt'


_PARENT_SKIP = ('_render', '_源稿', '复习资料', '_迁移', '_locks', '_生产', '_成品')


def _plain_stem(path):
    stem = os.path.splitext(os.path.basename(path))[0]
    if stem.endswith('_复习资料'):
        stem = stem[:-len('_复习资料')]
    return stem.strip() or 'run'


def _split(path):
    """拆成 (科目, 章节)：程序设计_第1章_复习资料.docx -> ('程序设计', '第1章')。

    科目优先从文件名取；文件名没有 `<科目>_第` 结构时退回所在目录名（成品都放在
    `复习资料\\<科目>\\` 下）。两边都取不到就返回空科目，由 prepare 打警告。
    """
    stem = _plain_stem(path)
    subj, chap = '', stem
    if '_第' in stem:
        subj, rest = stem.rsplit('_第', 1)
        chap = '第' + rest
    if not subj:
        parent = os.path.basename(os.path.dirname(os.path.abspath(path)))
        if parent and parent not in _PARENT_SKIP:
            subj = parent
    return subj, chap


def _slug(path):
    """过程目录名，必须带科目：未来跨科目并发时，两个「第1章」才撞不到一起。

    程序设计_第1章_复习资料.docx -> 程序设计_第1章
    """
    subj, chap = _split(path)
    return f'{subj}_{chap}' if subj else chap


def _run_dir(name=None):
    """定位过程目录：显式名字直接用；否则必须能唯一确定，多个就报错不猜。"""
    if name:
        d = name if os.path.isabs(name) else os.path.join(BASE, name)
        if not os.path.isdir(d):
            raise SystemExit(f'目录不存在：{d}')
        return d
    if not os.path.isdir(BASE):
        raise SystemExit('渲染目录还不存在，先 --prepare')
    runs = [d for d in sorted(os.listdir(BASE))
            if os.path.isdir(os.path.join(BASE, d))
            and os.path.exists(os.path.join(BASE, d, LIST_NAME))]
    flat = os.path.exists(os.path.join(BASE, LIST_NAME))
    total = len(runs) + (1 if flat else 0)
    if total == 1:
        return BASE if flat and not runs else os.path.join(BASE, runs[0])
    if not total:
        raise SystemExit('没有待渲染的清单，先 --prepare')
    shown = runs + (['(平铺，旧式)'] if flat else [])
    raise SystemExit(
        '发现多个渲染目录，不猜你要哪个：\n  ' + '\n  '.join(shown) +
        '\n指定一个：--render <目录名>。并行会话下各用各的，别动别人的。')


def _assert_free(out):
    """目录里已有的 pN.docx 若正被 Word/预览占用，就别动——那多半是别人正在跑的一轮。

    以前这里是 `shutil.rmtree(out, ignore_errors=True)`：两个会话做同一章时，后
    prepare 的那个把人家的 pN.docx、清单、日志、PNG 一起删掉；`ignore_errors=True`
    又让「Word 占着删不掉」的副本留下、清单却被删掉，对方半截状态更难查。
    """
    busy = []
    for p in sorted(glob.glob(os.path.join(out, 'p*.docx'))):
        try:
            with open(p, 'r+b'):
                pass
        except PermissionError:
            busy.append(os.path.basename(p))
    if busy:
        raise SystemExit(
            '过程目录正被占用：%s\n  %s\n'
            '另一个会话可能正在渲染同一章（Word 打开着这些副本）。等它跑完再 --prepare，'
            '或设 CRB_RENDER_DIR 指向你自己的私有目录。' % (out, '、'.join(busy)))


def prepare(paths):
    srcs = []
    for raw in paths:
        src = os.path.abspath(raw)
        if os.path.exists(src):
            srcs.append(src)
        else:
            print('跳过（不存在）:', src)
    if not srcs:
        raise SystemExit('没有可渲染的 docx')

    subs = {_split(s)[0] for s in srcs}
    if len(subs) > 1:
        raise SystemExit(
            '一次 --prepare 只能放同一科目，这次混了：%s\n'
            '分两批跑——过程目录一科目一个，跨科目混在一起就失去隔离的意义了。'
            % '、'.join(sorted(subs)))
    if '' in subs:
        print('WARN 这批 docx 的文件名和所在目录都没带科目，过程目录可能与别的科目撞名；'
              '建议命名成 <科目>_第X章_复习资料.docx')

    out = os.path.join(BASE, _slug(srcs[0]))
    name = os.path.basename(out)
    os.makedirs(out, exist_ok=True)
    _assert_free(out)                           # 别人正在用这一章就别动
    ps1 = os.path.join(out, 'render.ps1')
    shutil.copy2(PS1_SRC, ps1)                  # ps1 靠 $PSScriptRoot 定位清单，必须同目录

    listing, mapping = [], []
    for i, src in enumerate(srcs, 1):
        d = os.path.join(out, 'p%d.docx' % i)
        p = os.path.join(out, 'p%d.pdf' % i)
        shutil.copy2(src, d)
        if os.path.exists(p):
            os.remove(p)
        listing += [d, p]
        mapping.append('p%d = %s' % (i, os.path.basename(src)))

    # 上一轮多出来的 pN（这次册数变少）要清掉，否则 --render 会把旧页一起数进去。
    # 只删自己这个目录里的 p*.docx / p*.pdf，不碰任何别的目录。
    want = len(srcs)
    for p in glob.glob(os.path.join(out, 'p*.*')):
        b = os.path.basename(p)
        core, ext = os.path.splitext(b)
        if ext not in ('.docx', '.pdf') or not core.startswith('p'):
            continue
        idx = core[1:]
        if idx.isdigit() and int(idx) > want:
            try:
                os.remove(p)
            except OSError:
                pass

    # 上一轮的日志必须删：里面可能留着 TIMEOUT| 或上一批的 ok，看着像本轮结果。
    for stale in ('.renderlog.txt',):
        p = os.path.join(out, stale)
        if os.path.exists(p):
            try:
                os.remove(p)
            except OSError:
                pass
    shutil.rmtree(os.path.join(out, '_render_cache'), ignore_errors=True)  # PNG 一起清，免新旧混页

    with open(os.path.join(out, LIST_NAME), 'w', encoding='utf-8') as f:
        f.write('\n'.join(listing))
    with open(os.path.join(out, MAP_NAME), 'w', encoding='utf-8') as f:
        f.write('\n'.join(mapping))

    print('过程目录:', out)
    print('清单:', os.path.join(out, LIST_NAME), '共', len(listing), '行')
    for m in mapping:
        print(' ', m)
    print('下一步：用 PowerShell 工具执行  & "%s"' % ps1)
    print('渲染：python 复习资料_渲染.py --render %s' % name)
    print('清理：python 复习资料_渲染.py --clean %s' % name)


def _page_risks(page, margin=8):
    """逐页挑「必须亲眼看」的风险，其余页信任指标即可。

    看着像指标微调，其实是省视觉 token 的关键：一册 10 页全看要 10 张图、
    每张上千 token；标出风险页后通常只看 2–3 张。判定用实测分布定的阈值
    （A4 595×842，版心 x∈[62,547]；字/页中位 860–1270，最高实测 1847）：

      - 空页：文字 <40 字（容器整块换页留下的空白页）
      - 越界：文本块 x1 > 560 或 x0 < 52（矩阵/长公式顶出版心，视觉上被裁）
      - 图贴底：图片底边距页面底部 <40 pt（图注会被推到下一页）
    """
    risks = []
    rect = page.rect
    txt = page.get_text().strip()
    if len(txt) < 40:
        risks.append('空页')
    try:
        for b in page.get_text('blocks'):
            if b[2] > rect.width - 35 or b[0] < 30:
                risks.append('越界')
                break
    except Exception:
        pass
    try:
        for im in page.get_image_info():
            if rect.height - im['bbox'][3] < 40:
                risks.append('图贴底')
                break
    except Exception:
        pass
    return risks


def render(name=None, dpi=140):
    import pymupdf
    out = _run_dir(name)
    listing = os.path.join(out, LIST_NAME)
    cache = os.path.join(out, '_render_cache')
    lines = [x for x in open(listing, encoding='utf-8').read().splitlines() if x.strip()]
    os.makedirs(cache, exist_ok=True)
    for i in range(0, len(lines), 2):
        pdf = lines[i + 1]
        if not os.path.exists(pdf):
            print(os.path.basename(pdf), 'MISSING')
            continue
        d = pymupdf.open(pdf)
        print('===', os.path.basename(pdf), '|', d.page_count, '页 |', os.path.getsize(pdf), 'bytes')
        watch = []
        for pi in range(d.page_count):
            page = d[pi]
            t = page.get_text().strip()
            risks = _page_risks(page)
            if risks:
                watch.append('p%02d(%s)' % (pi + 1, '/'.join(risks)))
            flag = ('  << ' + ' / '.join(risks)) if risks else ''
            print('  p%02d: %4d 字%s' % (pi + 1, len(t), flag))
            pix = page.get_pixmap(dpi=dpi)
            png = os.path.join(cache, '%s_p%02d.png' % (os.path.splitext(os.path.basename(pdf))[0], pi + 1))
            pix.save(png)
        # 只给需要看的页，避免「全册逐页看」把视觉 token 花在没问题的页上
        print('  看图清单:', '、'.join(watch) if watch else '无风险页（看首页与速查页即可）')
        d.close()
    print('PNG 输出:', cache)


def clean(name=None):
    out = _run_dir(name)
    if os.path.abspath(out) == os.path.abspath(BASE):
        # 老式平铺目录：只删过程文件，目录本身留着（它可能是别人的父目录）
        for f in (LIST_NAME, MAP_NAME, 'render.ps1', '.renderlog.txt', '.renderpid.txt'):
            p = os.path.join(out, f)
            if os.path.exists(p):
                os.remove(p)
        for pat in ('p*.docx', 'p*.pdf'):
            for p in glob.glob(os.path.join(out, pat)):
                try:
                    os.remove(p)
                except OSError:
                    pass
        shutil.rmtree(os.path.join(out, '_render_cache'), ignore_errors=True)
    else:
        shutil.rmtree(out, ignore_errors=True)
    print('已清理:', out)


if __name__ == '__main__':
    a = sys.argv[1:]
    if a and a[0] == '--prepare':
        prepare(a[1:])
    elif a and a[0] == '--render':
        dpi, rest = 140, a[1:]
        if '--dpi' in rest:
            i = rest.index('--dpi')
            dpi = int(rest[i + 1])
            del rest[i:i + 2]
        render(rest[0] if rest else None, dpi)
    elif a and a[0] == '--clean':
        clean(a[1] if len(a) > 1 else None)
    else:
        print(__doc__)
