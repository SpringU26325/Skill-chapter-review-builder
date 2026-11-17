# -*- coding: utf-8 -*-
"""从主干 PDF 裁示意图，落到 _源稿\\figures\\<科目>\\ 并登记图索引。

只做一件事：把主干里「本身就是图」的内容原样裁出来。表格与公式不进这条通道——
表格写成 md 文字表格，公式走 LaTeX 重建，两者都可编辑、可搜索、放大不糊。

用法（<py> 指项目级解释器，见 SKILL.md「运行环境与依赖」）：
  # 1) 定位：渲染带坐标网格的整页图，按网格线读矩形（扫描版主干没有矢量信息，
  #    整页就是一张位图，无法自动算出子区域的框，只能目视量）
  "<py>" 复习资料_裁图.py probe <主干.pdf> --page 90 --grid 50

  # 2) 裁切并登记
  "<py>" 复习资料_裁图.py crop <主干.pdf> --page 90 --box 233,267,346,318 \
        --subject 物理 --file 第3章_例3.12_球在纸上 --printed 74 --note "例 3.12 用图"

  # 3) 查已登记
  "<py>" 复习资料_裁图.py list [--subject 物理]

--page 一律用「物理页」（台账口径，1 基），内部按 page-1 取页；印刷页另用
--printed 记录，扫描版读不出页码，只能目视填。
"""
import argparse
import datetime
import os
import sys
import io

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

import _env

ENV_HINT = ('先把课件转成 PDF 再裁：PowerPoint 另存为 PDF，或走 render.md 里的 '
            'PowerPoint COM 导出；本脚本只认 PDF。')


def _open_pdf(path):
    import pymupdf as fitz
    if not os.path.isfile(path):
        raise SystemExit('找不到主干文件：%s' % path)
    if not path.lower().endswith('.pdf'):
        raise SystemExit('只支持 PDF。%s\n%s' % (path, ENV_HINT))
    return fitz.open(path)


def _page_of(doc, page_no):
    """物理页（1 基）→ Page 对象。越界直接报，不静默夹取。"""
    n = len(doc)
    if not (1 <= page_no <= n):
        raise SystemExit('物理页 %d 超出范围：这本 PDF 共 %d 页（1–%d）。' % (page_no, n, n))
    return doc[page_no - 1]


def _effective_dpi(page):
    """扫描页整页是一张位图；按「原图像素宽 ÷ 页面宽(pt) × 72」算有效 dpi。
    低于 200 时裁出来的字会发虚，提醒换更清的源或缩小裁切范围。"""
    try:
        for img in page.get_images(full=True):
            xref, w = img[0], img[2]
            if w and page.rect.width:
                dpi = w / page.rect.width * 72.0
                if dpi > 0:
                    return dpi
    except Exception:
        pass
    return None


def cmd_probe(args):
    import pymupdf as fitz

    doc = _open_pdf(args.pdf)
    page = _page_of(doc, args.page)
    w, h = page.rect.width, page.rect.height

    out = fitz.open()
    canvas = out.new_page(width=w, height=h)
    canvas.insert_image(canvas.rect, pixmap=page.get_pixmap(dpi=args.dpi))

    step = args.grid
    shape = canvas.new_shape()
    x = step
    while x < w:
        shape.draw_line(fitz.Point(x, 0), fitz.Point(x, h))
        x += step
    y = step
    while y < h:
        shape.draw_line(fitz.Point(0, y), fitz.Point(w, y))
        y += step
    shape.finish(color=(1, 0, 0), width=0.35)
    shape.commit()

    x = step
    while x < w:
        for y0 in (10, h / 2, h - 6):
            canvas.insert_text((x + 1, y0), str(int(x)), fontsize=6, color=(0.85, 0, 0))
        x += step
    y = step
    while y < h:
        for x0 in (2, w / 2):
            canvas.insert_text((x0, y - 2), str(int(y)), fontsize=6, color=(0.85, 0, 0))
        y += step

    dest = os.path.join(os.environ.get('TEMP') or '.', 'crb_probe_p%d.png' % args.page)
    canvas.get_pixmap(dpi=args.dpi).save(dest)
    dpi = _effective_dpi(page)
    print('PAGE_SIZE=%.1f x %.1f pt' % (w, h))
    print('GRID=%d pt' % step)
    print('EFFECTIVE_DPI=%s' % ('%.0f' % dpi if dpi else 'unknown'))
    if dpi and dpi < 200:
        print('WARN 原图有效 dpi 只有 %.0f，裁出来会发虚；换更清的源或缩小裁切范围。' % dpi)
    print('PROBE_FILE=%s' % dest)
    out.close()
    doc.close()


def _fig_dir(root, subj):
    d = os.path.join(_env.src_dir(root), 'figures', subj)
    os.makedirs(d, exist_ok=True)
    return d


def cmd_crop(args):
    import pymupdf as fitz

    box = [float(v) for v in args.box.replace(' ', '').split(',')]
    if len(box) != 4:
        raise SystemExit('--box 要四个数：x0,y0,x1,y1（单位 pt，原点在左上）。')
    x0, y0, x1, y1 = box
    if x1 <= x0 or y1 <= y0:
        raise SystemExit('--box 的 x1/y1 必须大于 x0/y0：%s' % args.box)

    doc = _open_pdf(args.pdf)
    page = _page_of(doc, args.page)
    if x1 > page.rect.width + 0.5 or y1 > page.rect.height + 0.5:
        raise SystemExit('裁切框超出页面（页面 %.0f x %.0f pt）：%s'
                         % (page.rect.width, page.rect.height, args.box))

    root = _env.require_root()
    folder = _fig_dir(root, args.subject)
    name = args.file if args.file.lower().endswith('.png') else args.file + '.png'
    dest = os.path.join(folder, name)
    if os.path.exists(dest) and not args.force:
        raise SystemExit('同名图已存在：%s\n确认要覆盖就加 --force。' % dest)

    clip = fitz.Rect(x0, y0, x1, y1)
    pm = page.get_pixmap(clip=clip, dpi=args.dpi, alpha=False)
    pm.save(dest)

    cm_w = (x1 - x0) / 72.0 * 2.54
    cm_h = (y1 - y0) / 72.0 * 2.54
    print('OUT=%s' % dest)
    print('PIXELS=%d x %d' % (pm.width, pm.height))
    print('PRINT_SIZE=%.2f x %.2f cm' % (cm_w, cm_h))
    src_dpi = _effective_dpi(page)
    if src_dpi and src_dpi < 200:
        print('WARN 主干整页有效 dpi %.0f，这张图是按原图放大到 %d dpi 的，放大后会虚。'
              % (src_dpi, args.dpi))
    if cm_w > 16.6:
        print('WARN 宽 %.1f cm 超过版心 16.6 cm，插入时要写 {w=} 收窄。' % cm_w)

    _register(root, args, name, x0, y0, x1, y1)
    doc.close()


INDEX_HEAD = """# 图索引

`figures\\<科目>\\` 下每张图一条记录。主干换版后按「物理页」回查原图；裁错了重裁并更新本行。
`key` 即文件名（不含 .png），md 里写 `![图注](figures/<科目>/<key>.png)`。

| 文件 | 主干 | 物理页 | 印刷页 | 裁切框 (pt) | 内容 | 登记日期 |
|---|---|---|---|---|---|---|
"""


def _register(root, args, name, x0, y0, x1, y1):
    """同名图重裁时替换原行，不追加——否则一版图会留下七八条裁切框记录，
    下次核对分不清哪条才是当前用的那张。"""
    idx = os.path.join(_env.src_dir(root), 'figures', '_图索引.md')
    printed = str(args.printed) if args.printed else '—'
    main = os.path.basename(args.pdf)
    key = '%s/%s' % (args.subject, name)
    row = '| %s | %s | %d | %s | %.0f,%.0f,%.0f,%.0f | %s | %s |\n' % (
        key, main, args.page, printed, x0, y0, x1, y1,
        (args.note or '—').replace('|', '\\vert'),
        datetime.date.today().isoformat())

    lines = []
    if os.path.exists(idx):
        with open(idx, encoding='utf-8') as f:
            lines = f.read().splitlines()
    if not lines:
        lines = INDEX_HEAD.rstrip('\n').splitlines()

    out, replaced = [], False
    for ln in lines:
        if ln.startswith('| %s |' % key):
            if not replaced:
                out.append(row.rstrip('\n'))
                replaced = True
            continue
        out.append(ln)
    if not replaced:
        out.append(row.rstrip('\n'))

    with open(idx, 'w', encoding='utf-8', newline='\n') as f:
        f.write('\n'.join(out) + '\n')
    print('INDEX=%s%s' % (idx, ' (replaced)' if replaced else ''))


def cmd_list(args):
    root = _env.require_root()
    idx = os.path.join(_env.src_dir(root), 'figures', '_图索引.md')
    if not os.path.exists(idx):
        print('还没有任何登记：%s' % idx)
        return
    with open(idx, encoding='utf-8') as f:
        for ln in f:
            if args.subject and ln.startswith('| ') and not ln.startswith('| %s/' % args.subject):
                continue
            sys.stdout.write(ln)


def main():
    ap = argparse.ArgumentParser(description='主干 PDF 裁图与图索引')
    sub = ap.add_subparsers(dest='cmd', required=True)

    p1 = sub.add_parser('probe', help='渲染带网格的整页图，用来量裁切框')
    p1.add_argument('pdf')
    p1.add_argument('--page', type=int, required=True, help='物理页（1 基）')
    p1.add_argument('--dpi', type=int, default=110)
    p1.add_argument('--grid', type=int, default=50, help='网格间距，单位 pt，默认 50')
    p1.set_defaults(func=cmd_probe)

    p2 = sub.add_parser('crop', help='裁切并登记')
    p2.add_argument('pdf')
    p2.add_argument('--page', type=int, required=True, help='物理页（1 基）')
    p2.add_argument('--box', required=True, help='x0,y0,x1,y1（pt，原点左上）')
    p2.add_argument('--subject', required=True, help='科目名，与 复习资料\\<科目>\\ 一致')
    p2.add_argument('--file', required=True, help='文件名（不含扩展名）')
    p2.add_argument('--printed', help='印刷页，扫描版读不出，目视填')
    p2.add_argument('--note', help='这张图是什么内容')
    p2.add_argument('--dpi', type=int, default=300)
    p2.add_argument('--force', action='store_true', help='覆盖同名图')
    p2.set_defaults(func=cmd_crop)

    p3 = sub.add_parser('list', help='列出已登记的图')
    p3.add_argument('--subject')
    p3.set_defaults(func=cmd_list)

    args = ap.parse_args()
    args.func(args)


if __name__ == '__main__':
    main()
