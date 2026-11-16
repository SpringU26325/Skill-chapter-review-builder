# -*- coding: utf-8 -*-
"""原件盘点：扫各科目目录，报四类问题。

  1. 未知后缀 —— 不在已知后缀表里，必须按 references/sources.md 判档或问用户
  2. 未登记   —— 磁盘上有、`_原件.md` 里没有
  3. 疑似重复 —— 同大小 + 同首页文本（或同大小同名干）
  4. 登记失效 —— `_原件.md` 里写了、磁盘上没有

用法（在复习资料工作区根目录执行）：
    python 复习资料_原件.py              # 全部科目
    python 复习资料_原件.py 线性代数     # 单科目
    python 复习资料_原件.py --write      # 把未登记文件追加进 _原件.md（状态「未使用」）
    python 复习资料_原件.py --quiet      # 只报问题，不印统计

退出码：0 无 ERROR 级问题（未知后缀、登记失效），1 有。
"""

import io
import os
import re
import sys
import warnings

warnings.filterwarnings('ignore', category=FutureWarning)

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import _env

# 后缀 -> (档位, 通道摘要)；与 references/sources.md 同步维护
EXT = {
    # A 档：直读
    '.pdf': ('A', 'pdf，按是否有文字层分流'),
    '.md': ('A', '直接读'),
    '.txt': ('A', '直接读'),
    '.tex': ('A', '读源码，公式即源码'),
    '.epub': ('A', 'zip 解包读 OEBPS/*.xhtml'),
    '.html': ('A', '直接读；MathJax 保留 \\(…\\)'),
    '.htm': ('A', '同上'),
    '.mhtml': ('A', '按 MIME 段取 text/html'),
    '.mht': ('A', '同上'),
    '.srt': ('A', '字幕，按时间戳定位'),
    '.vtt': ('A', '同上'),
    '.ass': ('A', '同上'),
    '.csv': ('A', '表格，走 sheetagent 或另存'),
    '.json': ('A', '直接读'),
    '.yaml': ('A', '直接读'),
    '.yml': ('A', '直接读'),
    '.ini': ('A', '直接读'),
    '.c': ('A', '源码，仅对齐语法点，不作题源'),
    '.cpp': ('A', '同上'),
    '.h': ('A', '同上'),
    '.py': ('A', '同上'),
    '.java': ('A', '同上'),
    # A 档但需专用库
    '.docx': ('A', '直接读'),
    '.xlsx': ('A', '需 openpyxl 或 sheetagent'),
    '.xls': ('A', '需 openpyxl 或 sheetagent'),
    '.pptx': ('A', 'slideN.xml 的 a:t 定位，公式看渲染图'),
    # B 档：需转换
    '.doc': ('B', 'Word COM SaveAs2(路径,16) 转 docx'),
    '.wps': ('B', '不转：请用户另存 docx/pdf，禁用 WPS COM'),
    '.rar': ('B', '7z x 解到 _迁移\\conv\\'),
    '.7z': ('B', '同上'),
    '.zip': ('B', '解包后按内层后缀重判；原始包勿动'),
    '.tif': ('B', '多页 tiff，fitz 逐页 get_pixmap'),
    '.tiff': ('B', '同上'),
    '.heic': ('B', '需 pillow-heif，否则请用户导 jpg'),
    '.heif': ('B', '同上'),
    '.dat': ('B', '微信加密缓存，需解密；否则请用户另存'),
    '.djvu': ('B', '需 djvulibre，通常没有'),
    '.apkg': ('B', 'zip+sqlite，notes.flds 取正反面'),
    '.ppt': ('B', '旧 ppt 无 XML 文字层，转 pptx 或目视'),
    # C 档：只能看图
    '.png': ('C', '逐张看图'),
    '.jpg': ('C', '逐张看图'),
    '.jpeg': ('C', '逐张看图'),
    '.webp': ('C', '逐张看图'),
    '.bmp': ('C', '逐张看图'),
    '.gif': ('C', '逐张看图'),
    # C 档但多半不可处理
    '.note': ('C', 'GoodNotes/Notability 原生，只能请用户导 pdf'),
    # D 档
    '.mp4': ('D', '抽帧或转写，不整段看'),
    '.mkv': ('D', '同上'),
    '.mp3': ('D', '转写，标时间点'),
    '.m4a': ('D', '同上'),
    '.wav': ('D', '同上'),
    '.url': ('D', '网盘/网页链接，先下载'),
}

SKIP_DIRS = {'_成品', '_生产', '_旧版单册'}
SKIP_SUFFIX = {'.lock'}
SKIP_PREFIX = '~$'

HEAD = ['文件', '类型', '来源 / 日期', '覆盖', '内容可信度', '状态', '已用']


def walk_files(d):
    """递归列出 d 下文件，跳过 _成品/_生产/_旧版单册 与临时件。返回相对路径列表。"""
    out = []
    for dirpath, dirs, files in os.walk(d):
        dirs[:] = [x for x in dirs if x not in SKIP_DIRS and not x.startswith('~')]
        for f in files:
            if f.startswith(SKIP_PREFIX) or os.path.splitext(f)[1].lower() in SKIP_SUFFIX:
                continue
            if f == '_原件.md':
                continue
            full = os.path.join(dirpath, f)
            out.append(os.path.relpath(full, d))
    return sorted(out)


def fingerprint(full):
    """首页/开头文本指纹，用于判重。取不到返回 None。"""
    ext = os.path.splitext(full)[1].lower()
    try:
        if ext == '.pdf':
            import fitz
            doc = fitz.open(full)
            t = doc[0].get_text() if doc.page_count else ''
            doc.close()
            return re.sub(r'\s+', '', t)[:300] or None
        if ext == '.docx':
            from docx import Document
            ps = [p.text for p in Document(full).paragraphs[:20]]
            return re.sub(r'\s+', '', ''.join(ps))[:300] or None
    except Exception:
        return None
    return None


def read_registry(mdpath):
    """解析 _原件.md 首个表格里第一列的文件名。返回 (集合, 是否已有「已用」列)。

    只认第一个 `## ` 标题之前的表格——「非原件」等附录表不是原件登记。
    """
    if not os.path.exists(mdpath):
        return set(), False
    txt = io.open(mdpath, encoding='utf-8').read()
    names, has_used = set(), False
    for line in txt.split('\n'):
        if line.startswith('## '):
            break
        if not line.startswith('|'):
            continue
        cells = [c.strip() for c in line.strip().strip('|').split('|')]
        if len(cells) >= 6 and cells[-1] == '已用':
            has_used = True
        m = re.match(r'^`([^`]+)`\s*$', cells[0]) if cells else None
        if m:
            names.add(m.group(1).replace('/', '\\').strip('\\'))
    return names, has_used


def subject_dirs(root, only=None):
    base = os.path.join(root, _env.WORKSPACE_MARK)
    d = []
    for n in sorted(os.listdir(base)):
        p = os.path.join(base, n)
        if not os.path.isdir(p) or n.startswith('_'):
            continue
        if only and n not in only:
            continue
        d.append((n, p))
    return d


def scan(only=None, quiet=False, write=False):
    root = _env.require_root()
    unk, unreg, dup, dead = [], [], [], []
    total = 0
    for name, path in subject_dirs(root, only):
        files = walk_files(path)
        total += len(files)
        names, has_used = read_registry(os.path.join(path, '_原件.md'))
        on_disk = set()
        for rel in files:
            ext = os.path.splitext(rel)[1].lower()
            norm = rel.replace('/', '\\')
            on_disk.add(norm)
            full = os.path.join(path, rel)
            info = EXT.get(ext)
            if info is None:
                unk.append((name, rel, ext, os.path.getsize(full)))
            if norm not in names and not any(norm.startswith(n + '\\') for n in names):
                unreg.append((name, rel, ext, os.path.getsize(full), info))
        # 登记失效：表里有、磁盘没有（目录项按前缀匹配）
        for n in sorted(names):
            if n in on_disk:
                continue
            if any(x.startswith(n + '\\') for x in on_disk):
                continue
            dead.append((name, n))
        # 判重：同大小 + 同指纹
        bysize = {}
        for rel in files:
            full = os.path.join(path, rel)
            bysize.setdefault(os.path.getsize(full), []).append(rel)
        for size, group in bysize.items():
            if len(group) < 2:
                continue
            fps = {}
            for rel in group:
                fp = fingerprint(os.path.join(path, rel)) or ''
                fps.setdefault(fp, []).append(rel)
            for fp, g in fps.items():
                if len(g) > 1 or (fp == '' and len(group) > 1 and len(g) == len(group)):
                    dup.append((name, size, g, bool(fp)))
        if write and not os.path.exists(os.path.join(path, '_原件.md')):
            write_registry(os.path.join(path, '_原件.md'), name, [], create_only=True)

    # 输出
    print('原件盘点  根目录 %s' % root)
    print('科目 %d 个，文件 %d 个（已排除 _成品/_生产/临时件）' %
          (len(subject_dirs(root, only)), total))

    if unk:
        print('\n[未知后缀] %d 个 —— 读 references/sources.md 判档，不许猜' % len(unk))
        for name, rel, ext, size in unk:
            print('  [%s] %s  %s  %s' % (name, rel, ext or '(无后缀)', sizeof(size)))
    else:
        print('\n[未知后缀] 0 个')

    if unreg:
        print('\n[未登记] %d 个 —— 磁盘有、_原件.md 无' % len(unreg))
        for name, rel, ext, size, info in unreg:
            print('  [%s] %s  %s档 %s' % (name, rel, info[0] if info else '?', sizeof(size)))
    else:
        print('\n[未登记] 0 个')

    if dup:
        print('\n[疑似重复] %d 组 —— 保留一份为主，其余标「重复，未用」，不删' % len(dup))
        for name, size, g, hasfp in dup:
            print('  [%s] %s  %s' % (name, sizeof(size), '同首页文本' if hasfp else '仅同大小'))
            for x in g:
                print('        %s' % x)
    else:
        print('\n[疑似重复] 0 组')

    if dead:
        print('\n[登记失效] %d 条 —— 表里有、磁盘没有' % len(dead))
        for name, n in dead:
            print('  [%s] %s' % (name, n))
    else:
        print('\n[登记失效] 0 条')

    if write and unreg:
        for name, path in subject_dirs(root, only):
            rows = [(rel, info[0] if info else '?') for s, rel, ext, sz, info in unreg if s == name]
            if rows:
                write_registry(os.path.join(path, '_原件.md'), name, rows)
        print('\n已把 %d 条未登记行追加进各科目 _原件.md（状态「未使用」，其余列待填）' % len(unreg))

    err = len(unk) + len(dead)
    print('\n%s  ERROR %d（未知后缀 %d、登记失效 %d）'
          % ('存在' if err else '无', err, len(unk), len(dead)))
    return 1 if err else 0


def sizeof(n):
    for u in ('B', 'KB', 'MB'):
        if n < 1024 or u == 'MB':
            return '%.0f%s' % (n, u) if u == 'B' else '%.1f%s' % (n, u)
        n /= 1024.0


def write_registry(mdpath, subject, rows, create_only=False):
    """追加未登记行；表不存在时建模板。"""
    hdr = ('# %s 原件索引\n\n'
           '> 本科目的非教材来源资料。新拿到资料先登记再动：文件 → 来源 → 覆盖 → 可信度 → 状态。\n'
           '> 可信度口径见 skill 的 `references/intake.md` 第三节；类型判档见 `references/sources.md`。\n\n'
           '| 文件 | 类型 | 来源 / 日期 | 覆盖 | 内容可信度 | 状态 | 已用 |\n'
           '|---|---|---|---|---|---|---|\n') % subject
    if not os.path.exists(mdpath):
        io.open(mdpath, 'w', encoding='utf-8', newline='').write(hdr)
    if create_only or not rows:
        return
    txt = io.open(mdpath, encoding='utf-8', newline='').read()
    lines = txt.split('\n')
    # 插到首个表格的最后一行之后，不能写到文件尾——尾部常有「说明」列表
    end = None
    for i, l in enumerate(lines):
        if l.startswith('## '):
            break
        if l.startswith('|'):
            end = i
    add = ['| `%s` | 待填（%s档） | 待填 | 待填 | 待填 | 未使用 | 未用 |'
           % (rel.replace('/', '\\'), tag if tag != '?' else '未知') for rel, tag in rows]
    if end is None:
        lines.extend(add)
    else:
        lines[end + 1:end + 1] = add
    io.open(mdpath, 'w', encoding='utf-8', newline='').write('\n'.join(lines))


if __name__ == '__main__':
    args = [a for a in sys.argv[1:] if not a.startswith('-')]
    sys.exit(scan(only=args or None,
                  quiet='--quiet' in sys.argv,
                  write='--write' in sys.argv))
