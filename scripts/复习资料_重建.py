# -*- coding: utf-8 -*-
"""扫描 md 源稿，增量重建复习资料（可选转 PDF 报页数）。

用法：
    python 复习资料_重建.py                  # 只重建 md 比 docx 新的（增量）
    python 复习资料_重建.py --dry            # 预演：只打印待重建清单，不写文件、不加锁
    python 复习资料_重建.py --all            # 全部重建（改了引擎后用这个）
    python 复习资料_重建.py 高等数学          # 只处理该科目（子串匹配，会含该科全部章节）
    python 复习资料_重建.py 高等数学 --all --pdf
    python 复习资料_重建.py --only 程序设计_第1章 --pdf   # 只这一册，不连坐别的章
    python 复习资料_重建.py --root <工作区根>  # 手动指定工作区（默认从当前目录上溯找「复习资料」）

并行会话提醒：科目关键字是子串匹配，「程序设计」会命中该科所有章节；别人正在改的
章节若 md 比 docx 新，也会被你一起重出。只关心自己那册时用 --only <章节名>（精确到
stem，可带「上/下/中」后缀），它优先于位置参数。

为什么这样设计（多科目并发安全）：
  1. **不维护清单文件**：输出路径由命名约定推导（<科目>_第X章.md →
     复习资料/<科目>/<同名>_复习资料.docx），页眉写在 md 的 @header 里。
     没有共享可变配置，两个会话就不可能抢同一个文件。
  2. **文件锁**：写 docx / 调 Word 前抢 .build.lock，两个会话同时跑会排队，
     不会互相覆盖，也不会抢 Word COM 实例。
  3. **默认增量**：各科只重建自己改过的章节，互不干扰。
"""
import os, sys, io, time, glob, subprocess
from contextlib import contextmanager

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import _env

HERE = os.path.dirname(os.path.abspath(__file__))              # 本脚本所在目录
BUILDER = os.path.join(HERE, '复习资料_build.py')
PY = sys.executable

# 以下四项在 main() 里按 --root / 当前目录确定（脚本不猜安装位置，也不写死绝对路径）
ROOT = None          # 工作区根
SRC = None           # 源稿目录：<工作区>\复习资料\_生产\_源稿
LOCK_DIR = None      # 锁目录：<工作区>\复习资料\_生产\_locks
PDF_LIST = None      # 转 PDF 用的临时清单（放锁目录，不污染包）


@contextmanager
def _file_lock(path, timeout=900, what='构建'):
    """跨进程互斥：msvcrt 字节锁，进程退出自动释放（不会留下死锁）。"""
    import msvcrt
    fd = os.open(path, os.O_CREAT | os.O_RDWR)
    t0 = time.time()
    try:
        while True:
            try:
                msvcrt.locking(fd, msvcrt.LK_NBLCK, 1)
                break
            except OSError:
                if time.time() - t0 > timeout:
                    raise TimeoutError(f'等待{what}锁超时：另一个会话正在占用，稍后重试')
                time.sleep(0.5)
        yield
    finally:
        try:
            msvcrt.locking(fd, msvcrt.LK_UNLCK, 1)
        except OSError:
            pass
        os.close(fd)


def build_lock(subj):
    """按科目分锁：线代与高数可以同时构建，互不等待（写的是不同文件）。"""
    return _file_lock(os.path.join(LOCK_DIR, f'.build_{subj}.lock'), what=f'{subj}构建')


def pdf_lock():
    """全局锁：Word 是单实例资源，转 PDF 必须跨科目串行。"""
    return _file_lock(os.path.join(LOCK_DIR, '.pdf.lock'), what='Word 转 PDF')


def discover():
    """按命名约定发现所有章节源稿，推导输出路径。"""
    items = []
    for md in sorted(glob.glob(os.path.join(SRC, '*_第*.md'))):
        base = os.path.basename(md)[:-3]                       # 线性代数_第1章
        subj = base.split('_')[0]
        out = os.path.join(_env.out_dir(ROOT, subj), base + '_复习资料.docx')
        items.append((subj, md, out))
    return items


def build_one(md, out, subj):
    cmd = [PY, BUILDER, md, out]
    notation = os.path.join(SRC, subj + '_notation.md')
    if os.path.exists(notation):
        cmd.append('')          # 占位：HEADER 留空，让引擎读 md 的 @header
        cmd.append(notation)
    # 锁已经在这里拿着（build_lock），子进程不要再抢同一把：msvcrt 是跨进程锁，
    # 子进程重复加锁会自锁 900 秒然后失败。用环境变量告诉它「锁在祖先进程手上」。
    env = dict(os.environ, CRB_LOCK_HELD=subj)
    r = subprocess.run(cmd, capture_output=True, text=True, encoding='utf-8',
                       errors='replace', env=env)
    if r.returncode != 0:
        tail = (r.stderr or r.stdout or '').strip().splitlines()[-3:]
        return False, ' / '.join(tail)
    return True, ''


def to_pdf(docx_list):
    """一个 Word 实例内串行转换，路径经 UTF-8 清单文件传递（避免命令行中文乱码）。

    清单名带 pid：`.pdf_list.txt` 是固定名，并行会话虽然被 .pdf.lock 串行，
    但一旦异常路径（PS 超时、Ctrl-C）漏删，下一个会话就会读到上一个人的清单。
    """
    try:
        with open(PDF_LIST, 'w', encoding='utf-8') as f:
            f.write('\n'.join(docx_list))
        ps = (
            "$ErrorActionPreference='Stop';"
            "$w=New-Object -ComObject Word.Application;$w.Visible=$false;$w.DisplayAlerts=0;"
            "try{"
            f"$ps=Get-Content -LiteralPath '{PDF_LIST}' -Encoding UTF8;"
            "foreach($p in $ps){if(-not $p.Trim()){continue};"
            "$d=$w.Documents.Open($p,$false,$true);"
            "$n=$d.ComputeStatistics(2);$c=$d.ComputeStatistics(3);"
            "$d.Close($false);"
            # 只回传数字，路径由 Python 按顺序对应——避免 PowerShell 输出编码把中文路径弄成乱码
            "Write-Output ($n.ToString()+'|'+$c.ToString())}"
            "}finally{$w.Quit();[System.Runtime.InteropServices.Marshal]::ReleaseComObject($w)|Out-Null}"
        )
        r = subprocess.run(['powershell', '-NoProfile', '-NonInteractive', '-Command', ps],
                           capture_output=True, text=True, encoding='utf-8', errors='replace')
    finally:
        try:
            os.remove(PDF_LIST)
        except OSError:
            pass
    nums = [ln.strip() for ln in (r.stdout or '').splitlines() if ln.strip().count('|') == 1]
    res = {}
    for path, ln in zip(docx_list, nums):
        try:
            pages, chars = (int(x) for x in ln.split('|'))
        except ValueError:
            continue
        res[path] = (pages, chars)
    if r.returncode != 0 and not res:
        print('  PDF 转换失败:', (r.stderr or '').strip()[:200])
    return res


def _stem(md):
    """程序设计_第1章.md -> 程序设计_第1章"""
    return os.path.basename(md)[:-3]


def _only_match(stem, only):
    """--only 精确到章节 stem，但允许把分册一起带上。

    `程序设计_第1章` 命中 第1章 / 第1章上 / 第1章下，不命中 第2章。
    与位置参数的子串匹配不同，这里不做 substring：`程序设计` 不会命中任何章节
    （要整科就干脆不给 --only，直接用位置参数）。
    """
    if stem == only:
        return True
    if not stem.startswith(only):
        return False
    return stem[len(only):] in ('上', '下', '中')


def main():
    global ROOT, SRC, LOCK_DIR, PDF_LIST
    argv = sys.argv[1:]
    root_arg = only = None
    for flag in ('--root', '--only'):
        if flag in argv:                        # 先摘掉带值参数，别让它混进位置参数
            i = argv.index(flag)
            if i + 1 >= len(argv):
                raise SystemExit(f'{flag} 需要一个参数')
            val = argv[i + 1]
            del argv[i:i + 2]
            if flag == '--root':
                root_arg = val
            else:
                only = val.strip()
    flags = {a for a in argv if a.startswith('--')}
    args = [a for a in argv if not a.startswith('--')]
    kw = args[0] if args else ''
    force = '--all' in flags
    want_pdf = '--pdf' in flags
    dry = '--dry' in flags

    ROOT = os.path.abspath(root_arg) if root_arg else _env.require_root()
    SRC = _env.src_dir(ROOT)
    LOCK_DIR = _env.lock_dir(ROOT)
    PDF_LIST = os.path.join(LOCK_DIR, f'.pdf_list.{os.getpid()}.txt')
    if not os.path.isdir(SRC):
        print(f'源稿目录不存在：{SRC}')
        print('新科目接入时，先把 md 放进该目录（文件名 <科目>_第X章.md）')
        return

    all_items = discover()
    if only:
        items = [it for it in all_items if _only_match(_stem(it[1]), only)]
        if not items:
            print(f'--only「{only}」没有匹配的章节。现有源稿：')
            for it in all_items[:20]:
                print('  ' + _stem(it[1]))
            return
        print(f'--only 命中 {len(items)} 份：' + '、'.join(_stem(it[1]) for it in items))
    else:
        items = [it for it in all_items if kw in it[1]]
        if items:
            names = [_stem(it[1]) for it in items]
            shown = '、'.join(names[:8]) + (f' …等 {len(names)} 份' if len(names) > 8 else '')
            print(f'关键字「{kw or "（空）"}」命中 {len(items)} 份：{shown}')
            if not kw:
                print('  未给关键字＝全部科目。并行会话下请用科目名或 --only 收窄。')
    if not items:
        print('没有找到匹配的章节源稿（文件名应为 <科目>_第X章.md）')
        return

    todo = []
    for subj, md, out in items:
        # 记号表是每章共用的输入（构建时注入「记号约定」框），它变新了也要重建，
        # 否则补完 notation.md 再跑增量会「无需重建」，成品里的记号框还是旧的。
        src_mtime = os.path.getmtime(md)
        nota = os.path.join(SRC, subj + '_notation.md')
        if os.path.exists(nota):
            src_mtime = max(src_mtime, os.path.getmtime(nota))
        if force or not os.path.exists(out) or src_mtime > os.path.getmtime(out):
            todo.append((subj, md, out))

    if not todo:
        print(f'无需重建（{len(items)} 份都已是最新）。改过引擎请加 --all。')
    else:
        print(f'待重建 {len(todo)} 份 / 共 {len(items)} 份')
        for subj, md, out in todo:
            print(f'  - {_stem(md)}')
        if want_pdf and not force and not only and len(todo) > 1:
            print(f'  WARN --pdf 会对这 {len(todo)} 份全部统计页数；只关心一册请加'
                  f' --only {_stem(todo[0][1])}')

    if dry:
        print('\n【--dry 预演】不写文件、不加锁：')
        for subj, md, out in todo:
            print(f'  {os.path.basename(md)[:-3]}  ->  {os.path.relpath(out, ROOT)}')
        if not todo:
            print('  无')
        return

    ok = fail = 0
    done_paths = []
    for subj, md, out in todo:
        with build_lock(subj):                 # 按科目分锁，不同科目可并行
            good, msg = build_one(md, out, subj)
        name = os.path.basename(md)[:-3]
        if good:
            print(f'  OK   {name}')
            ok += 1
            done_paths.append(out)
        else:
            print(f'  FAIL {name}  {msg}')
            fail += 1

    if want_pdf and done_paths:
        print('\n转 PDF 统计（Word 单实例，跨科目串行）：')
        with pdf_lock():
            stats = to_pdf(done_paths)
        for p, (pages, chars) in stats.items():
            flag = '  << 超 15 页' if pages > 15 else ''
            print(f'  {os.path.basename(p)[:-5]:32s} PAGES={pages:3d} CHARS={chars}{flag}')

    print(f'\n完成：成功 {ok}，失败 {fail}')


if __name__ == '__main__':
    main()
