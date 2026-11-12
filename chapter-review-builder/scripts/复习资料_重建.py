# -*- coding: utf-8 -*-
"""扫描 md 源稿，增量重建复习资料（可选转 PDF 报页数）。

用法：
    python 复习资料_重建.py                  # 只重建 md 比 docx 新的（增量）
    python 复习资料_重建.py --dry            # 预演：只打印待重建清单，不写文件、不加锁
    python 复习资料_重建.py --all            # 全部重建（改了引擎后用这个）
    python 复习资料_重建.py 高等数学          # 只处理该科目
    python 复习资料_重建.py 高等数学 --all --pdf
    python 复习资料_重建.py --root <工作区根>  # 手动指定工作区（默认从当前目录上溯找「复习资料」）

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
SRC = None           # 源稿目录：<工作区>\复习资料\_源稿
LOCK_DIR = None      # 锁目录：<工作区>\复习资料\_locks
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
        out = os.path.join(ROOT, '复习资料', subj, base + '_复习资料.docx')
        items.append((subj, md, out))
    return items


def build_one(md, out, subj):
    cmd = [PY, BUILDER, md, out]
    notation = os.path.join(SRC, subj + '_notation.md')
    if os.path.exists(notation):
        cmd.append('')          # 占位：HEADER 留空，让引擎读 md 的 @header
        cmd.append(notation)
    r = subprocess.run(cmd, capture_output=True, text=True, encoding='utf-8', errors='replace')
    if r.returncode != 0:
        tail = (r.stderr or r.stdout or '').strip().splitlines()[-3:]
        return False, ' / '.join(tail)
    return True, ''


def to_pdf(docx_list):
    """一个 Word 实例内串行转换，路径经 UTF-8 清单文件传递（避免命令行中文乱码）。"""
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
    os.remove(PDF_LIST)
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


def main():
    global ROOT, SRC, LOCK_DIR, PDF_LIST
    argv = sys.argv[1:]
    root_arg = None
    if '--root' in argv:
        i = argv.index('--root')
        if i + 1 >= len(argv):
            raise SystemExit('--root 需要一个路径参数')
        root_arg = argv[i + 1]
        del argv[i:i + 2]
    flags = {a for a in argv if a.startswith('--')}
    args = [a for a in argv if not a.startswith('--')]
    kw = args[0] if args else ''
    force = '--all' in flags
    want_pdf = '--pdf' in flags
    dry = '--dry' in flags

    ROOT = os.path.abspath(root_arg) if root_arg else _env.require_root()
    SRC = _env.src_dir(ROOT)
    LOCK_DIR = _env.lock_dir(ROOT)
    PDF_LIST = os.path.join(LOCK_DIR, '.pdf_list.txt')
    if not os.path.isdir(SRC):
        print(f'源稿目录不存在：{SRC}')
        print('新科目接入时，先把 md 放进该目录（文件名 <科目>_第X章.md）')
        return

    items = [it for it in discover() if kw in it[1]]
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
