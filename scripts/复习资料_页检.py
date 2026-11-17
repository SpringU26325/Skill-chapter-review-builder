# -*- coding: utf-8 -*-
"""只读 Word COM 页数统计 + 逐页字符空页筛查（2026-11-10 新增）。

背景：11-09 四会话并行时，各会话手写 PowerShell COM 互相干扰（挂死、互关文档、孤儿进程）。
本脚本把「页数 / 空页」这类只读检查统一收口：
  - 走全局 .pdf.lock，与转 PDF 串行（Word 是单实例资源）；
  - 只做 ComputeStatistics / GoTo 读操作，**绝不调 ExportAsFixedFormat / SaveAs2**；
  - **一律新建独立实例**，不复用用户已打开的 Word。借用时 `Visible=$false`、`Close(0)`、
    `GoTo` 全落在用户那个进程上——11-16 把他正看的文档关掉，就是这么发生的；
  - 实例隐藏、文档只读打开，读完逐个关闭再 `Quit`；用户窗口不在这个进程里，天然不受影响；
  - **超时不按进程名杀 Word**：只把本次拉起的 PID 落盘，超时后报告 PID（标注「疑似」），
    结束与否由人决定；同一批超时后不再继续后续批次，避免堆出多个看不见的实例。
  - **重复参数去重**：同一份 docx 传两次会二次 `Close(0)` 抛错，整批 FAIL。

用法：python 复习资料_页检.py <docx> [docx2 ...]
输出：每份文件的页数、每页字符数（<40 标记为疑似空页）。退出码：0 正常 / 2 有疑似空页 / 1 COM 失败。
"""
import io
import os
import subprocess
import sys
import time
from contextlib import contextmanager

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import _env

ROOT = None
LOCK_DIR = None
LIST = None                                     # 每进程一份清单，见 _bind_paths()
PIDFILE = None                                  # 疑似本次拉起的 WINWORD PID，超时后只报告不杀
LOCK = None


def _root_from_docx(paths):
    """工作区外被调用时（find_root 为 None），从 docx 路径反推工作区根。

    锁目录必须落在工作区同一个 `复习资料\\_locks` 下：`lock_dir(None)` 会退到
    `TEMP\\crb_locks`，那是另一把 `.pdf.lock`，等于没上锁——工作区内的重建正在
    转 PDF 时，这里的页检会同时开第二个 Word 实例。
    """
    d = os.path.dirname(os.path.abspath(paths[0]))
    for _ in range(6):
        if os.path.isdir(os.path.join(d, _env.WORKSPACE_MARK)):
            return d
        parent = os.path.dirname(d)
        if parent == d:
            break
        d = parent
    return None


def _bind_paths(paths):
    """确定工作区 → 锁目录 → 本进程清单名。清单带 pid：并行页检不会互相删清单。"""
    global ROOT, LOCK_DIR, LIST, PIDFILE, LOCK
    ROOT = _env.find_root() or _root_from_docx(paths)
    LOCK_DIR = _env.lock_dir(ROOT)
    if not ROOT:
        print('WARN 定位不到工作区，锁目录退到 %s；与工作区内的会话不是同一把锁，'
              '可能同时开第二个 Word。建议在复习资料工作区里执行。' % LOCK_DIR)
    LIST = os.path.join(LOCK_DIR, '.页检清单.%d.txt' % os.getpid())
    PIDFILE = os.path.join(LOCK_DIR, '.页检实例.%d.txt' % os.getpid())
    LOCK = os.path.join(LOCK_DIR, '.pdf.lock')
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')


@contextmanager
def _file_lock(path, timeout=900, what='Word 只读页检'):
    """跨进程互斥（与 复习资料_重建.py 同一套 msvcrt 字节锁）。"""
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
                    raise TimeoutError(f'{what}锁等待超时（{timeout}s）：另一会话正在用 Word')
                time.sleep(0.5)
        yield
    finally:
        try:
            msvcrt.locking(fd, msvcrt.LK_UNLCK, 1)
        except OSError:
            pass
        os.close(fd)


PS_TEMPLATE = (
    "$ErrorActionPreference='Stop';"
    # 快照只为识别「本次新起的实例」；快照紧贴 New-Object，窗口越窄越不容易把
    # 用户此刻手动开的 Word 算进来。它不是杀进程的依据。
    "$pre=@(Get-Process WINWORD -ErrorAction SilentlyContinue|ForEach-Object{$_.Id});"
    "$w=$null;"
    "try{"
    # 一律新建实例。复用用户实例时，Visible/Close/GoTo 全作用在他正用的那个进程上
    "$w=New-Object -ComObject Word.Application;"
    "$w.Visible=$false;$w.DisplayAlerts=0;"
    # 把本次新起的 PID 落盘：超时后 Python 只报告这些 PID，不按名字杀全部 WINWORD。
    # 同时新起多个时取 StartTime 最早的那个——本次实例先于用户任何手动操作创建。
    "$mine=@();for($k=0;$k -lt 10;$k++){"
    "$new=@(Get-Process WINWORD -ErrorAction SilentlyContinue|"
    "Where-Object{$pre -notcontains $_.Id});"
    "if($new.Count -gt 0){$pick=$new[0];"
    "try{$pick=@($new|Sort-Object StartTime)[0]}catch{};"
    "$mine=@($pick.Id);break};"
    "Start-Sleep -Milliseconds 300};"
    "if($mine.Count -gt 0){"
    "Set-Content -LiteralPath '__PIDFILE__' -Value ($mine -join ' ') -Encoding ASCII};"
    "$lines=Get-Content -LiteralPath '__LIST__' -Encoding UTF8;"
    "foreach($p in $lines){"
    "if(-not $p.Trim()){continue};"
    "$full=[IO.Path]::GetFullPath($p);"
    "$d=$w.Documents.Open($full,$false,$true);"        # ReadOnly
    "$n=$d.ComputeStatistics(2);"                      # wdStatisticPages
    "Write-Output ('DOC|'+$n);"
    "for($i=1;$i -le $n;$i++){"
    "$s=$d.GoTo(1,1,$i).Start;"                        # 本页起点（GoTo 返回折叠光标）
    "if($i -lt $n){$e=$d.GoTo(1,1,$i+1).Start}else{$e=$d.Content.End};"
    "Write-Output ('PAGE|'+$i+'|'+($e-$s));"           # 页内字符数 = 下页起点 - 本页起点
    "};"
    # 文档都是本实例打开的，直接关；用户那份在别的进程里，碰不到
    "$d.Close(0);"
    "};"
    "Write-Output 'ok';"
    "}finally{"
    "if($w){try{$w.Quit()}catch{};"
    "try{[Runtime.InteropServices.Marshal]::ReleaseComObject($w)|Out-Null}catch{};"
    "[GC]::Collect();[GC]::WaitForPendingFinalizers()}"
    "}"
)


CHUNK = 8          # 每次 PowerShell 调用最多查几份；46 份一次跑会顶到 180s 看门狗


def _read_pids():
    """读疑似本次拉起的 WINWORD PID。只用于超时报告，不拿来杀进程。"""
    try:
        with open(PIDFILE, encoding='ascii') as f:
            return [x for x in f.read().split() if x.isdigit()]
    except OSError:
        return []


def _run_chunk(paths):
    """返回 (rows, err, hang)。rows = [(path, pages, [每页字符数…])]。"""
    with open(LIST, 'w', encoding='utf-8') as f:
        f.write('\n'.join(os.path.abspath(p) for p in paths))
    try:
        os.remove(PIDFILE)                  # 上一批的 PID 不能带到这一批
    except OSError:
        pass
    ps = (PS_TEMPLATE
          .replace('__LIST__', LIST)        # PS 单引号串内反斜杠是字面量，无需转义
          .replace('__PIDFILE__', PIDFILE))
    try:
        r = subprocess.run(['powershell', '-NoProfile', '-NonInteractive', '-Command', ps],
                           capture_output=True, text=True, encoding='utf-8',
                           errors='replace', timeout=180)
    except subprocess.TimeoutExpired:
        # 不按进程名杀 Word：只报出疑似本次拉起的 PID，结束与否交给用户
        stray = _read_pids()
        msg = '180s 内没跑完（这批 %d 份）' % len(paths)
        if stray:
            msg += ('；疑似本次拉起的 Word 未退出，PID %s（未自动结束；若期间你手动开过 Word，'
                    '也会被算进来）' % ','.join(stray))
        return [], msg, True
    if 'ok' not in (r.stdout or ''):
        return [], ('PowerShell 未回 ok：' + (r.stderr or r.stdout or '')[:200]), False
    rows, cur = [], None
    for line in (r.stdout or '').splitlines():
        parts = line.split('|')
        if parts[0] == 'DOC':
            cur = (None, int(parts[1]), [])
            rows.append(cur)
        elif parts[0] == 'PAGE' and cur is not None:
            cur[2].append(int(parts[2]))
        elif parts[0] == 'ok':
            break
    # 按顺序补上路径：PowerShell 是顺序打开的
    return [(paths[i], rows[i][1], rows[i][2]) for i in range(min(len(paths), len(rows)))], None, False


def check(paths):
    total_flag = 0
    done = 0
    fails = []
    for i in range(0, len(paths), CHUNK):
        batch = paths[i:i + CHUNK]
        rows, err, hang = _run_chunk(batch)
        if err:
            for p in batch:
                fails.append((p, err))
            print('FAIL %s：%s' % (os.path.basename(batch[0]), err))
            if hang:
                # 超时说明有一个看不见的 Word 卡住了。继续跑后续批次只会再起新的实例，
                # 把卡死实例越堆越多，所以这里直接停手，剩下的报未执行。
                rest = paths[i + len(batch):]
                for p in rest:
                    fails.append((p, '前一批超时，未执行'))
                if rest:
                    print('STOP 前一批超时，剩余 %d 份未执行；先确认残留 Word 是否需要结束' % len(rest))
                break
            continue
        for p, pages, lens in rows:
            done += 1
            print(f'--- {os.path.basename(p)}：{pages} 页')
            for k, ln in enumerate(lens, 1):
                flag = '  << 疑似空页(len<40)' if ln < 40 else ''
                if flag:
                    total_flag += 1
                print(f'  p{k:>2}: {ln}{flag}')
    print(f'\n共 {done} 份，疑似空页 {total_flag} 页')
    if fails:
        print(f'未完成 {len(fails)} 份：')
        for p, why in fails[:8]:
            print('  - %s：%s' % (os.path.basename(p), why))
        return 1
    return 2 if total_flag else 0


def main():
    args = [a for a in sys.argv[1:] if a.strip()]
    if not args:
        print(__doc__)
        raise SystemExit(1)
    # 去重：同一份 docx 传两次，第二次 Documents.Open 返回同一个对象，再 Close(0) 会抛错，
    # 整批直接 FAIL。
    seen, uniq = set(), []
    for a in args:
        k = os.path.normcase(os.path.abspath(a))
        if k in seen:
            print('SKIP 重复参数：%s' % a)
            continue
        seen.add(k)
        uniq.append(a)
    args = uniq
    _bind_paths(args)
    with _file_lock(LOCK):
        try:
            code = check(args)
        finally:
            # 必须在锁内删：锁外删会踩到下一个会话刚写进去的清单，
            # 对方的 PowerShell 读不到文件就整批 FAIL。
            for p in (LIST, PIDFILE):
                try:
                    os.remove(p)
                except OSError:
                    pass
    raise SystemExit(code)


if __name__ == '__main__':
    main()
