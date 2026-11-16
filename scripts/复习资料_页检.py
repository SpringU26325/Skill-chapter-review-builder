# -*- coding: utf-8 -*-
"""只读 Word COM 页数统计 + 逐页字符空页筛查（2026-11-10 新增）。

背景：11-09 四会话并行时，各会话手写 PowerShell COM 互相干扰（挂死、互关文档、孤儿进程）。
本脚本把「页数 / 空页」这类只读检查统一收口：
  - 走全局 .pdf.lock，与转 PDF 串行（Word 是单实例资源）；
  - 只做 ComputeStatistics / GoTo 读操作，**绝不调 ExportAsFixedFormat / SaveAs2**；
  - 优先复用已打开的 Word 实例（GetActiveObject），只 Close 自己打开的文档；
    自己新建的实例才 Quit，不碰别人的文档。
  - **不隐藏用户窗口**：Visible=0 只对本脚本新建的实例设——11-16 对用户实例设过一次，
    他的窗口整个消失，看起来就是「Word 被关了」；
  - **用户已打开的文档一律不关**：`Documents.Open` 对已开的文档返回同一个对象，
    后面无条件 `Close(0)` 关掉的就是他那一份，未保存的改动直接丢；
  - **测完还原活动文档与光标**：`GoTo` 会挪走光标，不还原等于擅自动了他看的位置。

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
    global ROOT, LOCK_DIR, LIST, LOCK
    ROOT = _env.find_root() or _root_from_docx(paths)
    LOCK_DIR = _env.lock_dir(ROOT)
    if not ROOT:
        print('WARN 定位不到工作区，锁目录退到 %s；与工作区内的会话不是同一把锁，'
              '可能同时开第二个 Word。建议在复习资料工作区里执行。' % LOCK_DIR)
    LIST = os.path.join(LOCK_DIR, '.页检清单.%d.txt' % os.getpid())
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
    "$own=$false;$prevDoc='';$prevSel=0;"
    "try{$w=[Runtime.InteropServices.Marshal]::GetActiveObject('Word.Application')}"
    "catch{$w=New-Object -ComObject Word.Application;$own=$true};"
    # 只隐藏自己起的实例。对用户正在用的实例设 Visible=0，会把他的窗口整个藏掉，
    # 看起来就是「Word 被关了」——11-16 真实踩过。
    "if($own){$w.Visible=$false};$w.DisplayAlerts=0;"
    # 记住用户当前看的文档与光标位置，测完原样还回去
    "if(-not $own){try{$prevDoc=$w.ActiveDocument.FullName;$prevSel=$w.Selection.Start}catch{}};"
    "try{"
    # 先记下已经打开的文档：Documents.Open 对已开的文档返回同一个对象，
    # 后面若无条件 Close(0)，关掉的就是用户那份（他的未保存改动直接丢）。只关自己开的。
    "$open=@{};foreach($od in @($w.Documents)){$open[$od.FullName]=$true};"
    "$ps=Get-Content -LiteralPath '__LIST__' -Encoding UTF8;"
    "foreach($p in $ps){"
    "if(-not $p.Trim()){continue};"
    "$full=[IO.Path]::GetFullPath($p);$had=$open.ContainsKey($full);"
    "$d=$w.Documents.Open($full,$false,$true);"        # ReadOnly
    "$n=$d.ComputeStatistics(2);"                      # wdStatisticPages
    "Write-Output ('DOC|'+$n);"
    "for($i=1;$i -le $n;$i++){"
    "$s=$d.GoTo(1,1,$i).Start;"                        # 本页起点（GoTo 返回折叠光标）
    "if($i -lt $n){$e=$d.GoTo(1,1,$i+1).Start}else{$e=$d.Content.End};"
    "Write-Output ('PAGE|'+$i+'|'+($e-$s));"           # 页内字符数 = 下页起点 - 本页起点
    "};"
    "if(-not $had){$d.Close(0)};"                       # wdDoNotSaveChanges，只关自己打开的
    "};"
    "Write-Output 'ok'"
    "}finally{"
    # 还原用户的活动文档与光标：GoTo 会把光标挪走，不还原等于擅自动了他的位置
    "if($prevDoc){foreach($od in @($w.Documents)){"
    "if($od.FullName -eq $prevDoc){$od.Activate();$w.Selection.SetRange($prevSel,$prevSel);break}}}"
    "if($own){$w.Quit()};"                              # 只关自己起的实例
    "[Runtime.InteropServices.Marshal]::ReleaseComObject($w)|Out-Null;"
    "[GC]::Collect();[GC]::WaitForPendingFinalizers()"
    "}"
)


CHUNK = 8          # 每次 PowerShell 调用最多查几份；46 份一次跑会顶到 180s 看门狗


def _kill_stray_word():
    """只清无主窗口的 Word（自动化起的那种），用户自己开着的窗口留着。"""
    ps = ("Get-Process WINWORD -ErrorAction SilentlyContinue | "
          "Where-Object { $_.MainWindowHandle -eq 0 } | Stop-Process -Force "
          "-ErrorAction SilentlyContinue")
    try:
        subprocess.run(['powershell', '-NoProfile', '-NonInteractive', '-Command', ps],
                       capture_output=True, timeout=60)
    except Exception:
        pass


def _run_chunk(paths):
    """返回 (rows, err)。rows = [(path, pages, [每页字符数…])]。"""
    with open(LIST, 'w', encoding='utf-8') as f:
        f.write('\n'.join(os.path.abspath(p) for p in paths))
    ps = PS_TEMPLATE.replace('__LIST__', LIST)  # PS 单引号串内反斜杠是字面量，无需转义
    try:
        r = subprocess.run(['powershell', '-NoProfile', '-NonInteractive', '-Command', ps],
                           capture_output=True, text=True, encoding='utf-8',
                           errors='replace', timeout=180)
    except subprocess.TimeoutExpired:
        _kill_stray_word()
        return [], '180s 内没跑完（这批 %d 份）；已清理残留的无窗口 Word' % len(paths)
    if 'ok' not in (r.stdout or ''):
        return [], ('PowerShell 未回 ok：' + (r.stderr or r.stdout or '')[:200])
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
    return [(paths[i], rows[i][1], rows[i][2]) for i in range(min(len(paths), len(rows)))], None


def check(paths):
    total_flag = 0
    done = 0
    fails = []
    for i in range(0, len(paths), CHUNK):
        batch = paths[i:i + CHUNK]
        rows, err = _run_chunk(batch)
        if err:
            for p in batch:
                fails.append((p, err))
            print('FAIL %s：%s' % (os.path.basename(batch[0]), err))
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
    _bind_paths(args)
    with _file_lock(LOCK):
        try:
            code = check(args)
        finally:
            # 必须在锁内删：锁外删会踩到下一个会话刚写进去的清单，
            # 对方的 PowerShell 读不到文件就整批 FAIL。
            try:
                os.remove(LIST)
            except OSError:
                pass
    raise SystemExit(code)


if __name__ == '__main__':
    main()
