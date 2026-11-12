# -*- coding: utf-8 -*-
"""只读 Word COM 页数统计 + 逐页字符空页筛查（2026-11-10 新增）。

背景：11-09 四会话并行时，各会话手写 PowerShell COM 互相干扰（挂死、互关文档、孤儿进程）。
本脚本把「页数 / 空页」这类只读检查统一收口：
  - 走全局 .pdf.lock，与转 PDF 串行（Word 是单实例资源）；
  - 只做 ComputeStatistics / GoTo 读操作，**绝不调 ExportAsFixedFormat / SaveAs2**；
  - 优先复用已打开的 Word 实例（GetActiveObject），只 Close 自己打开的文档；
    自己新建的实例才 Quit，不碰别人的文档。

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

ROOT = _env.find_root()
LOCK_DIR = _env.lock_dir(ROOT)
LIST = os.path.join(LOCK_DIR, '.页检清单.txt')
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
    "$own=$false;"
    "try{$w=[Runtime.InteropServices.Marshal]::GetActiveObject('Word.Application')}"
    "catch{$w=New-Object -ComObject Word.Application;$own=$true};"
    "$w.Visible=$false;$w.DisplayAlerts=0;"
    "try{"
    "$ps=Get-Content -LiteralPath '__LIST__' -Encoding UTF8;"
    "foreach($p in $ps){"
    "if(-not $p.Trim()){continue};"
    "$d=$w.Documents.Open($p,$false,$true);"          # ReadOnly
    "$n=$d.ComputeStatistics(2);"                      # wdStatisticPages
    "Write-Output ('DOC|'+$n);"
    "for($i=1;$i -le $n;$i++){"
    "$s=$d.GoTo(1,1,$i).Start;"                        # 本页起点（GoTo 返回折叠光标）
    "if($i -lt $n){$e=$d.GoTo(1,1,$i+1).Start}else{$e=$d.Content.End};"
    "Write-Output ('PAGE|'+$i+'|'+($e-$s));"           # 页内字符数 = 下页起点 - 本页起点
    "};"
    "$d.Close(0);"                                     # wdDoNotSaveChanges，只关自己打开的
    "};"
    "Write-Output 'ok'"
    "}finally{"
    "if($own){$w.Quit()};"                             # 只关自己起的实例
    "[Runtime.InteropServices.Marshal]::ReleaseComObject($w)|Out-Null;"
    "[GC]::Collect();[GC]::WaitForPendingFinalizers()"
    "}"
)


def check(paths):
    with open(LIST, 'w', encoding='utf-8') as f:
        f.write('\n'.join(os.path.abspath(p) for p in paths))
    ps = PS_TEMPLATE.replace('__LIST__', LIST)  # PS 单引号串内反斜杠是字面量，无需转义
    r = subprocess.run(['powershell', '-NoProfile', '-NonInteractive', '-Command', ps],
                       capture_output=True, text=True, encoding='utf-8',
                       errors='replace', timeout=180)
    if 'ok' not in (r.stdout or ''):
        print('FAIL', (r.stderr or '')[:300])
        return 1
    suspicious = 0
    doc_idx = 0
    total_flag = 0
    for line in (r.stdout or '').splitlines():
        parts = line.split('|')
        if parts[0] == 'DOC':
            doc_idx += 1
            suspicious = 0
            pages = int(parts[1])
            print(f'--- {os.path.basename(paths[doc_idx - 1])}：{pages} 页')
        elif parts[0] == 'PAGE':
            idx, ln = int(parts[1]), int(parts[2])
            flag = '  << 疑似空页(len<40)' if ln < 40 else ''
            if flag:
                suspicious += 1
                total_flag += 1
            print(f'  p{idx:>2}: {ln}{flag}')
        elif parts[0] == 'ok':
            break
    print(f'\n共 {doc_idx} 份，疑似空页 {total_flag} 页')
    return 2 if total_flag else 0


def main():
    args = [a for a in sys.argv[1:] if a.strip()]
    if not args:
        print(__doc__)
        raise SystemExit(1)
    with _file_lock(LOCK):
        code = check(args)
    os.remove(LIST)
    raise SystemExit(code)


if __name__ == '__main__':
    main()
