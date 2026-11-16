# 渲染链、并发与 COM 纪律

> **按需读**。只在导出 PDF、看图验收、页检、多会话并发构建时打开；写 md、配题、交付话术都不需要它。
> 本文件是渲染侧唯一真相：`SKILL.md` 只留命令，`delivery.md` 只留协作纪律，两边都不再重复 COM 细节。

## 一、执行顺序

命令写在 `SKILL.md`「命令速查」，这里不重复，只写顺序与每步的产物：

1. `--prepare <docx>` 建本章私有过程目录，写入 docx 副本、清单，并把 `render.ps1` 复制进同一目录。
2. 用 PowerShell 工具执行该目录里的 `render.ps1`（打印出来的绝对路径可直接粘贴），**必须后台跑**再轮询 `.renderlog.txt`。
3. `--render` 读清单，输出逐页字符数与 PNG；`--clean` 收摊（目录不止一个时必须写目录名）。

**第 2 步是最容易错的一步**：前台调用拉起 GUI 后工具可能提前返回，留下孤儿 Word；两个实例抢同一份 docx 就弹「文件正在使用」模态框，之后所有 COM 调用（含 `Quit`）都返回 `RPC_E_CALL_REJECTED`，看起来像「导出挂死」。正常一份 8 页资料 8–10 秒出 PDF。

## 二、过程目录按「科目 + 章节」隔离

`--prepare` 不写共享的 `_生产\_render\p1.docx`，而是建 `_生产\_render\<科目>_<章节>\`（如 `_生产\_render\程序设计_第1章\`）放自己的副本、清单、`render.ps1` 与 PNG。**科目必须在目录名里**——只写「第1章」的话，将来两个科目同时做第 1 章就会撞（科目取文件名 `<科目>_第X章` 前缀，取不到时退回所在目录名，一次 `--prepare` 只允许同一科目）。并行会话各占一个目录，就不会出现「同一轮渲染里第 7 页忽然变成另一章」。

目录不止一个时 `--render` / `--clean` **拒绝猜**，列出候选让你指定；也不要对 `_生产\_render` 整体做清理，那会删掉别人正在用的过程文件。`CRB_RENDER_DIR` 可整体改到别处（私有目录），优先级最高。

`--prepare` **不删整个目录**：同名目录里的 `pN.docx` 若正被 Word 占着（另一个会话在渲染同一章），直接报错让你等或换私有目录，而不是把人家的清单删掉、副本留下半截。重跑只清自己那一轮的遗留：多余编号的 `pN`、旧 `.renderlog.txt`（里面可能留着上一轮的 `TIMEOUT|`）、旧 `_render_cache`。

## 三、并发：三把锁与 pid 临时文件

- **构建锁** `_生产\_locks\.build_<科目>.lock`：直接跑 `复习资料_build.py <md> <out.docx>` 也会按科目抢它；由 `复习资料_重建.py` 调用时脚本已持有同一把锁，靠 `CRB_LOCK_HELD` 传下去避免自锁。绕过重建脚本单册构建时，这是唯一挡在「两个会话同存一份 docx、后写覆盖前写」前面的东西。
- **PDF 锁** `_生产\_locks\.pdf.lock`：**Word COM 是唯一串行资源**，跨科目排队，同时转会互相关掉文档。`--pdf` 已包在锁内；手工转时先确认另一个会话没在转。
- **临时文件全部带 pid**：锁目录的 `.pdf_list.<pid>.txt` / `.页检清单.<pid>.txt`、渲染目录的写权限探针 `.w.<pid>`。固定名在并行下会互相删——探针被踩的后果是**静默退回 `TEMP\crb_render`**，随后 `--render` 找不到清单，看着像 `--prepare` 没生效。页检的锁目录按「当前目录 → docx 所在工作区」顺序定位，不随 cwd 漂到 `TEMP\crb_locks`（那等于另一把锁，两个 Word 实例会同时起来）。

> 迁移或换机时锁目录会变，新旧锁互不互斥。迁移窗口内禁止任何构建。

## 四、render.ps1 与看门狗

`render.ps1` 自带熔断：日志 180 秒没有新增就记 `TIMEOUT|`、强杀 Word 让阻塞的 COM 立刻抛错，然后收尾退出（阈值可用 `CRB_RENDER_TIMEOUT` 临时调整；自测可压到 `3` 验证熔断路径）。脚本退出前会把本次新起却没退干净的 WINWORD 强杀（日志 `CLEANED leftover`、`leftover PID:`），**只杀自己拉起的实例**，不动用户开着的窗口。

一条硬约束：**清单必须在 `New-Object -ComObject Word.Application` 之前读取并转成 `[string]`**。改到实例创建之后再读，`ExportAsFixedFormat` 会永久阻塞且不报错。

**不要往 `render.ps1` 的 `try/finally` 之后追加逻辑**。COM 抛错后 finally 之后的语句不会执行（11-11 连测两次：日志与 stdout 都收不到），收尾、判定、统计一律写在 finally 内。

前提检查：执行前 `WINWORD` 进程数必须为 0、过程目录内无 `~$` 锁文件；有就先让用户点掉模态框，再优雅 `Close(0)` → `Quit()`，不要连刷。

## 五、Word COM 纪律（防「文档恢复」残留）

**转 PDF 一律走重建脚本的 `--pdf`**（默认增量，只重建 md 比 docx 新的章节；`--all` 仅在改过构建引擎后使用，否则会把所有已交付成品重建一遍）。不要为每轮微调反复 `New-Object -ComObject Word.Application`。确需自定义导出（dpi/clip 裁切、单独出一份校验 PDF）时，**复用同一个实例**，别一轮一启停。

手工调用 COM 必须成套做完：

1. 优先 `GetActiveObject("Word.Application")` 复用实例；`Visible=$false`、`DisplayAlerts=0`、`Documents.Open(..., ReadOnly=$true)`。
2. 收尾顺序固定：`$d.Close(0)` → `$w.Quit()` → `ReleaseComObject($d)` → `ReleaseComObject($w)` → `[GC]::Collect()`。**只 release `$w` 会留下僵尸 WINWORD 进程**，之后构建写 docx 会静默失败（时间戳不更新）。
3. **默认禁止 `Stop-Process -Force` / `taskkill` 结束 WINWORD**。强杀或异常退出会让 Word 来不及清掉 `%APPDATA%\Microsoft\Word\*.asd` 自动恢复副本，并在 `HKCU\...\Word\Resiliency` 留下恢复项；用户下次开关 Word 就会弹出「文档恢复」任务窗格。恢复列表里是**过程版副本**，误点会把已交付的最终版覆盖回去。
   **仅两处例外，且都已写进 `render.ps1`，不要在别处模仿**：① 日志 180 秒不动（看门狗判挂死）；② 脚本自身 `Quit()` 后 15 秒仍有本次拉起的实例（`$base` 之外的 PID，留着会撞下一次导出）。两处都只杀本次新起的实例、不碰用户自己开的 Word。
4. 遇 `~$xxx.docx` 锁文件：先 `GetActiveObject` 查 `Saved`，`Saved=True` 才允许 `Close(0)`；`Saved=False` 停下交给用户。确属自己起的隐藏实例用 `$w.Quit()` 优雅退出，仍不退就报告用户，不强杀。
5. **交付前只读确认**：`Get-Process WINWORD` 无自己起的残留、`复习资料\<科目>\` 及其 `_成品\` 下无 `~$` 锁文件。若 Word 已弹出「文档恢复」询问，指导用户选 **「否，删除这些文件。我已保存了需要的文件」**——成品已落盘且四项硬指标通过，恢复副本是旧版；**不要手动去 AppData 删文件**，交给 Word 自己清。
6. **预览进程可直接强杀**：预览面板会拉起文档预览引擎并**锁住 docx**，后果是构建脚本打印「正式文件被占用」并把结果写到源稿目录（**退出码 3，非 0**）。遇到预览进程占锁直接结束它（用户已授权）。查真实持有者用 `scripts\_wholock.py <docx>`（Restart Manager，直接打印 pid 与进程名），**不要靠猜**。`WINWORD` 仍按第 3 条禁止强杀。
7. **COM 只读检查统一走页检脚本，不要手写 PowerShell COM**：页数、逐页字符数、空页筛查全在脚本里；它优先复用已开的 Word 实例、只关自己的文档、带全局锁。手工 COM 仍是例外通道，用完必须成套收尾。

## 六、导出的三条前提（纠正「通道整体不可用」的旧误判）

`ExportAsFixedFormat` / `SaveAs2(…,17)` 本身可用，前提是三条同时满足：

1. **由 PowerShell 工具直接执行**。同样命令包进 `python subprocess → powershell -NonInteractive` 必挂（90–240s 超时）。这是过去误判的真正原因，不是 docx 内容、打包方式或样式的问题。
2. **`Visible=$true` + `Documents.Open(..., ReadOnly=$false)`**。`ReadOnly=$true` 或隐藏实例会让导出死循环；而同实例的 `ComputeStatistics(2)`、`Range.Text` 正常，所以页检能过、导出却挂。
3. **路径必须经 UTF-8 清单文件传入**，不要写进 `.ps1` 字面量、也不要放进命令行参数：PowerShell 5.1 无 BOM 时按 ANSI 解码，中文字面量会乱码。**中文路径本身不影响导出**（11-11 实测：源与目标都含中文，8 页 10.3s 成功），历史上「中文路径挂死」的结论是把清单读取时序的缺陷误归因到路径上。

## 七、导出失败的判读

**只认日志，不认退出码**。超时同样打印 `FAIL|…远程过程调用失败` 并以 0 退出——与「只挂了一个文件」长得一样，`TIMEOUT|` 才是挂死的证据，出现即判本次导出失败、PDF 不完整。

症状：出现 `~$xxx.docx` 锁文件且无 PDF 产出 = 挂在导出。先 `GetActiveObject` → `Quit()` 优雅退出再重试，**不要连刷**——每次超时留一个卡死实例，后续调用会连到坏实例，造成「越试越挂」的假象。

## 八、PowerPoint COM

与 Word COM 同一套：

- 单实例；`Presentations.Open(path,$true,$false,$false)`；
- 导出后 `Close()` → `Quit()` → `ReleaseComObject` → `GC`；
- **禁止强杀 POWERPNT**；
- 中文路径**不能写进命令行参数、也不能写进 .ps1 字面量**（跨进程传参 / ANSI 解码都会乱码，表现为找不到文件或静默挂住），必须用 **UTF-8 清单文件**传路径，命令行只留 ASCII 的清单路径。**路径本身含中文不影响导出**——Word COM 已实测（11-11），PowerPoint 同样适用；误判常源于把「传参乱码」当成「路径不支持」；
- `Slide.Export` 的输出路径必须是**反斜杠**，正斜杠报「无法将 ^0 保存到 ^1」；
- 在 Bash 里调 powershell 会被安全策略拦截，须用 PowerShell 工具。
