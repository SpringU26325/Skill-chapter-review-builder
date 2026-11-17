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

`render.ps1` 自带熔断：日志 180 秒没有新增就记 `TIMEOUT|`，只结束**本次实例自己的 PID**（`New-Object` 之后立刻写进同目录 `.renderpid.txt`），让阻塞的 COM 立刻抛错，然后收尾退出（阈值可用 `CRB_RENDER_TIMEOUT` 临时调整；自测可压到 `2` 验证熔断路径，日志写 `own pid=…; killing N`）。脚本退出前同样只清自己的 PID（日志 `CLEANED leftover`、`leftover PID:`），**不按进程名杀全部 WINWORD**。

三道额外保险，改动时不要删：① 杀之前核对 `ProcessName -eq 'WINWORD'`（PID 可能已被回收，属于别人）；② 熔断前先重试读 PID 文件约 2 秒，读完再复查日志——这 2 秒里日志若推进，说明没真挂，继续等；③ 结束宿主进程前比对 `StartTime.Ticks`，避免 PID 复用时杀错进程。

**归属只认 PID 文件，不认「开工前快照的差集」**。差集会把导出期间你自己新开的 Word 一并算成「本次实例」——正常导出收尾就会把它强杀。所以：`New-Object` 之前紧贴地再取一次快照（窗口越窄越好），`New-Object` 之后立刻把该实例 PID 落盘（新起多个时取 `StartTime` 最早的那个，本次实例先于任何手动操作创建）；看门狗超时时才去读这个文件，读不到先等约 2 秒再读一次（短超时下判定与落盘会撞在同一瞬间）；**仍读不到就不杀，只报**。代价是拿不到 PID 时会留一个孤儿，这比杀错窗口轻。

`.renderpid.txt` 在每次开工前删除，防止上一轮的文件给这一轮授权。目录里还有上一轮的 `.renderlog.txt` 时不要拿它判读，`--prepare` 会先删旧日志。

### DDE 会把用户文档吞进渲染实例（11-17 实测）

导出期间用户在资源管理器里双击任意 `.docx`，Windows 走 DDE 把文件交给**已运行的 Word 实例**，不新建进程——渲染实例就是最近的接收者。随后脚本收尾 `Quit()` 关的是整个应用，用户文档跟着消失；而 `DisplayAlerts=0` 让「是否保存」的提示不弹，未保存内容**静默丢失**。

实测证据（基线 0 个 Word）：

```
PRE_COUNT=0
AFTER_NEW  count=1  ids=30448        ← 渲染实例
AFTER_OPEN count=1  ids=30448        ← 用户双击，进程数没变（DDE 复用）
DOCS_IN_RENDER_INSTANCE=1
  DOC ...\crb_user_sim.docx          ← 用户文档在渲染实例里
AFTER_QUIT count=0                   ← 一起被关掉
USER_DOC_SURVIVED=NO
```

**隐藏或最小化都挡不住，而且会挂死**：把渲染实例设成 `Visible=$false` 再测，`DDE_REUSED=YES`、`DOCS_IN_HIDDEN=1`，照样被复用；设成最小化（`WindowState=1`）后 `ExportAsFixedFormat` 直接挂死——PDF 不产出，`CloseMainWindow()` 与 COM `Quit()` 都无效，只能强制结束进程。对应地，**渲染实例必须保持 `Visible=$true` 的正常窗口**，不要去「优化」观感。

代价是导出期间那个 Word 窗口会每 3–5 秒关一份、开下一份，标题在 `p1.docx`、`p2.docx` 之间跳。**这个窗口不能用来工作**：手速再快也来不及落字，打开文档还会被 DDE 并进去。要告诉用户「这段时间别在那台 Word 上操作」，而不要试图让它安静。

现方案的收尾逻辑（`render.ps1` 的 `finally`）：

1. 逐份 `Documents.Open` 时把绝对路径登记进 `$ownDocs`；
2. `Quit()` 之前遍历 `Documents`，凡不在 `$ownDocs` 里的都算外来文档；
3. 有外来文档 → 记 `FOREIGN DOC|…` + `QUIT SKIPPED|…` + `cleanup skipped`，把 `DisplayAlerts` 还原成 `-1`、`Visible=$true`，**不 Quit、不强杀**，实例留给用户自己保存关闭；
4. 没有外来文档 → 正常 `quit sent`。

因此：渲染前要提醒用户**不要在导出期间双击 Word 文档**；跑完看到 `QUIT SKIPPED` 时，必须让用户手动关闭那个 Word 窗口，**绝不按进程名或 PID 去杀**——那正是他文档所在的进程。中间每份之间若检测到实例里还有别的文档，会打一次 `FOREIGN DOC STAYS OPEN|…`。

**改完 `render.ps1` 必须重新 `--prepare`**：`--prepare` 会把它复制进各章的渲染目录（`$PSScriptRoot` 决定清单位置，必须同目录），已存在的目录里跑的是复制时的快照。11-17 就有一份 11-11 的旧副本留在 `_生产\_render\`，里面还是按进程名全杀的版本。

一条硬约束：**清单必须在 `New-Object -ComObject Word.Application` 之前读取并转成 `[string]`**。改到实例创建之后再读，`ExportAsFixedFormat` 会永久阻塞且不报错。

**不要往 `render.ps1` 的 `try/finally` 之后追加逻辑**。COM 抛错后 finally 之后的语句不会执行（11-11 连测两次：日志与 stdout 都收不到），收尾、判定、统计一律写在 finally 内。

看门狗 `Stop-Process -Id $parent` 杀的是**运行脚本的那个 PowerShell 宿主进程**，不是脚本本身。这是有意的（只有这样才能把被 COM 卡死的调用方一起解掉），代价是在自己的交互式窗口里直接跑这个文件，熔断时窗口会跟着关闭。

前提检查：过程目录内不能有 `~$` 锁文件，目标 docx 不能被别的实例占着。**不要求 `WINWORD` 数为 0**——脚本自建实例，用户开着 Word 也能跑；遇到模态框先让用户点掉，再优雅 `Close(0)` → `Quit()`，不要连刷。

## 五、Word COM 纪律（防「文档恢复」残留）

**转 PDF 一律走重建脚本的 `--pdf`**（默认增量，只重建 md 比 docx 新的章节；`--all` 仅在改过构建引擎后使用，否则会把所有已交付成品重建一遍）。不要为每轮微调反复 `New-Object -ComObject Word.Application`。确需自定义导出（dpi/clip 裁切、单独出一份校验 PDF）时，**复用同一个实例**，别一轮一启停。

手工调用 COM 必须成套做完：

1. **自建实例，不复用**：`New-Object -ComObject Word.Application`，再设 `Visible=$false`、`DisplayAlerts=0`、`Documents.Open(..., ReadOnly=$true)`。`GetActiveObject` 借到的是用户正在用的那个进程，隐藏窗口、关文档、挪光标都会落在他身上。
2. 收尾顺序固定：`$d.Close(0)` → `$w.Quit()` → `ReleaseComObject($d)` → `ReleaseComObject($w)` → `[GC]::Collect()`。**只 release `$w` 会留下僵尸 WINWORD 进程**，之后构建写 docx 会静默失败（时间戳不更新）。
3. **默认禁止 `Stop-Process -Force` / `taskkill` 结束 WINWORD**。强杀或异常退出会让 Word 来不及清掉 `%APPDATA%\Microsoft\Word\*.asd` 自动恢复副本，并在 `HKCU\...\Word\Resiliency` 留下恢复项；用户下次开关 Word 就会弹出「文档恢复」任务窗格。恢复列表里是**过程版副本**，误点会把已交付的最终版覆盖回去。
   **仅两处例外，且都已写进 `render.ps1`，不要在别处模仿**：① 日志 180 秒不动（看门狗判挂死）；② 脚本自身 `Quit()` 后 15 秒仍残留（留着会撞下一次导出）。两处**只杀 `.renderpid.txt` 里记的那个 PID**——不按进程名、不按差集；拿不到 PID 时只报不杀。**页检脚本超时不属于例外**：它只报出疑似本次拉起的 PID，由人决定是否结束。
4. 遇 `~$xxx.docx` 锁文件：**不要 `GetActiveObject` 去碰用户实例**。用 `scripts\_wholock.py <docx>` 查持有者；确认是本次自己拉起的隐藏实例才 `Quit()`，仍不退就报告用户，不强杀；持有者是用户就停下，请他保存并关闭。
5. **交付前只读确认**：`Get-Process WINWORD` 无自己起的残留、`复习资料\<科目>\` 及其 `_成品\` 下无 `~$` 锁文件。若 Word 已弹出「文档恢复」询问，指导用户选 **「否，删除这些文件。我已保存了需要的文件」**——成品已落盘且四项硬指标通过，恢复副本是旧版；**不要手动去 AppData 删文件**，交给 Word 自己清。
6. **预览进程可直接强杀**：预览面板会拉起文档预览引擎并**锁住 docx**，后果是构建脚本打印「正式文件被占用」并把结果写到源稿目录（**退出码 3，非 0**）。遇到预览进程占锁直接结束它（用户已授权）。查真实持有者用 `scripts\_wholock.py <docx>`（Restart Manager，直接打印 pid 与进程名），**不要靠猜**。`WINWORD` 仍按第 3 条禁止强杀。
7. **COM 只读检查统一走页检脚本，不要手写 PowerShell COM**：页数、逐页字符数、空页筛查全在脚本里；它**自建隐藏实例**，不碰用户已开的 Word，读完全部关闭并 `Quit`，带全局锁。超时只报出**疑似**本次拉起的 PID（这 3 秒里手动开的 Word 也可能被算进来），不自动结束进程，并立即中止剩余批次——继续跑只会再起实例，把卡死实例越堆越多。手工 COM 仍是例外通道，用完必须成套收尾。

## 六、导出的三条前提（纠正「通道整体不可用」的旧误判）

`ExportAsFixedFormat` / `SaveAs2(…,17)` 本身可用，前提是三条同时满足：

1. **由 PowerShell 工具直接执行**。同样命令包进 `python subprocess → powershell -NonInteractive` 必挂（90–240s 超时）。这是过去误判的真正原因，不是 docx 内容、打包方式或样式的问题。
2. **`Visible=$true` + `Documents.Open(..., ReadOnly=$false)`**。`ReadOnly=$true` 或隐藏实例会让导出死循环；而同实例的 `ComputeStatistics(2)`、`Range.Text` 正常，所以页检能过、导出却挂。
3. **路径必须经 UTF-8 清单文件传入**，不要写进 `.ps1` 字面量、也不要放进命令行参数：PowerShell 5.1 无 BOM 时按 ANSI 解码，中文字面量会乱码。**中文路径本身不影响导出**（11-11 实测：源与目标都含中文，8 页 10.3s 成功），历史上「中文路径挂死」的结论是把清单读取时序的缺陷误归因到路径上。

## 七、导出失败的判读

**只认日志，不认退出码**。超时同样打印 `FAIL|…远程过程调用失败` 并以 0 退出——与「只挂了一个文件」长得一样，`TIMEOUT|` 才是挂死的证据，出现即判本次导出失败、PDF 不完整。

症状：出现 `~$xxx.docx` 锁文件且无 PDF 产出 = 挂在导出。**进程退出后已经算不出「差集」**，判据只有三个：日志有没有 `TIMEOUT|`、`WINWORD_LEFT=` 后面的数、`leftover PID:` 这一行。`WINWORD_LEFT=unknown (own pid not identified)` 配上 `WARN cannot prove ownership` 表示这次没认出自己的实例，看到的那几个 PID 一个都没杀。有自己残留的 PID 就 `Quit()` 优雅退出再重试，**不要连刷**——每次超时留一个卡死实例，后续调用会连到坏实例，造成「越试越挂」的假象。上一轮漏杀的实例会进入下一轮的「开工前快照」，从此不再被自动清理，需要人工结束。

另两种收尾结局要分清：`QUIT SKIPPED` + `cleanup skipped` 表示实例里有一份不是本次打开的文档（`FOREIGN DOC|…` 会给出路径），脚本**故意不退出**以保住它——这不是故障，让人去保存并关闭那个 Word 窗口即可，**不要杀进程**；只有 `quit sent` + `WINWORD_LEFT=0` 才是干净收尾。

## 八、PowerPoint COM

与 Word COM 同一套：

- 单实例；`Presentations.Open(path,$true,$false,$false)`；
- 导出后 `Close()` → `Quit()` → `ReleaseComObject` → `GC`；
- **禁止强杀 POWERPNT**；
- 中文路径**不能写进命令行参数、也不能写进 .ps1 字面量**（跨进程传参 / ANSI 解码都会乱码，表现为找不到文件或静默挂住），必须用 **UTF-8 清单文件**传路径，命令行只留 ASCII 的清单路径。**路径本身含中文不影响导出**——Word COM 已实测（11-11），PowerPoint 同样适用；误判常源于把「传参乱码」当成「路径不支持」；
- `Slide.Export` 的输出路径必须是**反斜杠**，正斜杠报「无法将 ^0 保存到 ^1」；
- 在 Bash 里调 powershell 会被安全策略拦截，须用 PowerShell 工具。
