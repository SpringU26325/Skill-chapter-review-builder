# Render docx -> PDF via Word COM. Must be executed by the PowerShell tool.
#   1. All paths come from a UTF-8 list file (.renderlist.txt); no CJK on the command line.
#   2. Work dir is chosen by `--prepare` (see CRB_RENDER_DIR); absolute backslash paths.
#   3. Visible=$true + Documents.Open(..., ReadOnly=$false); ReadOnly=$true loops forever.
#   4. Per-file try/catch so one failure does not stop the rest.
#   5. The list is read and normalised BEFORE any COM object exists; the log is appended
#      after every step so a hang is visible while it happens.
#   6. A watchdog job kills the whole thing when the log stops moving (see below). A blocked
#      ExportAsFixedFormat call cannot be interrupted from inside this script, because this
#      thread is the blocked one; only a separate process can notice and act. The kill list is
#      NOT a PID diff against a pre-run snapshot: it is the single PID in `.renderpid.txt`,
#      written right after New-Object. A diff would also contain any Word window the user
#      opens while the export runs, and killing that is exactly what this file must not do.
#      No pid file => kill nothing, just report.
#   7. Probe the outcome from the log, not from the exit code: a per-file FAIL still lets the
#      other PDFs through, so a hang and an ordinary failure look alike. `TIMEOUT|` in the
#      log is the only reliable hang marker.
# This file is intentionally ASCII-only.
#
# Hard-won detail: keep the list read and the string normalisation ahead of the
# New-Object call. Reading the list after the Word instance is created reproducibly
# made ExportAsFixedFormat block forever on this machine, while the same script with
# the list read first exports in ~8 s.

$ErrorActionPreference = 'Stop'
$listPath = Join-Path $PSScriptRoot '.renderlist.txt'
$logPath = Join-Path $PSScriptRoot '.renderlog.txt'
# PID of the one Word instance this run owns. Written right after New-Object; read lazily by
# the watchdog at timeout and by the cleanup below. Removed up-front so a stale file from an
# earlier run can never authorise a kill.
$pidFile = Join-Path $PSScriptRoot '.renderpid.txt'
if (Test-Path -LiteralPath $pidFile) { Remove-Item -LiteralPath $pidFile -Force -ErrorAction SilentlyContinue }
if (-not (Test-Path -LiteralPath $listPath)) { throw "List not found: $listPath. Run --prepare first." }

function W($m) { Add-Content -LiteralPath $logPath -Value ((Get-Date -Format 'HH:mm:ss') + ' ' + $m) -Encoding UTF8 }

# Watchdog threshold: seconds of log silence before we call it a hang. A healthy
# 8-page document exports in ~10 s end to end, so this is >10x headroom.
$TimeoutSec = 180
if ($env:CRB_RENDER_TIMEOUT) { $TimeoutSec = [int]$env:CRB_RENDER_TIMEOUT }

# normalise up-front: plain BSTR strings, no PSObject wrappers
$raw = Get-Content -LiteralPath $listPath -Encoding UTF8
$items = @()
for ($i = 0; $i -lt $raw.Count; $i += 2) {
    $s = [string]$raw[$i]
    $o = [string]$raw[$i + 1]
    if (-not $s.Trim()) { continue }
    $items += , @($s.Trim(), $o.Trim())
}

# Diagnostic only, so the log records whether the user had Word open when the run started.
# Identification of our own instance uses $pre, snapshotted right before New-Object below.
$base = @(Get-Process WINWORD -ErrorAction SilentlyContinue | ForEach-Object { $_.Id })
W ('start; existing WINWORD=' + $base.Count + '; items=' + $items.Count)

# Watchdog runs in its own process, so it keeps ticking while this thread sits inside a
# blocked COM call. When the log has not moved for $TimeoutSec it records TIMEOUT, kills only
# the PID recorded in $pidFile, then kills this script so the caller gets a dead job instead
# of a task that shows "running" forever.
#
# $PID here is the PowerShell host that runs this script, not the script itself. Killing it is
# intentional (that is what unblocks the caller), but it means the watchdog also ends that
# host: running this file inside your own interactive window will close that window too.
$watch = $null
# StartTime ticks of this host process, so the watchdog cannot kill an unrelated process that
# happens to receive the same PID after this one exits (Windows recycles PIDs).
$parentTicks = 0
try { $parentTicks = (Get-Process -Id $PID).StartTime.Ticks } catch {}
if ($TimeoutSec -gt 0) {
    try {
        # The job gets the pid FILE, not a PID list: at start-up our own PID is not known yet,
        # and it must be read at timeout, not now.
        $watch = Start-Job -ArgumentList $logPath, $TimeoutSec, $PID, $pidFile, $parentTicks -ScriptBlock {
            param($log, $limit, $parent, $pidFile, $parentTicks)
            $enc = New-Object System.Text.UTF8Encoding($false)
            $last = (Get-Item -LiteralPath $log).LastWriteTime
            $stall = Get-Date
            while ($true) {
                Start-Sleep -Seconds 2
                $cur = (Get-Item -LiteralPath $log).LastWriteTime
                if ($cur -gt $last) { $last = $cur; $stall = Get-Date; continue }
                if (((Get-Date) - $stall).TotalSeconds -gt $limit) {
                    # The pid file is written right after New-Object returns. On a very short
                    # timeout the watchdog can fire in the same instant, read nothing, kill
                    # nothing - and leave an orphan behind. Retry briefly before giving up.
                    $stray = @()
                    for ($r = 0; $r -lt 3; $r++) {
                        if (Test-Path -LiteralPath $pidFile) {
                            $stray = @(Get-Content -LiteralPath $pidFile -Encoding ASCII -ErrorAction SilentlyContinue |
                                       Where-Object { $_ -match '^\s*\d+\s*$' } |
                                       ForEach-Object { [int]$_.Trim() })
                            if ($stray.Count -gt 0) { break }
                        }
                        Start-Sleep -Milliseconds 700
                    }
                    # The retry costs ~2 s. If the run wrote to the log meanwhile it was not
                    # hung after all - keep waiting instead of killing a healthy export.
                    $now = (Get-Item -LiteralPath $log).LastWriteTime
                    if ($now -gt $last) { $last = $now; $stall = Get-Date; continue }
                    $who = if ($stray.Count -gt 0) { $stray -join ',' } else { 'unknown' }
                    $msg = (Get-Date -Format 'HH:mm:ss') + ' TIMEOUT|log silent ' + $limit + 's; own pid=' +
                           $who + '; killing ' + $stray.Count + [Environment]::NewLine
                    [System.IO.File]::AppendAllText($log, $msg, $enc)
                    foreach ($sp in $stray) {
                        # PID may already belong to something else; only end a WINWORD.
                        $sp0 = Get-Process -Id $sp -ErrorAction SilentlyContinue
                        if ($sp0 -and $sp0.ProcessName -eq 'WINWORD') {
                            Stop-Process -Id $sp -Force -ErrorAction SilentlyContinue
                        }
                    }
                    Start-Sleep -Milliseconds 800
                    # Same PID-reuse guard for the host: compare StartTime, not just the number.
                    $pp = Get-Process -Id $parent -ErrorAction SilentlyContinue
                    $same = $false
                    if ($pp) {
                        if ($parentTicks -gt 0) {
                            try { $same = ($pp.StartTime.Ticks -eq $parentTicks) } catch { $same = $false }
                        } else {
                            $same = ($pp.ProcessName -like 'powershell*')
                        }
                    }
                    if ($same) { Stop-Process -Id $parent -Force -ErrorAction SilentlyContinue }
                    break
                }
            }
        }
        W ('watchdog on; timeout=' + $TimeoutSec + 's; job=' + $watch.Id)
    } catch {
        $watch = $null
        W ('watchdog FAILED|' + $_.Exception.Message + '; continuing without it')
    }
}

$w = $null
$mine = @()   # PID(s) of the instance this run owns; empty means "unknown, do not kill"
$ownDocs = @{}      # absolute path of every document THIS run opened -> $true
$refuseQuit = $false # set when the instance holds a document this run did not open
$warnedForeign = $false
# Taken immediately before New-Object, not at the top of the file: the narrower the window
# between snapshot and instance creation, the smaller the chance of claiming a Word window
# the user opened in the meantime.
$pre = @(Get-Process WINWORD -ErrorAction SilentlyContinue | ForEach-Object { $_.Id })
try {
    $w = New-Object -ComObject Word.Application
    $w.Visible = $true
    $w.DisplayAlerts = 0
    # Claim the instance immediately: everything killed later is filtered through $mine, so a
    # Word window the user opens during the run can never be mistaken for ours. Among the new
    # PIDs the earliest StartTime is ours (New-Object returned before any user action).
    for ($k = 0; $k -lt 10; $k++) {
        $new = @(Get-Process WINWORD -ErrorAction SilentlyContinue | Where-Object { $pre -notcontains $_.Id })
        if ($new.Count -gt 0) {
            $pick = $new[0]
            try { $pick = @($new | Sort-Object StartTime)[0] } catch {}
            $mine = @($pick.Id)
            break
        }
        Start-Sleep -Milliseconds 300
    }
    if ($mine.Count -gt 0) {
        Set-Content -LiteralPath $pidFile -Value ($mine -join "`r`n") -Encoding ASCII
        W ('instance pid=' + ($mine -join ','))
    } else {
        W 'instance pid unknown; a leftover would be reported, not killed'
    }
    W 'instance ready'
    foreach ($it in $items) {
        $src = $it[0]
        $dst = $it[1]
        $leaf = [IO.Path]::GetFileName($dst)
        $d = $null
        $ownDocs[[IO.Path]::GetFullPath($src)] = $true
        # Between two items the instance should be empty. Anything still open belongs to the
        # user: Word hands a double-clicked .docx to a running instance over DDE instead of
        # starting a new process, so the user's document can land in here.
        try {
            if ($w.Documents.Count -gt 0 -and -not $warnedForeign) {
                $names = @()
                for ($di = 1; $di -le $w.Documents.Count; $di++) {
                    try { $names += [string]$w.Documents.Item($di).FullName } catch {}
                }
                $warnedForeign = $true
                W ('FOREIGN DOC STAYS OPEN|' + ($names -join '; '))
                Write-Output ('WARN a document is open in the render instance: ' + ($names -join '; ') + ' - this run will not close it')
            }
        } catch { W ('DOCSCAN_FAIL|' + $_.Exception.Message) }
        try {
            W ('open: exists=' + (Test-Path -LiteralPath $src))
            $d = $w.Documents.Open($src, $false, $false)
            W ('opened; pages=' + $d.ComputeStatistics(2))
            W ('export: ' + $leaf)
            $sw = [Diagnostics.Stopwatch]::StartNew()
            $d.ExportAsFixedFormat($dst, 17)
            $sw.Stop()
            $exists = Test-Path -LiteralPath $dst
            W ('OK|' + $leaf + '|' + $sw.Elapsed.TotalSeconds.ToString('0.0') + 's|exists=' + $exists)
            Write-Output ('OK|' + $dst)
        }
        catch {
            W ('FAIL|' + $leaf + '|' + $_.Exception.Message)
            Write-Output ('FAIL|' + $dst + '|' + $_.Exception.Message)
        }
        finally {
            if ($d) { try { $d.Close(0); W 'closed' } catch { W ('CLOSE_FAIL|' + $_.Exception.Message) } }
        }
    }
}
catch {
    W ('FATAL|' + $_.Exception.Message)
    Write-Output ('FATAL|' + $_.Exception.Message)
}
finally {
    if ($w) {
        # Quit() closes the whole application, not just the documents this run opened. Word
        # routes a double-clicked .docx to an already running instance through DDE, so a
        # document the user opened during the export can be sitting in here - and with
        # DisplayAlerts=0 the "save changes?" prompt never appears, so unsaved work would be
        # discarded silently. Scan first; refuse to quit if anything foreign is loaded.
        $foreign = @()
        try {
            $cnt = 0
            try { $cnt = $w.Documents.Count } catch { $cnt = 0 }
            for ($di = 1; $di -le $cnt; $di++) {
                $fp = ''
                try { $fp = [string]$w.Documents.Item($di).FullName } catch { continue }
                if (-not $fp) { continue }
                $key = [IO.Path]::GetFullPath($fp)
                if (-not $ownDocs.ContainsKey($key)) { $foreign += $fp }
            }
        } catch { W ('DOCSCAN_FAIL|' + $_.Exception.Message) }

        if ($foreign.Count -gt 0) {
            $refuseQuit = $true
            W ('FOREIGN DOC|' + ($foreign -join '; '))
            W 'QUIT SKIPPED|instance holds a document this run did not open; leaving it alive'
            try { $w.DisplayAlerts = -1 } catch {}      # give the user back the save prompt
            try { $w.Visible = $true } catch {}
            Write-Output ('WARN quit skipped: your document is open in the render instance (' + ($foreign -join '; ') + '); it was NOT closed. Save it and close that Word window yourself.')
        } else {
            try { $w.Quit(); W 'quit sent' } catch { W ('QUIT_FAIL|' + $_.Exception.Message) }
        }
        try { [Runtime.InteropServices.Marshal]::ReleaseComObject($w) | Out-Null } catch {}
        [GC]::Collect(); [GC]::WaitForPendingFinalizers()
    }
    $left = @()
    if ($refuseQuit) {
        W 'cleanup skipped|instance kept alive on purpose (holds a user document)'
    }
    elseif ($mine.Count -gt 0) {
        for ($k = 0; $k -lt 15; $k++) {
            # Filter on the name too: the PID could have been recycled by now, in which case
            # it is somebody else's process and must not be touched.
            $left = @(Get-Process -Id $mine -ErrorAction SilentlyContinue | Where-Object { $_.ProcessName -eq 'WINWORD' })
            if ($left.Count -eq 0) { break }
            Start-Sleep -Milliseconds 1000
        }
        W ('WINWORD_LEFT=' + $left.Count)
    } else {
        # Ownership unknown: report what looks new, kill nothing. Leaving a stray is bad;
        # killing a window the user opened is worse.
        $new = @(Get-Process WINWORD -ErrorAction SilentlyContinue | Where-Object { $pre -notcontains $_.Id })
        W ('WINWORD_LEFT=unknown (own pid not identified); new WINWORD=' + $new.Count)
        if ($new.Count -gt 0) {
            Write-Output ('WARN cannot prove ownership; not killing WINWORD: ' + (($new | ForEach-Object { $_.Id }) -join ','))
        }
    }
    if ($left.Count -gt 0) {
        # Quit() is not always enough: the instance lingers for minutes and the next run
        # then collides with it ("file in use" modal, then RPC_E_CALL_REJECTED everywhere).
        # Only the PID recorded for this run is killed; anything else on the desktop is
        # invisible to this block.
        $ids = ($left | ForEach-Object { $_.Id }) -join ','
        W ('leftover PID: ' + $ids + '; forcing stop')
        Stop-Process -Id ($left | ForEach-Object { $_.Id }) -Force -ErrorAction SilentlyContinue
        Start-Sleep -Milliseconds 800
        $still = @(Get-Process -Id $mine -ErrorAction SilentlyContinue | Where-Object { $_.ProcessName -eq 'WINWORD' })
        W ('after force; WINWORD_LEFT=' + $still.Count)
        if ($still.Count -gt 0) {
            Write-Output ('WARN leftover WINWORD: ' + (($still | ForEach-Object { $_.Id }) -join ','))
        } else {
            Write-Output ('CLEANED leftover WINWORD: ' + $ids + ' (forced; Word may offer a recovery pane next launch - choose No)')
        }
    }
    if ($watch) {
        try { Stop-Job $watch -ErrorAction SilentlyContinue } catch {}
        try { Remove-Job $watch -Force -ErrorAction SilentlyContinue } catch {}
        W 'watchdog off'
    }
}
# NOTE: statements placed after this point never run once a COM call has thrown (measured
# twice on PS 5.1). Keep every step that matters inside the try/finally above.
