# Render docx -> PDF via Word COM. Must be executed by the PowerShell tool.
#   1. All paths come from a UTF-8 list file (.renderlist.txt); no CJK on the command line.
#   2. Work dir is chosen by `--prepare` (see CRB_RENDER_DIR); absolute backslash paths.
#   3. Visible=$true + Documents.Open(..., ReadOnly=$false); ReadOnly=$true loops forever.
#   4. Per-file try/catch so one failure does not stop the rest.
#   5. The list is read and normalised BEFORE any COM object exists; the log is appended
#      after every step so a hang is visible while it happens.
#   6. A watchdog job kills the whole thing when the log stops moving (see below). A blocked
#      ExportAsFixedFormat call cannot be interrupted from inside this script, because this
#      thread is the blocked one; only a separate process can notice and act.
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

$base = @(Get-Process WINWORD -ErrorAction SilentlyContinue | ForEach-Object { $_.Id })
W ('start; existing WINWORD=' + $base.Count + '; items=' + $items.Count)

# Watchdog runs in its own process, so it keeps ticking while this thread sits inside a
# blocked COM call. When the log has not moved for $TimeoutSec it records TIMEOUT, kills
# the stray WINWORD, then kills this script so the caller gets a dead job instead of a
# task that shows "running" forever.
$watch = $null
if ($TimeoutSec -gt 0) {
    try {
        $watch = Start-Job -ArgumentList $logPath, $TimeoutSec, $PID -ScriptBlock {
            param($log, $limit, $parent)
            $enc = New-Object System.Text.UTF8Encoding($false)
            $last = (Get-Item -LiteralPath $log).LastWriteTime
            $stall = Get-Date
            while ($true) {
                Start-Sleep -Seconds 2
                $cur = (Get-Item -LiteralPath $log).LastWriteTime
                if ($cur -gt $last) { $last = $cur; $stall = Get-Date; continue }
                if (((Get-Date) - $stall).TotalSeconds -gt $limit) {
                    $msg = (Get-Date -Format 'HH:mm:ss') + ' TIMEOUT|log silent ' + $limit + 's; killing WINWORD then pid ' + $parent + [Environment]::NewLine
                    [System.IO.File]::AppendAllText($log, $msg, $enc)
                    Get-Process WINWORD -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
                    Start-Sleep -Milliseconds 800
                    Stop-Process -Id $parent -Force -ErrorAction SilentlyContinue
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
try {
    $w = New-Object -ComObject Word.Application
    $w.Visible = $true
    $w.DisplayAlerts = 0
    W 'instance ready'
    foreach ($it in $items) {
        $src = $it[0]
        $dst = $it[1]
        $leaf = [IO.Path]::GetFileName($dst)
        $d = $null
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
        try { $w.Quit(); W 'quit sent' } catch { W ('QUIT_FAIL|' + $_.Exception.Message) }
        try { [Runtime.InteropServices.Marshal]::ReleaseComObject($w) | Out-Null } catch {}
        [GC]::Collect(); [GC]::WaitForPendingFinalizers()
    }
    $left = @()
    for ($k = 0; $k -lt 15; $k++) {
        $left = @(Get-Process WINWORD -ErrorAction SilentlyContinue | Where-Object { $base -notcontains $_.Id })
        if ($left.Count -eq 0) { break }
        Start-Sleep -Milliseconds 1000
    }
    W ('WINWORD_LEFT=' + $left.Count)
    if ($left.Count -gt 0) {
        # Quit() is not always enough: the instance lingers for minutes and the next run
        # then collides with it ("file in use" modal, then RPC_E_CALL_REJECTED everywhere).
        # Only instances this script started are killed; a Word window the user opened
        # themselves is in $base and is never touched.
        $ids = ($left | ForEach-Object { $_.Id }) -join ','
        W ('leftover PID: ' + $ids + '; forcing stop')
        Stop-Process -Id ($left | ForEach-Object { $_.Id }) -Force -ErrorAction SilentlyContinue
        Start-Sleep -Milliseconds 800
        $still = @(Get-Process WINWORD -ErrorAction SilentlyContinue | Where-Object { $base -notcontains $_.Id })
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
