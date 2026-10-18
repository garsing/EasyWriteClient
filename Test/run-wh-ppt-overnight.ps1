# Word / WPS 文字：PPT 结构 300 + 属性 250。先 Word 再 WPS，互不并行。
# 日志：EasyWriteClient/Test/overnight-logs/wh-ppt/
$ErrorActionPreference = "Continue"
try { chcp 65001 > $null } catch { }
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$OutputEncoding = [System.Text.Encoding]::UTF8
$exe = Join-Path $PSScriptRoot "WordDocumentRoundtripTest\bin\Debug\net472\WordDocumentRoundtripTest.exe"
$logDir = Join-Path $PSScriptRoot "overnight-logs\wh-ppt"
New-Item -ItemType Directory -Force -Path $logDir | Out-Null
$lock = Join-Path $logDir "runner.pid"
if (Test-Path $lock) {
    $old = 0
    [void][int]::TryParse((Get-Content $lock -ErrorAction SilentlyContinue | Select-Object -First 1), [ref]$old)
    if ($old -gt 0 -and (Get-Process -Id $old -ErrorAction SilentlyContinue)) {
        exit 0
    }
}
Set-Content -Path $lock -Value $PID -Encoding ASCII
Set-Content -Path (Join-Path $logDir "summary.txt") -Value ("START " + (Get-Date -Format o))

function Run-Host([string]$hostName) {
    $log = Join-Path $logDir ($hostName + ".log")
    $stamp = Get-Date -Format o
    "START $stamp host=$hostName" | Set-Content -Path $log -Encoding UTF8
    & $exe --case=wh-ppt --host=$hostName *>&1 | Out-File -FilePath $log -Append -Encoding utf8
    $code = $LASTEXITCODE
    "EXIT $code $(Get-Date -Format o)" | Tee-Object -FilePath $log -Append
    Add-Content -Path (Join-Path $logDir "summary.txt") -Encoding UTF8 -Value ("host=$hostName exit=$code $(Get-Date -Format o)")
}

if (-not (Test-Path $exe)) {
    Add-Content -Path (Join-Path $logDir "summary.txt") -Encoding UTF8 -Value "MISSING $exe"
    exit 1
}

Run-Host "word"
Run-Host "wps"
Add-Content -Path (Join-Path $logDir "summary.txt") -Encoding UTF8 -Value ("ALL DONE " + (Get-Date -Format o))
exit 0
