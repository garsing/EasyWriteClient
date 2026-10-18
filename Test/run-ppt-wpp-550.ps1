# PPT 演示再 WPP 演示：结构 300 + 属性 250，共 550。先后台串行，互不并行。
# 不跑 Word / WPS 文字。日志：EasyWriteClient/Test/overnight-logs/ppt-wpp-550/
$ErrorActionPreference = "Continue"
try { chcp 65001 > $null } catch { }
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$OutputEncoding = [System.Text.Encoding]::UTF8
$exe = Join-Path $PSScriptRoot "PptChartRoundtripTest\bin\Debug\net472\PptChartRoundtripTest.exe"
$logDir = Join-Path $PSScriptRoot "overnight-logs\ppt-wpp-550"
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
$summary = Join-Path $logDir "summary.txt"
Set-Content -Path $summary -Value ("START " + (Get-Date -Format o)) -Encoding UTF8

function Run-Host([string]$hostName, [string[]]$extraArgs) {
    $log = Join-Path $logDir ($hostName + ".log")
    $stamp = Get-Date -Format o
    "START $stamp host=$hostName" | Set-Content -Path $log -Encoding UTF8
    & $exe @extraArgs *>&1 | Out-File -FilePath $log -Append -Encoding utf8
    $code = $LASTEXITCODE
    "EXIT $code $(Get-Date -Format o)" | Add-Content -Path $log -Encoding UTF8
    Add-Content -Path $summary -Encoding UTF8 -Value ("host=$hostName exit=$code $(Get-Date -Format o)")
}

if (-not (Test-Path $exe)) {
    Add-Content -Path $summary -Encoding UTF8 -Value "MISSING $exe"
    exit 1
}

Run-Host "ppt" @("--suite=all")
Run-Host "wpp" @("--suite=all", "--host=wpp")
Add-Content -Path $summary -Encoding UTF8 -Value ("ALL DONE " + (Get-Date -Format o))
Remove-Item $lock -ErrorAction SilentlyContinue
exit 0
