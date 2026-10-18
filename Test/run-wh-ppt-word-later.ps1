# 等 WPS 套件进程退出，并且没有测试 exe 还在跑，再单独跑 Word。
# 只有「一打开就被拒、一条都没过」才重试。用例中途失败不重跑。
param([int]$WaitPid = 0)
$ErrorActionPreference = "Continue"
try { chcp 65001 > $null } catch { }
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$OutputEncoding = [System.Text.Encoding]::UTF8
$exe = Join-Path $PSScriptRoot "WordDocumentRoundtripTest\bin\Debug\net472\WordDocumentRoundtripTest.exe"
$logDir = Join-Path $PSScriptRoot "overnight-logs\wh-ppt"
New-Item -ItemType Directory -Force -Path $logDir | Out-Null

function Test-SuiteBusy {
    if ($WaitPid -gt 0 -and (Get-Process -Id $WaitPid -ErrorAction SilentlyContinue)) {
        return $true
    }
    $hit = Get-CimInstance Win32_Process -Filter "Name='WordDocumentRoundtripTest.exe'" |
        Where-Object { $_.CommandLine -like '*--case=wh-ppt*' }
    return [bool]$hit
}

while (Test-SuiteBusy) {
    Start-Sleep -Seconds 30
}
Start-Sleep -Seconds 20

$log = Join-Path $logDir "word.log"
$sum = Join-Path $logDir "summary.txt"
for ($try = 1; $try -le 12; $try++) {
    if ($try -gt 1) { Start-Sleep -Seconds 30 }
    "START $(Get-Date -Format o) host=word try=$try" | Set-Content -Path $log -Encoding UTF8
    & $exe --case=wh-ppt --host=word *>&1 | Out-File -FilePath $log -Append -Encoding utf8
    $code = $LASTEXITCODE
    Add-Content -Path $log -Encoding UTF8 -Value "EXIT $code $(Get-Date -Format o)"
    Add-Content -Path $sum -Encoding UTF8 -Value "host=word try=$try exit=$code $(Get-Date -Format o)"
    if ($code -eq 0) { break }
    $text = ""
    try { $text = Get-Content $log -Raw -ErrorAction SilentlyContinue } catch { }
    if ($text -match '\[OK\]') { break }
    if ($text -notmatch 'RPC_E_CALL_REJECTED' -and $text -notmatch 'wh-ppt-open') { break }
}
Add-Content -Path $sum -Encoding UTF8 -Value ("WORD LATER DONE " + (Get-Date -Format o))
exit 0
