# 从过夜日志抽出最新一条进度，写到 NOW.txt。主进程结束后自己退出。
$ErrorActionPreference = 'Continue'
$dir = Split-Path -Parent $MyInvocation.MyCommand.Path
$log = (Get-Content -LiteralPath (Join-Path $dir 'log-path.txt') -Raw).Trim()
$now = Join-Path $dir 'NOW.txt'
$mainPid = [int](Get-Content -LiteralPath (Join-Path $dir 'pid.txt') -Raw).Trim()

function Read-Tail([string]$path) {
    $fs = [IO.File]::Open($path, 'Open', 'Read', 'ReadWrite')
    try {
        $len = [Math]::Min(64000, [int]$fs.Length)
        if ($len -le 0) { return '' }
        $fs.Seek(-$len, 'End') | Out-Null
        $buf = New-Object byte[] $len
        [void]$fs.Read($buf, 0, $len)
        return [Text.Encoding]::UTF8.GetString($buf)
    }
    finally {
        $fs.Close()
    }
}

while ($true) {
    $alive = Get-Process -Id $mainPid -ErrorAction SilentlyContinue
    $page = ''
    $case = ''
    try {
        $lines = (Read-Tail $log) -split "`r?`n"
        foreach ($line in $lines) {
            if ($line -match '(\d+/\d+)') { $page = $Matches[1] }
            if ($line -match '\[(OK|FAIL|SKIP)\]\s*(.*)$') { $case = $Matches[1] + '  ' + $Matches[2].Trim() }
        }
    }
    catch { }

    $stamp = Get-Date -Format 'HH:mm:ss'
    if ($page -or $case) {
        Set-Content -LiteralPath $now -Value @($stamp, $page, $case) -Encoding utf8
    }
    if (-not $alive) {
        Add-Content -LiteralPath $now -Value '主进程已结束，看 STATUS.txt' -Encoding utf8
        break
    }
    Start-Sleep -Seconds 5
}
