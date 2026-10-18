param(
    [string]$Src,
    [string]$Dst
)
# 过夜日志里夹了 UTF-16 的编译输出，编辑器会当成二进制。把最新一份转成可打开的 UTF-8。
$ErrorActionPreference = "Continue"
$logDir = Join-Path $PSScriptRoot "overnight-logs"
if (-not $Dst) {
    $Dst = Join-Path $logDir "ppt-wpp-live.txt"
}

function Convert-MixedLog([string]$path) {
    $fs = [System.IO.File]::Open($path, "Open", "Read", "ReadWrite")
    try {
        $ms = New-Object System.IO.MemoryStream
        $fs.CopyTo($ms)
        $b = $ms.ToArray()
    }
    finally {
        $fs.Close()
    }

    $i0 = 0
    if ($b.Length -ge 3 -and $b[0] -eq 0xEF -and $b[1] -eq 0xBB -and $b[2] -eq 0xBF) {
        $i0 = 3
    }

    $utf16At = -1
    for ($i = $i0; $i -lt $b.Length - 1; $i++) {
        if ($b[$i + 1] -eq 0 -and $b[$i] -ne 0) {
            $utf16At = $i
            break
        }
    }

    $utf8 = New-Object System.Text.UTF8Encoding $false
    $uni = [System.Text.Encoding]::Unicode
    if ($utf16At -lt 0) {
        return $utf8.GetString($b, $i0, $b.Length - $i0)
    }

    $marker = [System.Text.Encoding]::ASCII.GetBytes("[OK] B")
    $back = -1
    for ($i = $utf16At; $i -le $b.Length - $marker.Length; $i++) {
        $hit = $true
        for ($k = 0; $k -lt $marker.Length; $k++) {
            if ($b[$i + $k] -ne $marker[$k]) {
                $hit = $false
                break
            }
        }
        if ($hit) {
            $back = $i
            break
        }
    }

    $head = $utf8.GetString($b, $i0, $utf16At - $i0)
    if ($back -lt 0) {
        $midLen = $b.Length - $utf16At
        if (($midLen % 2) -eq 1) { $midLen-- }
        return $head + $uni.GetString($b, $utf16At, $midLen)
    }

    $midLen = $back - $utf16At
    if (($midLen % 2) -eq 1) { $midLen-- }
    $mid = $uni.GetString($b, $utf16At, $midLen)
    $tail = $utf8.GetString($b, $back, $b.Length - $back)
    return $head + $mid + $tail
}

while ($true) {
    if (-not $Src) {
        $latest = Get-ChildItem -Path $logDir -Filter "ppt-wpp-20*.log" | Sort-Object LastWriteTime -Descending | Select-Object -First 1
        if (-not $latest) {
            Start-Sleep -Seconds 2
            continue
        }
        $SrcPath = $latest.FullName
    }
    else {
        $SrcPath = $Src
    }

    try {
        $text = Convert-MixedLog $SrcPath
        $out = New-Object System.Text.UTF8Encoding $true
        [System.IO.File]::WriteAllText($Dst, $text, $out)
    }
    catch {
    }

    if ($Src) { break }
    Start-Sleep -Seconds 3
}
