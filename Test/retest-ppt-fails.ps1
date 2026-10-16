# 把昨夜 PPT 失败用例按条重跑。过滤名会带上更长的同前缀用例，结果按用例名拆开。
# Object does not exist 的条再单独跑一次。
$ErrorActionPreference = 'Continue'
$exe = 'D:\project\EasyWrite\EasyWriteClient\Test\PptChartRoundtripTest\bin\Debug\net472\PptChartRoundtripTest.exe'
$dir = 'D:\project\EasyWrite\EasyWriteClient\Test\overnight-logs\ppt-retest'
New-Item -ItemType Directory -Force -Path $dir | Out-Null
$summary = Join-Path $dir 'summary.txt'
Set-Content -LiteralPath $summary -Value '' -Encoding utf8

$jobs = @(
    @{ Flag = '--formal-only'; Name = 'create-bar-right-r5' },
    @{ Flag = '--formal-only'; Name = 'create-pie2d-top-r6' },
    @{ Flag = '--formal-only'; Name = 'create-pie2d-bottom-r4' },
    @{ Flag = '--formal-only'; Name = 'create-pie2d-bottom-r5' },
    @{ Flag = '--formal-only'; Name = 'create-pie2d-bottom-r6' },
    @{ Flag = '--formal-only'; Name = 'replace-values-none-r4' },
    @{ Flag = '--formal-only'; Name = 'replace-values-top-r5' },
    @{ Flag = '--formal-only'; Name = 'replace-to-bar-bottom-r6' },
    @{ Flag = '--formal-only'; Name = 'replace-to-line-left-r4' },
    @{ Flag = '--formal-only'; Name = 'replace-to-pie2d-right-r5' },
    @{ Flag = '--formal-only'; Name = 'replace-to-pie3d-none-r6' },
    @{ Flag = '--formal-only'; Name = 'replace-to-combo-y2-top-r4' },
    @{ Flag = '--attrs-only'; Name = 'create-sink-line' },
    @{ Flag = '--attrs-only'; Name = 'create-mix-y2-skin-5' },
    @{ Flag = '--attrs-only'; Name = 'replace-sink-line' },
    @{ Flag = '--attrs-only'; Name = 'replace-sink-pie' },
    @{ Flag = '--attrs-only'; Name = 'replace-sink-combo' },
    @{ Flag = '--attrs-only'; Name = 'replace-title-text' },
    @{ Flag = '--attrs-only'; Name = 'replace-title-font-size' },
    @{ Flag = '--attrs-only'; Name = 'replace-title-font-bold' },
    @{ Flag = '--attrs-only'; Name = 'replace-title-off' },
    @{ Flag = '--attrs-only'; Name = 'replace-title-full' },
    @{ Flag = '--attrs-only'; Name = 'replace-title-long' },
    @{ Flag = '--attrs-only'; Name = 'replace-combo-gradient-label-gap' },
    @{ Flag = '--attrs-only'; Name = 'replace-mix-title-legend-top' },
    @{ Flag = '--attrs-only'; Name = 'replace-mix-title-legend-left' },
    @{ Flag = '--attrs-only'; Name = 'replace-mix-gap-ov-5' },
    @{ Flag = '--attrs-only'; Name = 'replace-mix-line-skin-2' },
    @{ Flag = '--attrs-only'; Name = 'replace-mix-pie-3' },
    @{ Flag = '--attrs-only'; Name = 'replace-mix-pie-4' },
    @{ Flag = '--attrs-only'; Name = 'replace-mix-label-chrome-1' },
    @{ Flag = '--attrs-only'; Name = 'replace-mix-label-chrome-5' },
    @{ Flag = '--attrs-only'; Name = 'replace-mix-style-chrome-1' },
    @{ Flag = '--attrs-only'; Name = 'replace-mix-style-chrome-2' },
    @{ Flag = '--attrs-only'; Name = 'replace-mix-style-chrome-3' },
    @{ Flag = '--attrs-only'; Name = 'replace-mix-y2-skin-1' }
)

function Invoke-One([string]$flag, [string]$name, [string]$log) {
    cmd /c "`"$exe`" $flag --case=$name > `"$log`" 2>&1"
    return $LASTEXITCODE
}

function Get-CaseLines([string]$text) {
    $list = @()
    $lines = $text -split "`r?`n"
    for ($i = 0; $i -lt $lines.Length; $i++) {
        if ($lines[$i] -match '\[(OK|FAIL)\]\s+\S+\s+#\d+\s+(\S+)') {
            $kind = $Matches[1]
            $case = $Matches[2]
            $reason = ''
            if ($kind -eq 'FAIL' -and ($i + 1) -lt $lines.Length) {
                $reason = $lines[$i + 1].Trim()
            }
            $list += [pscustomobject]@{ Kind = $kind; Case = $case; Reason = $reason }
        }
    }
    return $list
}

$seen = @{}
foreach ($job in $jobs) {
    $safe = $job.Name -replace '[^\w\-]', '_'
    $log = Join-Path $dir ($safe + '.log')
    $code = Invoke-One $job.Flag $job.Name $log
    $text = [IO.File]::ReadAllText($log, [Text.Encoding]::UTF8)
    $rows = @(Get-CaseLines $text)
    if ($rows.Count -eq 0) {
        $line = "NOCASE exit=$code filter=$($job.Name)"
        Add-Content -LiteralPath $summary -Value $line -Encoding utf8
        continue
    }
    foreach ($row in $rows) {
        $line = "$($row.Kind)`t$($row.Case)`t$($row.Reason)"
        Add-Content -LiteralPath $summary -Value $line -Encoding utf8
        $seen[$row.Case] = $row
    }
}

$rerun = @($seen.Values | Where-Object { $_.Kind -eq 'FAIL' -and $_.Reason -match 'Object does not exist' })
foreach ($row in $rerun) {
    $safe = 'retry-' + ($row.Case -replace '[^\w\-]', '_')
    $log = Join-Path $dir ($safe + '.log')
    $flag = '--formal-only'
    if ($row.Case -match 'sink-|title-|mix-|combo-') { $flag = '--attrs-only' }
    if ($row.Case -match '^create-bar-|^create-pie2d-|^replace-values-|^replace-to-') { $flag = '--formal-only' }
    [void](Invoke-One $flag $row.Case $log)
    $text = [IO.File]::ReadAllText($log, [Text.Encoding]::UTF8)
    $rows = @(Get-CaseLines $text)
    $hit = $rows | Where-Object { $_.Case -eq $row.Case } | Select-Object -First 1
    if ($hit) {
        Add-Content -LiteralPath $summary -Value ("RETRY`t$($hit.Kind)`t$($hit.Case)`t$($hit.Reason)") -Encoding utf8
    }
    else {
        Add-Content -LiteralPath $summary -Value ("RETRY`tNOCASE`t$($row.Case)") -Encoding utf8
    }
}

Add-Content -LiteralPath $summary -Value 'ALL DONE' -Encoding utf8
