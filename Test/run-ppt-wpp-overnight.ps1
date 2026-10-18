# PPT / WPP 图表套件过夜跑。结构 300 + 属性 250，两侧各一遍。失败不中断。
# 日志：EasyWriteClient/Test/overnight-logs/
$ErrorActionPreference = 'Continue'
$root = 'D:\project\EasyWrite'
$proj = Join-Path $root 'EasyWriteClient\Test\PptChartRoundtripTest\PptChartRoundtripTest.csproj'
$exe = Join-Path $root 'EasyWriteClient\Test\PptChartRoundtripTest\bin\Debug\net472\PptChartRoundtripTest.exe'
$logDir = Join-Path $root 'EasyWriteClient\Test\overnight-logs'
New-Item -ItemType Directory -Force -Path $logDir | Out-Null

$stamp = Get-Date -Format 'yyyyMMdd-HHmmss'
$log = Join-Path $logDir "ppt-wpp-$stamp.log"
$status = Join-Path $logDir 'STATUS.txt'
Set-Content -Path (Join-Path $logDir 'pid.txt') -Value $PID -Encoding utf8
Set-Content -Path (Join-Path $logDir 'log-path.txt') -Value $log -Encoding utf8

function Write-Status([string]$text) {
    $line = (Get-Date -Format 'yyyy-MM-dd HH:mm:ss') + '  ' + $text
    Add-Content -Path $log -Value $line -Encoding utf8
    Set-Content -Path $status -Value $line -Encoding utf8
}

Write-Status "START pid=$PID"
Write-Status 'BUILD PptChartRoundtripTest'
& dotnet build $proj -c Debug --nologo 2>&1 | ForEach-Object { Add-Content -Path $log -Value ([string]$_) -Encoding utf8 }
$build = $LASTEXITCODE
Write-Status "BUILD exit=$build"
if ($build -ne 0) {
    Write-Status 'STOP build failed'
    exit $build
}

if (-not (Test-Path -LiteralPath $exe)) {
    Write-Status "STOP missing $exe"
    exit 2
}

$steps = @(
    @{ Name = 'ppt-struct'; Args = '--formal-only' },
    @{ Name = 'ppt-attrs'; Args = '--attrs-only' },
    @{ Name = 'wpp-struct'; Args = '--formal-only --host=wpp' },
    @{ Name = 'wpp-attrs'; Args = '--attrs-only --host=wpp' }
)

foreach ($step in $steps) {
    Write-Status ('RUN ' + $step.Name)
    $sw = [Diagnostics.Stopwatch]::StartNew()
    cmd /c "`"$exe`" $($step.Args) >> `"$log`" 2>&1"
    $code = $LASTEXITCODE
    $sw.Stop()
    $mins = [math]::Round($sw.Elapsed.TotalMinutes, 1)
    Write-Status ("DONE " + $step.Name + " exit=$code minutes=$mins")
}

Write-Status 'ALL DONE'
exit 0
