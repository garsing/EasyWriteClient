# 侧载：登记 Native Messaging Host（Chrome + Edge）
param(
    [Parameter(Mandatory = $true)]
    [string]$ExtensionId,

    [Parameter(Mandatory = $false)]
    [string]$BridgeExe = ""
)

$ErrorActionPreference = "Stop"
$hostName = "com.yiwrite.browser_bridge"

if (-not $BridgeExe) {
    $candidates = @(
        (Join-Path $PSScriptRoot "..\Desktop\bin\Debug\net472\YiWriteBrowserBridge.exe"),
        (Join-Path $PSScriptRoot "..\Desktop\bin\Release\net472\YiWriteBrowserBridge.exe"),
        (Join-Path $PSScriptRoot "..\BrowserBridge\bin\Debug\net472\YiWriteBrowserBridge.exe"),
        (Join-Path $PSScriptRoot "..\BrowserBridge\bin\Release\net472\YiWriteBrowserBridge.exe")
    )
    foreach ($c in $candidates) {
        $full = [IO.Path]::GetFullPath($c)
        if (Test-Path $full) {
            $BridgeExe = $full
            break
        }
    }
}

if (-not $BridgeExe -or -not (Test-Path $BridgeExe)) {
    throw "找不到 YiWriteBrowserBridge.exe，请用 -BridgeExe 指定完整路径"
}

$BridgeExe = [IO.Path]::GetFullPath($BridgeExe)
$id = $ExtensionId.Trim()
if ($id.StartsWith("chrome-extension://")) {
    $origin = $id.TrimEnd('/') + "/"
} else {
    $origin = "chrome-extension://$id/"
}

$manifestDir = Join-Path $env:LOCALAPPDATA "YiWrite\native-messaging"
New-Item -ItemType Directory -Force -Path $manifestDir | Out-Null
$manifestPath = Join-Path $manifestDir "$hostName.json"

# 合并已有 allowed_origins
$origins = New-Object System.Collections.Generic.List[string]
$origins.Add($origin)
if (Test-Path $manifestPath) {
    try {
        $existing = Get-Content $manifestPath -Raw | ConvertFrom-Json
        foreach ($o in @($existing.allowed_origins)) {
            if ($o -and -not $origins.Contains($o)) { $origins.Add($o) }
        }
    } catch {}
}

$originJson = ($origins | ForEach-Object { '"' + ($_ -replace '\\','\\' -replace '"','\"') + '"' }) -join ", "
$bridgeEscaped = $BridgeExe -replace '\\','\\' -replace '"','\"'
$json = @"
{
  "name": "$hostName",
  "description": "YiWrite browser bridge",
  "path": "$bridgeEscaped",
  "type": "stdio",
  "allowed_origins": [ $originJson ]
}
"@
Set-Content -Path $manifestPath -Value $json -Encoding UTF8
Write-Host "Wrote $manifestPath"

function Set-NmKey([string]$rel) {
    $key = "HKCU:\$rel"
    New-Item -Path $key -Force | Out-Null
    Set-ItemProperty -Path $key -Name "(default)" -Value $manifestPath
    Write-Host "Registry $rel -> $manifestPath"
}

Set-NmKey "Software\Google\Chrome\NativeMessagingHosts\$hostName"
Set-NmKey "Software\Microsoft\Edge\NativeMessagingHosts\$hostName"

# 同步到 Desktop 输出目录的 id 文件（若存在）
$desktopOuts = @(
    (Join-Path $PSScriptRoot "..\Desktop\bin\Debug\net472"),
    (Join-Path $PSScriptRoot "..\Desktop\bin\Release\net472")
)
foreach ($dir in $desktopOuts) {
    $full = [IO.Path]::GetFullPath($dir)
    if (Test-Path $full) {
        $idFile = Join-Path $full "browser-extension-id.txt"
        $lines = @()
        if (Test-Path $idFile) {
            $lines = Get-Content $idFile | Where-Object { $_ -and -not $_.StartsWith("#") }
        }
        if ($lines -notcontains $id) {
            Add-Content -Path $idFile -Value $id
        }
        Write-Host "Updated $idFile"
    }
}

Write-Host "Done. Reload the extension and restart the browser if needed."
