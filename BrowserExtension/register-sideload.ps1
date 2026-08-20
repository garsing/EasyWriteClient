# Register Native Messaging Host for sideload (Chrome + Edge)
param(
    [Parameter(Mandatory = $true)]
    [string]$ExtensionId,

    [Parameter(Mandatory = $false)]
    [string]$BridgeExe = ""
)

$ErrorActionPreference = "Stop"
$NmHostName = "com.yiwrite.browser_bridge"

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
    throw "YiWriteBrowserBridge.exe not found. Pass -BridgeExe with full path."
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
$manifestPath = Join-Path $manifestDir "$NmHostName.json"

$origins = New-Object System.Collections.Generic.List[string]
[void]$origins.Add($origin)
if (Test-Path $manifestPath) {
    try {
        $existing = Get-Content $manifestPath -Raw -Encoding UTF8 | ConvertFrom-Json
        foreach ($o in @($existing.allowed_origins)) {
            if ($o -and -not $origins.Contains([string]$o)) {
                [void]$origins.Add([string]$o)
            }
        }
    } catch {
        # ignore bad existing manifest
    }
}

$originParts = @()
foreach ($o in $origins) {
    $escaped = $o.Replace('\', '\\').Replace('"', '\"')
    $originParts += ('"' + $escaped + '"')
}
$originJson = [string]::Join(", ", $originParts)
$bridgeEscaped = $BridgeExe.Replace('\', '\\').Replace('"', '\"')

$json = @"
{
  "name": "$NmHostName",
  "description": "YiWrite browser bridge",
  "path": "$bridgeEscaped",
  "type": "stdio",
  "allowed_origins": [ $originJson ]
}
"@
Set-Content -Path $manifestPath -Value $json -Encoding UTF8
Write-Host "Wrote $manifestPath"

function Set-NmKey {
    param([string]$Rel)
    $key = "HKCU:\$Rel"
    New-Item -Path $key -Force | Out-Null
    Set-ItemProperty -Path $key -Name "(default)" -Value $manifestPath
    Write-Host "Registry $Rel -> $manifestPath"
}

Set-NmKey -Rel "Software\Google\Chrome\NativeMessagingHosts\$NmHostName"
Set-NmKey -Rel "Software\Microsoft\Edge\NativeMessagingHosts\$NmHostName"

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
            $lines = @(Get-Content $idFile | Where-Object { $_ -and -not $_.StartsWith("#") })
        }
        if ($lines -notcontains $id) {
            Add-Content -Path $idFile -Value $id
        }
        Write-Host "Updated $idFile"
    }
}

Write-Host "Done. Reload the extension (and restart browser if needed)."
