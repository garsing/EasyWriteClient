# Build the Vue frontend, Desktop Release, then YiWrite-Setup.exe
$ErrorActionPreference = "Stop"
$InstallerDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$ClientRoot = Resolve-Path (Join-Path $InstallerDir "..\..")
$FrontendDir = Join-Path $ClientRoot "Shared\Frontend"
$DesktopProj = Join-Path $ClientRoot "Desktop\EasyWriteClient.Desktop.csproj"
$RedistDir = Join-Path $InstallerDir "redist"
$WebView2Setup = Join-Path $RedistDir "MicrosoftEdgeWebview2Setup.exe"
$Iss = Join-Path $InstallerDir "YiWrite.iss"

function Find-Iscc {
    $candidates = @(
        "${env:ProgramFiles(x86)}\Inno Setup 6\ISCC.exe",
        "$env:ProgramFiles\Inno Setup 6\ISCC.exe",
        "$env:LOCALAPPDATA\Programs\Inno Setup 6\ISCC.exe"
    )
    foreach ($path in $candidates) {
        if (Test-Path $path) { return $path }
    }
    $cmd = Get-Command ISCC.exe -ErrorAction SilentlyContinue
    if ($cmd) { return $cmd.Source }
    throw "ISCC.exe not found. Install Inno Setup 6."
}

Write-Host "build frontend"
Push-Location $FrontendDir
try {
    npm run build
    if ($LASTEXITCODE -ne 0) { throw "npm run build failed: $LASTEXITCODE" }
}
finally {
    Pop-Location
}

Write-Host "build Desktop Release"
dotnet build $DesktopProj -c Release --nologo
if ($LASTEXITCODE -ne 0) { throw "dotnet build failed: $LASTEXITCODE" }

$publishRoot = Join-Path $ClientRoot "Desktop\bin\Release\net472"
if (-not (Test-Path (Join-Path $publishRoot "EasyWriteClient.Desktop.exe"))) {
    throw "missing EasyWriteClient.Desktop.exe"
}
if (-not (Test-Path (Join-Path $publishRoot "wwwroot\index.html"))) {
    throw "missing wwwroot\index.html"
}

if (-not (Test-Path $WebView2Setup)) {
    Write-Host "download WebView2 bootstrapper"
    New-Item -ItemType Directory -Force -Path $RedistDir | Out-Null
    Invoke-WebRequest -Uri "https://go.microsoft.com/fwlink/p/?LinkId=2124703" -OutFile $WebView2Setup
}

$iscc = Find-Iscc
Write-Host "compile installer"
& $iscc $Iss
if ($LASTEXITCODE -ne 0) { throw "ISCC failed: $LASTEXITCODE" }

$setup = Join-Path $InstallerDir "output"
Get-ChildItem $setup -Filter *.exe | ForEach-Object { Write-Host ("done: " + $_.FullName) }
