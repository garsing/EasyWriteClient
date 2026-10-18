$ErrorActionPreference = "Continue"
try { chcp 65001 > $null } catch {}
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$OutputEncoding = [System.Text.Encoding]::UTF8
$exe = 'D:\project\EasyWrite\EasyWriteClient\Test\WordDocumentRoundtripTest\bin\Debug\net472\WordDocumentRoundtripTest.exe'
$log = 'D:\project\EasyWrite\EasyWriteClient\Test\overnight-logs\wh-ppt\word-from280-fix3.log'
"START $(Get-Date -Format o)" | Set-Content -Path $log -Encoding UTF8
& $exe --case=wh-ppt --host=word --wh-from=create-combo-title-legend-labels *>&1 | ForEach-Object { $_ | Add-Content -Path $log -Encoding utf8 }
