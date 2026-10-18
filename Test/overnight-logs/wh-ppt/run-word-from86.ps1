$ErrorActionPreference = "Continue"
try { chcp 65001 > $null } catch {}
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$OutputEncoding = [System.Text.Encoding]::UTF8
$exe = 'D:\project\EasyWrite\EasyWriteClient\Test\WordDocumentRoundtripTest\bin\Debug\net472\WordDocumentRoundtripTest.exe'
$log = 'D:\project\EasyWrite\EasyWriteClient\Test\overnight-logs\wh-ppt\word-from86.log'
$sum = 'D:\project\EasyWrite\EasyWriteClient\Test\overnight-logs\wh-ppt\summary.txt'
"START $(Get-Date -Format o) host=word from=create-line-right-r4" | Set-Content -Path $log -Encoding UTF8
& $exe --case=wh-ppt --host=word --wh-from=create-line-right-r4 *>&1 | ForEach-Object { $_ | Add-Content -Path $log -Encoding utf8 }
$code = $LASTEXITCODE
Add-Content -Path $log -Encoding UTF8 -Value "EXIT $code $(Get-Date -Format o)"
Add-Content -Path $sum -Encoding UTF8 -Value "host=word from86 exit=$code $(Get-Date -Format o)"
