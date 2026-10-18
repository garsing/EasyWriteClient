$ErrorActionPreference = "Continue"
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$exe = 'D:\project\EasyWrite\EasyWriteClient\Test\WordDocumentRoundtripTest\bin\Debug\net472\WordDocumentRoundtripTest.exe'
$log = 'D:\project\EasyWrite\EasyWriteClient\Test\overnight-logs\wh-ppt\word-class1-type-r5.log'
"START $(Get-Date -Format o)" | Set-Content -Path $log -Encoding UTF8
& $exe --case=wh-ppt --host=word --wh-ids=342 *>&1 | ForEach-Object { $_ | Add-Content -Path $log -Encoding utf8 }
