$ErrorActionPreference = "Continue"
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$exe = 'D:\project\EasyWrite\EasyWriteClient\Test\WordDocumentRoundtripTest\bin\Debug\net472\WordDocumentRoundtripTest.exe'
$log = 'D:\project\EasyWrite\EasyWriteClient\Test\overnight-logs\wh-ppt\word-class2-y2.log'
"START $(Get-Date -Format o)" | Set-Content -Path $log -Encoding UTF8
& $exe --case=wh-ppt --host=word --wh-ids=370,421,422,423,424,425,429,439,464,467,495,547,548,549,550 *>&1 | ForEach-Object { $_ | Add-Content -Path $log -Encoding utf8 }
