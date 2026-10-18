$ErrorActionPreference = "Continue"
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$exe = 'D:\project\EasyWrite\EasyWriteClient\Test\WordDocumentRoundtripTest\bin\Debug\net472\WordDocumentRoundtripTest.exe'
$log = 'D:\project\EasyWrite\EasyWriteClient\Test\overnight-logs\wh-ppt\wps-class2-y2.log'
"START $(Get-Date -Format o)" | Set-Content -Path $log -Encoding UTF8
& $exe --case=wh-ppt --host=wps --wh-ids=157,167,177,187,197,207,217,227,257,267,277,287,297,304,314,339,342,370,421,422,423,424 *>&1 | ForEach-Object { $_ | Add-Content -Path $log -Encoding utf8 }
