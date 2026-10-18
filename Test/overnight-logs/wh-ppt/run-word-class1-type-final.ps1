$ErrorActionPreference = "Continue"
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$exe = 'D:\project\EasyWrite\EasyWriteClient\Test\WordDocumentRoundtripTest\bin\Debug\net472\WordDocumentRoundtripTest.exe'
$log = 'D:\project\EasyWrite\EasyWriteClient\Test\overnight-logs\wh-ppt\word-class1-type-final.log'
"START $(Get-Date -Format o)" | Set-Content -Path $log -Encoding UTF8
& $exe --case=wh-ppt --host=word --wh-ids=158,167,168,177,178,187,188,197,198,207,208,217,218,227,228,237,238,247,248,257,258,267,268,278,287,288,297,298,302,304,314,339,342 *>&1 | ForEach-Object { $_ | Add-Content -Path $log -Encoding utf8 }
