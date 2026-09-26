Param (
    [parameter(Mandatory = $true)][String]$url,
    [bool]$open_folder = $true
)

$dl_target_url = $url.Trim();
$filename = $dl_target_url.Substring($dl_target_url.LastIndexOf('/') + 1);

if ($filename.IndexOf('?') -gt -1) {
    $filename = $filename.Substring(0, $filename.IndexOf('?'));
}

if ($filename.LastIndexOf('.') -gt -1) {
    $extension = $filename.Substring($filename.LastIndexOf('.'));
    $filename = $filename.Substring(0, $filename.LastIndexOf('.'));
}

if ($filename.Length -eq 0) {
    $filename = "download";
}

$dl_folder = $env:USERPROFILE;
$dl_folder = "${dl_folder}\Downloads";
$now_str = [DateTime]::Now.ToString('yyyyMMddHHmmss');

$filename = "${filename}_${now_str}"
$dest_path = "${dl_folder}\${filename}${extension}"

xh $dl_target_url -d -o $dest_path

Write-Output "Downloaded to: $dest_path"

if ($open_folder) {
    $argment = "/select,`"${dest_path}`""
    Start-Process explorer.exe -ArgumentList $argment
}
