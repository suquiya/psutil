Param (
    [parameter(Mandatory = $true)][String]$url,
    [Alias('out')][String]$out_filename = $null,
    [Alias('time')][bool]$timestamp = $true,
    [Alias('open')][bool]$open_folder = $true
)

$dl_target_url = $url.Trim();

if ($out_filename -ne $null) {
    $filename = $out_filename.Trim();
} else {
    $filename = $dl_target_url.Substring($dl_target_url.LastIndexOf('/') + 1);
    $extension = "";
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
    if ($timestamp) {
        $now_str = [DateTime]::Now.ToString('yyyyMMddHHmmss');
        $filename = "${filename}_${now_str}${extension}";
    }else{
        $filename = "${filename}${extension}";
    }
}


$dl_folder = $env:USERPROFILE;
$dl_folder = "${dl_folder}\Downloads";

$dest_path = "${dl_folder}\${filename}"

xh $dl_target_url -d -o $dest_path

Write-Output "Downloaded to: $dest_path"

if ($open_folder) {
    $argment = "/select,`"${dest_path}`""
    Start-Process explorer.exe -ArgumentList $argment
}
