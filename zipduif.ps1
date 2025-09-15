Param (
    [parameter(Mandatory = $true)][String]$url
)

$dl_target_url = $url.Trim();
$filename = $dl_target_url.Substring($dl_target_url.LastIndexOf('/') + 1);

if ($filename.IndexOf('?') -gt -1) {
    $filename = $filename.Substring(0, $filename.IndexOf('?'));
}

if ($filename.LastIndexOf('.') -gt -1) {
    $extension = $filename.Substring($filename.LastIndexOf('.'));
    if ($extension.ToLower() -eq '.zip') {
        $filename = $filename.Substring(0, $filename.LastIndexOf('.'));
    }
}

if ($filename.Length -eq 0) {
    $filename = "download";
}

$dl_folder = $env:USERPROFILE;
$dl_folder = "${dl_folder}\Downloads";
$now_str = [DateTime]::Now.ToString('yyyyMMddHHmmss');

$filename = "${filename}_${now_str}"
$zipPath = "${dl_folder}\$filename.zip"

xh $dl_target_url -d -o $zipPath

$dest_path = "${dl_folder}\${filename}"

Write-Output "Extracting ${zipPath} to ${dest_path}"
Expand-Archive -Path $zipPath -DestinationPath $dest_path
Write-Output "Extraction complete"

Get-ChildItem $dest_path -Recurse -File | Where-Object { $_.Extension -eq ".ttf" -or $_.Extension -eq ".otf" -or $_.Extension -eq ".ttc" } | ForEach-Object { Invoke-Item $_.FullName }