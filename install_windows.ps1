<#
Description: "Installs" the 3D Printed Metric thread profiles into your Fusion360 `ThreadData` directory
Author: Peter Lanoie https://github.com/peterlanoie
#>

$fusionDir = "$($env:LOCALAPPDATA)\Autodesk\webdeploy\Production\"
$threadFileName = "3DPrintedMetric.xml"

#<version ID>\Fusion\Server\Fusion\Configuration\ThreadData

Write-Host "Searching install path for ThreadData folder: $fusionDir"
$dirs = Get-ChildItem -Path $fusionDir -Recurse -Directory -Filter ThreadData
if ($dirs.Length -gt 1) {
    # This is uncommon and shouldn't happen, but want to account for it just in case
    Write-Warning "Found more than one 'ThreadData' folder:"
    Write-Host $dirs
    return -1
} else {
    Write-Host "Found fusion thread install path: "
    Write-Host "    $dirs"
}

# back up any existing copies
$existing = Get-ChildItem -Path $dirs -Filter "$threadFileName*"
if ($existing.Count -gt 0) {
    Write-Warning "Found existing thread profiles, backing up the current one"
    Move-Item -Path "$dirs\$threadFileName" -Destination "$dirs\$($threadFileName)_$($existing.Count)"
}

# copy it
Write-Host "Copying thread profile file to Fusion."
Write-Host "    $threadFileName => $dirs"
Copy-Item -Path .\3DPrintedMetric.xml -Destination $dirs

# verify copy
$existing = Get-ChildItem -Path $dirs -Filter $threadFileName
if ($existing.Count -eq 1) {
    Write-Host "File copied successfully. Happy threading!"
}

Write-Host ""
Write-Host "Be sure to restart fusion to load the new profile."