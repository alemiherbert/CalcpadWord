<#
.SYNOPSIS
  Removes the Calcpad add-in from Word for the current user (registry entries, and with
  -RemoveFiles the installed files too). Documents keep their last calculated results.
#>
param([switch]$RemoveFiles)
$ErrorActionPreference = 'Stop'

$clsid  = '{587D11AB-E524-453A-80FE-CCB3017CFA8B}'
$progId = 'CalcpadWord.Connect'

foreach ($key in "HKCU:\Software\Microsoft\Office\Word\Addins\$progId",
                 "HKCU:\Software\Classes\CLSID\$clsid",
                 "HKCU:\Software\Classes\$progId") {
    if (Test-Path $key) { Remove-Item $key -Recurse; Write-Host "Removed $key" }
}

if ($RemoveFiles) {
    $deploy = Join-Path $env:LOCALAPPDATA 'CalcpadWord\AddIn'
    if (Test-Path $deploy) {
        try { Remove-Item $deploy -Recurse -Force -ErrorAction Stop; Write-Host "Removed $deploy" }
        catch { Write-Host "Some files in $deploy are in use by a running Word. Close Word and run this again to remove them." }
    }
}
Write-Host 'Restart Word to finish.'
