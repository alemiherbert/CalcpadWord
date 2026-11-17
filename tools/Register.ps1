<#
.SYNOPSIS
  Builds the Calcpad add-in and installs it into desktop Word for the current user.
.DESCRIPTION
  Publishes to a new versioned folder under %LOCALAPPDATA%\CalcpadWord\AddIn and registers the
  COM class plus the Word add-in entry under HKCU only, so no administrator rights are needed.
  A Word that is already running keeps the version it loaded; Word picks up the new version the
  next time it starts. Older versions are deleted once no Word has them loaded.
  Remove with Unregister.ps1.
#>
param(
    [ValidateSet('Release', 'Debug')][string]$Configuration = 'Release'
)
$ErrorActionPreference = 'Stop'

$root     = Split-Path $PSScriptRoot -Parent
$project  = Join-Path $root 'src\CalcpadWord.AddIn\CalcpadWord.AddIn.csproj'
$installs = Join-Path $env:LOCALAPPDATA 'CalcpadWord\AddIn'
$deploy   = Join-Path $installs ('v' + (Get-Date -Format 'yyyyMMdd-HHmmss'))
$clsid    = '{587D11AB-E524-453A-80FE-CCB3017CFA8B}'
$progId   = 'CalcpadWord.Connect'
$name     = 'Calcpad for Word'

Write-Host "Publishing ($Configuration) to $deploy"
dotnet publish $project -c $Configuration -o $deploy --nologo -v q
if ($LASTEXITCODE -ne 0) { throw 'Publish failed.' }

$comhost = Join-Path $deploy 'CalcpadWord.AddIn.comhost.dll'
if (-not (Test-Path $comhost)) { throw "Not found: $comhost" }

function Set-Key([string]$path, [hashtable]$values) {
    if (-not (Test-Path $path)) { New-Item -Path $path -Force | Out-Null }
    foreach ($k in $values.Keys) {
        $v = $values[$k]
        $type = if ($v -is [int]) { 'DWord' } else { 'String' }
        New-ItemProperty -Path $path -Name $k -Value $v -PropertyType $type -Force | Out-Null
    }
}

$classes = 'HKCU:\Software\Classes'
Set-Key "$classes\CLSID\$clsid"                @{ '(default)' = $name }
Set-Key "$classes\CLSID\$clsid\InprocServer32" @{ '(default)' = $comhost; 'ThreadingModel' = 'Both' }
Set-Key "$classes\CLSID\$clsid\ProgID"         @{ '(default)' = $progId }
Set-Key "$classes\$progId"                     @{ '(default)' = $name }
Set-Key "$classes\$progId\CLSID"               @{ '(default)' = $clsid }
Set-Key "HKCU:\Software\Microsoft\Office\Word\Addins\$progId" @{
    FriendlyName = $name
    Description  = 'Evaluates Calcpad expressions and blocks in Word documents'
    LoadBehavior = 3   # load at startup
}
Write-Host "Registered $progId -> $comhost"

# Remove older versions (and files from the earlier single-folder layout) that no Word has loaded.
$kept = 0
Get-ChildItem $installs -Force | Where-Object { $_.FullName -ne $deploy } | ForEach-Object {
    try { Remove-Item $_.FullName -Recurse -Force -ErrorAction Stop }
    catch { $kept++ }
}
if ($kept -gt 0) { Write-Host "Kept $kept older file(s)/folder(s) still in use by a running Word; they are removed next time." }

if (Get-Process WINWORD -ErrorAction SilentlyContinue) {
    Write-Host 'Word is running: it keeps the version it loaded. Restart Word to use this version.'
}
