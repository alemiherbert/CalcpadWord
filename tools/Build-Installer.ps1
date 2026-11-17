<#
.SYNOPSIS
  Builds out\installer\CalcpadWord-Setup-<version>.exe: one file with everything inside.
.DESCRIPTION
  The payload, which Setup installs to %LOCALAPPDATA%\Programs\CalcpadWord\app-<version>-<time>:
    - the add-in, with Calcpad's engine (Calcpad.Core and Calcpad.OpenXml) and their libraries;
    - CalcpadWord.Loader.dll, which Word loads: it starts the runtime below, then the add-in;
    - runtime\: a private copy of the .NET runtime (host\fxr and shared\Microsoft.NETCore.App),
      taken from the .NET installed on this computer, so teammates need no .NET;
    - fonts\: the math fonts and Libre Franklin (installer\fonts), which Setup installs per user;
    - the third-party notices.
  Setup itself is self-contained, so it runs without .NET too. Everything installs per user,
  without administrator rights. Setup isn't code-signed, so SmartScreen may warn on other PCs.
  The loader DLL is kept prebuilt next to its source; tools\Build-Loader.ps1 rebuilds it.
#>
param([ValidateSet('Release', 'Debug')][string]$Configuration = 'Release')
$ErrorActionPreference = 'Stop'

$root    = Split-Path $PSScriptRoot -Parent
$out     = Join-Path $root 'out\installer'
$payload = Join-Path $out 'payload'
$zip     = Join-Path $out 'payload.zip'
$setup   = Join-Path $out 'setup'

# The loader: rebuilt only when its source is newer.
$loaderDir = Join-Path $root 'src\CalcpadWord.Loader'
$loader = Join-Path $loaderDir 'CalcpadWord.Loader.dll'
$loaderSource = Get-ChildItem $loaderDir -Include loader.c, loader.rc -Recurse | Sort-Object LastWriteTime | Select-Object -Last 1
if (-not (Test-Path $loader) -or (Get-Item $loader).LastWriteTime -lt $loaderSource.LastWriteTime) {
    & (Join-Path $PSScriptRoot 'Build-Loader.ps1')
}

# Empty the folder rather than delete it: an Explorer window open on it would block that.
New-Item -ItemType Directory -Force $out | Out-Null
Get-ChildItem $out -Force | Remove-Item -Recurse -Force

Write-Host 'Publishing the add-in...'
dotnet publish (Join-Path $root 'src\CalcpadWord.AddIn\CalcpadWord.AddIn.csproj') -c $Configuration -o $payload --nologo -v q
if ($LASTEXITCODE -ne 0) { throw 'Publishing the add-in failed.' }
Get-ChildItem $payload -Filter *.pdb | Remove-Item
# English only: leave out Calcpad's translated messages (bg, zh: folders of *.resources.dll).
Get-ChildItem $payload -Directory | Where-Object { @(Get-ChildItem $_.FullName -File | Where-Object Name -notlike '*.resources.dll').Count -eq 0 } |
    Remove-Item -Recurse
Copy-Item $loader $payload

# The .NET runtime the add-in targets (the newest 10.0.x installed here), as a private copy.
$config = Get-Content (Join-Path $payload 'CalcpadWord.AddIn.runtimeconfig.json') -Raw | ConvertFrom-Json
$framework = $config.runtimeOptions.framework
$major = ([version]$framework.version).Major
$dotnetRoot = Split-Path (Get-Command dotnet).Source
function Newest([string]$dir) {
    Get-ChildItem $dir -Directory | Where-Object { $_.Name -match "^$major\.\d+\.\d+$" } | Sort-Object { [version]$_.Name } | Select-Object -Last 1
}
$shared = Newest (Join-Path $dotnetRoot "shared\$($framework.name)")
$fxr = Newest (Join-Path $dotnetRoot 'host\fxr')
if (-not $shared -or -not $fxr) { throw "The .NET $major runtime ($($framework.name)) isn't installed in $dotnetRoot." }
Write-Host "Bundling $($framework.name) $($shared.Name) (hostfxr $($fxr.Name))..."
$runtime = Join-Path $payload 'runtime'
Copy-Item $shared.FullName (New-Item -ItemType Directory -Force (Join-Path $runtime "shared\$($framework.name)")).FullName -Recurse
Copy-Item $fxr.FullName (New-Item -ItemType Directory -Force (Join-Path $runtime 'host\fxr')).FullName -Recurse
Copy-Item (Join-Path $dotnetRoot 'LICENSE.txt'), (Join-Path $dotnetRoot 'ThirdPartyNotices.txt') $runtime

# Fonts and notices.
$fonts = Join-Path $root 'installer\fonts'
foreach ($line in Get-Content (Join-Path $fonts 'fonts.txt') -Encoding UTF8) {
    $file = ($line -split "`t")[0]
    if (-not (Test-Path (Join-Path $fonts $file))) { throw "installer\fonts\fonts.txt lists $file, which isn't there." }
}
Copy-Item $fonts (Join-Path $payload 'fonts') -Recurse
Copy-Item (Join-Path $root 'installer\THIRD-PARTY-NOTICES.txt') $payload

Add-Type -AssemblyName System.IO.Compression.FileSystem
[IO.Compression.ZipFile]::CreateFromDirectory($payload, $zip, [IO.Compression.CompressionLevel]::Optimal, $false)
Write-Host ("Payload: {0:N1} MB unpacked, {1:N1} MB zipped" -f `
    ((Get-ChildItem $payload -Recurse -File | Measure-Object Length -Sum).Sum / 1MB), ((Get-Item $zip).Length / 1MB))

Write-Host 'Building the setup...'
dotnet publish (Join-Path $root 'src\CalcpadWord.Setup\CalcpadWord.Setup.csproj') -c $Configuration -o $setup "-p:PayloadZip=$zip" --nologo -v q
if ($LASTEXITCODE -ne 0) { throw 'Building the setup failed.' }

$exe = Join-Path $setup 'CalcpadWord-Setup.exe'
$version = (Get-Item $exe).VersionInfo.ProductVersion -replace '\+.*$', ''
$final = Join-Path $out "CalcpadWord-Setup-$version.exe"
Copy-Item $exe $final
Write-Host ("Built {0} ({1:N1} MB)" -f $final, ((Get-Item $final).Length / 1MB))
