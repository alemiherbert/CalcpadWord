<#
.SYNOPSIS
  Builds src\CalcpadWord.Loader\CalcpadWord.Loader.dll from loader.c.
.DESCRIPTION
  The loader is the small native DLL Word loads; it starts the .NET runtime bundled with the
  add-in (see loader.c). The built DLL is kept next to its source, so building the installer
  needs no C compiler: run this only after changing loader.c or loader.rc.
  It uses Zig's C compiler, downloaded once to %LOCALAPPDATA%\CalcpadWord-build (no install, no
  administrator rights) and checked against the published SHA-256.
#>
$ErrorActionPreference = 'Stop'

$zigVersion = '0.17.1'
$zigUrl     = "https://ziglang.org/download/$zigVersion/zig-x86_64-windows-$zigVersion.zip"
$zigSha256  = 'b03f477a1943673e02aca0f570e8030dffd536ce1f45cbc4d3326ec31159562e'

$root   = Split-Path $PSScriptRoot -Parent
$source = Join-Path $root 'src\CalcpadWord.Loader'
$cache  = Join-Path $env:LOCALAPPDATA 'CalcpadWord-build'
$zigDir = Join-Path $cache "zig-x86_64-windows-$zigVersion"
$zig    = Join-Path $zigDir 'zig.exe'

if (-not (Test-Path $zig)) {
    New-Item -ItemType Directory -Force $cache | Out-Null
    $zip = Join-Path $cache "zig-$zigVersion.zip"
    Write-Host "Downloading Zig $zigVersion (C compiler, about 100 MB)..."
    $ProgressPreference = 'SilentlyContinue'
    Invoke-WebRequest $zigUrl -OutFile $zip -UseBasicParsing
    $hash = (Get-FileHash $zip -Algorithm SHA256).Hash
    if ($hash -ne $zigSha256) { Remove-Item $zip; throw "The Zig download doesn't match its published SHA-256 ($hash)." }
    Expand-Archive $zip -DestinationPath $cache -Force
    Remove-Item $zip
}

$dll = Join-Path $source 'CalcpadWord.Loader.dll'
Push-Location $source
try {
    # No C runtime (-nostdlib), so Zig's Windows headers are named explicitly.
    $headers = Join-Path $zigDir 'lib\libc\include\any-windows-any'
    & $zig cc -target x86_64-windows-gnu -shared -O2 -s -nostdlib '-Wl,--entry=DllMain' -isystem $headers `
        -o $dll loader.c loader.rc -lkernel32
    if ($LASTEXITCODE -ne 0) { throw 'Compiling the loader failed.' }
}
finally {
    Pop-Location
    # Zig also writes an import library (named after the source) that nothing needs.
    Get-ChildItem $source -File | Where-Object Extension -in '.lib', '.pdb' | Remove-Item
}
Write-Host ("Built {0} ({1:N0} bytes)" -f $dll, (Get-Item $dll).Length)
