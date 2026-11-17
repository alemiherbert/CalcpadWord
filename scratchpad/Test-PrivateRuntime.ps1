# Does what CalcpadWord.Loader.dll will do, in this process: load hostfxr from a private runtime,
# then create the add-in through COM. Reports where .NET was loaded from.
param([Parameter(Mandatory)][string]$Fxr)
Add-Type -Namespace W -Name K -MemberDefinition '[DllImport("kernel32.dll", CharSet = CharSet.Unicode, SetLastError = true)] public static extern IntPtr LoadLibraryW(string path);'
$h = [W.K]::LoadLibraryW($Fxr)
"hostfxr preloaded: $($h -ne [IntPtr]::Zero)"
$o = [Activator]::CreateInstance([Type]::GetTypeFromCLSID([guid]'587D11AB-E524-453A-80FE-CCB3017CFA8B'))
"add-in object created: $($null -ne $o)"
Get-Process -Id $PID -Module | Where-Object { $_.ModuleName -match '^(hostfxr|hostpolicy|coreclr|CalcpadWord.*)\.dll$' } |
    ForEach-Object { '  ' + $_.ModuleName + '  <-  ' + $_.FileName }
