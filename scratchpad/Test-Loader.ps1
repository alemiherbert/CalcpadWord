# Calls CalcpadWord.Loader.dll's DllGetClassObject for the add-in's class, as COM does when Word
# starts the add-in, then reports where .NET was loaded from. No registry involved.
param([Parameter(Mandatory)][string]$Loader)
Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;
public static class LoaderTest
{
    [DllImport("kernel32.dll", CharSet = CharSet.Unicode, SetLastError = true)] static extern IntPtr LoadLibraryW(string path);
    [DllImport("kernel32.dll", CharSet = CharSet.Ansi)] static extern IntPtr GetProcAddress(IntPtr module, string name);
    [UnmanagedFunctionPointer(CallingConvention.StdCall)] delegate int GetClassObject(ref Guid clsid, ref Guid iid, out IntPtr factory);
    [UnmanagedFunctionPointer(CallingConvention.StdCall)] delegate int CanUnloadNow();
    public static string Run(string loader, Guid clsid)
    {
        IntPtr module = LoadLibraryW(loader);
        if (module == IntPtr.Zero) return "LoadLibrary failed: " + Marshal.GetLastWin32Error();
        var get = (GetClassObject)Marshal.GetDelegateForFunctionPointer(GetProcAddress(module, "DllGetClassObject"), typeof(GetClassObject));
        var unload = (CanUnloadNow)Marshal.GetDelegateForFunctionPointer(GetProcAddress(module, "DllCanUnloadNow"), typeof(CanUnloadNow));
        Guid classFactory = new Guid("00000001-0000-0000-C000-000000000046");
        IntPtr factory;
        int hr = get(ref clsid, ref classFactory, out factory);
        string result = string.Format("DllGetClassObject: hr=0x{0:X8}, class factory {1}", hr, factory != IntPtr.Zero ? "returned" : "missing");
        if (factory != IntPtr.Zero) Marshal.Release(factory);
        Guid other = Guid.NewGuid();
        hr = get(ref other, ref classFactory, out factory);
        result += string.Format("; unknown class: hr=0x{0:X8}; DllCanUnloadNow: {1}", hr, unload());
        return result;
    }
}
'@
[LoaderTest]::Run($Loader, [guid]'587D11AB-E524-453A-80FE-CCB3017CFA8B')
Get-Process -Id $PID -Module | Where-Object { $_.ModuleName -match '^(hostfxr|hostpolicy|coreclr|CalcpadWord.*)\.dll$' } |
    ForEach-Object { '  ' + $_.ModuleName + '  <-  ' + $_.FileName }
