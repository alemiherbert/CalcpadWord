/*
 * Calcpad for Word: the in-process COM server that Word loads (InprocServer32).
 *
 * It points .NET at the runtime bundled next to it (runtime\), so the computer needs no .NET
 * installed, then hands over to the add-in's own .NET COM host (CalcpadWord.AddIn.comhost.dll).
 *
 * Why it is needed: the .NET COM host looks only for a machine-wide .NET install (Program Files,
 * or DOTNET_ROOT). But if hostfxr is already loaded in the process, the COM host uses it and the
 * runtime it belongs to. So this loads runtime\host\fxr\<version>\hostfxr.dll first. If another
 * add-in has already started .NET in Word, that hostfxr is kept: a process holds only one runtime.
 *
 * Built by tools\Build-Loader.ps1 (Zig's C compiler). Uses kernel32 only: no C runtime.
 */
#define WIN32_LEAN_AND_MEAN
#include <windows.h>

#define COM_HOST L"CalcpadWord.AddIn.comhost.dll"
#define FXR_DIR L"runtime\\host\\fxr\\"
#define PATH_CAPACITY 2048

typedef HRESULT(WINAPI *GetClassObjectFn)(REFCLSID, REFIID, LPVOID *);

static HMODULE self;
static INIT_ONCE once = INIT_ONCE_STATIC_INIT;
static GetClassObjectFn comHostGetClassObject;
static HRESULT loadResult = E_UNEXPECTED;

/* Appends text at *at; FALSE if it doesn't fit. Keeps buffer terminated. */
static BOOL Append(WCHAR *buffer, SIZE_T *at, const WCHAR *text)
{
    while (*text) {
        if (*at + 1 >= PATH_CAPACITY)
            return FALSE;
        buffer[(*at)++] = *text++;
    }
    buffer[*at] = 0;
    return TRUE;
}

static BOOL Join(WCHAR *buffer, const WCHAR *a, const WCHAR *b, const WCHAR *c, const WCHAR *d)
{
    SIZE_T at = 0;
    buffer[0] = 0;
    return Append(buffer, &at, a) && Append(buffer, &at, b) && Append(buffer, &at, c) && Append(buffer, &at, d);
}

/* Loads the bundled hostfxr (unless one is loaded already), then the COM host. */
static BOOL CALLBACK Load(PINIT_ONCE initOnce, PVOID parameter, PVOID *context)
{
    static WCHAR folder[PATH_CAPACITY], file[PATH_CAPACITY];
    DWORD length = GetModuleFileNameW(self, folder, PATH_CAPACITY);
    (void)initOnce; (void)parameter; (void)context;

    if (length == 0 || length >= PATH_CAPACITY) {
        loadResult = HRESULT_FROM_WIN32(length == 0 ? GetLastError() : ERROR_FILENAME_EXCED_RANGE);
        return TRUE;
    }
    while (length > 0 && folder[length - 1] != L'\\')
        --length;
    folder[length] = 0; /* ends with a backslash */

    if (GetModuleHandleW(L"hostfxr.dll") == NULL && Join(file, folder, FXR_DIR, L"*", L"")) {
        WIN32_FIND_DATAW found;
        HANDLE search = FindFirstFileW(file, &found);
        if (search != INVALID_HANDLE_VALUE) {
            do {
                if ((found.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY) && found.cFileName[0] != L'.') {
                    if (Join(file, folder, FXR_DIR, found.cFileName, L"\\hostfxr.dll"))
                        LoadLibraryExW(file, NULL, LOAD_WITH_ALTERED_SEARCH_PATH);
                    break;
                }
            } while (FindNextFileW(search, &found));
            FindClose(search);
        }
        /* No bundled runtime (a developer build): the COM host looks for an installed one. */
    }

    if (!Join(file, folder, COM_HOST, L"", L"")) {
        loadResult = HRESULT_FROM_WIN32(ERROR_FILENAME_EXCED_RANGE);
        return TRUE;
    }
    HMODULE comHost = LoadLibraryExW(file, NULL, LOAD_WITH_ALTERED_SEARCH_PATH);
    if (comHost == NULL) {
        loadResult = HRESULT_FROM_WIN32(GetLastError());
        return TRUE;
    }
    comHostGetClassObject = (GetClassObjectFn)(void *)GetProcAddress(comHost, "DllGetClassObject");
    loadResult = comHostGetClassObject != NULL ? S_OK : HRESULT_FROM_WIN32(GetLastError());
    return TRUE;
}

__declspec(dllexport) HRESULT WINAPI DllGetClassObject(REFCLSID classId, REFIID interfaceId, LPVOID *result)
{
    if (result == NULL)
        return E_POINTER;
    *result = NULL;
    InitOnceExecuteOnce(&once, Load, NULL, NULL);
    return comHostGetClassObject != NULL ? comHostGetClassObject(classId, interfaceId, result) : loadResult;
}

/* .NET can't be unloaded from a process, so neither can this. */
__declspec(dllexport) HRESULT WINAPI DllCanUnloadNow(void)
{
    return S_FALSE;
}

BOOL WINAPI DllMain(HINSTANCE instance, DWORD reason, LPVOID reserved)
{
    (void)reserved;
    if (reason == DLL_PROCESS_ATTACH) {
        self = instance;
        DisableThreadLibraryCalls(instance);
    }
    return TRUE;
}
