#include <windows.h>

// A simple exported function
__declspec(dllexport) void DummyFunction() {
    MessageBoxA(NULL, "DummyFunction called from test DLL!", "Test DLL", MB_OK | MB_ICONINFORMATION);
}

// DLL entry point
BOOL APIENTRY DllMain(HMODULE hModule, DWORD ul_reason_for_call, LPVOID lpReserved) {
    switch (ul_reason_for_call) {
        case DLL_PROCESS_ATTACH:
            // Code executed when the DLL is loaded
            MessageBoxA(NULL, "DLL Loaded Successfully!", "Test DLL", MB_OK | MB_ICONINFORMATION);
            break;
        case DLL_PROCESS_DETACH:
            // Code executed when the DLL is unloaded
            MessageBoxA(NULL, "DLL Unloaded Successfully!", "Test DLL", MB_OK | MB_ICONINFORMATION);
            break;
    }
    return TRUE;
}
