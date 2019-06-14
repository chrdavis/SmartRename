#pragma once

#include "targetver.h"

#define WIN32_LEAN_AND_MEAN             // Exclude rarely-used stuff from Windows headers
// Windows Header Files:
#include <windows.h>
#include <unknwn.h>
#include <shlwapi.h>
#include <atlbase.h>
#include <Shobjidl.h>
#include <Shlobj.h>

void DllAddRef();
void DllRelease();

#define INITGUID
#include <guiddef.h>

// {81ADB5B6-F9A4-4320-87B3-D9360F82EC50}
DEFINE_GUID(CLSID_SmartRenameMenu, 0x81ADB5B6, 0xF9A4, 0x4320, 0x87, 0xB3, 0xD9, 0x36, 0x0F, 0x82, 0xEC, 0x50);

#pragma comment(linker,"/manifestdependency:\"type='win32' name='Microsoft.Windows.Common-Controls' version='6.0.0.0' processorArchitecture='*' publicKeyToken='6595b64144ccf1df' language='*'\"")

