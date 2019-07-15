// SmartRenameTest.cpp : Defines the entry point for the application.
//

#include "stdafx.h"
#include "SmartRenameTest.h"
#include <SmartRenameInterfaces.h>
#include <SmartRenameItem.h>
#include <SmartRenameUI.h>
#include <SmartRenameManager.h>
#include <Shobjidl.h>

#pragma comment(linker,"/manifestdependency:\"type='win32' name='Microsoft.Windows.Common-Controls' version='6.0.0.0' processorArchitecture='*' publicKeyToken='6595b64144ccf1df' language='*'\"")

HINSTANCE g_hInst;

// {81ADB5B6-F9A4-4320-87B3-D9360F82EC50}
//DEFINE_GUID(CLSID_SmartRenameMenu, 0x81ADB5B6, 0xF9A4, 0x4320, 0x87, 0xB3, 0xD9, 0x36, 0x0F, 0x82, 0xEC, 0x50);

class __declspec(uuid("{81ADB5B6-F9A4-4320-87B3-D9360F82EC50}")) Foo;
static const CLSID CLSID_SmartRenameMenu = __uuidof(Foo);

DEFINE_GUID(BHID_DataObject, 0xb8c0bd9f, 0xed24, 0x455c, 0x83, 0xe6, 0xd5, 0x39, 0xc, 0x4f, 0xe8, 0xc4);

int APIENTRY wWinMain(
    _In_ HINSTANCE hInstance,
    _In_opt_ HINSTANCE hPrevInstance,
    _In_ PWSTR lpCmdLine,
    _In_ int nCmdShow)
{
    g_hInst = hInstance;
    HRESULT hr = CoInitializeEx(nullptr, COINIT_APARTMENTTHREADED | COINIT_DISABLE_OLE1DDE);
    if (SUCCEEDED(hr))
    {
        // Create the smart rename manager
        CComPtr<ISmartRenameManager> spsrm;
        if (SUCCEEDED(CSmartRenameManager::s_CreateInstance(&spsrm)))
        {
            // Create the factory for our items
            CComPtr<ISmartRenameItemFactory> spsrif;
            if (SUCCEEDED(CSmartRenameItem::s_CreateInstance(nullptr, IID_PPV_ARGS(&spsrif))))
            {
                // Pass the factory to the manager
                if (SUCCEEDED(spsrm->put_smartRenameItemFactory(spsrif)))
                {
                    // Create the smart rename UI instance and pass the manager
                    CComPtr<ISmartRenameUI> spsrui;
                    if (SUCCEEDED(CSmartRenameUI::s_CreateInstance(spsrm, nullptr, true, &spsrui)))
                    {
                        // Call blocks until we are done
                        spsrui->Show();
                        spsrui->Close();

                        // Need to call shutdown to break circular dependencies
                        spsrm->Shutdown();
                    }
                }
            }
        }
        CoUninitialize();
    }
    return 0;
}