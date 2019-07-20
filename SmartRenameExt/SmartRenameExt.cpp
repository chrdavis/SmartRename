#include "stdafx.h"
#include "SmartRenameExt.h"
#include <SmartRenameUI.h>
#include <SmartRenameItem.h>
#include <SmartRenameManager.h>
#include "resource.h"

extern HINSTANCE g_hInst;

HWND g_hwndParent = 0;

CSmartRenameMenu::CSmartRenameMenu()
{
    DllAddRef();
}

CSmartRenameMenu::~CSmartRenameMenu()
{
    m_spdo = nullptr;
    DllRelease();
}

HRESULT CSmartRenameMenu::s_CreateInstance(_In_opt_ IUnknown*, _In_ REFIID riid, _Outptr_ void **ppv)
{
    *ppv = nullptr;
    HRESULT hr = E_OUTOFMEMORY;
    CSmartRenameMenu *pprm = new CSmartRenameMenu();
    if (pprm)
    {
        hr = pprm->QueryInterface(riid, ppv);
        pprm->Release();
    }
    return hr;
}

// IShellExtInit
HRESULT CSmartRenameMenu::Initialize(_In_opt_ PCIDLIST_ABSOLUTE, _In_ IDataObject *pdtobj, HKEY)
{
    // Cache the data object to be used later
    m_spdo = pdtobj;
    return S_OK;
}

// IContextMenu
HRESULT CSmartRenameMenu::QueryContextMenu(HMENU hMenu, UINT index, UINT uIDFirst, UINT, UINT uFlags)
{
    HRESULT hr = E_UNEXPECTED;
    if (m_spdo)
    {
        if ((uFlags & ~CMF_OPTIMIZEFORINVOKE) && (uFlags & ~(CMF_DEFAULTONLY | CMF_VERBSONLY)))
        {
            wchar_t menuName[64] = { 0 };
            LoadString(g_hInst, IDS_SMARTRENAME, menuName, ARRAYSIZE(menuName));
            InsertMenu(hMenu, index, MF_STRING | MF_BYPOSITION, uIDFirst++, menuName);
            hr = MAKE_HRESULT(SEVERITY_SUCCESS, FACILITY_NULL, 1);
        }
    }

    return hr;
}

HRESULT CSmartRenameMenu::InvokeCommand(_In_ LPCMINVOKECOMMANDINFO pici)
{
    HRESULT hr = E_FAIL;

    if ((IS_INTRESOURCE(pici->lpVerb)) &&
        (LOWORD(pici->lpVerb) == 0))
    {
        IStream* pstrm = nullptr;
        if (SUCCEEDED(CoMarshalInterThreadInterfaceInStream(__uuidof(m_spdo), m_spdo, &pstrm)))
        {
            if (!SHCreateThread(s_SmartRenameUIThreadProc, pstrm, CTF_COINIT | CTF_PROCESS_REF, nullptr))
            {
                pstrm->Release(); // if we failed to create the thread, then we must release the stream
            }
        }
    }

    return hr;
}

DWORD WINAPI CSmartRenameMenu::s_SmartRenameUIThreadProc(_In_ void* pData)
{
    IStream* pstrm = static_cast<IStream*>(pData);
    CComPtr<IDataObject> spdo;
    if (SUCCEEDED(CoGetInterfaceAndReleaseStream(pstrm, IID_PPV_ARGS(&spdo))))
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
                    // Create the smart rename UI instance and pass the smart rename manager
                    CComPtr<ISmartRenameUI> spsrui;
                    if (SUCCEEDED(CSmartRenameUI::s_CreateInstance(spsrm, spdo, false, &spsrui)))
                    {
                        // Call blocks until we are done
                        spsrui->Show();
                        spsrui->Close();
                    }
                }
            }

            // Need to call shutdown to break circular dependencies
            spsrm->Shutdown();
        }
    }

    return 0;
}
