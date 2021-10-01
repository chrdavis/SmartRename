#include "stdafx.h"
#include "SmartRenameExt.h"
#include <SmartRenameUI.h>
#include <SmartRenameItem.h>
#include <SmartRenameManager.h>
#include <Helpers.h>
#include <Settings.h>
#include "resource.h"

extern HINSTANCE g_hInst;

struct InvokeStruct
{
    HWND hwndParent;
    IStream* pstrm;
};

CSmartRenameMenu::CSmartRenameMenu()
{
    DllAddRef();
}

CSmartRenameMenu::~CSmartRenameMenu()
{
    m_spia = nullptr;
    DeleteObject(m_hbmpIcon);
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
    // Check if we have disabled ourselves
    if (!CSettings::GetEnabled())
        return E_FAIL;

    // Cache the data object to be used later
    GetShellItemArrayFromUnknown(pdtobj, &m_spia);
    return S_OK;
}

// IContextMenu
HRESULT CSmartRenameMenu::QueryContextMenu(HMENU hMenu, UINT index, UINT uIDFirst, UINT, UINT uFlags)
{
    // Check if we have disabled ourselves
    if (!CSettings::GetEnabled())
        return E_FAIL;

    // Check if we should only be on the extended context menu
    if (CSettings::GetExtendedContextMenuOnly() && (!(uFlags & CMF_EXTENDEDVERBS)))
        return E_FAIL;

    HRESULT hr = E_UNEXPECTED;
    if (m_spia && !(uFlags & (CMF_DEFAULTONLY | CMF_VERBSONLY | CMF_OPTIMIZEFORINVOKE)))
    {
        wchar_t menuName[64] = { 0 };
        LoadString(g_hInst, IDS_SMARTRENAME, menuName, ARRAYSIZE(menuName));

        MENUITEMINFO mii;
        mii.cbSize = sizeof(MENUITEMINFO);
        mii.fMask = MIIM_STRING | MIIM_FTYPE | MIIM_ID | MIIM_STATE;
        mii.wID = uIDFirst++;
        mii.fType = MFT_STRING;
        mii.dwTypeData = (PWSTR)menuName;
        mii.fState = MFS_ENABLED;

        if (CSettings::GetShowIconOnMenu())
        {
            HICON hIcon = (HICON)LoadImage(g_hInst, MAKEINTRESOURCE(IDI_RENAME), IMAGE_ICON, 16, 16, 0);
            if (hIcon)
            {
                mii.fMask |= MIIM_BITMAP;
                if (m_hbmpIcon == NULL)
                {
                    m_hbmpIcon = CreateBitmapFromIcon(hIcon);
                }
                mii.hbmpItem = m_hbmpIcon;
                DestroyIcon(hIcon);
            }
        }

        if (!InsertMenuItem(hMenu, index, TRUE, &mii))
        {
            hr = HRESULT_FROM_WIN32(GetLastError());
        }
        else
        {
            hr = MAKE_HRESULT(SEVERITY_SUCCESS, FACILITY_NULL, 1);
        }
    }

    return hr;
}

HRESULT CSmartRenameMenu::InvokeCommand(_In_ LPCMINVOKECOMMANDINFO pici)
{
    HRESULT hr = E_FAIL;

    if (CSettings::GetEnabled() && 
        (IS_INTRESOURCE(pici->lpVerb)) &&
        (LOWORD(pici->lpVerb) == 0))
    {
        hr = _InvokeInternal(pici->hwnd);
    }

    return hr;
}

// IExplorerCommand
IFACEMETHODIMP CSmartRenameMenu::GetTitle(_In_opt_ IShellItemArray* psia, _Outptr_result_nullonfailure_ PWSTR* name)
{
    *name = nullptr;
    wchar_t menuName[64] = { 0 };
    LoadString(g_hInst, IDS_SMARTRENAME, menuName, ARRAYSIZE(menuName));
    return SHStrDup(menuName, name);
}

IFACEMETHODIMP CSmartRenameMenu::GetIcon(_In_opt_ IShellItemArray*, _Outptr_result_nullonfailure_ PWSTR* icon)
{
    *icon = nullptr;
#ifdef _WIN64
    return SHStrDup(L"SmartRenameExt64.dll,-132", icon);
#else
    return SHStrDup(L"SmartRenameExt32.dll,-132", icon);
#endif
}

IFACEMETHODIMP CSmartRenameMenu::GetState(_In_opt_ IShellItemArray* psia, _In_ BOOL okToBeSlow, _Out_ EXPCMDSTATE* cmdState)
{
    // Check if we have disabled ourselves
    *cmdState = (CSettings::GetEnabled()) ? ECS_ENABLED : ECS_HIDDEN;
    return S_OK;
}

IFACEMETHODIMP CSmartRenameMenu::Invoke(_In_opt_ IShellItemArray* psia, _In_opt_ IBindCtx*)
{
    HRESULT hr = E_FAIL;
    if (CSettings::GetEnabled() && psia)
    {
        m_spia = psia;
        hr = _InvokeInternal(nullptr);
    }

    return hr;
}

// IObjectWithSites
IFACEMETHODIMP CSmartRenameMenu::SetSite(_In_ IUnknown* punk)
{
    m_spSite = punk;
    return S_OK;
}

IFACEMETHODIMP CSmartRenameMenu::GetSite(_In_ REFIID riid, _COM_Outptr_ void** ppunk)
{
    *ppunk = nullptr;
    HRESULT hr = E_FAIL;
    if (m_spSite)
    {
        hr = m_spSite->QueryInterface(riid, ppunk);
    }
    return hr;
}

HRESULT CSmartRenameMenu::_InvokeInternal(_In_opt_ HWND hwndParent)
{
    InvokeStruct* pInvokeData = new InvokeStruct;
    HRESULT hr = pInvokeData ? S_OK : E_OUTOFMEMORY;
    if (SUCCEEDED(hr))
    {
        pInvokeData->hwndParent = hwndParent;
        hr = CoMarshalInterThreadInterfaceInStream(__uuidof(m_spia), m_spia, &(pInvokeData->pstrm));
        if (SUCCEEDED(hr))
        {
            hr = SHCreateThread(s_SmartRenameUIThreadProc, pInvokeData, CTF_COINIT | CTF_PROCESS_REF, nullptr) ? S_OK : E_FAIL;
            if (FAILED(hr))
            {
                pInvokeData->pstrm->Release(); // if we failed to create the thread, then we must release the stream
            }
        }

        if (FAILED(hr))
        {
            delete pInvokeData;
        }
    }

    return hr;
}

DWORD WINAPI CSmartRenameMenu::s_SmartRenameUIThreadProc(_In_ void* pData)
{
    InvokeStruct* pInvokeData = static_cast<InvokeStruct*>(pData);
    CComPtr<IUnknown> spunk;
    if (SUCCEEDED(CoGetInterfaceAndReleaseStream(pInvokeData->pstrm, IID_PPV_ARGS(&spunk))))
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
                if (SUCCEEDED(spsrm->put_renameItemFactory(spsrif)))
                {
                    // Create the smart rename UI instance and pass the smart rename manager
                    CComPtr<ISmartRenameUI> spsrui;
                    if (SUCCEEDED(CSmartRenameUI::s_CreateInstance(spsrm, spunk, false, &spsrui)))
                    {
                        // Call blocks until we are done
                        spsrui->Show(pInvokeData->hwndParent);
                        spsrui->Close();
                    }
                }
            }

            // Need to call shutdown to break circular dependencies
            spsrm->Shutdown();
        }
    }

    delete pInvokeData;

    return 0;
}
