#include "stdafx.h"
#include "SmartRenameExt.h"
#include <SmartRename.h>
#include <SmartRenameModel.h>
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

HRESULT CSmartRenameMenu::s_CreateInstance(_In_opt_ IUnknown*, _In_ REFIID riid, _Out_ void **ppv)
{
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

bool CSmartRenameMenu::_IsFolder()
{
    bool isFolder = false;

    CComPtr<IShellItemArray> spsia;
    HRESULT hr = SHCreateShellItemArrayFromDataObject(m_spdo, IID_PPV_ARGS(&spsia));
    if (SUCCEEDED(hr))
    {
        CComPtr<IShellItem> spsi;
        hr = spsia->GetItemAt(0, &spsi);
        if (SUCCEEDED(hr))
        {
            PWSTR filePath = nullptr;
            hr = spsi->GetDisplayName(SIGDN_FILESYSPATH, &filePath);
            if (SUCCEEDED(hr))
            {
                if (GetFileAttributes(filePath) & FILE_ATTRIBUTE_DIRECTORY)
                {
                    isFolder = true;
                }
                else
                {
                    PCWSTR fileExt = PathFindExtension(filePath);
                    if (fileExt)
                    {
                        isFolder = false;
                    }
                }

                CoTaskMemFree(filePath);
            }
            else
            {
                isFolder = true;
            }
        }
    }

    return isFolder;
}

DWORD WINAPI CSmartRenameMenu::s_SmartRenameUIThreadProc(_In_ void* pData)
{
    IStream* pstrm = (IStream*)pData;
    CComPtr<IDataObject> spdo;
    if (SUCCEEDED(CoGetInterfaceAndReleaseStream(pstrm, IID_PPV_ARGS(&spdo))))
    {
        // Create the smart rename model
        CComPtr<ISmartRenameModel> spsrm;
        if (SUCCEEDED(CSmartRenameModel::s_CreateInstance(&spsrm)))
        {
            // Create the smart rename UI instance and pass the smart rename model
            CComPtr<ISmartRenameView> spsrui;
            if (SUCCEEDED(CSmartRenameDlg::s_CreateInstance(spsrm, spdo, &spsrui)))
            {
                // Call blocks until we are done
                spsrui->Show();
                spsrui->Close();
            }
        }
    }

    return 0; // ignored
}
