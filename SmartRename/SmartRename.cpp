// SmartRename.cpp : Defines the entry point for the application.
//


#include "stdafx.h"
#include "resource.h"
#include "SmartRename.h"

extern HINSTANCE g_hInst;

HRESULT CSmartRenameDlg::_DoModal(__in_opt HWND hwnd)
{
    HRESULT hr = S_OK;
    INT_PTR ret = DialogBoxParam(g_hInst, MAKEINTRESOURCE(IDD_MAIN), hwnd, s_DlgProc, (LPARAM)this);
    if (ret < 0)
    {
        hr = HRESULT_FROM_WIN32(GetLastError());
    }
    return hr;
}


HRESULT CSmartRenameDlg::s_CreateInstance(_In_ ISmartRenameModel* psrm, _In_opt_ IDataObject* pdo, _COM_Outptr_ ISmartRenameView** ppsrui)
{
    *ppsrui = nullptr;
    CSmartRenameDlg *prui = new CSmartRenameDlg();
    HRESULT hr = prui ? S_OK : E_OUTOFMEMORY;
    if (SUCCEEDED(hr))
    {
        // Pass the ISmartRenameModel to the SmartRenameView so it can subscribe to events
        hr = prui->_Initialize(psrm, pdo);
        if (SUCCEEDED(hr))
        {
            hr = prui->QueryInterface(IID_PPV_ARGS(ppsrui));
        }
        prui->Release();
    }
    return hr;
}


// ISmartRenameView

IFACEMETHODIMP CSmartRenameDlg::Show()
{
    return _DoModal(NULL);
}

IFACEMETHODIMP CSmartRenameDlg::Close()
{
    return S_OK;
}

IFACEMETHODIMP CSmartRenameDlg::Update()
{
    return S_OK;
}

IFACEMETHODIMP CSmartRenameDlg::OnItemAdded(_In_ ISmartRenameItem*)
{
    return S_OK;
}

IFACEMETHODIMP CSmartRenameDlg::OnUpdate(_In_ ISmartRenameItem*)
{
    return S_OK;
}

IFACEMETHODIMP CSmartRenameDlg::OnError(_In_ ISmartRenameItem*)
{
    return S_OK;
}

IFACEMETHODIMP CSmartRenameDlg::OnRegExStarted()
{
    return S_OK;
}

IFACEMETHODIMP CSmartRenameDlg::OnRegExCanceled()
{
    return S_OK;
}

IFACEMETHODIMP CSmartRenameDlg::OnRegExCompleted()
{
    return S_OK;
}

IFACEMETHODIMP CSmartRenameDlg::OnRenameStarted()
{
    return S_OK;
}

IFACEMETHODIMP CSmartRenameDlg::OnRenameCompleted()
{
    return S_OK;
}

HRESULT CSmartRenameDlg::_Initialize(_In_ ISmartRenameModel* psrm, _In_opt_ IDataObject* pdo)
{
    // Cache the smart rename model
    m_spsrm = psrm;

    // Cache the data object for enumeration later
    m_spdo = pdo;

    // Subscribe to smart rename model events
    HRESULT hr = m_spsrm->Advise(this, &m_cookie);
    
    if (FAILED(hr))
    {
        //_Cleanup();
    }

    return hr;
}

// TODO: persist settings made in the UI

HRESULT CSmartRenameDlg::_ReadSettings()
{
   /* HKEY hKey = NULL;
    HRESULT hr = HRESULT_FROM_WIN32(RegOpenKeyEx(HKEY_CURRENT_USER, REG_DD_SETTINGS_KEY, 0, KEY_READ, &hKey));
    if (SUCCEEDED(hr))
    {
        WCHAR szBuff[MAX_PATH] = { 0 };
        DWORD cb = (DWORD)(sizeof(WCHAR) * ARRAYSIZE(szBuff));
        DWORD dwType = REG_SZ;
        if (ERROR_SUCCESS == RegQueryValueEx(hKey, REG_DD_SETTING_SOURCE_MODULE, NULL, &dwType, (LPBYTE)szBuff, &cb))
        {
            SetDlgItemText(m_hwnd, IDC_EDIT_SOURCE_MODULE, szBuff);
        }

        szBuff[0] = L'\0';
        cb = (DWORD)(sizeof(WCHAR) * ARRAYSIZE(szBuff));
        dwType = REG_SZ;
        if (ERROR_SUCCESS == RegQueryValueEx(hKey, REG_DD_SETTING_DEST_PATH, NULL, &dwType, (LPBYTE)szBuff, &cb))
        {
            SetDlgItemText(m_hwnd, IDC_EDIT_SCREENSHOT_DESTINATION, szBuff);
        }

        cb = (DWORD)sizeof(DWORD);
        dwType = REG_DWORD;
        DWORD dwVal = 0;
        if (ERROR_SUCCESS == RegQueryValueEx(hKey, REG_DD_SETTING_SHOW_PROGRESS, NULL, &dwType, (LPBYTE)&dwVal, &cb))
        {
            CheckDlgButton(m_hwnd, IDC_SHOW_PROGRESS, (dwVal == 1) ? BST_CHECKED : BST_UNCHECKED);
        }

        RegCloseKey(hKey);
    }
    return hr;*/

    return S_OK;
}

HRESULT CSmartRenameDlg::_WriteSettings()
{
    /*HKEY hKey = NULL;
    HRESULT hr = HRESULT_FROM_WIN32(RegCreateKeyEx(HKEY_CURRENT_USER, REG_DD_SETTINGS_KEY, NULL, NULL, NULL,
        KEY_WRITE, NULL, &hKey, NULL));
    if (SUCCEEDED(hr))
    {
        WCHAR szBuff[MAX_PATH] = { 0 };
        GetDlgItemText(m_hwnd, IDC_EDIT_SOURCE_MODULE, szBuff, ARRAYSIZE(szBuff));
        RegSetValueEx(hKey, REG_DD_SETTING_SOURCE_MODULE, 0, REG_SZ, (LPBYTE)szBuff, (DWORD)(sizeof(WCHAR) * wcslen(szBuff)));

        GetDlgItemText(m_hwnd, IDC_EDIT_SCREENSHOT_DESTINATION, szBuff, ARRAYSIZE(szBuff));
        RegSetValueEx(hKey, REG_DD_SETTING_DEST_PATH, 0, REG_SZ, (LPBYTE)szBuff, (DWORD)(sizeof(WCHAR) * wcslen(szBuff)));

        DWORD dwVal = (IsDlgButtonChecked(m_hwnd, IDC_SHOW_PROGRESS) == BST_CHECKED) ? 1 : 0;
        RegSetValueEx(hKey, REG_DD_SETTING_SHOW_PROGRESS, 0, REG_DWORD, (LPBYTE)&dwVal, (DWORD)(sizeof(DWORD)));

        RegCloseKey(hKey);
    }

    return hr;*/

    return S_OK;
}

void CSmartRenameDlg::_OnClear()
{
    // Clear input fields, reset options to defaults and reinit the contents of the list view
    // Clear the contents of the edit boxes
    //SetDlgItemText(m_hwnd, IDC_EDIT_SOURCE_MODULE, L"");
    //SetDlgItemText(m_hwnd, IDC_EDIT_SCREENSHOT_DESTINATION, L"");

    //_ToggleContent(TRUE);
}

void CSmartRenameDlg::_OnInitDlg()
{
    // Initialize from stored settings
    _ReadSettings();

    // Load the main icon
    /*if (SUCCEEDED(CGraphicsHelper::LoadIconFromModule(g_hInst, IDI_RENAME, 32, 32, &m_iconMain)))
    {
        // Convert the icon to a bitmap
        if (SUCCEEDED(CGraphicsHelper::HBITMAPFromHICON(m_iconMain, &m_BitmapMain)))
        {
            // Set the bitmap as the contents of the control on our dialog
            HGDIOBJ hgdiOld = (HGDIOBJ)SendDlgItemMessage(m_hwnd, IDC_MAIN_IMAGE, STM_SETIMAGE, IMAGE_BITMAP, (LPARAM)_hBitmapMain);
            if (hgdiOld)
            {
                DeleteObject(hgdiOld);  // if there was an old one clean it up
            }
        }
        // Update the icon associated with our main app window
        SendMessage(m_hwnd, WM_SETICON, (WPARAM)ICON_SMALL, (LPARAM)m_iconMain);
        SendMessage(m_hwnd, WM_SETICON, (WPARAM)ICON_BIG, (LPARAM)m_iconMain);
    }*/
}

void CSmartRenameDlg::_OnCloseDlg()
{
    // Persist the current settings
    _WriteSettings();
    EndDialog(m_hwnd, 1);
}

void CSmartRenameDlg::_OnDestroyDlg()
{
    // TODO: teardown
}

void CSmartRenameDlg::_OnRename()
{
    // TODO: 
}

/*void CSmartRenameDlg::_OnRunRenamePreview()
{
    // TODO: We should have a interface and event interface that wraps the search/replace/regex inputs and settings/options
    // TODO: That way the model could respond to changes set via the UI and it will run the rename preview on its own
}*/

INT_PTR CSmartRenameDlg::_DlgProc(UINT uMsg, WPARAM wParam, LPARAM)
{
    INT_PTR bRet = TRUE;   // default for all handled cases in switch below

    switch (uMsg)
    {
    case WM_INITDIALOG:
        _OnInitDlg();
        break;

    case WM_COMMAND:
    {
        switch (LOWORD(wParam))
        {
        case IDOK:
        case IDCANCEL:
            _OnCloseDlg();
            break;

        case ID_RENAME:
            _OnRename();
            break;
        }
    }
    break;

    case WM_CLOSE:
        _OnCloseDlg();
        break;

    case WM_DESTROY:
        _OnDestroyDlg();
        break;

    default:
        bRet = FALSE;
    }
    return bRet;
}
