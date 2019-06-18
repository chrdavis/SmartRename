// SmartRename.cpp : Defines the entry point for the application.
//


#include "stdafx.h"
#include "resource.h"
#include "SmartRenameUI.h"

extern HINSTANCE g_hInst;

HRESULT CSmartRenameUI::_DoModal(__in_opt HWND hwnd)
{
    HRESULT hr = S_OK;
    INT_PTR ret = DialogBoxParam(g_hInst, MAKEINTRESOURCE(IDD_MAIN), hwnd, s_DlgProc, (LPARAM)this);
    if (ret < 0)
    {
        hr = HRESULT_FROM_WIN32(GetLastError());
    }
    return hr;
}

HRESULT CSmartRenameUI::s_CreateInstance(_In_ ISmartRenameManager* psrm, _In_opt_ IDataObject* pdo, _COM_Outptr_ ISmartRenameUI** ppsrui)
{
    *ppsrui = nullptr;
    CSmartRenameUI *prui = new CSmartRenameUI();
    HRESULT hr = prui ? S_OK : E_OUTOFMEMORY;
    if (SUCCEEDED(hr))
    {
        // Pass the ISmartRenameManager to the SmartRenameView so it can subscribe to events
        hr = prui->_Initialize(psrm, pdo);
        if (SUCCEEDED(hr))
        {
            hr = prui->QueryInterface(IID_PPV_ARGS(ppsrui));
        }
        prui->Release();
    }
    return hr;
}

// ISmartRenameUI

IFACEMETHODIMP CSmartRenameUI::Show()
{
    return _DoModal(NULL);
}

IFACEMETHODIMP CSmartRenameUI::Close()
{
    return S_OK;
}

IFACEMETHODIMP CSmartRenameUI::Update()
{
    return S_OK;
}

IFACEMETHODIMP CSmartRenameUI::OnItemAdded(_In_ ISmartRenameItem*)
{
    return S_OK;
}

IFACEMETHODIMP CSmartRenameUI::OnUpdate(_In_ ISmartRenameItem*)
{
    return S_OK;
}

IFACEMETHODIMP CSmartRenameUI::OnError(_In_ ISmartRenameItem*)
{
    return S_OK;
}

IFACEMETHODIMP CSmartRenameUI::OnRegExStarted()
{
    return S_OK;
}

IFACEMETHODIMP CSmartRenameUI::OnRegExCanceled()
{
    return S_OK;
}

IFACEMETHODIMP CSmartRenameUI::OnRegExCompleted()
{
    return S_OK;
}

IFACEMETHODIMP CSmartRenameUI::OnRenameStarted()
{
    return S_OK;
}

IFACEMETHODIMP CSmartRenameUI::OnRenameCompleted()
{
    return S_OK;
}

HRESULT CSmartRenameUI::_Initialize(_In_ ISmartRenameManager* psrm, _In_opt_ IDataObject* pdo)
{
    // Cache the smart rename manager
    m_spsrm = psrm;

    // Cache the data object for enumeration later
    m_spdo = pdo;

    // Subscribe to smart rename manager events
    HRESULT hr = m_spsrm->Advise(this, &m_cookie);
    
    if (FAILED(hr))
    {
        //_Cleanup();
    }

    return hr;
}

// TODO: persist settings made in the UI

HRESULT CSmartRenameUI::_ReadSettings()
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

HRESULT CSmartRenameUI::_WriteSettings()
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

void CSmartRenameUI::_OnClear()
{
    // Clear input fields, reset options to defaults and reinit the contents of the list view
    // Clear the contents of the edit boxes
    //SetDlgItemText(m_hwnd, IDC_EDIT_SOURCE_MODULE, L"");
    //SetDlgItemText(m_hwnd, IDC_EDIT_SCREENSHOT_DESTINATION, L"");

    //_ToggleContent(TRUE);
}

void CSmartRenameUI::_OnInitDlg()
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

void CSmartRenameUI::_OnCloseDlg()
{
    // Persist the current settings
    _WriteSettings();
    EndDialog(m_hwnd, 1);
}

void CSmartRenameUI::_OnDestroyDlg()
{
    // TODO: teardown
}

void CSmartRenameUI::_OnRename()
{
    // TODO: 
}

/*void CSmartRenameUI::_OnRunRenamePreview()
{
    // TODO: We should have a interface and event interface that wraps the search/replace/regex inputs and settings/options
    // TODO: That way the manager could respond to changes set via the UI and it will run the rename preview on its own
}*/

INT_PTR CSmartRenameUI::_DlgProc(UINT uMsg, WPARAM wParam, LPARAM)
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
