#include "stdafx.h"
#include "resource.h"
#include "SmartRenameUI.h"
#include <commctrl.h>
#include <Shlobj.h>
#include <helpers.h>

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

HRESULT CSmartRenameUI::s_CreateInstance(_In_ ISmartRenameManager* psrm, _In_opt_ IDataObject* pdo, _Outptr_ ISmartRenameUI** ppsrui)
{
    *ppsrui = nullptr;
    CSmartRenameUI *prui = new CSmartRenameUI();
    HRESULT hr = prui ? S_OK : E_OUTOFMEMORY;
    if (SUCCEEDED(hr))
    {
        // Pass the ISmartRenameManager to the ISmartRenameUI so it can subscribe to events
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

IFACEMETHODIMP CSmartRenameUI::get_hwnd(_Out_ HWND* hwnd)
{
    *hwnd = m_hwnd;
    return S_OK;
}

IFACEMETHODIMP CSmartRenameUI::get_showUI(_Out_ bool* showUI)
{
    // Let callers know that it is OK to show UI (ex: progress dialog, error dialog and conflict dialog UI)
    *showUI = true;
    return S_OK;
}

IFACEMETHODIMP CSmartRenameUI::OnItemAdded(_In_ ISmartRenameItem* pItem)
{
    return m_listview.InsertItem(pItem);
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


// IDropTarget
IFACEMETHODIMP CSmartRenameUI::DragEnter(_In_ IDataObject* pdtobj, DWORD /* grfKeyState */, POINTL pt, _Inout_ DWORD* pdwEffect)
{
    if (m_spdth)
    {
        POINT ptT = { pt.x, pt.y };
        m_spdth->DragEnter(m_hwnd, pdtobj, &ptT, *pdwEffect);
    }

    return S_OK;
}

IFACEMETHODIMP CSmartRenameUI::DragOver(DWORD /* grfKeyState */, POINTL pt, _Inout_ DWORD* pdwEffect)
{
    if (m_spdth)
    {
        POINT ptT = { pt.x, pt.y };
        m_spdth->DragOver(&ptT, *pdwEffect);
    }

    return S_OK;
}

IFACEMETHODIMP CSmartRenameUI::DragLeave()
{
    if (m_spdth)
    {
        m_spdth->DragLeave();
    }

    return S_OK;
}

IFACEMETHODIMP CSmartRenameUI::Drop(_In_ IDataObject* pdtobj, DWORD, POINTL pt, _Inout_ DWORD* pdwEffect)
{
    if (m_spdth)
    {
        POINT ptT = { pt.x, pt.y };
        m_spdth->Drop(pdtobj, &ptT, *pdwEffect);
    }

    _OnClear();

    EnableWindow(GetDlgItem(m_hwnd, ID_RENAME), TRUE);
    EnableWindow(m_hwndLV, TRUE);

    // Enumerate the data object and popuplate the manager
    EnumerateDataObject(pdtobj, m_spsrm);

    return S_OK;
}

HRESULT CSmartRenameUI::_Initialize(_In_ ISmartRenameManager* psrm, _In_opt_ IDataObject* pdo)
{
    // Cache the smart rename manager
    m_spsrm = psrm;

    // Cache the data object for enumeration later
    m_spdo = pdo;

    HRESULT hr = CoCreateInstance(CLSID_DragDropHelper, NULL, CLSCTX_INPROC, IID_PPV_ARGS(&m_spdth));
    if (SUCCEEDED(hr))
    {
        // Subscribe to smart rename manager events
        hr = m_spsrm->Advise(this, &m_cookie);
    }

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
    // TODO: Update UI to show we are renaming (disable controls)
    m_spsrm->Rename(m_hwnd);
}

/*void CSmartRenameUI::_OnRunRenamePreview()
{
    // TODO: We should have a interface and event interface that wraps the search/replace/regex inputs and settings/options
    // TODO: That way the manager could respond to changes set via the UI and it will run the rename preview on its own
}*/

INT_PTR CSmartRenameUI::_DlgProc(UINT uMsg, WPARAM wParam, LPARAM lParam)
{
    INT_PTR bRet = TRUE;   // default for all handled cases in switch below

    switch (uMsg)
    {
    case WM_INITDIALOG:
        _OnInitDlg();
        break;

    case WM_COMMAND:
        _OnCommand(wParam, lParam);
        break;

    case WM_NOTIFY:
        bRet = _OnNotify(wParam, lParam);
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

void CSmartRenameUI::_OnInitDlg()
{
    // Initialize from stored settings
    _ReadSettings();

    m_hwndLV = GetDlgItem(m_hwnd, IDC_LIST_PREVIEW);

    m_listview.Init(m_hwndLV);

    if (m_spdo)
    {
        EnumerateDataObject(m_spdo, m_spsrm);
    }

    // TODO: Add dialog icon, image and description to top of dialog?

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

    // TODO: put this behind a setting?
    RegisterDragDrop(m_hwnd, this);

    // Disable until items added
    EnableWindow(m_hwndLV, FALSE);

    m_initialized = true;
}

void CSmartRenameUI::_OnCommand(_In_ WPARAM wParam, _In_ LPARAM lParam)
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

bool CSmartRenameUI::_OnNotify(_In_ WPARAM wParam, _In_ LPARAM lParam)
{
    bool ret = false;
    LPNMHDR          pnmdr = (LPNMHDR)lParam;
    LPNMLISTVIEW     pnmlv = (LPNMLISTVIEW)pnmdr;
    NMLVEMPTYMARKUP* pnmMarkup = NULL;

    if (pnmdr)
    {
        BOOL checked = FALSE;
        switch (pnmdr->code)
        {
        case HDN_ITEMSTATEICONCLICK:
            ListView_SetCheckState(m_hwndLV, -1, (!(((LPNMHEADER)lParam)->pitem->fmt & HDF_CHECKED)));
            break;

        case LVN_GETEMPTYMARKUP:
            pnmMarkup = (NMLVEMPTYMARKUP*)lParam;
            pnmMarkup->dwFlags = EMF_CENTERED;
            //LoadString(g_hInst, IDS_LISTVIEW_EMPTY, pnmMarkup->szMarkup, ARRAYSIZE(pnmMarkup->szMarkup));
            ret = true;
            break;

        case LVN_BEGINLABELEDIT:
            ret = true;
            break;

        case LVN_ITEMCHANGED:
            if ((m_initialized) &&
                (pnmlv->uChanged & LVIF_STATE) &&
                ((pnmlv->uNewState & LVIS_STATEIMAGEMASK) != (pnmlv->uOldState & LVIS_STATEIMAGEMASK)) &&
                (pnmlv->uOldState != 0))
            {
                // TODO: will the check state be set on the ISmartRenameItem in the List View class?
                m_listview.UpdateItemCheckState(m_spsrm, pnmlv->iItem);
                //PostThreadMessage(pDlg->_dwPreviewThreadId, REM_UPDATEITEM, (WPARAM)pnmlv->iItem, 0);
            }
            break;

        case NM_DBLCLK:
            checked = ListView_GetCheckState(m_hwndLV, pnmlv->iItem);
            ListView_SetCheckState(m_hwndLV, pnmlv->iItem, !checked);
            break;
        }
    }

    return ret;
}

CSmartRenameListView::CSmartRenameListView()
{
}

CSmartRenameListView::~CSmartRenameListView()
{
    Clear();
}

HRESULT CSmartRenameListView::Init(_In_ HWND hwndLV)
{
    HRESULT hr = E_INVALIDARG;

    if (hwndLV)
    {
        m_hwndLV = hwndLV;

        EnableWindow(m_hwndLV, TRUE);

        // Set the standard styles
        DWORD dwLVStyle = (DWORD)GetWindowLongPtr(m_hwndLV, GWL_STYLE);
        dwLVStyle |= LVS_ALIGNLEFT | LVS_REPORT | LVS_SHAREIMAGELISTS | LVS_SINGLESEL;
        SetWindowLongPtr(m_hwndLV, GWL_STYLE, dwLVStyle);

        // Set the extended view styles
        ListView_SetExtendedListViewStyle(m_hwndLV, LVS_EX_CHECKBOXES);

        // Get the system image lists.  Our list view is setup to not destroy
        // these since the image list belongs to the entire explorer process
        HIMAGELIST himlLarge;
        HIMAGELIST himlSmall;
        if (Shell_GetImageLists(&himlLarge, &himlSmall))
        {
            ListView_SetImageList(m_hwndLV, himlSmall, LVSIL_SMALL);
            ListView_SetImageList(m_hwndLV, himlLarge, LVSIL_NORMAL);
        }

        hr = _UpdateColumns();
    }

    return hr;
}

HRESULT CSmartRenameListView::ToggleAll(_In_ bool selected)
{
    HRESULT hr = E_FAIL;

    if (m_hwndLV)
    {
        int iCount = ListView_GetItemCount(m_hwndLV);
        for (int i = 0; i < iCount; i++)
        {
            // Set the item check state
            ListView_SetCheckState(m_hwndLV, i, selected);
        }
        hr = S_OK;
    }

    return hr;
}

HRESULT CSmartRenameListView::UpdateItemCheckState(_In_ ISmartRenameManager* psrm, _In_ int iItem)
{
    HRESULT hr = E_INVALIDARG;

    if ((m_hwndLV) && (iItem > -1))
    {
        CComPtr<ISmartRenameItem> spItem;
        hr = GetItemByIndex(psrm, iItem, &spItem);
        if (SUCCEEDED(hr))
        {
            bool checked = ListView_GetCheckState(m_hwndLV, iItem);
            spItem->put_shouldRename(checked);

            UINT uSelected = (checked) ? LVIS_SELECTED : 0;
            ListView_SetItemState(m_hwndLV, iItem, uSelected, LVIS_SELECTED);
        }

        // Get the total number of list items and compare it to what is selected
        // We need to update the column checkbox if all items are selected or if
        // not all of the items are selected.
        bool checkHeader = (ListView_GetSelectedCount(m_hwndLV) == (UINT)ListView_GetItemCount(m_hwndLV));
        _UpdateHeaderCheckState(checkHeader);
    }

    return hr;
}

// TODO: LPARAM should store id of item instead of pointer to interface
HRESULT CSmartRenameListView::GetItemByIndex(_In_ ISmartRenameManager* psrm, _In_ int nIndex, _Out_ ISmartRenameItem** ppItem)
{
    *ppItem = nullptr;
    HRESULT hr = E_FAIL;

    if (nIndex >= 0)
    {
        LVITEM lvItem = { 0 };
        lvItem.iItem = nIndex;
        lvItem.mask = LVIF_PARAM;
        if (ListView_GetItem(m_hwndLV, &lvItem))
        {
            int id = static_cast<int>(lvItem.lParam);
            hr = psrm->GetItemById(id, ppItem);
        }
    }

    return hr;
}

HRESULT CSmartRenameListView::Clear()
{
    ListView_DeleteAllItems(m_hwndLV);

    return S_OK;
}

HRESULT CSmartRenameListView::UpdateItems(_In_ ISmartRenameManager* psrm)
{
    HRESULT hr = E_INVALIDARG;
    if ((m_hwndLV) && (psrm))
    {
        // Clear the contents of the list view
        hr = Clear();
        if (SUCCEEDED(hr))
        {
            hr = _InsertItems(psrm);
        }

        _UpdateColumnSizes();
    }

    return hr;
}

HRESULT CSmartRenameListView::InsertItem(_In_ ISmartRenameItem* pItem)
{
    HRESULT hr = E_INVALIDARG;
    if ((m_hwndLV) && (pItem))
    {
        hr = E_FAIL;
        int iCount = ListView_GetItemCount(m_hwndLV);
 
        LVITEM lvitemNew = { 0 };

        lvitemNew.mask = LVIF_IMAGE | LVIF_PARAM | LVIF_INDENT;
        int id = -1;
        pItem->get_id(&id);
        lvitemNew.lParam = (LPARAM)id;
        lvitemNew.iItem = iCount + 1;
        pItem->get_iconIndex(&lvitemNew.iImage);
        UINT depth = 0;
        pItem->get_depth(&depth);
        lvitemNew.iIndent = static_cast<int>(depth);

        int iNum = ListView_InsertItem(m_hwndLV, &lvitemNew);
        if (iNum != -1)
        {
            // Update the sub items of the list item
            hr = _UpdateSubItems(pItem, iNum);

            // Set the check state
            bool shouldRename = false;
            pItem->get_shouldRename(&shouldRename);
            ListView_SetCheckState(m_hwndLV, iNum, shouldRename);
        }
    }

    return hr;
}

HRESULT CSmartRenameListView::UpdateItem(_In_ ISmartRenameItem* pItem)
{
    HRESULT hr = E_INVALIDARG;

    if ((m_hwndLV) && (pItem))
    {
        // Find the item in the list view
        int iIndex = -1;
        int id = -1;
        pItem->get_id(&id);
        hr = _FindItemByParam((LPARAM)id, &iIndex);
        if ((SUCCEEDED(hr)) && (iIndex != -1))
        {
            // Insert the client settings in the correct columns (if the column is visible)
            hr = _UpdateSubItems(pItem, iIndex);
        }
    }

    return hr;
}

HRESULT CSmartRenameListView::RemoveItem(_In_ ISmartRenameItem* pItem)
{
    HRESULT hr = E_INVALIDARG;

    if ((m_hwndLV) && (pItem))
    {
        // Find the item in the list view
        int iIndex = -1;
        int id = -1;
        pItem->get_id(&id);
        hr = _FindItemByParam((LPARAM)id, &iIndex);
        if ((SUCCEEDED(hr)) && (iIndex != -1))
        {
            // Remove the item
            ListView_DeleteItem(m_hwndLV, iIndex);
        }
    }

    return hr;
}

HRESULT CSmartRenameListView::_InsertItems(_In_ ISmartRenameManager* psrm)
{
    HRESULT hr = E_INVALIDARG;

    if (m_hwndLV)
    {
        // Loop through our list of items to rename
        UINT itemCount = 0;
        hr = psrm->GetItemCount(&itemCount);
        for (UINT i = 0; ((i < itemCount) && (SUCCEEDED(hr))); itemCount++)
        {
            CComPtr<ISmartRenameItem> spItem;
            hr = psrm->GetItemByIndex(i, &spItem);
            if (SUCCEEDED(hr))
            {
                // Are we including this type of item?
                if (_ShouldIncludeItem(spItem))
                {
                    hr = InsertItem(spItem);
                }

                if (SUCCEEDED(hr))
                {
                    // Are we displaying child items?
                    //if (dwFlags & OPT_INCLUDESUBFOLDERS)
                    {
                        // Does this item have children?
                        //HDPA dpaChildItems = pCurrItem->GetChildDPA();
                        //if (DPA_GetPtrCount(dpaItemList) > 0)
                        //{
                        //    hr = _InsertItems(dpaChildItems, dwFlags);
                        //}
                    }
                }
            }
        }
    }

    return hr;
}

bool CSmartRenameListView::_ShouldIncludeItem(_In_ ISmartRenameItem* pItem)
{
    bool include = true;

    if (pItem)
    {
        // TODO: Update by checking against flags
        // TODO: Should the LV class be notified of flags changing as well?
        // TODO: We should rethink the flags changed notification

        // Is this a folder and if so are we allowing folders?
        /*if ((!(dwOptions & OPT_EXCLUDEFOLDERS)) &&
            (pRenameItem->IsFolder()))
        {
            fInclude = TRUE;
        }

        // Is this a file and if so are we allowing files?
        if ((!(dwOptions & OPT_EXCLUDEFILES)) &&
            (!pRenameItem->IsFolder()))
        {
            fInclude = TRUE;
        }*/
    }

    return include;
}

HRESULT CSmartRenameListView::_UpdateSubItems(_In_ ISmartRenameItem* pItem, _In_ int iItem)
{
    HRESULT hr = E_INVALIDARG;

    if ((m_hwndLV) && (pItem))
    {
        LVITEM lvitemCurr = { 0 };

        int iInsertIndex = 0;

        lvitemCurr.iItem = iItem;
        lvitemCurr.mask = LVIF_TEXT;

        PWSTR originalName = nullptr;
        hr = pItem->get_originalName(&originalName);
        if (SUCCEEDED(hr))
        {
            // Add the original name
            lvitemCurr.iSubItem = iInsertIndex;
            lvitemCurr.pszText = originalName;

            ListView_SetItem(m_hwndLV, &lvitemCurr);

            iInsertIndex++;

            // Get the new name if we have one
            lvitemCurr.pszText = L"";
            PWSTR newName = nullptr;
            if (SUCCEEDED(pItem->get_newName(&newName)))
            {
                // We have a new name
                lvitemCurr.pszText = newName;
            }

            // Add the new name - if any
            lvitemCurr.iSubItem = iInsertIndex;
            ListView_SetItem(m_hwndLV, &lvitemCurr);

            CoTaskMemFree(originalName);
            CoTaskMemFree(newName);
        }

        hr = S_OK;
    }

    return hr;
}

HRESULT CSmartRenameListView::_UpdateColumns()
{
    HRESULT hr = E_INVALIDARG;

    if (m_hwndLV)
    {
        // And the list view columns
        int iInsertPoint = 0;

        LV_COLUMN lvc = { 0 };
        lvc.mask = LVCF_FMT | LVCF_ORDER | LVCF_WIDTH | LVCF_TEXT;
        lvc.fmt = LVCFMT_LEFT;
        lvc.iOrder = iInsertPoint;

        WCHAR buffer[64] = { 0 };
        LoadString(g_hInst, IDS_ORIGINAL, buffer, ARRAYSIZE(buffer));
        lvc.pszText = buffer;

        ListView_InsertColumn(m_hwndLV, iInsertPoint, &lvc);

        iInsertPoint++;

        lvc.iOrder = iInsertPoint;
        LoadString(g_hInst, IDS_RENAMED, buffer, ARRAYSIZE(buffer));
        lvc.pszText = buffer;

        ListView_InsertColumn(m_hwndLV, iInsertPoint, &lvc);

        // Get a handle to the header of the columns
        HWND hwndHeader = ListView_GetHeader(m_hwndLV);

        if (hwndHeader)
        {
            // Update the header style to allow checkboxes
            DWORD dwHeaderStyle = (DWORD)GetWindowLongPtr(hwndHeader, GWL_STYLE);
            dwHeaderStyle |= HDS_CHECKBOXES;
            SetWindowLongPtr(hwndHeader, GWL_STYLE, dwHeaderStyle);

            _UpdateHeaderCheckState(TRUE);
        }

        _UpdateColumnSizes();

        hr = S_OK;
    }

    return hr;
}

HRESULT CSmartRenameListView::_UpdateColumnSizes()
{
    if (m_hwndLV)
    {
        RECT rc;
        GetClientRect(m_hwndLV, &rc);

        ListView_SetColumnWidth(m_hwndLV, 0, (rc.right - rc.left) / 2);
        ListView_SetColumnWidth(m_hwndLV, 1, (rc.right - rc.left) / 2);
    }

    return S_OK;
}

HRESULT CSmartRenameListView::_UpdateHeaderCheckState(_In_ bool check)
{
    // Get a handle to the header of the columns
    HWND hwndHeader = ListView_GetHeader(m_hwndLV);
    if (hwndHeader)
    {
        WCHAR szBuff[MAX_PATH] = { 0 };

        // Retrieve the existing header first so we
        // don't trash the text already there
        HDITEM hdi = { 0 };
        hdi.mask = HDI_FORMAT | HDI_TEXT;
        hdi.pszText = szBuff;
        hdi.cchTextMax = ARRAYSIZE(szBuff);

        Header_GetItem(hwndHeader, 0, &hdi);

        // Set the first column to contain a checkbox
        hdi.fmt |= HDF_CHECKBOX;
        hdi.fmt |= (check) ? HDF_CHECKED : 0;

        Header_SetItem(hwndHeader, 0, &hdi);
    }

    return S_OK;
}

HRESULT CSmartRenameListView::_FindItemByParam(__in LPARAM lParam, __out int* piIndex)
{
    HRESULT hr = E_FAIL;
    *piIndex = -1;
    if ((m_hwndLV) && (lParam))
    {
        LVFINDINFO lvfi = { 0 };
        lvfi.flags = LVFI_PARAM;
        lvfi.lParam = lParam;
        *piIndex = ListView_FindItem(m_hwndLV, -1, &lvfi);
        hr = S_OK;
    }

    return hr;
}


