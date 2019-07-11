#include "stdafx.h"
#include "resource.h"
#include "SmartRenameUI.h"
#include <commctrl.h>
#include <Shlobj.h>
#include <helpers.h>
#include <windowsx.h>

extern HINSTANCE g_hInst;


int g_rgnMatchModeResIDs[] =
{
    IDS_ENTIREITEMNAME,
    IDS_NAMEONLY,
    IDS_EXTENSIONONLY
};

enum
{
    MATCHMODE_FULLNAME = 0,
    MATCHMODE_NAMEONLY,
    MATCHMODE_EXTENIONONLY
};

struct FlagCheckboxMap
{
    DWORD flag;
    DWORD id;
};

FlagCheckboxMap g_flagCheckboxMap[] =
{
    { UseRegularExpressions, IDC_CHECK_USEREGEX },
    { ExcludeSubfolders, IDC_CHECK_EXCLUDESUBFOLDERS },
    { EnumerateItems, IDC_CHECK_ENUMITEMS },
    { ExcludeFiles, IDC_CHECK_EXCLUDEFILES },
    { CaseSensitive, IDC_CHECK_CASESENSITIVE },
    { MatchAllOccurences, IDC_CHECK_MATCHALLOCCRENCES },
    { ExcludeFolders, IDC_CHECK_EXCLUDEFOLDERS },
    { NameOnly, IDC_CHECK_NAMEONLY },
    { ExtensionOnly, IDC_CHECK_EXTENSIONONLY }
};

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

HRESULT CSmartRenameUI::s_CreateInstance(_In_ ISmartRenameManager* psrm, _In_opt_ IDataObject* pdo, _In_ bool enableDragDrop, _Outptr_ ISmartRenameUI** ppsrui)
{
    *ppsrui = nullptr;
    CSmartRenameUI *prui = new CSmartRenameUI();
    HRESULT hr = prui ? S_OK : E_OUTOFMEMORY;
    if (SUCCEEDED(hr))
    {
        // Pass the ISmartRenameManager to the ISmartRenameUI so it can subscribe to events
        hr = prui->_Initialize(psrm, pdo, enableDragDrop);
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
    DWORD flags = 0;
    if (m_spsrm)
    {
        m_spsrm->get_flags(&flags);
    }

    m_listview.InsertItem(pItem, flags);
    _UpdateCounts();
    return S_OK;
}

IFACEMETHODIMP CSmartRenameUI::OnUpdate(_In_ ISmartRenameItem* pItem)
{
    DWORD flags = 0;
    if (m_spsrm)
    {
        m_spsrm->get_flags(&flags);
    }

    m_listview.UpdateItem(pItem, flags);
    _UpdateCounts();
    return S_OK;
}

IFACEMETHODIMP CSmartRenameUI::OnError(_In_ ISmartRenameItem*)
{
    return S_OK;
}

IFACEMETHODIMP CSmartRenameUI::OnRegExStarted()
{
    // Disable list view
    _UpdateCounts();
    return S_OK;
}

IFACEMETHODIMP CSmartRenameUI::OnRegExCanceled()
{
    // Enable list view
    _UpdateCounts();
    return S_OK;
}

IFACEMETHODIMP CSmartRenameUI::OnRegExCompleted()
{
    // Enable list view
    _UpdateCounts();
    return S_OK;
}

IFACEMETHODIMP CSmartRenameUI::OnRenameStarted()
{
    // Disable controls
    EnableWindow(m_hwnd, FALSE);
    return S_OK;
}

IFACEMETHODIMP CSmartRenameUI::OnRenameCompleted()
{
    // Enable controls
    EnableWindow(m_hwnd, TRUE);
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
    if (m_spsrm)
    {
        EnumerateDataObject(pdtobj, m_spsrm);
    }

    return S_OK;
}

HRESULT CSmartRenameUI::_Initialize(_In_ ISmartRenameManager* psrm, _In_opt_ IDataObject* pdo, _In_ bool enableDragDrop)
{
    // Cache the smart rename manager
    m_spsrm = psrm;

    // Cache the data object for enumeration later
    m_spdo = pdo;

    m_enableDragDrop = enableDragDrop;

    HRESULT hr = CoCreateInstance(CLSID_DragDropHelper, NULL, CLSCTX_INPROC, IID_PPV_ARGS(&m_spdth));
    if (SUCCEEDED(hr))
    {
        // Subscribe to smart rename manager events
        hr = m_spsrm->Advise(this, &m_cookie);
    }

    if (FAILED(hr))
    {
        _Cleanup();
    }

    return hr;
}

void CSmartRenameUI::_Cleanup()
{
    if (m_spsrm && m_cookie != 0)
    {
        m_spsrm->UnAdvise(m_cookie);
        m_cookie = 0;
        m_spsrm = nullptr;
    }

    m_spdo = nullptr;
    m_spdth = nullptr;

    if (m_enableDragDrop)
    {
        RevokeDragDrop(m_hwnd);
    }
}

// TODO: persist settings made in the UI
HRESULT CSmartRenameUI::_ReadSettings()
{
    return S_OK;
}

HRESULT CSmartRenameUI::_WriteSettings()
{
    return S_OK;
}

void CSmartRenameUI::_OnClear()
{
}

void CSmartRenameUI::_OnCloseDlg()
{
    // Persist the current settings
    _WriteSettings();
    EndDialog(m_hwnd, 1);
}

void CSmartRenameUI::_OnDestroyDlg()
{
    _Cleanup();
}

void CSmartRenameUI::_OnRename()
{
    // TODO: Update UI to show we are renaming (disable controls)
    if (m_spsrm)
    {
        m_spsrm->Rename(m_hwnd);
    }
}

void CSmartRenameUI::_OnAbout()
{
    // Launch github page
    SHELLEXECUTEINFO info = {0};
    info.cbSize = sizeof(SHELLEXECUTEINFO);
    info.lpVerb = L"open";
    info.lpFile = L"http://www.github.com/chrdavis";
    info.nShow = SW_SHOWDEFAULT;

    ShellExecuteEx(&info);
}

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

    // Initialize checkboxes from flags
    if (m_spsrm)
    {
        DWORD flags = 0;
        m_spsrm->get_flags(&flags);
        _SetCheckboxesFromFlags(flags);
    }

    if (m_spdo && m_spsrm)
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
    if (m_enableDragDrop)
    {
        RegisterDragDrop(m_hwnd, this);
    }

    // Disable rename button by default.  It will be enabled in _UpdateCounts if
    // there are tiems to be renamed
    EnableWindow(GetDlgItem(m_hwnd, ID_RENAME), FALSE);

    // Update UI elements that depend on number of items selected or to be renamed
    _UpdateCounts();

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

    case ID_ABOUT:
        _OnAbout();
        break;

    case IDC_EDIT_REPLACEWITH:
    case IDC_EDIT_SEARCHFOR:
        if (GET_WM_COMMAND_CMD(wParam, lParam) == EN_CHANGE)
        {
            _OnSearchReplaceChanged();
        }
        break;

    case IDC_CHECK_CASESENSITIVE:
    case IDC_CHECK_ENUMITEMS:
    case IDC_CHECK_EXCLUDEFILES:
    case IDC_CHECK_EXCLUDEFOLDERS:
    case IDC_CHECK_EXCLUDESUBFOLDERS:
    case IDC_CHECK_MATCHALLOCCRENCES:
    case IDC_CHECK_USEREGEX:
    case IDC_CHECK_EXTENSIONONLY:
    case IDC_CHECK_NAMEONLY:
        if (BN_CLICKED == HIWORD(wParam))
        {
            _ValidateFlagCheckbox(LOWORD(wParam));
            _GetFlagsFromCheckboxes();
        }
        break;
    }
}

BOOL CSmartRenameUI::_OnNotify(_In_ WPARAM wParam, _In_ LPARAM lParam)
{
    bool ret = FALSE;
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
            LoadString(g_hInst, IDS_LISTVIEW_EMPTY, pnmMarkup->szMarkup, ARRAYSIZE(pnmMarkup->szMarkup));
            ret = TRUE;
            break;

        case LVN_BEGINLABELEDIT:
            ret = TRUE;
            break;

        case LVN_ITEMCHANGED:
            if ((m_initialized) &&
                (pnmlv->uChanged & LVIF_STATE) &&
                ((pnmlv->uNewState & LVIS_STATEIMAGEMASK) != (pnmlv->uOldState & LVIS_STATEIMAGEMASK)) &&
                (pnmlv->uOldState != 0))
            {
                if (m_spsrm)
                {
                    m_listview.UpdateItemCheckState(m_spsrm, pnmlv->iItem);
                    _UpdateCounts();
                }
            }
            break;

        case NM_DBLCLK:
            {
                BOOL checked = ListView_GetCheckState(m_hwndLV, pnmlv->iItem);
                ListView_SetCheckState(m_hwndLV, pnmlv->iItem, !checked);
                break;
            }
        }
    }

    return ret;
}

void CSmartRenameUI::_OnSearchReplaceChanged()
{
    // Pass updated search and replace terms to the ISmartRenameRegEx handler
    CComPtr<ISmartRenameRegEx> spRegEx;
    if (m_spsrm && SUCCEEDED(m_spsrm->get_smartRenameRegEx(&spRegEx)))
    {
        wchar_t buffer[MAX_PATH] = { 0 };
        GetDlgItemText(m_hwnd, IDC_EDIT_SEARCHFOR, buffer, ARRAYSIZE(buffer));
        spRegEx->put_searchTerm(buffer);

        buffer[0] = L'\0';
        GetDlgItemText(m_hwnd, IDC_EDIT_REPLACEWITH, buffer, ARRAYSIZE(buffer));
        spRegEx->put_replaceTerm(buffer);
    }
}

DWORD CSmartRenameUI::_GetFlagsFromCheckboxes()
{
    DWORD flags = 0;
    for (int i = 0; i < ARRAYSIZE(g_flagCheckboxMap); i++)
    {
        if (Button_GetCheck(GetDlgItem(m_hwnd, g_flagCheckboxMap[i].id)) == BST_CHECKED)
        {
            flags |= g_flagCheckboxMap[i].flag;
        }
    }

    // Ensure we update flags
    if (m_spsrm)
    {
        m_spsrm->put_flags(flags);
    }

    return flags;
}

void CSmartRenameUI::_SetCheckboxesFromFlags(_In_ DWORD flags)
{
    for (int i = 0; i < ARRAYSIZE(g_flagCheckboxMap); i++)
    {
        Button_SetCheck(GetDlgItem(m_hwnd, g_flagCheckboxMap[i].id), flags & g_flagCheckboxMap[i].flag);
    }
}

void CSmartRenameUI::_ValidateFlagCheckbox(_In_ DWORD checkBoxId)
{
    if (checkBoxId == IDC_CHECK_NAMEONLY)
    {
        if (Button_GetCheck(GetDlgItem(m_hwnd, IDC_CHECK_NAMEONLY)) == BST_CHECKED)
        {
            Button_SetCheck(GetDlgItem(m_hwnd, IDC_CHECK_EXTENSIONONLY), FALSE);
        }
    }
    else if (checkBoxId == IDC_CHECK_EXTENSIONONLY)
    {
        if (Button_GetCheck(GetDlgItem(m_hwnd, IDC_CHECK_EXTENSIONONLY)) == BST_CHECKED)
        {
            Button_SetCheck(GetDlgItem(m_hwnd, IDC_CHECK_NAMEONLY), FALSE);
        }
    }
}

void CSmartRenameUI::_UpdateCounts()
{
    UINT selectedCount = 0;
    UINT renamingCount = 0;
    if (m_spsrm)
    {
        m_spsrm->GetSelectedItemCount(&selectedCount);
        m_spsrm->GetRenameItemCount(&renamingCount);
    }

    if (m_selectedCount != selectedCount ||
        m_renamingCount != renamingCount)
    {
        m_selectedCount = selectedCount;
        m_renamingCount = renamingCount;

        // Update selected and rename count label
        wchar_t countsLabelFormat[100] = { 0 };
        LoadString(g_hInst, IDS_COUNTSLABELFMT, countsLabelFormat, ARRAYSIZE(countsLabelFormat));

        wchar_t countsLabel[100] = { 0 };
        StringCchPrintf(countsLabel, ARRAYSIZE(countsLabel), countsLabelFormat, selectedCount, renamingCount);
        SetDlgItemText(m_hwnd, IDC_STATUS_MESSAGE, countsLabel);

        // Update Rename button state
        EnableWindow(GetDlgItem(m_hwnd, ID_RENAME), (renamingCount > 0));
    }
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
        ListView_SetExtendedListViewStyle(m_hwndLV, LVS_EX_CHECKBOXES | LVS_EX_DOUBLEBUFFER);

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

    if (psrm && m_hwndLV && (iItem > -1))
    {
        CComPtr<ISmartRenameItem> spItem;
        hr = GetItemByIndex(psrm, iItem, &spItem);
        if (SUCCEEDED(hr))
        {
            bool checked = ListView_GetCheckState(m_hwndLV, iItem);
            spItem->put_selected(checked);

            UINT uSelected = (checked) ? LVIS_SELECTED : 0;
            ListView_SetItemState(m_hwndLV, iItem, uSelected, LVIS_SELECTED);

            // Update the rename column if necessary
            DWORD flags = 0;
            psrm->get_flags(&flags);
            UpdateItem(spItem, flags);
        }

        // Get the total number of list items and compare it to what is selected
        // We need to update the column checkbox if all items are selected or if
        // not all of the items are selected.
        bool checkHeader = (ListView_GetSelectedCount(m_hwndLV) == ListView_GetItemCount(m_hwndLV));
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

HRESULT CSmartRenameListView::InsertItem(_In_ ISmartRenameItem* pItem, _In_ DWORD flags)
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
            hr = _UpdateSubItems(pItem, flags, iNum);

            // Set the check state
            bool selected = false;
            pItem->get_selected(&selected);
            ListView_SetCheckState(m_hwndLV, iNum, selected);
        }
    }

    return hr;
}

HRESULT CSmartRenameListView::UpdateItem(_In_ ISmartRenameItem* pItem, _In_ DWORD flags)
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
            hr = _UpdateSubItems(pItem, flags, iIndex);
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
        DWORD flags = 0;
        psrm->get_flags(&flags);

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
                hr = InsertItem(spItem, flags);
            }
        }
    }

    return hr;
}

HRESULT CSmartRenameListView::_UpdateSubItems(_In_ ISmartRenameItem* pItem, _In_ DWORD flags, _In_ int iItem)
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
            
            bool shouldRename = false;
            if (SUCCEEDED(pItem->ShouldRenameItem(flags, &shouldRename)) && shouldRename)
            {
                if (SUCCEEDED(pItem->get_newName(&newName)))
                {
                    // We have a new name
                    lvitemCurr.pszText = newName;
                }
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


