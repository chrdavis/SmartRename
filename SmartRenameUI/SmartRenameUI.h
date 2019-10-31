#pragma once
#include <SmartRenameInterfaces.h>

class CSmartRenameListView
{

public:
    CSmartRenameListView() = default;
    ~CSmartRenameListView() = default;

    void Init(_In_ HWND hwndLV);
    void ToggleAll(_In_ ISmartRenameManager* psrm, _In_ bool selected);
    void ToggleItem(_In_ ISmartRenameManager* psrm, _In_ int item);
    void UpdateItemCheckState(_In_ ISmartRenameManager* psrm, _In_ int iItem);
    void RedrawItems(_In_ int first, _In_ int last);
    void SetItemCount(_In_ UINT itemCount);
    void OnKeyDown(_In_ ISmartRenameManager* psrm, _In_ LV_KEYDOWN* lvKeyDown);
    void OnClickList(_In_ ISmartRenameManager* psrm, NM_LISTVIEW* pnmListView);
    void GetDisplayInfo(_In_ ISmartRenameManager* psrm, _Inout_ LV_DISPINFO* plvdi);
    HWND GetHWND() { return m_hwndLV; }

private:
    void _UpdateColumns();
    void _UpdateColumnSizes();
    void _UpdateHeaderCheckState(_In_ bool check);

    HWND m_hwndLV = nullptr;
};

class CSmartRenameUI :
    public IDropTarget,
    public ISmartRenameUI,
    public ISmartRenameManagerEvents
{
public:
    CSmartRenameUI() :
        m_refCount(1)
    {
        (void)OleInitialize(nullptr);
    }

    // IUnknown
    IFACEMETHODIMP QueryInterface(__in REFIID riid, __deref_out void** ppv);
    IFACEMETHODIMP_(ULONG) AddRef();
    IFACEMETHODIMP_(ULONG) Release();

    // ISmartRenameUI
    IFACEMETHODIMP Show(_In_opt_ HWND hwndParent);
    IFACEMETHODIMP Close();
    IFACEMETHODIMP Update();
    IFACEMETHODIMP get_hwnd(_Out_ HWND* hwnd);
    IFACEMETHODIMP get_showUI(_Out_ bool* showUI);

    // ISmartRenameManagerEvents
    IFACEMETHODIMP OnItemAdded(_In_ ISmartRenameItem* renameItem);
    IFACEMETHODIMP OnUpdate(_In_ ISmartRenameItem* renameItem);
    IFACEMETHODIMP OnError(_In_ ISmartRenameItem* renameItem);
    IFACEMETHODIMP OnRegExStarted(_In_ DWORD threadId);
    IFACEMETHODIMP OnRegExCanceled(_In_ DWORD threadId);
    IFACEMETHODIMP OnRegExCompleted(_In_ DWORD threadId);
    IFACEMETHODIMP OnRenameStarted();
    IFACEMETHODIMP OnRenameCompleted();

    // IDropTarget
    IFACEMETHODIMP DragEnter(_In_ IDataObject* pdtobj, DWORD grfKeyState, POINTL pt, _Inout_ DWORD* pdwEffect);
    IFACEMETHODIMP DragOver(DWORD grfKeyState, POINTL pt, _Inout_ DWORD* pdwEffect);
    IFACEMETHODIMP DragLeave();
    IFACEMETHODIMP Drop(_In_ IDataObject* pdtobj, DWORD grfKeyState, POINTL pt, _Inout_ DWORD* pdwEffect);

    static HRESULT s_CreateInstance(_In_ ISmartRenameManager* psrm, _In_opt_ IDataObject* pdo, _In_ bool enableDragDrop, _Outptr_ ISmartRenameUI** ppsrui);

private:
    ~CSmartRenameUI()
    {
        DeleteObject(m_iconMain);
        OleUninitialize();
    }

    HRESULT _DoModal(__in_opt HWND hwnd);

    static INT_PTR CALLBACK s_DlgProc(HWND hdlg, UINT uMsg, WPARAM wParam, LPARAM lParam)
    {
        CSmartRenameUI* pDlg = reinterpret_cast<CSmartRenameUI*>(GetWindowLongPtr(hdlg, DWLP_USER));
        if (uMsg == WM_INITDIALOG)
        {
            pDlg = reinterpret_cast<CSmartRenameUI*>(lParam);
            pDlg->m_hwnd = hdlg;
            SetWindowLongPtr(hdlg, DWLP_USER, reinterpret_cast<LONG_PTR>(pDlg));
        }
        return pDlg ? pDlg->_DlgProc(uMsg, wParam, lParam) : FALSE;
    }

    INT_PTR _DlgProc(UINT uMsg, WPARAM wParam, LPARAM lParam);
    void _OnCommand(_In_ WPARAM wParam, _In_ LPARAM lParam);
    BOOL _OnNotify(_In_ WPARAM wParam, _In_ LPARAM lParam);

    HRESULT _Initialize(_In_ ISmartRenameManager* psrm, _In_opt_ IDataObject* pdo, _In_ bool enableDragDrop);
    void _Cleanup();

    void _OnInitDlg();
    void _OnRename();
    void _OnAbout();
    void _OnCloseDlg();
    void _OnDestroyDlg();
    void _OnClear();
    void _OnSearchReplaceChanged();

    HRESULT _ReadSettings();
    HRESULT _WriteSettings();

    DWORD _GetFlagsFromCheckboxes();
    void _SetCheckboxesFromFlags(_In_ DWORD flags);
    void _ValidateFlagCheckbox(_In_ DWORD checkBoxId);

    void _EnumerateItems(_In_ IDataObject* pdtobj);
    void _UpdateCounts();

    long m_refCount = 0;
    bool m_initialized = false;
    bool m_enableDragDrop = false;
    bool m_disableCountUpdate = false;
    HWND m_hwnd = nullptr;
    HWND m_hwndLV = nullptr;
    HICON m_iconMain = nullptr;
    DWORD m_cookie = 0;
    DWORD m_currentRegExId = 0;
    UINT m_selectedCount = 0;
    UINT m_renamingCount = 0;
    CComPtr<ISmartRenameManager> m_spsrm;
    CComPtr<IDataObject> m_spdo;
    CComPtr<IDropTargetHelper> m_spdth;
    CSmartRenameListView m_listview;
};