#pragma once
#include <SmartRenameInterfaces.h>
#include <shldisp.h>
#include <atomic>

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
    void OnSize();
    HWND GetHWND() { return m_hwndLV; }

private:
    void _UpdateColumns();
    void _UpdateColumnSizes();
    void _UpdateHeaderCheckState(_In_ bool check);

    HWND m_hwndLV = nullptr;
};

class CSmartRenameProgressUI :
    public IUnknown
{
public:
    CSmartRenameProgressUI() :
        m_refCount(1)
    {
    }

    ~CSmartRenameProgressUI() = default;

    // IUnknown
    IFACEMETHODIMP QueryInterface(__in REFIID riid, __deref_out void** ppv);
    IFACEMETHODIMP_(ULONG) AddRef();
    IFACEMETHODIMP_(ULONG) Release();

    HRESULT Start();
    HRESULT Stop();
    bool IsCanceled() { return m_canceled; }

private:
    static LRESULT CALLBACK s_msgWndProc(_In_ HWND hwnd, _In_ UINT uMsg, _In_ WPARAM wParam, _In_ LPARAM lParam);
    LRESULT _WndProc(_In_ HWND hwnd, _In_ UINT msg, _In_ WPARAM wParam, _In_ LPARAM lParam);

    static DWORD WINAPI s_workerThread(_In_ void* pv);

    void _UpdateCancelState();
    void _Cleanup();

    long m_refCount = 0;
    bool m_canceled = false;
    std::atomic<bool> m_loadingThread{ false };
    HANDLE m_workerThreadHandle = nullptr;
    CComPtr<IProgressDialog> m_sppd;
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

    static HRESULT s_CreateInstance(_In_ ISmartRenameManager* psrm, _In_opt_ IUnknown* punk, _In_ bool enableDragDrop, _Outptr_ ISmartRenameUI** ppsrui);

private:
    ~CSmartRenameUI()
    {
        DeleteObject(m_iconMain);
        OleUninitialize();
    }

    HRESULT _DoModal(__in_opt HWND hwnd);
    HRESULT _DoModeless(__in_opt HWND hwnd);

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

    HRESULT _Initialize(_In_ ISmartRenameManager* psrm, _In_opt_ IUnknown* punk, _In_ bool enableDragDrop);
    HRESULT _InitAutoComplete();
    void _Cleanup();

    INT_PTR _DlgProc(UINT uMsg, WPARAM wParam, LPARAM lParam);
    void _OnCommand(_In_ WPARAM wParam, _In_ LPARAM lParam);
    BOOL _OnNotify(_In_ WPARAM wParam, _In_ LPARAM lParam);
    void _OnSize(_In_ WPARAM wParam);
    void _OnGetMinMaxInfo(_In_ LPARAM lParam);
    void _OnInitDlg();
    void _OnRename();
    void _OnAbout();
    void _OnCloseDlg();
    void _OnDestroyDlg();
    void _OnSearchReplaceChanged();
    void _MoveControl(_In_ DWORD id, _In_ DWORD repositionFlags, _In_ int xDelta, _In_ int yDelta);

    HRESULT _ReadSettings();
    HRESULT _WriteSettings();

    DWORD _GetFlagsFromCheckboxes();
    void _SetCheckboxesFromFlags(_In_ DWORD flags);
    void _ValidateFlagCheckbox(_In_ DWORD checkBoxId);

    HRESULT _EnumerateItems(_In_ IUnknown* punk);
    void _UpdateCounts();

    long m_refCount = 0;
    bool m_initialized = false;
    bool m_enableDragDrop = false;
    bool m_disableCountUpdate = false;
    bool m_modeless = true;
    HWND m_hwnd = nullptr;
    HWND m_hwndLV = nullptr;
    HICON m_iconMain = nullptr;
    DWORD m_cookie = 0;
    DWORD m_currentRegExId = 0;
    UINT m_selectedCount = 0;
    UINT m_renamingCount = 0;
    int m_initialWidth = 0;
    int m_initialHeight = 0;
    int m_lastWidth = 0;
    int m_lastHeight = 0;
    CSmartRenameProgressUI m_srpui;
    CComPtr<ISmartRenameManager> m_spsrm;
    CComPtr<ISmartRenameEnum> m_spsre;
    CComPtr<IUnknown> m_spunk;
    CComPtr<IDropTargetHelper> m_spdth;
    CComPtr<IAutoComplete2> m_spSearchAC;
    CComPtr<IUnknown> m_spSearchACL;
    CComPtr<IAutoComplete2> m_spReplaceAC;
    CComPtr<IUnknown> m_spReplaceACL;
    CSmartRenameListView m_listview;
};