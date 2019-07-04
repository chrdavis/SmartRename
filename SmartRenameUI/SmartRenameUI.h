#pragma once
#include <SmartRenameInterfaces.h>

class CSmartRenameListView
{

public:
    CSmartRenameListView();
    ~CSmartRenameListView();

    HRESULT Init(_In_ HWND hwndLV);
    HRESULT Clear();
    HRESULT UpdateItems(_In_ ISmartRenameManager* psrm);
    HRESULT InsertItem(_In_ ISmartRenameItem* pItem);
    HRESULT UpdateItem(_In_ ISmartRenameItem* pItem);
    HRESULT RemoveItem(_In_ ISmartRenameItem* pItem);
    HRESULT ToggleAll(_In_ bool selected);
    HRESULT UpdateItemCheckState(_In_ ISmartRenameManager* psrm, _In_ int iItem);
    HRESULT GetItemByIndex(_In_ ISmartRenameManager* psrm, _In_ int nIndex, _Out_ ISmartRenameItem** ppItem);
    HWND GetHWND() { return m_hwndLV; }

private:
    HRESULT _InsertItems(_In_ ISmartRenameManager* psrm);
    HRESULT _UpdateSubItems(_In_ ISmartRenameItem* pItem, _In_ int iItem);
    HRESULT _UpdateColumns();
    HRESULT _UpdateColumnSizes();
    HRESULT _UpdateHeaderCheckState(_In_ bool check);
    HRESULT _FindItemByParam(_In_ LPARAM lParam, _Out_ int* piIndex);

    bool _ShouldIncludeItem(_In_ ISmartRenameItem* pItem);

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
    IFACEMETHODIMP QueryInterface(__in REFIID riid, __deref_out void** ppv)
    {
        static const QITAB qit[] =
        {
            QITABENT(CSmartRenameUI, ISmartRenameUI),
            QITABENT(CSmartRenameUI, ISmartRenameManagerEvents),
            QITABENT(CSmartRenameUI, IDropTarget),
            { 0 },
        };
        return QISearch(this, qit, riid, ppv);
    }

    IFACEMETHODIMP_(ULONG) AddRef()
    {
        return InterlockedIncrement(&m_refCount);
    }

    IFACEMETHODIMP_(ULONG) Release()
    {
        long refCount = InterlockedDecrement(&m_refCount);
        if (refCount == 0)
        {
            delete this;
        }
        return refCount;
    }

    // ISmartRenameUI
    IFACEMETHODIMP Show();
    IFACEMETHODIMP Close();
    IFACEMETHODIMP Update();
    IFACEMETHODIMP get_hwnd(_Out_ HWND* hwnd);
    IFACEMETHODIMP get_showUI(_Out_ bool* showUI);

    // ISmartRenameManagerEvents
    IFACEMETHODIMP OnItemAdded(_In_ ISmartRenameItem* renameItem);
    IFACEMETHODIMP OnUpdate(_In_ ISmartRenameItem* renameItem);
    IFACEMETHODIMP OnError(_In_ ISmartRenameItem* renameItem);
    IFACEMETHODIMP OnRegExStarted();
    IFACEMETHODIMP OnRegExCanceled();
    IFACEMETHODIMP OnRegExCompleted();
    IFACEMETHODIMP OnRenameStarted();
    IFACEMETHODIMP OnRenameCompleted();

    // IDropTarget
    IFACEMETHODIMP DragEnter(_In_ IDataObject* pdtobj, DWORD grfKeyState, POINTL pt, _Inout_ DWORD* pdwEffect);
    IFACEMETHODIMP DragOver(DWORD grfKeyState, POINTL pt, _Inout_ DWORD* pdwEffect);
    IFACEMETHODIMP DragLeave();
    IFACEMETHODIMP Drop(_In_ IDataObject* pdtobj, DWORD grfKeyState, POINTL pt, _Inout_ DWORD* pdwEffect);

    static HRESULT s_CreateInstance(_In_ ISmartRenameManager* psrm, _In_opt_ IDataObject* pdo, _Outptr_ ISmartRenameUI** ppsrui);

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

    HRESULT _Initialize(_In_ ISmartRenameManager* psrm, _In_opt_ IDataObject* pdo);

    void _OnInitDlg();
    void _OnRename();
    void _OnCloseDlg();
    void _OnDestroyDlg();
    void _OnClear();
    void _OnSearchReplaceChanged();

    HRESULT _ReadSettings();
    HRESULT _WriteSettings();

    long m_refCount = 0;
    bool m_initialized = false;
    HWND m_hwnd = nullptr;
    HWND m_hwndLV = nullptr;
    HICON m_iconMain = nullptr;
    DWORD m_cookie = 0;
    CComPtr<ISmartRenameManager> m_spsrm;
    CComPtr<IDataObject> m_spdo;
    CComPtr<IDropTargetHelper> m_spdth;
    CSmartRenameListView m_listview;
};