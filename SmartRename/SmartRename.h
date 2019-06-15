#pragma once
#include <SmartRenameInterfaces.h>

class CSmartRenameDlg :
    public ISmartRenameView,
    public ISmartRenameModelEvents
{
public:
    CSmartRenameDlg() :
        m_refCount(1)
    {
    }

    // IUnknown
    IFACEMETHODIMP QueryInterface(__in REFIID riid, __deref_out void** ppv)
    {
        static const QITAB qit[] =
        {
            QITABENT(CSmartRenameDlg, ISmartRenameView),
            QITABENT(CSmartRenameDlg, ISmartRenameModelEvents),
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

    // ISmartRenameView
    IFACEMETHODIMP Show();
    IFACEMETHODIMP Close();
    IFACEMETHODIMP Update();

    // ISmartRenameModelEvents
    IFACEMETHODIMP OnItemAdded(_In_ ISmartRenameItem* renameItem);
    IFACEMETHODIMP OnUpdate(_In_ ISmartRenameItem* renameItem);
    IFACEMETHODIMP OnError(_In_ ISmartRenameItem* renameItem);
    IFACEMETHODIMP OnRegExStarted();
    IFACEMETHODIMP OnRegExCanceled();
    IFACEMETHODIMP OnRegExCompleted();
    IFACEMETHODIMP OnRenameStarted();
    IFACEMETHODIMP OnRenameCompleted();

    static HRESULT s_CreateInstance(_In_ ISmartRenameModel* psrm, _In_opt_ IDataObject* pdo, _COM_Outptr_ ISmartRenameView** ppsrui);

private:
    ~CSmartRenameDlg()
    {
        DeleteObject(m_iconMain);
    }

    HRESULT _DoModal(__in_opt HWND hwnd);

    static INT_PTR CALLBACK s_DlgProc(HWND hdlg, UINT uMsg, WPARAM wParam, LPARAM lParam)
    {
        CSmartRenameDlg* pDlg = reinterpret_cast<CSmartRenameDlg*>(GetWindowLongPtr(hdlg, DWLP_USER));
        if (uMsg == WM_INITDIALOG)
        {
            pDlg = reinterpret_cast<CSmartRenameDlg*>(lParam);
            pDlg->m_hwnd = hdlg;
            SetWindowLongPtr(hdlg, DWLP_USER, reinterpret_cast<LONG_PTR>(pDlg));
        }
        return pDlg ? pDlg->_DlgProc(uMsg, wParam, lParam) : FALSE;
    }

    INT_PTR _DlgProc(UINT uMsg, WPARAM wParam, LPARAM lParam);

    HRESULT _Initialize(_In_ ISmartRenameModel* psrm, _In_opt_ IDataObject* pdo);

    void _OnInitDlg();
    void _OnRename();
    void _OnCloseDlg();
    void _OnDestroyDlg();
    void _OnClear();

    HRESULT _ReadSettings();
    HRESULT _WriteSettings();

    long m_refCount = 0;
    HWND m_hwnd = nullptr;
    HICON m_iconMain = nullptr;
    DWORD m_cookie = 0;
    CComPtr<ISmartRenameModel> m_spsrm;
    CComPtr<IDataObject> m_spdo;
};