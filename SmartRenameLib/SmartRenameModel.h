#pragma once
#include "stdafx.h"
#include "SmartRenameInterfaces.h"
#include <vector>
#include "srwlock.h"

class CSmartRenameModel :
    public ISmartRenameModel,
    public ISmartRenameRegExEvents
{
public:
    // IUnknown
    IFACEMETHODIMP  QueryInterface(_In_ REFIID iid, _Outptr_ void** resultInterface);
    IFACEMETHODIMP_(ULONG) AddRef();
    IFACEMETHODIMP_(ULONG) Release();

    // ISmartRenameModel
    IFACEMETHODIMP Advise(_In_ ISmartRenameModelEvents* renameOpEvent, _Out_ DWORD *cookie);
    IFACEMETHODIMP UnAdvise(_In_ DWORD cookie);
    IFACEMETHODIMP Start();
    IFACEMETHODIMP Stop();
    IFACEMETHODIMP Reset();
    IFACEMETHODIMP Shutdown();
    IFACEMETHODIMP AddItem(_In_ ISmartRenameItem* pItem);
    IFACEMETHODIMP GetItem(_In_ UINT index, _COM_Outptr_ ISmartRenameItem** ppItem);
    IFACEMETHODIMP GetItemCount(_Out_ UINT* count);
    IFACEMETHODIMP get_smartRenameRegEx(_COM_Outptr_ ISmartRenameRegEx** ppRegEx);
    IFACEMETHODIMP put_smartRenameRegEx(_COM_Outptr_ ISmartRenameRegEx* pRegEx);
    IFACEMETHODIMP get_smartRenameItemFactory(_In_ ISmartRenameItemFactory** ppItemFactory);
    IFACEMETHODIMP put_smartRenameItemFactory(_In_ ISmartRenameItemFactory* pItemFactory);

    // ISmartRenameRegExEvents
    IFACEMETHODIMP OnSearchTermChanged(_In_ PCWSTR searchTerm);
    IFACEMETHODIMP OnReplaceTermChanged(_In_ PCWSTR replaceTerm);
    IFACEMETHODIMP OnFlagsChanged(_In_ DWORD flags);

    static HRESULT s_CreateInstance(_COM_Outptr_ ISmartRenameModel** pprm);

private:
    CSmartRenameModel();
    ~CSmartRenameModel();

    void _Cancel();

    void _OnItemAdded(_In_ ISmartRenameItem* renameItem);
    void _OnUpdate(_In_ ISmartRenameItem* renameItem);
    void _OnError(_In_ ISmartRenameItem* renameItem);
    void _OnRegExStarted();
    void _OnRegExCanceled();
    void _OnRegExCompleted();
    void _OnRenameStarted();
    void _OnRenameCompleted();

    // Thread proc for performing the regex rename of each item
    static DWORD WINAPI s_regexWorkerThread(_In_ void* pvoid);
    // Thread proc for performing the actual file operation that does the file rename
    static DWORD WINAPI s_fileOpWorkerThread(_In_ void* pvoid);

    HANDLE m_startRegExWorkerEvent = nullptr;
    HANDLE m_cancelRegExWorkerEvent = nullptr;

    HANDLE m_startFileOpWorkerEvent = nullptr;

    CSRWLock m_lockEvents;
    CSRWLock m_lockItems;

    DWORD m_cookie = 0;

    struct SMART_RENAME_MODEL_EVENT
    {
        ISmartRenameModelEvents* pEvents;
        DWORD cookie;
    };

    CComPtr<ISmartRenameItemFactory> m_spItemFactory;
    CComPtr<ISmartRenameRegEx> m_spRegEx;

    _Guarded_by_(m_lockEvents) std::vector<SMART_RENAME_MODEL_EVENT> m_smartRenameModelEvents;
    _Guarded_by_(m_lockItems) std::vector<ISmartRenameItem*> m_smartRenameItems;

    long m_refCount;
};