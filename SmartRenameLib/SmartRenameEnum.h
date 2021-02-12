#pragma once
#include "stdafx.h"
#include "SmartRenameInterfaces.h"
#include <vector>
#include "srwlock.h"

class CSmartRenameEnum :
    public ISmartRenameEnum
{
public:
    // IUnknown
    IFACEMETHODIMP  QueryInterface(_In_ REFIID iid, _Outptr_ void** resultInterface);
    IFACEMETHODIMP_(ULONG) AddRef();
    IFACEMETHODIMP_(ULONG) Release();

    // ISmartRenameEnum
    IFACEMETHODIMP Advise(_In_ ISmartRenameEnumEvents* events, _Out_ DWORD* cookie);
    IFACEMETHODIMP UnAdvise(_In_ DWORD cookie);
    IFACEMETHODIMP Start();
    IFACEMETHODIMP Cancel();

public:
    static HRESULT s_CreateInstance(_In_ IDataObject* pdo, _In_ _In_ ISmartRenameManager* psrm, _In_ REFIID iid, _Outptr_ void** resultInterface);

protected:
    CSmartRenameEnum();
    virtual ~CSmartRenameEnum();

    HRESULT _Init(_In_ IDataObject* pdo, _In_ ISmartRenameManager* psrm);
    HRESULT _ParseEnumItems(_In_ IEnumShellItems* pesi, _In_ int depth = 0);

    void _OnStarted();
    void _OnCompleted();
    void _OnFoundItem(_In_ ISmartRenameItem* item);

    struct RENAME_ENUM_EVENT
    {
        ISmartRenameEnumEvents* pEvents;
        DWORD cookie;
    };

    DWORD m_cookie = 0;

    CSRWLock m_lockEvents;
    _Guarded_by_(m_lockEvents) std::vector<RENAME_ENUM_EVENT> m_renameEnumEvents;

    CComPtr<ISmartRenameManager> m_spsrm;
    CComPtr<IDataObject> m_spdo;
    bool m_canceled = false;
    long m_refCount = 0;
};