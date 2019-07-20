#pragma once
#include <vector>
#include "srwlock.h"

class CMockSmartRenameRegExEvents:
    public ISmartRenameRegExEvents
{
public:
    // IUnknown
    IFACEMETHODIMP  QueryInterface(_In_ REFIID iid, _Outptr_ void** resultInterface);
    IFACEMETHODIMP_(ULONG) AddRef();
    IFACEMETHODIMP_(ULONG) Release();

    // ISmartRenameRegExEvents
    IFACEMETHODIMP OnSearchTermChanged(_In_ PCWSTR searchTerm);
    IFACEMETHODIMP OnReplaceTermChanged(_In_ PCWSTR replaceTerm);
    IFACEMETHODIMP OnFlagsChanged(_In_ DWORD flags);

    static HRESULT s_CreateInstance(_Outptr_ ISmartRenameRegExEvents** ppsrree);

    CMockSmartRenameRegExEvents() :
        m_refCount(1)
    {
    }

    ~CMockSmartRenameRegExEvents()
    {
        CoTaskMemFree(m_searchTerm);
        CoTaskMemFree(m_replaceTerm);
    }

    PWSTR m_searchTerm = nullptr;
    PWSTR m_replaceTerm = nullptr;
    DWORD m_flags = 0;
    long m_refCount;
};
