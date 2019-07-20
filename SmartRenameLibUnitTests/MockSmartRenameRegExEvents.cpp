#include "stdafx.h"
#include "MockSmartRenameRegExEvents.h"

IFACEMETHODIMP_(ULONG) CMockSmartRenameRegExEvents::AddRef()
{
    return InterlockedIncrement(&m_refCount);
}

IFACEMETHODIMP_(ULONG) CMockSmartRenameRegExEvents::Release()
{
    long refCount = InterlockedDecrement(&m_refCount);

    if (refCount == 0)
    {
        delete this;
    }
    return refCount;
}

IFACEMETHODIMP CMockSmartRenameRegExEvents::QueryInterface(_In_ REFIID riid, _Outptr_ void** ppv)
{
    static const QITAB qit[] = {
        QITABENT(CMockSmartRenameRegExEvents, ISmartRenameRegExEvents),
        { 0 }
    };
    return QISearch(this, qit, riid, ppv);
}

IFACEMETHODIMP CMockSmartRenameRegExEvents::OnSearchTermChanged(_In_ PCWSTR searchTerm)
{
    CoTaskMemFree(m_searchTerm);
    m_searchTerm = nullptr;
    if (searchTerm != nullptr)
    {
        SHStrDup(searchTerm, &m_searchTerm);
    }
    return S_OK;
}

IFACEMETHODIMP CMockSmartRenameRegExEvents::OnReplaceTermChanged(_In_ PCWSTR replaceTerm)
{
    CoTaskMemFree(m_replaceTerm);
    m_replaceTerm = nullptr;
    if (replaceTerm != nullptr)
    {
        SHStrDup(replaceTerm, &m_replaceTerm);
    }
    return S_OK;
}

IFACEMETHODIMP CMockSmartRenameRegExEvents::OnFlagsChanged(_In_ DWORD flags)
{
    m_flags = flags;
    return S_OK;
}

HRESULT CMockSmartRenameRegExEvents::s_CreateInstance(_Outptr_ ISmartRenameRegExEvents** ppsrree)
{
    *ppsrree = nullptr;
    CMockSmartRenameRegExEvents* psrree = new CMockSmartRenameRegExEvents();
    HRESULT hr = psrree ? S_OK : E_OUTOFMEMORY;
    if (SUCCEEDED(hr))
    {
        hr = psrree->QueryInterface(IID_PPV_ARGS(ppsrree));
        psrree->Release();
    }
    return hr;
}

