#include "stdafx.h"
#include "MockSmartRenameManagerEvents.h"

// IUnknown
IFACEMETHODIMP CMockSmartRenameManagerEvents::QueryInterface(__in REFIID riid, __deref_out void** ppv)
{
    static const QITAB qit[] =
    {
        QITABENT(CMockSmartRenameManagerEvents, ISmartRenameManagerEvents),
        { 0 },
    };
    return QISearch(this, qit, riid, ppv);
}

IFACEMETHODIMP_(ULONG) CMockSmartRenameManagerEvents::AddRef()
{
    return InterlockedIncrement(&m_refCount);
}

IFACEMETHODIMP_(ULONG) CMockSmartRenameManagerEvents::Release()
{
    long refCount = InterlockedDecrement(&m_refCount);
    if (refCount == 0)
    {
        delete this;
    }
    return refCount;
}

// ISmartRenameManagerEvents
IFACEMETHODIMP CMockSmartRenameManagerEvents::OnItemAdded(_In_ ISmartRenameItem* pItem)
{
    return S_OK;
}

IFACEMETHODIMP CMockSmartRenameManagerEvents::OnUpdate(_In_ ISmartRenameItem* pItem)
{
    return S_OK;
}

IFACEMETHODIMP CMockSmartRenameManagerEvents::OnError(_In_ ISmartRenameItem*)
{
    return S_OK;
}

IFACEMETHODIMP CMockSmartRenameManagerEvents::OnRegExStarted()
{
    return S_OK;
}

IFACEMETHODIMP CMockSmartRenameManagerEvents::OnRegExCanceled()
{
    return S_OK;
}

IFACEMETHODIMP CMockSmartRenameManagerEvents::OnRegExCompleted()
{
    return S_OK;
}

IFACEMETHODIMP CMockSmartRenameManagerEvents::OnRenameStarted()
{
    return S_OK;
}

IFACEMETHODIMP CMockSmartRenameManagerEvents::OnRenameCompleted()
{
    return S_OK;
}

HRESULT CMockSmartRenameManagerEvents::s_CreateInstance(_In_ ISmartRenameManager* psrm, _Outptr_ ISmartRenameUI** ppsrui)
{
    *ppsrui = nullptr;
    CMockSmartRenameManagerEvents* events = new CMockSmartRenameManagerEvents();
    HRESULT hr = events != nullptr ? S_OK : E_OUTOFMEMORY;
    if (SUCCEEDED(hr))
    {
        hr = events->QueryInterface(IID_PPV_ARGS(ppsrui));
        events->Release();
    }

    return hr;
}

