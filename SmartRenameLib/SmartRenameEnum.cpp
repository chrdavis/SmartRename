#include "stdafx.h"
#include "SmartRenameEnum.h"
#include <ShlGuid.h>

IFACEMETHODIMP_(ULONG) CSmartRenameEnum::AddRef()
{
    return InterlockedIncrement(&m_refCount);
}

IFACEMETHODIMP_(ULONG) CSmartRenameEnum::Release()
{
    long refCount = InterlockedDecrement(&m_refCount);

    if (refCount == 0)
    {
        delete this;
    }
    return refCount;
}

IFACEMETHODIMP CSmartRenameEnum::QueryInterface(_In_ REFIID riid, _Outptr_ void** ppv)
{
    static const QITAB qit[] = {
        QITABENT(CSmartRenameEnum, ISmartRenameEnum),
        { 0 }
    };
    return QISearch(this, qit, riid, ppv);
}

IFACEMETHODIMP CSmartRenameEnum::Advise(_In_ ISmartRenameEnumEvents* events, _Out_ DWORD* cookie)
{
    CSRWExclusiveAutoLock lock(&m_lockEvents);
    m_cookie++;
    RENAME_ENUM_EVENT srme;
    srme.cookie = m_cookie;
    srme.pEvents = events;
    events->AddRef();
    m_renameEnumEvents.push_back(srme);

    *cookie = m_cookie;

    return S_OK;
}

IFACEMETHODIMP CSmartRenameEnum::UnAdvise(_In_ DWORD cookie)
{
    HRESULT hr = E_FAIL;
    CSRWExclusiveAutoLock lock(&m_lockEvents);

    for (std::vector<RENAME_ENUM_EVENT>::iterator it = m_renameEnumEvents.begin(); it != m_renameEnumEvents.end(); ++it)
    {
        if (it->cookie == cookie)
        {
            hr = S_OK;
            it->cookie = 0;
            if (it->pEvents)
            {
                it->pEvents->Release();
                it->pEvents = nullptr;
            }
            break;
        }
    }

    return hr;
}


IFACEMETHODIMP CSmartRenameEnum::Start()
{
    _OnStarted();
    m_canceled = false;
    CComPtr<IShellItemArray> spsia;
    HRESULT hr = SHCreateShellItemArrayFromDataObject(m_spdo, IID_PPV_ARGS(&spsia));
    if (SUCCEEDED(hr))
    {
        CComPtr<IEnumShellItems> spesi;
        hr = spsia->EnumItems(&spesi);
        if (SUCCEEDED(hr))
        {
            hr = _ParseEnumItems(spesi);
        }
    }

    _OnCompleted();
    return hr;
}

IFACEMETHODIMP CSmartRenameEnum::Cancel()
{
    m_canceled = true;
    return S_OK;
}

HRESULT CSmartRenameEnum::s_CreateInstance(_In_ IDataObject* pdo, _In_ _In_ ISmartRenameManager* psrm, _In_ REFIID iid, _Outptr_ void** resultInterface)
{
    *resultInterface = nullptr;

    CSmartRenameEnum* newRenameEnum = new CSmartRenameEnum();
    HRESULT hr = newRenameEnum ? S_OK : E_OUTOFMEMORY;
    if (SUCCEEDED(hr))
    {
        hr = newRenameEnum->_Init(pdo, psrm);
        if (SUCCEEDED(hr))
        {
            hr = newRenameEnum->QueryInterface(iid, resultInterface);
        }

        newRenameEnum->Release();
    }
    return hr;
}

CSmartRenameEnum::CSmartRenameEnum() :
    m_refCount(1)
{
}

CSmartRenameEnum::~CSmartRenameEnum()
{
}


void CSmartRenameEnum::_OnStarted()
{
    CSRWSharedAutoLock lock(&m_lockEvents);

    for (std::vector<RENAME_ENUM_EVENT>::iterator it = m_renameEnumEvents.begin(); it != m_renameEnumEvents.end(); ++it)
    {
        if (it->pEvents)
        {
            it->pEvents->OnStarted();
        }
    }
}

void CSmartRenameEnum::_OnCompleted()
{
    CSRWSharedAutoLock lock(&m_lockEvents);

    for (std::vector<RENAME_ENUM_EVENT>::iterator it = m_renameEnumEvents.begin(); it != m_renameEnumEvents.end(); ++it)
    {
        if (it->pEvents)
        {
            it->pEvents->OnCompleted(m_canceled);
        }
    }
}

void CSmartRenameEnum::_OnFoundItem(_In_ ISmartRenameItem* item)
{
    CSRWSharedAutoLock lock(&m_lockEvents);

    for (std::vector<RENAME_ENUM_EVENT>::iterator it = m_renameEnumEvents.begin(); it != m_renameEnumEvents.end(); ++it)
    {
        if (it->pEvents)
        {
            it->pEvents->OnFoundItem(item);
        }
    }
}

HRESULT CSmartRenameEnum::_Init(_In_ IDataObject* pdo, _In_ ISmartRenameManager* psrm)
{
    m_spdo = pdo;
    m_spsrm = psrm;
    return S_OK;
}

HRESULT CSmartRenameEnum::_ParseEnumItems(_In_ IEnumShellItems* pesi, _In_ int depth)
{
    HRESULT hr = E_INVALIDARG;

    // We shouldn't get this deep since we only enum the contents of
    // regular folders but adding just in case
    if ((pesi) && (depth < (MAX_PATH / 2)))
    {
        hr = S_OK;

        ULONG celtFetched;
        CComPtr<IShellItem> spsi;
        while ((S_OK == pesi->Next(1, &spsi, &celtFetched)) && (SUCCEEDED(hr)))
        {
            if (m_canceled)
            {
                return E_ABORT;
            }

            CComPtr<ISmartRenameItemFactory> spsrif;
            hr = m_spsrm->get_renameItemFactory(&spsrif);
            if (SUCCEEDED(hr))
            {
                CComPtr<ISmartRenameItem> spNewItem;
                // Failure may be valid if we come across a shell item that does
                // not support a file system path.  In that case we simply ignore
                // the item.
                if (SUCCEEDED(spsrif->Create(spsi, &spNewItem)))
                {
                    spNewItem->put_depth(depth);
                    _OnFoundItem(spNewItem);
                    hr = m_spsrm->AddItem(spNewItem);
                    if (SUCCEEDED(hr))
                    {
                        bool isFolder = false;
                        if (SUCCEEDED(spNewItem->get_isFolder(&isFolder)) && isFolder)
                        {
                            // Bind to the IShellItem for the IEnumShellItems interface
                            CComPtr<IEnumShellItems> spesiNext;
                            hr = spsi->BindToHandler(nullptr, BHID_EnumItems, IID_PPV_ARGS(&spesiNext));
                            if (SUCCEEDED(hr))
                            {
                                // Parse the folder contents recursively
                                hr = _ParseEnumItems(spesiNext, depth + 1);
                            }
                        }
                    }
                }
            }

            spsi = nullptr;
        }
    }

    return hr;
}

