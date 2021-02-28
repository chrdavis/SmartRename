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

IFACEMETHODIMP CSmartRenameEnum::Start()
{
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

    return hr;
}

IFACEMETHODIMP CSmartRenameEnum::Cancel()
{
    m_canceled = true;
    return S_OK;
}

HRESULT CSmartRenameEnum::s_CreateInstance(_In_ IDataObject* pdo, _In_ ISmartRenameManager* psrm, _In_ REFIID iid, _Outptr_ void** resultInterface)
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

