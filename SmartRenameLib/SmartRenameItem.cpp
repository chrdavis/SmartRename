#include "stdafx.h"
#include "SmartRenameItem.h"

IFACEMETHODIMP_(ULONG) CSmartRenameItem::AddRef()
{
    return InterlockedIncrement(&m_refCount);
}

IFACEMETHODIMP_(ULONG) CSmartRenameItem::Release()
{
    long refCount = InterlockedDecrement(&m_refCount);

    if (refCount == 0)
    {
        delete this;
    }
    return refCount;
}

IFACEMETHODIMP CSmartRenameItem::QueryInterface(_In_ REFIID riid, _Outptr_ void **ppv)
{
    static const QITAB qit[] = {
        QITABENT(CSmartRenameItem, ISmartRenameItem),
        { 0 }
    };
    return QISearch(this, qit, riid, ppv);
}

IFACEMETHODIMP CSmartRenameItem::GetOriginalName(_Outptr_ PWSTR * originalName)
{
    return E_NOTIMPL;
}

IFACEMETHODIMP CSmartRenameItem::SetNewName(_In_ PCWSTR newName)
{
    CoTaskMemFree(m_newName);
    m_newName = StrDup(newName);
    return m_newName ? S_OK : E_OUTOFMEMORY;
}

IFACEMETHODIMP CSmartRenameItem::GetNewName(_Outptr_ PWSTR * newName)
{
    HRESULT hr = m_newName ? S_OK : E_FAIL;
    if (SUCCEEDED(hr))
    {
        *newName = StrDup(m_newName);
        hr = (*newName) ? S_OK : E_OUTOFMEMORY;
    }
    return hr;
}

IFACEMETHODIMP CSmartRenameItem::GetIsFolder(_Out_ bool * isFolder)
{
    *isFolder = m_isFolder;
    return S_OK;
}

IFACEMETHODIMP CSmartRenameItem::GetIsDirty(_Out_ bool * isDirty)
{
    *isDirty = m_isDirty;
    return S_OK;
}

IFACEMETHODIMP CSmartRenameItem::GetId(_Out_ int * id)
{
    *id = m_id;
    return S_OK;
}

IFACEMETHODIMP CSmartRenameItem::Reset()
{
    CoTaskMemFree(m_newName);
    m_newName = nullptr;
    m_isDirty = false;
    return S_OK;
}

HRESULT CSmartRenameItem::CreateInstance(
    _In_ PCWSTR fullPath,
    _In_ PCWSTR oldName,
    _In_ bool isFolder,
    _In_ int id,
    _Outptr_ ISmartRenameItem ** renameItem)
{
    *renameItem = nullptr;

    CSmartRenameItem *newRenameItem = new CSmartRenameItem();
    HRESULT hr = newRenameItem ? S_OK : E_OUTOFMEMORY;
    if (SUCCEEDED(hr))
    {
        newRenameItem->_SetId(id);
        newRenameItem->_SetIsFolder(isFolder);
        hr = newRenameItem->_SetFullPath(fullPath);
        if (SUCCEEDED(hr))
        {
            hr = newRenameItem->_SetOldName(oldName);
        }

        if (SUCCEEDED(hr))
        {
            hr = newRenameItem->QueryInterface(IID_PPV_ARGS(renameItem));
        }

        newRenameItem->Release();
    }
    return hr;
}

CSmartRenameItem::CSmartRenameItem() :
    m_fullPath(nullptr),
    m_oldName(nullptr),
    m_newName(nullptr),
    m_error(S_OK),
    m_isFolder(false),
    m_isDirty(false),
    m_id(-1),
    m_refCount(1)
{
}

CSmartRenameItem::~CSmartRenameItem()
{
    CoTaskMemFree(m_fullPath);
    CoTaskMemFree(m_oldName);
    CoTaskMemFree(m_newName);
}

void CSmartRenameItem::_SetId(_In_ int id)
{
    m_id = id;
}

void CSmartRenameItem::_SetIsFolder(_In_ bool isFolder)
{
    m_isFolder = isFolder;
}

HRESULT CSmartRenameItem::_SetFullPath(_In_ PCWSTR fullPath)
{
    HRESULT hr = (fullPath && wcslen(fullPath) > 0) ? S_OK : E_INVALIDARG;
    if (SUCCEEDED(hr))
    {
        CoTaskMemFree(m_fullPath);
        m_fullPath = StrDup(fullPath);
        hr = m_fullPath ? S_OK : E_OUTOFMEMORY;
    }
    return hr;
}

HRESULT CSmartRenameItem::_SetOldName(_In_ PCWSTR oldName)
{
    HRESULT hr = (oldName && wcslen(oldName) > 0) ? S_OK : E_INVALIDARG;
    if (SUCCEEDED(hr))
    {
        CoTaskMemFree(m_oldName);
        m_oldName = StrDup(oldName);
        hr = m_oldName ? S_OK : E_OUTOFMEMORY;
    }
    return hr;
}