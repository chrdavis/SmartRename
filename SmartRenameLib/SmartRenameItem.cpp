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

IFACEMETHODIMP CSmartRenameItem::QueryInterface(_In_ REFIID riid, _Outptr_ void** ppv)
{
    static const QITAB qit[] = {
        QITABENT(CSmartRenameItem, ISmartRenameItem),
        { 0 }
    };
    return QISearch(this, qit, riid, ppv);
}

IFACEMETHODIMP CSmartRenameItem::get_Path(_Outptr_ PWSTR* path)
{
    HRESULT hr = m_fullPath ? S_OK : E_FAIL;
    if (SUCCEEDED(hr))
    {
        *path = StrDup(m_fullPath);
        hr = (*path) ? S_OK : E_OUTOFMEMORY;
    }
    return hr;
}

IFACEMETHODIMP CSmartRenameItem::get_originalName(_Outptr_ PWSTR* originalName)
{
    HRESULT hr = m_originalName ? S_OK : E_FAIL;
    if (SUCCEEDED(hr))
    {
        *originalName = StrDup(m_originalName);
        hr = (*originalName) ? S_OK : E_OUTOFMEMORY;
    }
    return hr;
}

IFACEMETHODIMP CSmartRenameItem::put_newName(_In_ PCWSTR newName)
{
    CoTaskMemFree(m_newName);
    m_newName = StrDup(newName);
    return m_newName ? S_OK : E_OUTOFMEMORY;
}

IFACEMETHODIMP CSmartRenameItem::get_newName(_Outptr_ PWSTR* newName)
{
    HRESULT hr = m_newName ? S_OK : E_FAIL;
    if (SUCCEEDED(hr))
    {
        *newName = StrDup(m_newName);
        hr = (*newName) ? S_OK : E_OUTOFMEMORY;
    }
    return hr;
}

IFACEMETHODIMP CSmartRenameItem::get_isFolder(_Out_ bool* isFolder)
{
    *isFolder = m_isFolder;
    return S_OK;
}

IFACEMETHODIMP CSmartRenameItem::get_isDirty(_Out_ bool* isDirty)
{
    *isDirty = m_isDirty;
    return S_OK;
}

IFACEMETHODIMP CSmartRenameItem::get_shouldRename(_Out_ bool* shouldRename)
{
    *shouldRename = m_shouldRename;
    return S_OK;
}

IFACEMETHODIMP CSmartRenameItem::put_shouldRename(_In_ bool shouldRename)
{
    m_shouldRename = shouldRename;
    return S_OK;
}
IFACEMETHODIMP CSmartRenameItem::get_id(_Out_ int* id)
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
    _In_ bool isFolder,
    _In_ int id,
    _Outptr_ ISmartRenameItem** renameItem)
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
            hr = newRenameItem->_GetOriginalNameFromFullPath();
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
    m_refCount(1)
{
}

CSmartRenameItem::~CSmartRenameItem()
{
    CoTaskMemFree(m_fullPath);
    CoTaskMemFree(m_originalName);
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

HRESULT CSmartRenameItem::_GetOriginalNameFromFullPath()
{
    CoTaskMemFree(m_originalName);
    m_originalName = nullptr;
    PWSTR fileName = PathFindFileName(m_fullPath);
    if (fileName != nullptr)
    {
        m_originalName = StrDup(fileName);
    }
    return m_originalName ? S_OK : E_OUTOFMEMORY;;
}