#include "stdafx.h"
#include "MockSmartRenameItem.h"

HRESULT CMockSmartRenameItem::CreateInstance(_In_opt_ PCWSTR path, _In_opt_ PCWSTR originalName, _In_ UINT depth, _In_ bool isFolder, _Outptr_ ISmartRenameItem** ppItem)
{
    *ppItem = nullptr;
    CMockSmartRenameItem* newItem = new CMockSmartRenameItem();
    HRESULT hr = newItem ? S_OK : E_OUTOFMEMORY;
    if (SUCCEEDED(hr))
    {
        newItem->Init(path, originalName, depth, isFolder);
        hr = newItem->QueryInterface(IID_PPV_ARGS(ppItem));
        newItem->Release();
    }

    return hr;
}

void CMockSmartRenameItem::Init(_In_opt_ PCWSTR path, _In_opt_ PCWSTR originalName, _In_ UINT depth, _In_ bool isFolder)
{
    if (path != nullptr)
    {
        SHStrDup(path, &m_path);
    }

    if (originalName != nullptr)
    {
        SHStrDup(originalName, &m_originalName);
    }

    m_depth = depth;
    m_isFolder = isFolder;
}

