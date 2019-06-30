#include "stdafx.h"
#include "Helpers.h"
#include <ShlGuid.h>

HRESULT GetIconIndexFromPath(_In_ PCWSTR path, _Out_ int* index)
{
    HRESULT hr = E_FAIL;

    SHFILEINFO shFileInfo = { 0 };

    if (!PathIsRelative(path))
    {
        DWORD attrib = GetFileAttributes(path);
        HIMAGELIST himl = (HIMAGELIST)SHGetFileInfo(path, attrib, &shFileInfo, sizeof(shFileInfo), (SHGFI_SYSICONINDEX | SHGFI_SMALLICON | SHGFI_USEFILEATTRIBUTES));
        if (himl && shFileInfo.iIcon)
        {
            *index = shFileInfo.iIcon;
            // I don't need to free the HIMAGELIST.
            hr = S_OK;
        }
    }

    return hr;
}

HRESULT _ParseEnumItems(_In_ IEnumShellItems* pesi, _In_ ISmartRenameManager* psrm, _In_ int depth = 0)
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
            CComPtr<ISmartRenameItemFactory> spsrif;
            hr = psrm->get_smartRenameItemFactory(&spsrif);
            if (SUCCEEDED(hr))
            {
                CComPtr<ISmartRenameItem> spNewItem;
                hr = spsrif->Create(spsi, &spNewItem);
                if (SUCCEEDED(hr))
                {
                    spNewItem->put_depth(depth);
                    hr = psrm->AddItem(spNewItem);
                }

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
                            hr = _ParseEnumItems(spesiNext, psrm, ++depth);
                        }
                    }
                }
            }

            spsi = nullptr;
        }
    }

    return hr;
}

// Iterate through the data object and add paths to the rotation manager
HRESULT EnumerateDataObject(_In_ IDataObject* pdo, _In_ ISmartRenameManager* psrm)
{
    CComPtr<IShellItemArray> spsia;
    HRESULT hr = SHCreateShellItemArrayFromDataObject(pdo, IID_PPV_ARGS(&spsia));
    if (SUCCEEDED(hr))
    {
        CComPtr<IEnumShellItems> spesi;
        hr = spsia->EnumItems(&spesi);
        if (SUCCEEDED(hr))
        {
            hr = _ParseEnumItems(spesi, psrm);
        }
    }

    return hr;
}
