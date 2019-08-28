#pragma once
#include "stdafx.h"
#include <SmartRenameItem.h>
#include "srwlock.h"

class CMockSmartRenameItem :
    public CSmartRenameItem
{
public:
    static HRESULT CreateInstance(_In_opt_ PCWSTR path, _In_opt_ PCWSTR originalName, _In_ UINT depth, _In_ bool isFolder, _Outptr_ ISmartRenameItem** ppItem);
    void Init(_In_opt_ PCWSTR path, _In_opt_ PCWSTR originalName, _In_ UINT depth, _In_ bool isFolder);
};