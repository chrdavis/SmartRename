#pragma once
#include "stdafx.h"
#include "SmartRenameInterfaces.h"

class CSmartRenameRegEx : public ISmartRenameRegEx
{
public:
    // IUnknown
    IFACEMETHODIMP  QueryInterface(_In_ REFIID iid, _Outptr_ void** resultInterface);
    IFACEMETHODIMP_(ULONG) AddRef();
    IFACEMETHODIMP_(ULONG) Release();

    // ISmartRenameRegEx
    IFACEMETHODIMP Replace(_In_ PCWSTR source, _In_ PCWSTR match, _In_ PCWSTR replace, _Outptr_ PWSTR *result);

    static HRESULT CreateInstance(_Outptr_ ISmartRenameRegEx **renameRegEx);

private:
    CSmartRenameRegEx();
    ~CSmartRenameRegEx();

    long m_refCount;
};