#include "stdafx.h"
#include "SmartRenameRegEx.h"
#include <regex>
#include <string>

using namespace std;

IFACEMETHODIMP_(ULONG) CSmartRenameRegEx::AddRef()
{
    return InterlockedIncrement(&m_refCount);
}

IFACEMETHODIMP_(ULONG) CSmartRenameRegEx::Release()
{
    long refCount = InterlockedDecrement(&m_refCount);

    if (refCount == 0)
    {
        delete this;
    }
    return refCount;
}

IFACEMETHODIMP CSmartRenameRegEx::QueryInterface(_In_ REFIID riid, _Outptr_ void **ppv)
{
    static const QITAB qit[] = {
        QITABENT(CSmartRenameRegEx, ISmartRenameRegEx),
        { 0 }
    };
    return QISearch(this, qit, riid, ppv);
}


HRESULT CSmartRenameRegEx::CreateInstance(_Outptr_ ISmartRenameRegEx ** renameRegEx)
{
    *renameRegEx = nullptr;

    CSmartRenameRegEx *newRenameRegEx = new CSmartRenameRegEx();
    HRESULT hr = newRenameRegEx ? S_OK : E_OUTOFMEMORY;
    if (SUCCEEDED(hr))
    {
        hr = newRenameRegEx->QueryInterface(IID_PPV_ARGS(renameRegEx));
        newRenameRegEx->Release();
    }
    return hr;
}


CSmartRenameRegEx::CSmartRenameRegEx() :
    m_refCount(1)
{
}

CSmartRenameRegEx::~CSmartRenameRegEx()
{
}

HRESULT CSmartRenameRegEx::Replace(_In_ PCWSTR source, _In_ PCWSTR match, _In_ PCWSTR replace, _Outptr_ PWSTR *result)
{
    *result = nullptr;
    HRESULT hr = (source && wcslen(source) > 0 && match && wcslen(match) > 0) ? S_OK : E_INVALIDARG;
    if (SUCCEEDED(hr))
    {
        // TODO: optimize so we aren't having to convert to wstring all the time
        // TODO: also be sure we handle exceptions here
        wstring src = wstring(source);
        wstring rep = wstring(replace);
        wstring res = regex_replace(src, wregex(wstring(match)), rep);

        *result = StrDup(res.c_str());
        hr = (*result) ? S_OK : E_OUTOFMEMORY;
    }
    return hr;
}