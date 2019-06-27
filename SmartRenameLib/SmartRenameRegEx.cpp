#include "stdafx.h"
#include "SmartRenameRegEx.h"
#include <regex>
#include <string>


using namespace std;
using std::regex_error;

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

IFACEMETHODIMP CSmartRenameRegEx::QueryInterface(_In_ REFIID riid, _Outptr_ void** ppv)
{
    static const QITAB qit[] = {
        QITABENT(CSmartRenameRegEx, ISmartRenameRegEx),
        { 0 }
    };
    return QISearch(this, qit, riid, ppv);
}

IFACEMETHODIMP CSmartRenameRegEx::Advise(_In_ ISmartRenameRegExEvents* regExEvents, _Out_ DWORD* cookie)
{
    CSRWExclusiveAutoLock lock(&m_lockEvents);
    m_cookie++;
    SMART_RENAME_REGEX_EVENT srre;
    srre.cookie = m_cookie;
    srre.pEvents = regExEvents;
    regExEvents->AddRef();
    m_smartRenameRegExEvents.push_back(srre);

    *cookie = m_cookie;

    return S_OK;
}

IFACEMETHODIMP CSmartRenameRegEx::UnAdvise(_In_ DWORD cookie)
{
    HRESULT hr = E_FAIL;
    CSRWExclusiveAutoLock lock(&m_lockEvents);

    for (std::vector<SMART_RENAME_REGEX_EVENT>::iterator it = m_smartRenameRegExEvents.begin(); it != m_smartRenameRegExEvents.end(); ++it)
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

IFACEMETHODIMP CSmartRenameRegEx::get_searchTerm(_Outptr_ PWSTR* searchTerm)
{
    CSRWSharedAutoLock lock(&m_lock);
    *searchTerm = nullptr;
    HRESULT hr = m_searchTerm ? S_OK : E_FAIL;
    if (SUCCEEDED(hr))
    {
        hr = SHStrDup(m_searchTerm, searchTerm);
    }
    return hr;
}

IFACEMETHODIMP CSmartRenameRegEx::put_searchTerm(_In_ PCWSTR searchTerm)
{
    CSRWExclusiveAutoLock lock(&m_lock);
    HRESULT hr = searchTerm ? S_OK : E_INVALIDARG;
    if (SUCCEEDED(hr))
    {
        if (m_searchTerm == nullptr || lstrcmp(searchTerm, m_searchTerm) != 0)
        {
            CoTaskMemFree(m_searchTerm);
            hr = SHStrDup(searchTerm, &m_searchTerm);
            if (SUCCEEDED(hr))
            {
                _OnSearchTermChanged();
            }
        }
    }
    return hr;
}

IFACEMETHODIMP CSmartRenameRegEx::get_replaceTerm(_Outptr_ PWSTR* replaceTerm)
{
    CSRWSharedAutoLock lock(&m_lock);
    *replaceTerm = nullptr;
    HRESULT hr = m_replaceTerm ? S_OK : E_FAIL;
    if (SUCCEEDED(hr))
    {
        hr = SHStrDup(m_replaceTerm, replaceTerm);
    }
    return hr;
}

IFACEMETHODIMP CSmartRenameRegEx::put_replaceTerm(_In_ PCWSTR replaceTerm)
{
    CSRWExclusiveAutoLock lock(&m_lock);
    HRESULT hr = replaceTerm ? S_OK : E_INVALIDARG;
    if (SUCCEEDED(hr))
    {
        if (m_replaceTerm == nullptr || lstrcmp(replaceTerm, m_replaceTerm) != 0)
        {
            CoTaskMemFree(m_replaceTerm);
            hr = SHStrDup(replaceTerm, &m_replaceTerm);
            if (SUCCEEDED(hr))
            {
                _OnReplaceTermChanged();
            }
        }
    }
    return hr;
}

IFACEMETHODIMP CSmartRenameRegEx::get_flags(_Out_ DWORD* flags)
{
    *flags = m_flags;
    return S_OK;
}

IFACEMETHODIMP CSmartRenameRegEx::put_flags(_In_ DWORD flags)
{
    if (m_flags != flags)
    {
        m_flags = flags;
    }
    return S_OK;
}

HRESULT CSmartRenameRegEx::s_CreateInstance(_Outptr_ ISmartRenameRegEx** renameRegEx)
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
    CoTaskMemFree(m_searchTerm);
    CoTaskMemFree(m_replaceTerm);
}

HRESULT CSmartRenameRegEx::Replace(_In_ PCWSTR source, _Outptr_ PWSTR* result)
{
    *result = nullptr;
    HRESULT hr = (source && wcslen(source) > 0 && m_searchTerm && wcslen(m_searchTerm) > 0) ? S_OK : E_INVALIDARG;
    if (SUCCEEDED(hr))
    {
        wstring res = source;
        try
        {
            res = regex_replace(wstring(source), wregex(wstring(m_searchTerm)), m_replaceTerm ? wstring(m_replaceTerm) : wstring(L""));
             *result = StrDup(res.c_str());
            hr = (*result) ? S_OK : E_OUTOFMEMORY;
        }
        catch (regex_error e)
        {
            hr = E_FAIL;
        }
    }
    return hr;
}

void CSmartRenameRegEx::_OnSearchTermChanged()
{
    CSRWExclusiveAutoLock lock(&m_lockEvents);

    for (std::vector<SMART_RENAME_REGEX_EVENT>::iterator it = m_smartRenameRegExEvents.begin(); it != m_smartRenameRegExEvents.end(); ++it)
    {
        if (it->pEvents)
        {
            it->pEvents->OnSearchTermChanged(m_searchTerm);
        }
    }
}

void CSmartRenameRegEx::_OnReplaceTermChanged()
{
    CSRWExclusiveAutoLock lock(&m_lockEvents);

    for (std::vector<SMART_RENAME_REGEX_EVENT>::iterator it = m_smartRenameRegExEvents.begin(); it != m_smartRenameRegExEvents.end(); ++it)
    {
        if (it->pEvents)
        {
            it->pEvents->OnReplaceTermChanged(m_replaceTerm);
        }
    }
}
