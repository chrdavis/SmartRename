#include "stdafx.h"
#include "SmartRenameRegEx.h"
#include <regex>
#include <string>
#include <algorithm>


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
    RENAME_REGEX_EVENT srre;
    srre.cookie = m_cookie;
    srre.pEvents = regExEvents;
    regExEvents->AddRef();
    m_renameRegExEvents.push_back(srre);

    *cookie = m_cookie;

    return S_OK;
}

IFACEMETHODIMP CSmartRenameRegEx::UnAdvise(_In_ DWORD cookie)
{
    HRESULT hr = E_FAIL;
    CSRWExclusiveAutoLock lock(&m_lockEvents);

    for (std::vector<RENAME_REGEX_EVENT>::iterator it = m_renameRegExEvents.begin(); it != m_renameRegExEvents.end(); ++it)
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
    *searchTerm = nullptr;
    HRESULT hr = m_searchTerm ? S_OK : E_FAIL;
    if (SUCCEEDED(hr))
    {
        CSRWSharedAutoLock lock(&m_lock);
        hr = SHStrDup(m_searchTerm, searchTerm);
    }
    return hr;
}

IFACEMETHODIMP CSmartRenameRegEx::put_searchTerm(_In_ PCWSTR searchTerm)
{
    bool changed = false;
    HRESULT hr = searchTerm ? S_OK : E_INVALIDARG;
    if (SUCCEEDED(hr))
    {
        CSRWExclusiveAutoLock lock(&m_lock);
        if (m_searchTerm == nullptr || lstrcmp(searchTerm, m_searchTerm) != 0)
        {
            changed = true;
            CoTaskMemFree(m_searchTerm);
            hr = SHStrDup(searchTerm, &m_searchTerm);
        }
    }

    if (SUCCEEDED(hr) && changed)
    {
        _OnSearchTermChanged();
    }

    return hr;
}

IFACEMETHODIMP CSmartRenameRegEx::get_replaceTerm(_Outptr_ PWSTR* replaceTerm)
{
    *replaceTerm = nullptr;
    HRESULT hr = m_replaceTerm ? S_OK : E_FAIL;
    if (SUCCEEDED(hr))
    {
        CSRWSharedAutoLock lock(&m_lock);
        hr = SHStrDup(m_replaceTerm, replaceTerm);
    }
    return hr;
}

IFACEMETHODIMP CSmartRenameRegEx::put_replaceTerm(_In_ PCWSTR replaceTerm)
{
    bool changed = false;
    HRESULT hr = replaceTerm ? S_OK : E_INVALIDARG;
    if (SUCCEEDED(hr))
    {
        CSRWExclusiveAutoLock lock(&m_lock);
        if (m_replaceTerm == nullptr || lstrcmp(replaceTerm, m_replaceTerm) != 0)
        {
            changed = true;
            CoTaskMemFree(m_replaceTerm);
            hr = SHStrDup(replaceTerm, &m_replaceTerm);
        }
    }

    if (SUCCEEDED(hr) && changed)
    {
        _OnReplaceTermChanged();
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
        _OnFlagsChanged();
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
    // Init to empty strings
    SHStrDup(L"", &m_searchTerm);
    SHStrDup(L"", &m_replaceTerm);
}

CSmartRenameRegEx::~CSmartRenameRegEx()
{
    CoTaskMemFree(m_searchTerm);
    CoTaskMemFree(m_replaceTerm);
}

HRESULT CSmartRenameRegEx::Replace(_In_ PCWSTR source, _Outptr_ PWSTR* result)
{
    *result = nullptr;

    CSRWSharedAutoLock lock(&m_lock);
    HRESULT hr = (source && wcslen(source) > 0 && m_searchTerm && wcslen(m_searchTerm) > 0) ? S_OK : E_INVALIDARG;
    if (SUCCEEDED(hr))
    {
        wstring res = source;
        try
        {
            // TODO: creating the regex could be costly.  May want to cache this.
            std::wstring sourceToUse(source);
            std::wstring searchTerm(m_searchTerm);
            std::wstring replaceTerm(m_replaceTerm ? wstring(m_replaceTerm) : wstring(L""));

            if (m_flags & UseRegularExpressions)
            {
                std::wregex pattern(m_searchTerm, (!(m_flags & CaseSensitive)) ? regex_constants::icase | regex_constants::ECMAScript : regex_constants::ECMAScript);
                res = regex_replace(sourceToUse, pattern, replaceTerm, (m_flags & MatchAllOccurences) ? regex_constants::match_default : regex_constants::format_first_only);
            }
            else
            {
                // Simple search and replace
                size_t pos = 0;
                do
                {
                    pos = _Find(sourceToUse, searchTerm, (!(m_flags & CaseSensitive)), pos);
                    if (pos != std::string::npos)
                    {
                        res = sourceToUse.replace(pos, searchTerm.length(), replaceTerm);
                        pos += replaceTerm.length();
                    }

                    if (!(m_flags & MatchAllOccurences))
                    {
                        break;
                    }
                } while (pos != std::string::npos);
            }

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

size_t CSmartRenameRegEx::_Find(std::wstring data, std::wstring toSearch, bool caseInsensitive, size_t pos)
{
    if (caseInsensitive)
    {
        // Convert to lower
        std::transform(data.begin(), data.end(), data.begin(), ::towlower);
        std::transform(toSearch.begin(), toSearch.end(), toSearch.begin(), ::towlower);
    }

    // Find sub string position in given string starting at position pos
    return data.find(toSearch, pos);
}

void CSmartRenameRegEx::_OnSearchTermChanged()
{
    CSRWSharedAutoLock lock(&m_lockEvents);

    for (std::vector<RENAME_REGEX_EVENT>::iterator it = m_renameRegExEvents.begin(); it != m_renameRegExEvents.end(); ++it)
    {
        if (it->pEvents)
        {
            it->pEvents->OnSearchTermChanged(m_searchTerm);
        }
    }
}

void CSmartRenameRegEx::_OnReplaceTermChanged()
{
    CSRWSharedAutoLock lock(&m_lockEvents);

    for (std::vector<RENAME_REGEX_EVENT>::iterator it = m_renameRegExEvents.begin(); it != m_renameRegExEvents.end(); ++it)
    {
        if (it->pEvents)
        {
            it->pEvents->OnReplaceTermChanged(m_replaceTerm);
        }
    }
}

void CSmartRenameRegEx::_OnFlagsChanged()
{
    CSRWSharedAutoLock lock(&m_lockEvents);

    for (std::vector<RENAME_REGEX_EVENT>::iterator it = m_renameRegExEvents.begin(); it != m_renameRegExEvents.end(); ++it)
    {
        if (it->pEvents)
        {
            it->pEvents->OnFlagsChanged(m_flags);
        }
    }
}
