#pragma once
#include "stdafx.h"
#include "SmartRenameInterfaces.h"

class CSmartRenameItem : public ISmartRenameItem
{
public:
    // IUnknown
    IFACEMETHODIMP  QueryInterface(_In_ REFIID iid, _Outptr_ void** resultInterface);
    IFACEMETHODIMP_(ULONG) AddRef();
    IFACEMETHODIMP_(ULONG) Release();

    // ISmartRenameItem
    IFACEMETHODIMP GetOriginalName(_Outptr_ PWSTR *originalName);
    IFACEMETHODIMP SetNewName(_In_ PCWSTR newName);
    IFACEMETHODIMP GetNewName(_Outptr_ PWSTR *newName);
    IFACEMETHODIMP GetIsFolder(_Out_ bool *isFolder);
    IFACEMETHODIMP GetIsDirty(_Out_ bool *isDirty);
    IFACEMETHODIMP GetId(_Out_ int *id);
    IFACEMETHODIMP Reset();

public:
    static HRESULT CreateInstance(
        _In_ PCWSTR fullPath,
        _In_ PCWSTR oldName,
        _In_ bool isFolder,
        _In_ int id,
        _Outptr_ ISmartRenameItem **renameItem);

private:
    CSmartRenameItem();
    ~CSmartRenameItem();

    void    _SetId(_In_ int id);
    void    _SetIsFolder(_In_ bool isFolder);
    HRESULT _SetFullPath(_In_ PCWSTR fullPath);
    HRESULT _SetOldName(_In_ PCWSTR oldName);

private:
    bool    m_isDirty;
    bool    m_isFolder;
    int     m_id;
    HRESULT m_error;
    PWSTR   m_oldName;
    PWSTR   m_fullPath;
    PWSTR   m_newName;
    long    m_refCount;
};