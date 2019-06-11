#pragma once
#include "stdafx.h"
#include "SmartRenameInterfaces.h"

class CSmartRenameItem :
    public ISmartRenameItem,
    public ISmartRenameItemFactory
{
public:
    // IUnknown
    IFACEMETHODIMP  QueryInterface(_In_ REFIID iid, _Outptr_ void** resultInterface);
    IFACEMETHODIMP_(ULONG) AddRef();
    IFACEMETHODIMP_(ULONG) Release();

    // ISmartRenameItem
    IFACEMETHODIMP get_path(_Outptr_ PWSTR* path);
    IFACEMETHODIMP put_path(_In_ PCWSTR path);
    IFACEMETHODIMP get_originalName(_Outptr_ PWSTR* originalName);
    IFACEMETHODIMP put_newName(_In_ PCWSTR newName);
    IFACEMETHODIMP get_newName(_Outptr_ PWSTR* newName);
    IFACEMETHODIMP get_isFolder(_Out_ bool* isFolder);
    IFACEMETHODIMP get_isDirty(_Out_ bool* isDirty);
    IFACEMETHODIMP get_shouldRename(_Out_ bool* shouldRename);
    IFACEMETHODIMP put_shouldRename(_In_ bool shouldRename);
    IFACEMETHODIMP get_id(_Out_ int* id);
    IFACEMETHODIMP Reset();

    // ISmartRenameItemFactory
    IFACEMETHODIMP Create(_COM_Outptr_ ISmartRenameItem** ppItem)
    {
        return CSmartRenameItem::s_CreateInstance(ppItem);
    }

public:
    static HRESULT s_CreateInstance(_Outptr_ ISmartRenameItem** renameItem);

private:
    CSmartRenameItem();
    ~CSmartRenameItem();

    HRESULT _OnItemAdded(_In_ UINT index);
    HRESULT _OnItemProcessed(_In_ UINT index);
    HRESULT _OnProgress(_In_ UINT completed, _In_ UINT total);
    HRESULT _OnCanceled();
    HRESULT _OnCompleted();

    void    _SetId(_In_ int id);
    void    _SetIsFolder(_In_ bool isFolder);
    HRESULT _GetOriginalNameFromFullPath();

private:
    bool    m_isDirty = false;
    bool    m_isFolder = false;
    bool    m_shouldRename = true;
    int     m_id = -1;
    HRESULT m_error = S_OK;
    PWSTR   m_originalName = nullptr;
    PWSTR   m_fullPath = nullptr;
    PWSTR   m_newName = nullptr;
    long    m_refCount = 0;
};