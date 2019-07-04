#pragma once
#include "stdafx.h"
#include "SmartRenameInterfaces.h"
#include "srwlock.h"

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
    IFACEMETHODIMP get_parentPath(_Outptr_ PWSTR* parentPath);
    IFACEMETHODIMP get_shellItem(_Outptr_ IShellItem** ppsi);
    IFACEMETHODIMP get_originalName(_Outptr_ PWSTR* originalName);
    IFACEMETHODIMP put_newName(_In_opt_ PCWSTR newName);
    IFACEMETHODIMP get_newName(_Outptr_ PWSTR* newName);
    IFACEMETHODIMP get_isFolder(_Out_ bool* isFolder);
    IFACEMETHODIMP get_isSubFolderContent(_Out_ bool* isSubFolderContent);
    IFACEMETHODIMP put_isSubFolderContent(_In_ bool isSubFolderContent);
    IFACEMETHODIMP get_isDirty(_Out_ bool* isDirty);
    IFACEMETHODIMP get_shouldRename(_Out_ bool* shouldRename);
    IFACEMETHODIMP put_shouldRename(_In_ bool shouldRename);
    IFACEMETHODIMP get_id(_Out_ int* id);
    IFACEMETHODIMP get_iconIndex(_Out_ int* iconIndex);
    IFACEMETHODIMP get_depth(_Out_ UINT* depth);
    IFACEMETHODIMP put_depth(_In_ int depth);
    IFACEMETHODIMP Reset();

    // ISmartRenameItemFactory
    IFACEMETHODIMP Create(_In_ IShellItem* psi, _Outptr_ ISmartRenameItem** ppItem)
    {
        return CSmartRenameItem::s_CreateInstance(psi, IID_PPV_ARGS(ppItem));
    }

public:
    static HRESULT s_CreateInstance(_In_opt_ IShellItem* psi, _In_ REFIID iid, _Outptr_ void** resultInterface);

private:
    static int s_id;
    CSmartRenameItem();
    ~CSmartRenameItem();

    HRESULT _Init(_In_ IShellItem* psi);

private:
    bool     m_isDirty = false;
    bool     m_isSubFolderContent = false;
    bool     m_shouldRename = true;
    bool     m_isFolder = false;
    int      m_id = -1;
    int      m_iconIndex = -1;
    UINT     m_depth = 0;
    HRESULT  m_error = S_OK;
    PWSTR    m_path = nullptr;
    PWSTR    m_parentPath = nullptr;
    PWSTR    m_originalName = nullptr;
    PWSTR    m_newName = nullptr;
    CSRWLock m_lock;
    long     m_refCount = 0;
};