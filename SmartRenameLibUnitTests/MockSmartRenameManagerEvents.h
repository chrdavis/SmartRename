#pragma once
#include <SmartRenameInterfaces.h>

class CMockSmartRenameManagerEvents :
    public ISmartRenameManagerEvents
{
public:
    CMockSmartRenameManagerEvents() :
        m_refCount(1)
    {
    }

    // IUnknown
    IFACEMETHODIMP QueryInterface(__in REFIID riid, __deref_out void** ppv);
    IFACEMETHODIMP_(ULONG) AddRef();
    IFACEMETHODIMP_(ULONG) Release();
    
    // ISmartRenameManagerEvents
    IFACEMETHODIMP OnItemAdded(_In_ ISmartRenameItem* renameItem);
    IFACEMETHODIMP OnUpdate(_In_ ISmartRenameItem* renameItem);
    IFACEMETHODIMP OnError(_In_ ISmartRenameItem* renameItem);
    IFACEMETHODIMP OnRegExStarted();
    IFACEMETHODIMP OnRegExCanceled();
    IFACEMETHODIMP OnRegExCompleted();
    IFACEMETHODIMP OnRenameStarted();
    IFACEMETHODIMP OnRenameCompleted();

    static HRESULT s_CreateInstance(_In_ ISmartRenameManager* psrm, _Outptr_ ISmartRenameUI** ppsrui);

    ~CMockSmartRenameManagerEvents()
    {
    }

    long m_refCount = 0;
};