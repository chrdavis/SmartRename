#pragma once
#include "stdafx.h"
#include "SmartRenameInterfaces.h"

class CSmartRenameModel : public ISmartRenameModel
{
public:
    // IUnknown
    IFACEMETHODIMP  QueryInterface(_In_ REFIID iid, _Outptr_ void** resultInterface);
    IFACEMETHODIMP_(ULONG) AddRef();
    IFACEMETHODIMP_(ULONG) Release();

    // ISmartRenameModel
    IFACEMETHODIMP Advise(_In_ ISmartRenameModelEvent *renameOpEvent, _Out_ DWORD *cookie);
    IFACEMETHODIMP UnAdvise(_In_ DWORD cookie);
    IFACEMETHODIMP Start();
    IFACEMETHODIMP Wait(_In_ DWORD timeoutMS);
    IFACEMETHODIMP Stop();
    IFACEMETHODIMP Reset();

private:
    CSmartRenameModel();
    ~CSmartRenameModel();

    static DWORD WINAPI s_regexWorkerThread(_In_ void *pvoid);

    // TODO: Create an array of these for multiple views to subscribe to?
    CComPtr<ISmartRenameModelEvent> m_renameModelEvent;
    DWORD m_renameModelEventCookie;    // Unique id associated with a ISmartRenameModelEvent

    HANDLE m_workerThreadHandle;  // Thread handle for worker thread
    HANDLE m_workReadyEvent;      // Signaled to let the worker thread know that work is ready to be processed
    HANDLE m_cancelEvent;         // Signaled to cancel the current worker thread pass
    HANDLE m_exitEvent;           // Signaled when requesting to exit the worker thread
    long m_refCount;
};