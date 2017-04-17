#include "stdafx.h"
#include "SmartRenameModel.h"

IFACEMETHODIMP_(ULONG) CSmartRenameModel::AddRef()
{
    return InterlockedIncrement(&m_refCount);
}

IFACEMETHODIMP_(ULONG) CSmartRenameModel::Release()
{
    long refCount = InterlockedDecrement(&m_refCount);

    if (refCount == 0)
    {
        delete this;
    }
    return refCount;
}

IFACEMETHODIMP CSmartRenameModel::QueryInterface(_In_ REFIID riid, _Outptr_ void **ppv)
{
    static const QITAB qit[] = {
        QITABENT(CSmartRenameModel, ISmartRenameModel),
        { 0 }
    };
    return QISearch(this, qit, riid, ppv);
}

IFACEMETHODIMP CSmartRenameModel::Advise(ISmartRenameModelEvent * renameOpEvent, DWORD * cookie)
{
    return E_NOTIMPL;
}

IFACEMETHODIMP CSmartRenameModel::UnAdvise(DWORD cookie)
{
    return E_NOTIMPL;
}

IFACEMETHODIMP CSmartRenameModel::Start()
{
    return E_NOTIMPL;
}

IFACEMETHODIMP CSmartRenameModel::Wait(DWORD timeoutMS)
{
    // Wait until the worker thread has completed its running task
    // or the supplied timeout has occurred.
    return E_NOTIMPL;
}

IFACEMETHODIMP CSmartRenameModel::Stop()
{
    return E_NOTIMPL;
}

IFACEMETHODIMP CSmartRenameModel::Reset()
{
    return E_NOTIMPL;
}

CSmartRenameModel::CSmartRenameModel() :
    m_renameModelEvent(nullptr),
    m_renameModelEventCookie(-1),
    m_refCount(1)
{
    m_workReadyEvent = CreateEvent(nullptr, TRUE, FALSE, nullptr);
    m_cancelEvent = CreateEvent(nullptr, TRUE, FALSE, nullptr);
}

CSmartRenameModel::~CSmartRenameModel()
{
}

DWORD WINAPI CSmartRenameModel::s_regexWorkerThread(void * pvoid)
{
    // Wait for multiple objects for cancel event or worker ready
    // If cancel then exit the thread
    return 0;
}
