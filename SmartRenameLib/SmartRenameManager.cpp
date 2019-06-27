#include "stdafx.h"
#include "SmartRenameManager.h"
#include <shobjidl.h>
#include <shellapi.h>
#include <algorithm>


// The default FOF flags to use in the rename operations
#define FOF_DEFAULTFLAGS (FOF_ALLOWUNDO | FOFX_SHOWELEVATIONPROMPT | FOF_RENAMEONCOLLISION)

IFACEMETHODIMP_(ULONG) CSmartRenameManager::AddRef()
{
    return InterlockedIncrement(&m_refCount);
}

IFACEMETHODIMP_(ULONG) CSmartRenameManager::Release()
{
    long refCount = InterlockedDecrement(&m_refCount);

    if (refCount == 0)
    {
        delete this;
    }
    return refCount;
}

IFACEMETHODIMP CSmartRenameManager::QueryInterface(_In_ REFIID riid, _Outptr_ void** ppv)
{
    static const QITAB qit[] = {
        QITABENT(CSmartRenameManager, ISmartRenameManager),
        QITABENT(CSmartRenameManager, ISmartRenameRegExEvents),
        { 0 }
    };
    return QISearch(this, qit, riid, ppv);
}

IFACEMETHODIMP CSmartRenameManager::Advise(_In_ ISmartRenameManagerEvents* renameOpEvents, _Out_ DWORD* cookie)
{
    CSRWExclusiveAutoLock lock(&m_lockEvents);
    m_cookie++;
    SMART_RENAME_MGR_EVENT srme;
    srme.cookie = m_cookie;
    srme.pEvents = renameOpEvents;
    renameOpEvents->AddRef();
    m_SmartRenameManagerEvents.push_back(srme);

    *cookie = m_cookie;

    return S_OK;
}

IFACEMETHODIMP CSmartRenameManager::UnAdvise(_In_ DWORD cookie)
{
    HRESULT hr = E_FAIL;
    CSRWExclusiveAutoLock lock(&m_lockEvents);

    for (std::vector<SMART_RENAME_MGR_EVENT>::iterator it = m_SmartRenameManagerEvents.begin(); it != m_SmartRenameManagerEvents.end(); ++it)
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

IFACEMETHODIMP CSmartRenameManager::Start()
{
    return E_NOTIMPL;
}

IFACEMETHODIMP CSmartRenameManager::Stop()
{
    return E_NOTIMPL;
}

IFACEMETHODIMP CSmartRenameManager::Rename(_In_ HWND hwndParent)
{
    m_hwndParent = hwndParent;
    return _PerformFileOperation();
}

IFACEMETHODIMP CSmartRenameManager::Reset()
{
    // Stop all threads and wait
    // Reset all rename items
    return E_NOTIMPL;
}

IFACEMETHODIMP CSmartRenameManager::Shutdown()
{
    return E_NOTIMPL;
}

IFACEMETHODIMP CSmartRenameManager::AddItem(_In_ ISmartRenameItem* pItem)
{
    // Scope lock
    {
        CSRWExclusiveAutoLock lock(&m_lockItems);
        m_smartRenameItems.push_back(pItem);
        pItem->AddRef();
    }

    _OnItemAdded(pItem);

    return S_OK;
}

IFACEMETHODIMP CSmartRenameManager::GetItemByIndex(_In_ UINT index, _COM_Outptr_ ISmartRenameItem** ppItem)
{
    *ppItem = nullptr;
    CSRWSharedAutoLock lock(&m_lockItems);
    HRESULT hr = E_FAIL;
    if (index < m_smartRenameItems.size())
    {
        *ppItem = m_smartRenameItems.at(index);
        (*ppItem)->AddRef();
        hr = S_OK;
    }

    return hr;
}

IFACEMETHODIMP CSmartRenameManager::GetItemById(_In_ int id, _COM_Outptr_ ISmartRenameItem** ppItem)
{
    *ppItem = nullptr;

    CSRWSharedAutoLock lock(&m_lockItems);
    HRESULT hr = E_FAIL;
    auto iterator = std::find_if(m_smartRenameItems.begin(), m_smartRenameItems.end(), [id](_In_ ISmartRenameItem* currentItem)
        {
            int idCurrent;
            currentItem->get_id(&idCurrent);
            return (idCurrent == id);
        });

    if (iterator != m_smartRenameItems.end())
    {
        *ppItem = (*iterator);
        (*ppItem)->AddRef();
        hr = S_OK;
    }

    return hr;
}

IFACEMETHODIMP CSmartRenameManager::GetItemCount(_Out_ UINT* count)
{
    CSRWSharedAutoLock lock(&m_lockItems);
    *count = static_cast<UINT>(m_smartRenameItems.size());
    return S_OK;
}

IFACEMETHODIMP CSmartRenameManager::get_flags(_Out_ DWORD* flags)
{
    *flags = m_flags;
    return S_OK;
}

IFACEMETHODIMP CSmartRenameManager::put_flags(_In_ DWORD flags)
{
    if (flags != m_flags)
    {
        m_flags = flags;
        m_spRegEx->put_flags(flags);
    }
    return S_OK;
}

IFACEMETHODIMP CSmartRenameManager::get_smartRenameRegEx(_COM_Outptr_ ISmartRenameRegEx** ppRegEx)
{
    *ppRegEx = nullptr;
    HRESULT hr = E_FAIL;
    if (m_spRegEx)
    {
        hr = S_OK;
        *ppRegEx = m_spRegEx;
        (*ppRegEx)->AddRef();
    }
    return hr;
}

IFACEMETHODIMP CSmartRenameManager::put_smartRenameRegEx(_In_ ISmartRenameRegEx* pRegEx)
{
    m_spRegEx = pRegEx;
    return S_OK;
}

IFACEMETHODIMP CSmartRenameManager::get_smartRenameItemFactory(_COM_Outptr_ ISmartRenameItemFactory** ppItemFactory)
{
    *ppItemFactory = nullptr;
    HRESULT hr = E_FAIL;
    if (m_spItemFactory)
    {
        hr = S_OK;
        *ppItemFactory = m_spItemFactory;
        (*ppItemFactory)->AddRef();
    }
    return hr;
}

IFACEMETHODIMP CSmartRenameManager::put_smartRenameItemFactory(_In_ ISmartRenameItemFactory* pItemFactory)
{
    m_spItemFactory = pItemFactory;
    return S_OK;
}

IFACEMETHODIMP CSmartRenameManager::OnSearchTermChanged(_In_ PCWSTR /*searchTerm*/)
{
    // TODO: Cancel and restart rename regex thread
    return S_OK;
}

IFACEMETHODIMP CSmartRenameManager::OnReplaceTermChanged(_In_ PCWSTR /*replaceTerm*/)
{
    // TODO: Cancel and restart rename regex thread
    return S_OK;
}

HRESULT CSmartRenameManager::s_CreateInstance(_Outptr_ ISmartRenameManager** ppsrm)
{
    *ppsrm = nullptr;
    CSmartRenameManager *psrm = new CSmartRenameManager();
    HRESULT hr = psrm ? S_OK : E_OUTOFMEMORY;
    if (SUCCEEDED(hr))
    {
        hr = psrm->_Init();
        if (SUCCEEDED(hr))
        {
            hr = psrm->QueryInterface(IID_PPV_ARGS(ppsrm));
        }
        psrm->Release();
    }
    return hr;
}

CSmartRenameManager::CSmartRenameManager() :
    m_refCount(1)
{
}

CSmartRenameManager::~CSmartRenameManager()
{
}

HRESULT CSmartRenameManager::_Init()
{
    // Guaranteed to succeed
    m_startFileOpWorkerEvent = CreateEvent(nullptr, TRUE, FALSE, nullptr);
    m_startRegExWorkerEvent = CreateEvent(nullptr, TRUE, FALSE, nullptr);
    m_cancelRegExWorkerEvent = CreateEvent(nullptr, TRUE, FALSE, nullptr);
   
    return S_OK;
}

// Custom messages for worker threads
enum
{
    SRM_REGEX_ITEM_PROCESSED = (WM_APP + 1),  // Single smart rename item processed by regex worker thread
    SRM_REGEX_CANCELED,                       // Regex operation was canceled
    SRM_REGEX_COMPLETE,                       // Regex worker thread completed
    SRM_FILEOP_COMPLETE                       // File Operation worker thread completed
};

struct WorkerThreadData
{
    DWORD managerThreadId = 0;
    HANDLE startEvent = nullptr;
    HANDLE cancelEvent = nullptr;
    HWND hwndParent = nullptr;
    CComPtr<ISmartRenameManager> spsrm;
};

HRESULT CSmartRenameManager::_PerformFileOperation()
{
    // Create worker thread which will perform the actual rename
    HRESULT hr = _CreateFileOpWorkerThread();
    if (SUCCEEDED(hr))
    {
        _OnRenameStarted();

        // Signal the worker thread that they can start working. We needed to wait until we
        // were ready to process thread messages.
        SetEvent(m_startFileOpWorkerEvent);

        while (true)
        {
            // Check if worker thread has exited
            if (WaitForSingleObject(m_fileOpWorkerThreadHandle, 0) == WAIT_OBJECT_0)
            {
                break;
            }

            MSG msg;
            while (PeekMessage(&msg, nullptr, 0, 0, PM_REMOVE))
            {
                if (msg.message == SRM_FILEOP_COMPLETE)
                {
                    // Worker thread completed
                    break;
                }
                else
                {
                    TranslateMessage(&msg);
                    DispatchMessage(&msg);
                }
            }
        }

        _OnRenameCompleted();
    }

    return 0;
}

HRESULT CSmartRenameManager::_CreateFileOpWorkerThread()
{
    WorkerThreadData* pwtd = new WorkerThreadData;
    HRESULT hr = pwtd ? S_OK : E_OUTOFMEMORY;
    if (SUCCEEDED(hr))
    {
        pwtd->managerThreadId = GetCurrentThreadId();
        pwtd->startEvent = m_startRegExWorkerEvent;
        pwtd->cancelEvent = nullptr;
        pwtd->spsrm = this;
        m_fileOpWorkerThreadHandle = CreateThread(nullptr, 0, s_fileOpWorkerThread, pwtd, 0, nullptr);
        hr = (m_fileOpWorkerThreadHandle) ? S_OK : E_FAIL;
        if (FAILED(hr))
        {
            delete pwtd;
        }
    }

    return hr;
}

DWORD WINAPI CSmartRenameManager::s_fileOpWorkerThread(_In_ void* pv)
{
    if (SUCCEEDED(CoInitializeEx(NULL, 0)))
    {
        WorkerThreadData* pwtd = reinterpret_cast<WorkerThreadData*>(pv);
        if (pwtd)
        {
            // Wait to be told we can begin
            if (WaitForSingleObject(pwtd->startEvent, INFINITE) == WAIT_OBJECT_0)
            {
                CComPtr<ISmartRenameRegEx> spRenameRegEx;
                if (SUCCEEDED(pwtd->spsrm->get_smartRenameRegEx(&spRenameRegEx)))
                {
                    // Create IFileOperation interface
                    CComPtr<IFileOperation> spFileOp;
                    if (SUCCEEDED(CoCreateInstance(CLSID_FileOperation, nullptr, CLSCTX_INPROC_SERVER, IID_PPV_ARGS(&spFileOp))))
                    {
                        DWORD flags = 0;
                        spRenameRegEx->get_flags(&flags);

                        UINT itemCount = 0;
                        pwtd->spsrm->GetItemCount(&itemCount);
                        // Add each rename operation
                        for (UINT u = 0; u <= itemCount; u++)
                        {
                            CComPtr<ISmartRenameItem> spItem;
                            if (SUCCEEDED(pwtd->spsrm->GetItemByIndex(u, &spItem)))
                            {
                                if (_ShouldRenameItem(spItem, flags))
                                {
                                    PWSTR path = nullptr;
                                    if (SUCCEEDED(spItem->get_path(&path)))
                                    {
                                        // TODO: handle extensions and enumerating items
                                        PWSTR newName = nullptr;
                                        if (SUCCEEDED(spItem->get_newName(&newName)))
                                        {
                                            CComPtr<IShellItem> spShellItem;
                                            if (SUCCEEDED(SHCreateItemFromParsingName(path, nullptr, IID_PPV_ARGS(&spShellItem))))
                                            {
                                                spFileOp->RenameItem(spShellItem, newName, nullptr);
                                            }
                                            CoTaskMemFree(newName);
                                        }
                                        CoTaskMemFree(path);
                                    }
                                }

                                // TODO: WE HAVE NOT RENAMED THIS YET!  WE SHOULD USING THE PROGRESS SYNK AND HAVE A MAP FROM PATH TO INDEX OR SOMETHING
                                // Send the manager thread the item processed message
                                PostThreadMessage(pwtd->managerThreadId, SRM_REGEX_ITEM_PROCESSED, GetCurrentThreadId(), u);
                            }
                        }

                        // Set the operation flags
                        if (SUCCEEDED(spFileOp->SetOperationFlags(FOF_DEFAULTFLAGS)))
                        {
                            // TODO: Update with hwnd of UI
                            // Set the parent window
                            if (pwtd->hwndParent)
                            {
                                spFileOp->SetOwnerWindow(pwtd->hwndParent);
                            }
                            
                            // Perform the operation
                            // We don't care about the return code here. We would rather
                            // return control back to explorer so the user can cleanly
                            // undo the operation if it failed halfway through.
                            spFileOp->PerformOperations();
                        }
                    }
                }
            }

            // Send the manager thread the completion message
            PostThreadMessage(pwtd->managerThreadId, SRM_REGEX_COMPLETE, GetCurrentThreadId(), 0);

            delete pwtd;
        }
        CoUninitialize();
    }

    return 0;
}

bool CSmartRenameManager::_ShouldRenameItem(_In_ ISmartRenameItem* item, _In_ DWORD flags)
{
    // Should we perform a rename on this item given its
    // state and the options that were set?
    bool isDirty = false;
    bool shouldRename = false;
    bool isFolder = false;
    bool isSubFolderContent = false;
    item->get_isDirty(&isDirty);
    item->get_shouldRename(&shouldRename);
    item->get_isFolder(&isFolder);
    item->get_isSubFolderContent(&isSubFolderContent);
    return (shouldRename && isDirty &&
        (!(isFolder && (flags & SmartRenameFlags::ExcludeFolders))) &&
        (!(!isFolder && (flags & SmartRenameFlags::ExcludeFiles))) &&
        (!isSubFolderContent && (flags & SmartRenameFlags::ExcludeSubfolders)));
}

HRESULT CSmartRenameManager::_PerformRegExRename()
{
    // Create worker thread which will message us progress and completion.
    HRESULT hr = _CreateRegExWorkerThread();
    if (SUCCEEDED(hr))
    {
        ResetEvent(m_cancelRegExWorkerEvent);

        _OnRegExStarted();

        // Signal the worker thread that they can start working. We needed to wait until we
        // were ready to process thread messages.
        SetEvent(m_startRegExWorkerEvent);

        while (true)
        {
            // Check if worker thread has exited
            if (WaitForSingleObject(m_regExWorkerThreadHandle, 0) == WAIT_OBJECT_0)
            {
                break;
            }

            MSG msg;
            while (PeekMessage(&msg, nullptr, 0, 0, PM_REMOVE))
            {
                if (msg.message == SRM_REGEX_ITEM_PROCESSED)
                {
                }
                else if (msg.message == SRM_REGEX_CANCELED)
                {
                    _OnRegExCanceled();
                }
                else if (msg.message == SRM_REGEX_COMPLETE)
                {
                    // Worker thread completed
                    break;
                }
                else
                {
                    TranslateMessage(&msg);
                    DispatchMessage(&msg);
                }
            }
        }

        _OnRegExCompleted();
    }

    return 0;
}

HRESULT CSmartRenameManager::_CreateRegExWorkerThread()
{
    WorkerThreadData* pwtd = new WorkerThreadData;
    HRESULT hr = pwtd ? S_OK : E_OUTOFMEMORY;
    if (SUCCEEDED(hr))
    {
        pwtd->managerThreadId = GetCurrentThreadId();
        pwtd->startEvent = m_startRegExWorkerEvent;
        pwtd->cancelEvent = m_cancelRegExWorkerEvent;
        pwtd->hwndParent = m_hwndParent;
        pwtd->spsrm = this;
        m_regExWorkerThreadHandle = CreateThread(nullptr, 0, s_regexWorkerThread, pwtd, 0, nullptr);
        hr = (m_regExWorkerThreadHandle) ? S_OK : E_FAIL;
        if (FAILED(hr))
        {
            delete pwtd;
        }
    }

    return hr;
}

DWORD WINAPI CSmartRenameManager::s_regexWorkerThread(_In_ void* pv)
{
    if (SUCCEEDED(CoInitializeEx(NULL, 0)))
    {
        WorkerThreadData* pwtd = reinterpret_cast<WorkerThreadData*>(pv);
        if (pwtd)
        {
            // Wait to be told we can begin
            if (WaitForSingleObject(pwtd->startEvent, INFINITE) == WAIT_OBJECT_0)
            {
                CComPtr<ISmartRenameRegEx> spRenameRegEx;
                if (SUCCEEDED(pwtd->spsrm->get_smartRenameRegEx(&spRenameRegEx)))
                {
                    UINT itemCount = 0;
                    pwtd->spsrm->GetItemCount(&itemCount);
                    for (UINT u = 0; u <= itemCount; u++)
                    {
                        // Check if cancel event is signaled
                        if (WaitForSingleObject(pwtd->cancelEvent, 0) == WAIT_OBJECT_0)
                        {
                            // Canceled from manager
                            // Send the manager thread the canceled message
                            PostThreadMessage(pwtd->managerThreadId, SRM_REGEX_CANCELED, GetCurrentThreadId(), 0);
                            break;
                        }

                        CComPtr<ISmartRenameItem> spItem;
                        if (SUCCEEDED(pwtd->spsrm->GetItemByIndex(u, &spItem)))
                        {
                            PWSTR originalName = nullptr;
                            if (SUCCEEDED(spItem->get_originalName(&originalName)))
                            {
                                PWSTR newName = nullptr;
                                if (SUCCEEDED(spRenameRegEx->Replace(originalName, &newName)))
                                {
                                    spItem->put_newName(newName);
                                    CoTaskMemFree(newName);
                                }
                                CoTaskMemFree(originalName);
                            }
                            // Send the manager thread the item processed message
                            PostThreadMessage(pwtd->managerThreadId, SRM_REGEX_ITEM_PROCESSED, GetCurrentThreadId(), u);
                        }
                    }
                }
            }

            // Send the manager thread the completion message
            PostThreadMessage(pwtd->managerThreadId, SRM_REGEX_COMPLETE, GetCurrentThreadId(), 0);

            delete pwtd;
        }
        CoUninitialize();
    }

    return 0;
}

void CSmartRenameManager::_Cancel()
{
    SetEvent(m_startFileOpWorkerEvent);
    SetEvent(m_startRegExWorkerEvent);
    SetEvent(m_cancelRegExWorkerEvent);
}

void CSmartRenameManager::_OnItemAdded(_In_ ISmartRenameItem* renameItem)
{
    CSRWExclusiveAutoLock lock(&m_lockEvents);

    for (std::vector<SMART_RENAME_MGR_EVENT>::iterator it = m_SmartRenameManagerEvents.begin(); it != m_SmartRenameManagerEvents.end(); ++it)
    {
        if (it->pEvents)
        {
            it->pEvents->OnItemAdded(renameItem);
        }
    }
}

void CSmartRenameManager::_OnUpdate(_In_ ISmartRenameItem* renameItem)
{
    CSRWExclusiveAutoLock lock(&m_lockEvents);

    for (std::vector<SMART_RENAME_MGR_EVENT>::iterator it = m_SmartRenameManagerEvents.begin(); it != m_SmartRenameManagerEvents.end(); ++it)
    {
        if (it->pEvents)
        {
            it->pEvents->OnUpdate(renameItem);
        }
    }
}

void CSmartRenameManager::_OnError(_In_ ISmartRenameItem* renameItem)
{
    CSRWExclusiveAutoLock lock(&m_lockEvents);

    for (std::vector<SMART_RENAME_MGR_EVENT>::iterator it = m_SmartRenameManagerEvents.begin(); it != m_SmartRenameManagerEvents.end(); ++it)
    {
        if (it->pEvents)
        {
            it->pEvents->OnError(renameItem);
        }
    }
}

void CSmartRenameManager::_OnRegExStarted()
{
    CSRWExclusiveAutoLock lock(&m_lockEvents);

    for (std::vector<SMART_RENAME_MGR_EVENT>::iterator it = m_SmartRenameManagerEvents.begin(); it != m_SmartRenameManagerEvents.end(); ++it)
    {
        if (it->pEvents)
        {
            it->pEvents->OnRegExStarted();
        }
    }
}

void CSmartRenameManager::_OnRegExCanceled()
{
    CSRWExclusiveAutoLock lock(&m_lockEvents);

    for (std::vector<SMART_RENAME_MGR_EVENT>::iterator it = m_SmartRenameManagerEvents.begin(); it != m_SmartRenameManagerEvents.end(); ++it)
    {
        if (it->pEvents)
        {
            it->pEvents->OnRegExCanceled();
        }
    }
}

void CSmartRenameManager::_OnRegExCompleted()
{
    CSRWExclusiveAutoLock lock(&m_lockEvents);

    for (std::vector<SMART_RENAME_MGR_EVENT>::iterator it = m_SmartRenameManagerEvents.begin(); it != m_SmartRenameManagerEvents.end(); ++it)
    {
        if (it->pEvents)
        {
            it->pEvents->OnRegExCompleted();
        }
    }
}

void CSmartRenameManager::_OnRenameStarted()
{
    CSRWExclusiveAutoLock lock(&m_lockEvents);

    for (std::vector<SMART_RENAME_MGR_EVENT>::iterator it = m_SmartRenameManagerEvents.begin(); it != m_SmartRenameManagerEvents.end(); ++it)
    {
        if (it->pEvents)
        {
            it->pEvents->OnRenameStarted();
        }
    }
}

void CSmartRenameManager::_OnRenameCompleted()
{
    CSRWExclusiveAutoLock lock(&m_lockEvents);

    for (std::vector<SMART_RENAME_MGR_EVENT>::iterator it = m_SmartRenameManagerEvents.begin(); it != m_SmartRenameManagerEvents.end(); ++it)
    {
        if (it->pEvents)
        {
            it->pEvents->OnRenameCompleted();
        }
    }
}

// Just in case setup a maximum folder depth
#define MAX_ENUM_DEPTH 300

bool CSmartRenameManager::_PathIsDotOrDotDot(_In_ PCWSTR path)
{
    return ((path[0] == L'.') && ((path[1] == L'\0') || ((path[1] == L'.') && (path[2] == L'\0'))));
}

bool CSmartRenameManager::_EnumeratePath(_In_ PCWSTR path, _In_ UINT depth)
{
    bool ret = false;
    if (depth < MAX_ENUM_DEPTH)
    {
        wchar_t searchPath[MAX_PATH] = { 0 };
        wchar_t parent[MAX_PATH] = { 0 };

        StringCchCopy(searchPath, ARRAYSIZE(searchPath), path);
        StringCchCopy(parent, ARRAYSIZE(parent), path);

        if (PathIsDirectory(searchPath))
        {
            // Add wildcard to end of folder path so we can enumerate its contents
            PathCchAddBackslash(searchPath, ARRAYSIZE(searchPath));
            StringCchCat(searchPath, ARRAYSIZE(searchPath), L"*");
        }
        else
        {
            PathCchRemoveFileSpec(parent, ARRAYSIZE(parent));
        }

        // TODO: Read from flags
        bool enumSubFolders = false;

        WIN32_FIND_DATA findData = { 0 };
        HANDLE findHandle = FindFirstFile(searchPath, &findData);
        if (findHandle != INVALID_HANDLE_VALUE)
        {
            do
            {
                if ((findData.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY) == FILE_ATTRIBUTE_DIRECTORY)
                {
                    // Ensure the directory is not . or ..
                    if (enumSubFolders && !_PathIsDotOrDotDot(findData.cFileName))
                    {
                        wchar_t pathSubFolder[MAX_PATH] = { 0 };
                        if (SUCCEEDED(PathCchCombine(pathSubFolder, ARRAYSIZE(pathSubFolder), parent, findData.cFileName)))
                        {
                            PathCchAddBackslash(pathSubFolder, ARRAYSIZE(pathSubFolder));
                            ret = _EnumeratePath(pathSubFolder, ++depth) || ret;
                        }
                    }
                }
                else
                {
                    wchar_t pathFile[MAX_PATH] = { 0 };
                    if (SUCCEEDED(PathCchCombine(pathFile, ARRAYSIZE(pathFile), parent, findData.cFileName)))
                    {
                        // Use the ISmartRenameItemFactory to create a new ISmartRenameItem
                        CComPtr<ISmartRenameItem> spsriNew;
                        if (SUCCEEDED(m_spItemFactory->Create(&spsriNew)))
                        {
                            if (SUCCEEDED(spsriNew->put_path(pathFile)))
                            {
                                // Add the item to the manager
                                ret = SUCCEEDED(AddItem(spsriNew)) || ret;
                            }
                        }
                    }
                }
            } while (FindNextFile(findHandle, &findData));

            FindClose(findHandle);
        }
    }

    return ret;
}

void CSmartRenameManager::_ClearEventHandlers()
{
    CSRWExclusiveAutoLock lock(&m_lockEvents);

    // Cleanup event handlers
    for (std::vector<SMART_RENAME_MGR_EVENT>::iterator it = m_SmartRenameManagerEvents.begin(); it != m_SmartRenameManagerEvents.end(); ++it)
    {
        it->cookie = 0;
        if (it->pEvents)
        {
            it->pEvents->Release();
            it->pEvents = nullptr;
        }
    }

    m_SmartRenameManagerEvents.clear();
}

void CSmartRenameManager::_ClearSmartRenameItems()
{
    CSRWExclusiveAutoLock lock(&m_lockItems);

    // Cleanup smart rename items
    for (std::vector<ISmartRenameItem*>::iterator it = m_smartRenameItems.begin(); it != m_smartRenameItems.end(); ++it)
    {
        ISmartRenameItem* pItem = *it;
        pItem->Release();
    }

    m_smartRenameItems.clear();
}

void CSmartRenameManager::_Cleanup()
{
    CloseHandle(m_startFileOpWorkerEvent);
    m_startFileOpWorkerEvent = nullptr;

    CloseHandle(m_startRegExWorkerEvent);
    m_startRegExWorkerEvent = nullptr;

    CloseHandle(m_cancelRegExWorkerEvent);
    m_cancelRegExWorkerEvent = nullptr;

    _ClearEventHandlers();
    _ClearSmartRenameItems();
}
