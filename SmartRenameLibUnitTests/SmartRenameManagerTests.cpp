#include "stdafx.h"
#include "CppUnitTest.h"
#include <SmartRenameInterfaces.h>
#include <SmartRenameManager.h>
#include <SmartRenameItem.h>
#include "MockSmartRenameItem.h"
#include "MockSmartRenameManagerEvents.h"
#include "TestFileHelper.h"

using namespace Microsoft::VisualStudio::CppUnitTestFramework;

EXTERN_C IMAGE_DOS_HEADER __ImageBase;

#define HINST_THISCOMPONENT ((HINSTANCE)&__ImageBase)

HINSTANCE g_hInst = HINST_THISCOMPONENT;

namespace SmartRenameManagerTests
{
    TEST_CLASS(SimpleTests)
    {
    public:
        TEST_METHOD(CreateTest)
        {
            CComPtr<ISmartRenameManager> mgr;
            Assert::IsTrue(CSmartRenameManager::s_CreateInstance(&mgr) == S_OK);
        }

        TEST_METHOD(CreateAndShutdownTest)
        {
            CComPtr<ISmartRenameManager> mgr;
            Assert::IsTrue(CSmartRenameManager::s_CreateInstance(&mgr) == S_OK);
            Assert::IsTrue(mgr->Shutdown() == S_OK);
        }

        TEST_METHOD(AddItemTest)
        {
            CComPtr<ISmartRenameManager> mgr;
            Assert::IsTrue(CSmartRenameManager::s_CreateInstance(&mgr) == S_OK);
            CComPtr<ISmartRenameItem> item;
            CMockSmartRenameItem::CreateInstance(L"foo", L"foo", 0, false, &item);
            mgr->AddItem(item);
            Assert::IsTrue(mgr->Shutdown() == S_OK);
        }

        TEST_METHOD(VerifySmartManagerEvents)
        {
            CComPtr<ISmartRenameManager> mgr;
            Assert::IsTrue(CSmartRenameManager::s_CreateInstance(&mgr) == S_OK);
            CMockSmartRenameManagerEvents* mockMgrEvents = new CMockSmartRenameManagerEvents();
            CComPtr<ISmartRenameManagerEvents> mgrEvents;
            Assert::IsTrue(mockMgrEvents->QueryInterface(IID_PPV_ARGS(&mgrEvents)) == S_OK);
            DWORD cookie = 0;
            Assert::IsTrue(mgr->Advise(mgrEvents, &cookie) == S_OK);
            CComPtr<ISmartRenameItem> item;
            CMockSmartRenameItem::CreateInstance(L"foo", L"foo", 0, false, &item);
            int itemId = 0;
            Assert::IsTrue(item->get_id(&itemId) == S_OK);
            mgr->AddItem(item);

            // Verify the item we added is the same from the event
            Assert::IsTrue(mockMgrEvents->m_itemAdded != nullptr && mockMgrEvents->m_itemAdded == item);
            int eventItemId = 0;
            Assert::IsTrue(mockMgrEvents->m_itemAdded->get_id(&eventItemId) == S_OK);
            Assert::IsTrue(itemId == eventItemId);
            Assert::IsTrue(mgr->Shutdown() == S_OK);

            mockMgrEvents->Release();
        }

        TEST_METHOD(VerifySingleRename)
        {
            // TODO: Create a single item (in a temp directory) and verify rename works as expected
            // TODO: Also verify events
            CTestFileHelper testFileHelper;
            Assert::IsTrue(testFileHelper.AddFile(L"foo.txt"));

            CComPtr<ISmartRenameManager> mgr;
            Assert::IsTrue(CSmartRenameManager::s_CreateInstance(&mgr) == S_OK);
            CMockSmartRenameManagerEvents* mockMgrEvents = new CMockSmartRenameManagerEvents();
            CComPtr<ISmartRenameManagerEvents> mgrEvents;
            Assert::IsTrue(mockMgrEvents->QueryInterface(IID_PPV_ARGS(&mgrEvents)) == S_OK);
            DWORD cookie = 0;
            Assert::IsTrue(mgr->Advise(mgrEvents, &cookie) == S_OK);
            CComPtr<ISmartRenameItem> item;
            CMockSmartRenameItem::CreateInstance(testFileHelper.GetFullPath(L"foo.txt").c_str(), L"foo.txt", 0, false, &item);
            int itemId = 0;
            Assert::IsTrue(item->get_id(&itemId) == S_OK);
            mgr->AddItem(item);

            // Verify the item we added is the same from the event
            Assert::IsTrue(mockMgrEvents->m_itemAdded != nullptr && mockMgrEvents->m_itemAdded == item);
            int eventItemId = 0;
            Assert::IsTrue(mockMgrEvents->m_itemAdded->get_id(&eventItemId) == S_OK);
            Assert::IsTrue(itemId == eventItemId);

            // TODO: Setup match and replace parameters

            // TODO: Use this function in a generic way - data driven

            // Perform the rename
            Assert::IsTrue(mgr->Rename(0) == S_OK);


            Assert::IsTrue(mgr->Shutdown() == S_OK);

            mockMgrEvents->Release();
        }

        TEST_METHOD(VerifyMultiRename)
        {

        }

        TEST_METHOD(VerifyFilesOnlyRename)
        {

        }

        TEST_METHOD(VerifyFoldersOnlyRename)
        {

        }

        TEST_METHOD(VerifyFileNameOnlyRename)
        {

        }

        TEST_METHOD(VerifyFileExtensionOnlyRename)
        {

        }

        TEST_METHOD(VerifySubFoldersRename)
        {

        }

    };
}