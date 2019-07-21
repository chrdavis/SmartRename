#include "stdafx.h"
#include "CppUnitTest.h"
#include <SmartRenameInterfaces.h>
#include <SmartRenameManager.h>
#include <SmartRenameItem.h>
#include "MockSmartRenameItem.h"
#include "MockSmartRenameManagerEvents.h"

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
            CMockSmartRenameItem::CreateInstance(L"foo", L"parent", L"foo", 0, false, &item);
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
            CMockSmartRenameItem::CreateInstance(L"foo", L"parent", L"foo", 0, false, &item);
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

    };
}