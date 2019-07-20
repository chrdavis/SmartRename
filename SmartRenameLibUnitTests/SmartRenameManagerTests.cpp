#include "stdafx.h"
#include "CppUnitTest.h"
#include <SmartRenameInterfaces.h>
#include <SmartRenameManager.h>
#include <SmartRenameItem.h>
#include "MockSmartRenameItem.h"

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

    };
}