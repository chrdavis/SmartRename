#include "stdafx.h"
#include "CppUnitTest.h"
#include <SmartRenameInterfaces.h>
#include <SmartRenameRegEx.h>

using namespace Microsoft::VisualStudio::CppUnitTestFramework;

namespace SmartRenameRegExTests
{        
    TEST_CLASS(UnitTest1)
    {
    public:
        
        TEST_METHOD(GeneralReplaceTest)
        {
            CComPtr<ISmartRenameRegEx> renameRegEx;
            Assert::IsTrue(CSmartRenameRegEx::CreateInstance(&renameRegEx) == S_OK);
            PWSTR result = nullptr;
            Assert::IsTrue(renameRegEx->Replace(L"foobar", L"foo", L"big", &result) == S_OK);
            Assert::IsTrue(wcscmp(result, L"bigbar") == 0);
            //Logger::WriteMessage(L"Result was ");
            //Logger::WriteMessage(result);
        }
    };
}