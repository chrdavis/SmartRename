#include "stdafx.h"
#include "CppUnitTest.h"
#include <SmartRenameInterfaces.h>
#include <SmartRenameRegEx.h>

using namespace Microsoft::VisualStudio::CppUnitTestFramework;

namespace SmartRenameRegExTests
{        
    TEST_CLASS(SimpleTests)
    {
    public:
        
        TEST_METHOD(GeneralReplaceTest)
        {
            CComPtr<ISmartRenameRegEx> renameRegEx;
            Assert::IsTrue(CSmartRenameRegEx::CreateInstance(&renameRegEx) == S_OK);
            PWSTR result = nullptr;
            Assert::IsTrue(renameRegEx->Replace(L"foobar", L"foo", L"big", &result) == S_OK);
            Assert::IsTrue(wcscmp(result, L"bigbar") == 0);
        }

        TEST_METHOD(ReplaceNoMatch)
        {
            CComPtr<ISmartRenameRegEx> renameRegEx;
            Assert::IsTrue(CSmartRenameRegEx::CreateInstance(&renameRegEx) == S_OK);
            PWSTR result = nullptr;
            Assert::IsTrue(renameRegEx->Replace(L"foobar", L"notfound", L"big", &result) == S_OK);
            Assert::IsTrue(wcscmp(result, L"foobar") == 0);
        }
    };
}