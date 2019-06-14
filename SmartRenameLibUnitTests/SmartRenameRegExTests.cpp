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
            Assert::IsTrue(CSmartRenameRegEx::s_CreateInstance(&renameRegEx) == S_OK);
            PWSTR result = nullptr;
            Assert::IsTrue(renameRegEx->put_searchTerm(L"foo") == S_OK);
            Assert::IsTrue(renameRegEx->put_replaceTerm(L"big") == S_OK);
            Assert::IsTrue(renameRegEx->Replace(L"foobar", &result) == S_OK);
            Assert::IsTrue(wcscmp(result, L"bigbar") == 0);
        }

        TEST_METHOD(ReplaceNoMatch)
        {
            CComPtr<ISmartRenameRegEx> renameRegEx;
            Assert::IsTrue(CSmartRenameRegEx::s_CreateInstance(&renameRegEx) == S_OK);
            PWSTR result = nullptr;
            Assert::IsTrue(renameRegEx->put_searchTerm(L"notfound") == S_OK);
            Assert::IsTrue(renameRegEx->put_replaceTerm(L"big") == S_OK);
            Assert::IsTrue(renameRegEx->Replace(L"foobar", &result) == S_OK);
            Assert::IsTrue(wcscmp(result, L"foobar") == 0);
        }

        TEST_METHOD(ReplaceNoSearchOrReplaceTerm)
        {
            CComPtr<ISmartRenameRegEx> renameRegEx;
            Assert::IsTrue(CSmartRenameRegEx::s_CreateInstance(&renameRegEx) == S_OK);
            PWSTR result = nullptr;
            Assert::IsTrue(renameRegEx->Replace(L"foobar", &result) != S_OK);
            Assert::IsTrue(result == nullptr);
        }

        TEST_METHOD(ReplaceNoReplaceTerm)
        {
            CComPtr<ISmartRenameRegEx> renameRegEx;
            Assert::IsTrue(CSmartRenameRegEx::s_CreateInstance(&renameRegEx) == S_OK);
            PWSTR result = nullptr;
            Assert::IsTrue(renameRegEx->put_searchTerm(L"foo") == S_OK);
            Assert::IsTrue(renameRegEx->Replace(L"foobar", &result) == S_OK);
            Assert::IsTrue(wcscmp(result, L"bar") == 0);
        }

        TEST_METHOD(ReplaceEmptyStringReplaceTerm)
        {
            CComPtr<ISmartRenameRegEx> renameRegEx;
            Assert::IsTrue(CSmartRenameRegEx::s_CreateInstance(&renameRegEx) == S_OK);
            PWSTR result = nullptr;
            Assert::IsTrue(renameRegEx->put_searchTerm(L"foo") == S_OK);
            Assert::IsTrue(renameRegEx->put_replaceTerm(L"") == S_OK);
            Assert::IsTrue(renameRegEx->Replace(L"foobar", &result) == S_OK);
            Assert::IsTrue(wcscmp(result, L"bar") == 0);
        }
    };
}