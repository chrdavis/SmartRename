#pragma once

#include <filesystem>
#include <string>
#include <windows.h>


class CTestFileHelper
{
public:
    CTestFileHelper();
    ~CTestFileHelper();

    bool AddFile(_In_ PCWSTR path);
    bool AddFolder(_In_ PCWSTR path);
    const std::filesystem::path GetTempDirectory() { return _tempDirectory; }
    bool PathExists(_In_ PCWSTR path);
    std::filesystem::path GetFullPath(_In_ PCWSTR path);

private:
    bool _CreateTempDirectory();
    void _DeleteTempDirectory();

    std::filesystem::path _tempDirectory;
};