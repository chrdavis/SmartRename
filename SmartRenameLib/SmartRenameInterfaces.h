#pragma once
#include "stdafx.h"

enum SmartRenameFlags
{
    CaseSensitive = 0x1,
    MatchAllOccurences = 0x2,
    UseRegularExpressions = 0x4,
    EnumerateItems = 0x8,
    ExcludeFiles = 0x10,
    ExcludeFolders = 0x20,
    ExcludeSubfolders = 0x40
};

interface __declspec(uuid("3ECBA62B-E0F0-4472-AA2E-DEE7A1AA46B9")) ISmartRenameRegExEvents : public IUnknown
{
public:
    IFACEMETHOD(OnSearchTermChanged)(_In_ PCWSTR searchTerm) = 0;
    IFACEMETHOD(OnReplaceTermChanged)(_In_ PCWSTR replaceTerm) = 0;
    IFACEMETHOD(OnFlagsChanged)(_In_ DWORD flags) = 0;
};

interface __declspec(uuid("E3ED45B5-9CE0-47E2-A595-67EB950B9B72")) ISmartRenameRegEx : public IUnknown
{
public:
    IFACEMETHOD(Advise)(_In_ ISmartRenameRegExEvents* regExEvents, _Out_ DWORD* cookie) = 0;
    IFACEMETHOD(UnAdvise)(_In_ DWORD cookie) = 0;
    IFACEMETHOD(get_searchTerm)(_Outptr_ PWSTR* searchTerm) = 0;
    IFACEMETHOD(put_searchTerm)(_In_ PCWSTR searchTerm) = 0;
    IFACEMETHOD(get_replaceTerm)(_Outptr_ PWSTR* replaceTerm) = 0;
    IFACEMETHOD(put_replaceTerm)(_In_ PCWSTR replaceTerm) = 0;
    IFACEMETHOD(get_flags)(_Out_ DWORD* flags) = 0;
    IFACEMETHOD(put_flags)(_In_ DWORD flags) = 0;
    IFACEMETHOD(Replace)(_In_ PCWSTR source, _Outptr_ PWSTR* result) = 0;
};

interface __declspec(uuid("C7F59201-4DE1-4855-A3A2-26FC3279C8A5")) ISmartRenameItem : public IUnknown
{
public:
    IFACEMETHOD(get_path)(_Outptr_ PWSTR* path) = 0;
    IFACEMETHOD(put_path)(_In_ PCWSTR path) = 0;
    IFACEMETHOD(get_originalName)(_Outptr_ PWSTR* originalName) = 0;
    IFACEMETHOD(get_newName)(_Outptr_ PWSTR* newName) = 0;
    IFACEMETHOD(put_newName)(_In_ PCWSTR newName) = 0;
    IFACEMETHOD(get_isFolder)(_Out_ bool* isFolder) = 0;
    IFACEMETHOD(get_isSubFolderContent)(_Out_ bool* isSubFolderContent) = 0;
    IFACEMETHOD(put_isSubFolderContent)(_In_ bool isSubFolderContent) = 0;
    IFACEMETHOD(get_isDirty)(_Out_ bool* isDirty) = 0;
    IFACEMETHOD(get_shouldRename)(_Out_ bool* shouldRename) = 0;
    IFACEMETHOD(put_shouldRename)(_In_ bool shouldRename) = 0;
    IFACEMETHOD(get_id)(_Out_ int *id) = 0;
    IFACEMETHOD(Reset)() = 0;
};

interface __declspec(uuid("{26CBFFD9-13B3-424E-BAC9-D12B0539149C}")) ISmartRenameItemFactory : public IUnknown
{
public:
    IFACEMETHOD(Create)(_COM_Outptr_ ISmartRenameItem** ppItem) = 0;
};

interface __declspec(uuid("87FC43F9-7634-43D9-99A5-20876AFCE4AD")) ISmartRenameManagerEvents : public IUnknown
{
public:
    IFACEMETHOD(OnItemAdded)(_In_ ISmartRenameItem* renameItem) = 0;
    IFACEMETHOD(OnUpdate)(_In_ ISmartRenameItem* renameItem) = 0;
    IFACEMETHOD(OnError)(_In_ ISmartRenameItem* renameItem) = 0;
    IFACEMETHOD(OnRegExStarted)() = 0;
    IFACEMETHOD(OnRegExCanceled)() = 0;
    IFACEMETHOD(OnRegExCompleted)() = 0;
    IFACEMETHOD(OnRenameStarted)() = 0;
    IFACEMETHOD(OnRenameCompleted)() = 0;
};

interface __declspec(uuid("001BBD88-53D2-4FA6-95D2-F9A9FA4F9F70")) ISmartRenameManager : public IUnknown
{
public:
    IFACEMETHOD(Advise)(_In_ ISmartRenameManagerEvents* renameManagerEvent, _Out_ DWORD* cookie) = 0;
    IFACEMETHOD(UnAdvise)(_In_ DWORD cookie) = 0;
    IFACEMETHOD(Start)() = 0;
    IFACEMETHOD(Stop)() = 0;
    IFACEMETHOD(Reset)() = 0;
    IFACEMETHOD(Shutdown)() = 0;
    IFACEMETHOD(AddItem)(_In_ ISmartRenameItem* pItem) = 0;
    IFACEMETHOD(GetItem)(_In_ UINT index, _COM_Outptr_ ISmartRenameItem** ppItem) = 0;
    IFACEMETHOD(GetItemCount)(_Out_ UINT* count) = 0;
    IFACEMETHOD(get_smartRenameRegEx)(_COM_Outptr_ ISmartRenameRegEx** ppRegEx) = 0;
    IFACEMETHOD(put_smartRenameRegEx)(_In_ ISmartRenameRegEx* pRegEx) = 0;
    IFACEMETHOD(get_smartRenameItemFactory)(_COM_Outptr_ ISmartRenameItemFactory** ppItemFactory) = 0;
    IFACEMETHOD(put_smartRenameItemFactory)(_In_ ISmartRenameItemFactory* pItemFactory) = 0;
};

interface __declspec(uuid("E6679DEB-460D-42C1-A7A8-E25897061C99")) ISmartRenameUI : public IUnknown
{
public:
    IFACEMETHOD(Show)() = 0;
    IFACEMETHOD(Close)() = 0;
    IFACEMETHOD(Update)() = 0;
};

