#pragma once
#include "stdafx.h"

interface __declspec(uuid("C7F59201-4DE1-4855-A3A2-26FC3279C8A5")) ISmartRenameItem : public IUnknown
{
public:
    IFACEMETHOD(GetOriginalName)(_Outptr_ PWSTR *originalName) = 0;
    IFACEMETHOD(SetNewName)(_In_ PCWSTR newName) = 0;
    IFACEMETHOD(GetNewName)(_Outptr_ PWSTR *newName) = 0;
    IFACEMETHOD(GetIsFolder)(_Out_ bool *isFolder) = 0;
    IFACEMETHOD(GetIsDirty)(_Out_ bool *isDirty) = 0;
    IFACEMETHOD(GetId)(_Out_ int *id) = 0;
    IFACEMETHOD(Reset)() = 0;
};

interface __declspec(uuid("87FC43F9-7634-43D9-99A5-20876AFCE4AD")) ISmartRenameModelEvent : public IUnknown
{
public:
    IFACEMETHOD(OnUpdate)(_Outptr_ ISmartRenameItem **renameItem) = 0;
    IFACEMETHOD(OnError)(_Outptr_ ISmartRenameItem **renameItem) = 0;
    IFACEMETHOD(OnComplete)() = 0;
};

interface __declspec(uuid("001BBD88-53D2-4FA6-95D2-F9A9FA4F9F70")) ISmartRenameModel : public IUnknown
{
public:
    IFACEMETHOD(Advise)(_In_ ISmartRenameModelEvent *renameModelEvent, _Out_ DWORD *cookie) = 0;
    IFACEMETHOD(UnAdvise)(_In_ DWORD cookie) = 0;
    IFACEMETHOD(Start)() = 0;
    IFACEMETHOD(Wait)(_In_ DWORD timeoutMS) = 0;
    IFACEMETHOD(Stop)() = 0;
    IFACEMETHOD(Reset)() = 0;
};

interface __declspec(uuid("E6679DEB-460D-42C1-A7A8-E25897061C99")) ISmartRenameView : public IUnknown
{
public:
    IFACEMETHOD(Update)() = 0;
};

interface __declspec(uuid("E3ED45B5-9CE0-47E2-A595-67EB950B9B72")) ISmartRenameRegEx : public IUnknown
{
public:
    IFACEMETHOD(Replace)(_In_ PCWSTR source, _In_ PCWSTR match, _In_ PCWSTR replace, _Outptr_ PWSTR *result) = 0;
};
