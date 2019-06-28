#pragma once
#include "stdafx.h"

HRESULT EnumerateDataObject(_In_ IDataObject* pdo, _In_ ISmartRenameManager* psrm);
HRESULT GetIconIndexFromPath(_In_ PCWSTR path, _Out_ int* index);