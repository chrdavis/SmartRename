#include "stdafx.h"
#include "Helpers.h"
#include <ShlGuid.h>

HRESULT GetIconIndexFromPath(_In_ PCWSTR path, _Out_ int* index)
{
    *index = 0;

    HRESULT hr = E_FAIL;

    SHFILEINFO shFileInfo = { 0 };

    if (!PathIsRelative(path))
    {
        DWORD attrib = GetFileAttributes(path);
        HIMAGELIST himl = (HIMAGELIST)SHGetFileInfo(path, attrib, &shFileInfo, sizeof(shFileInfo), (SHGFI_SYSICONINDEX | SHGFI_SMALLICON | SHGFI_USEFILEATTRIBUTES));
        if (himl)
        {
            *index = shFileInfo.iIcon;
            // We shouldn't free the HIMAGELIST.
            hr = S_OK;
        }
    }

    return hr;
}

HBITMAP CreateBitmapFromIcon(_In_ HICON hIcon, _In_opt_ UINT width, _In_opt_ UINT height)
{
    HBITMAP hBitmapResult = NULL;

    // Create compatible DC
    HDC hDC = CreateCompatibleDC(NULL);
    if (hDC != NULL)
    {
        // Get bitmap rectangle size
        RECT rc = { 0 };
        rc.left = 0;
        rc.right = (width != 0) ? width : GetSystemMetrics(SM_CXSMICON);
        rc.top = 0;
        rc.bottom = (height != 0) ? height : GetSystemMetrics(SM_CYSMICON);

        // Create bitmap compatible with DC
        BITMAPINFO BitmapInfo;
        ZeroMemory(&BitmapInfo, sizeof(BITMAPINFO));

        BitmapInfo.bmiHeader.biSize = sizeof(BITMAPINFOHEADER);
        BitmapInfo.bmiHeader.biWidth = rc.right;
        BitmapInfo.bmiHeader.biHeight = rc.bottom;
        BitmapInfo.bmiHeader.biPlanes = 1;
        BitmapInfo.bmiHeader.biBitCount = 32;
        BitmapInfo.bmiHeader.biCompression = BI_RGB;

        HDC hDCBitmap = GetDC(NULL);

        HBITMAP hBitmap = CreateDIBSection(hDCBitmap, &BitmapInfo, DIB_RGB_COLORS, NULL, NULL, 0);

        ReleaseDC(NULL, hDCBitmap);

        if (hBitmap != NULL)
        {
            // Select bitmap into DC
            HBITMAP hBitmapOld = (HBITMAP)SelectObject(hDC, hBitmap);
            if (hBitmapOld != NULL)
            {
                // Draw icon into DC
                if (DrawIconEx(hDC, 0, 0, hIcon, rc.right, rc.bottom, 0, NULL, DI_NORMAL))
                {
                    // Restore original bitmap in DC
                    hBitmapResult = (HBITMAP)SelectObject(hDC, hBitmapOld);
                    hBitmapOld = NULL;
                    hBitmap = NULL;
                }

                if (hBitmapOld != NULL)
                {
                    SelectObject(hDC, hBitmapOld);
                }
            }

            if (hBitmap != NULL)
            {
                DeleteObject(hBitmap);
            }
        }

        DeleteDC(hDC);
    }
  
    return hBitmapResult;
}

HWND CreateMsgWindow(_In_ HINSTANCE hInst, _In_ WNDPROC pfnWndProc, _In_ void* p)
{
    WNDCLASS wc = { 0 };
    PWSTR wndClassName = L"MsgWindow";

    wc.lpfnWndProc = DefWindowProc;
    wc.cbWndExtra = sizeof(void*);
    wc.hInstance = hInst;
    wc.hbrBackground = (HBRUSH)(COLOR_BTNFACE + 1);
    wc.lpszClassName = wndClassName;

    RegisterClass(&wc);

    HWND hwnd = CreateWindowEx(
        0, wndClassName, nullptr, 0,
        0, 0, 0, 0, HWND_MESSAGE,
        0, hInst, nullptr);
    if (hwnd)
    {
        SetWindowLongPtr(hwnd, 0, (LONG_PTR)p);
        if (pfnWndProc)
        {
            SetWindowLongPtr(hwnd, GWLP_WNDPROC, (LONG_PTR)pfnWndProc);
        }
    }

    return hwnd;
}

HRESULT GetShellItemArrayFromUnknown(_In_ IUnknown* punk, _COM_Outptr_ IShellItemArray** ppsia)
{
    *ppsia = nullptr;
    CComPtr<IDataObject> spdo;
    HRESULT hr;
    if (SUCCEEDED(punk->QueryInterface(IID_PPV_ARGS(&spdo))))
    {
        hr = SHCreateShellItemArrayFromDataObject(spdo, IID_PPV_ARGS(ppsia));
    }
    else
    {
        hr = punk->QueryInterface(IID_PPV_ARGS(ppsia));
    }

    return hr;
}

BOOL GetEnumeratedFileName(__out_ecount(cchMax) PWSTR pszUniqueName, UINT cchMax,
    __in PCWSTR pszTemplate, __in_opt PCWSTR pszDir, unsigned long ulMinLong,
    __inout unsigned long* pulNumUsed)
{
    PWSTR pszName = nullptr;
    HRESULT hr = S_OK;
    BOOL fRet = FALSE;
    int cchDir = 0;

    if (0 != cchMax && pszUniqueName)
    {
        *pszUniqueName = 0;
        if (pszDir)
        {
            hr = StringCchCopy(pszUniqueName, cchMax, pszDir);
            if (SUCCEEDED(hr))
            {
                size_t cch = wcslen(pszUniqueName);
                pszName = pszUniqueName + cch;
                size_t cchRemaining = cchMax - cch;
                if (cch && (L'\\' != *(pszName - 1)))
                {
                    hr = StringCchCopyEx(pszName, cchRemaining, L"\\", &pszName, &cchRemaining, 0);
                }

                if (SUCCEEDED(hr))
                {
                    cchDir = lstrlen(pszDir);
                }
            }
        }
        else
        {
            cchDir = 0;
            pszName = pszUniqueName;
        }
    }
    else
    {
        hr = E_INVALIDARG;
    }

    int cchTmp = 0;
    int cchStem = 0;
    PCWSTR pszStem = nullptr;
    PCWSTR pszRest = nullptr;
    wchar_t szFormat[MAX_PATH] = { 0 };

    if (SUCCEEDED(hr))
    {
        pszStem = pszTemplate;

        pszRest = StrChr(pszTemplate, L'(');
        while (pszRest)
        {
            PCWSTR pszEndUniq = CharNext(pszRest);
            while (*pszEndUniq && *pszEndUniq >= L'0' && *pszEndUniq <= L'9')
            {
                pszEndUniq++;
            }

            if (*pszEndUniq == L')')
            {
                break;
            }

            pszRest = StrChr(CharNext(pszRest), L'(');
        }

        if (!pszRest)
        {
            pszRest = PathFindExtension(pszTemplate);
            cchStem = (int)(pszRest - pszTemplate);

            hr = StringCchCopy(szFormat, ARRAYSIZE(szFormat), L" (%lu)");
        }
        else
        {
            pszRest++;

            cchStem = (int)(pszRest - pszTemplate);

            while (*pszRest && *pszRest >= L'0' && *pszRest <= L'9')
            {
                pszRest++;
            }

            hr = StringCchCopy(szFormat, ARRAYSIZE(szFormat), L"%lu");
        }
    }

    unsigned long ulMax = 0;
    unsigned long ulMin = 0;
    if (SUCCEEDED(hr))
    {
        int cchFormat = lstrlen(szFormat);
        if (cchFormat < 3)
        {
            *pszUniqueName = L'\0';
            return FALSE;
        }
        ulMin = ulMinLong;
        cchTmp = cchMax - cchDir - cchStem - (cchFormat - 3);
        switch (cchTmp)
        {
        case 1:
            ulMax = 10;
            break;
        case 2:
            ulMax = 100;
            break;
        case 3:
            ulMax = 1000;
            break;
        case 4:
            ulMax = 10000;
            break;
        case 5:
            ulMax = 100000;
            break;
        default:
            if (cchTmp <= 0)
            {
                ulMax = ulMin;
            }
            else
            {
                ulMax = 1000000;
            }
            break;
        }
    }

    if (SUCCEEDED(hr))
    {
        hr = StringCchCopyN(pszName, pszUniqueName + cchMax - pszName, pszStem, cchStem);
        if (SUCCEEDED(hr))
        {
            PWSTR pszDigit = pszName + cchStem;

            for (unsigned long ul = ulMin; ((ul < ulMax) && (!fRet)); ul++)
            {
                wchar_t szTemp[MAX_PATH] = { 0 };
                hr = StringCchPrintf(szTemp, ARRAYSIZE(szTemp), szFormat, ul);
                if (SUCCEEDED(hr))
                {
                    hr = StringCchCat(szTemp, ARRAYSIZE(szTemp), pszRest);
                    if (SUCCEEDED(hr))
                    {
                        hr = StringCchCopy(pszDigit, pszUniqueName + cchMax - pszDigit, szTemp);
                        if (SUCCEEDED(hr))
                        {
                            if (!PathFileExists(pszUniqueName))
                            {
                                (*pulNumUsed) = ul;
                                fRet = TRUE;
                            }
                        }
                    }
                }
            }
        }
    }

    if (!fRet)
    {
        *pszUniqueName = L'\0';
    }

    return fRet;
}
