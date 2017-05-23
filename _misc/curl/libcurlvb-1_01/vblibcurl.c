/***************************************************************************
 *
 * Project: libcurl.vb
 *
 * Copyright (c) 2005 Jeff Phillips (jeff@jeffp.net)
 *
 * This software is licensed as described in the file COPYING, which you
 * should have received as part of this distribution.
 *
 * You may opt to use, copy, modify, merge, publish, distribute and/or sell
 * copies of this Software, and permit persons to whom the Software is
 * furnished to do so, under the terms of the COPYING file.
 *
 * This software is distributed on an "AS IS" basis, WITHOUT WARRANTY OF
 * ANY KIND, either express or implied.
 *
 * $Id: vblibcurl.c,v 1.1 2005/03/01 00:06:25 jeffreyphillips Exp $
 **************************************************************************/

#include <windows.h>
#include <curl/curl.h>
#include "easy.h"

BOOL WINAPI DllMain(HINSTANCE hInstance, DWORD dwReason, LPVOID lpvReserved)
{
    switch(dwReason)
    {
        case DLL_PROCESS_ATTACH:
            curl_global_init(CURL_GLOBAL_ALL);
            easy_create_context_table();

            break;

        case DLL_PROCESS_DETACH:
            easy_free_context_table();
            curl_global_cleanup();
            break;

        case DLL_THREAD_ATTACH:
        case DLL_THREAD_DETACH:
            break;
    }
    return TRUE;
}

BSTR __stdcall vbcurl_string_escape(BSTR bsInput, int length)
{
    char *szInput, *szOutput;
    int len;
    wchar_t* wcOutput;

    len = wcslen((wchar_t*)bsInput);
    szInput = _alloca(len + 1);
    wcstombs(szInput, (wchar_t*)bsInput, len + 1);
    szOutput = curl_escape(szInput, length);
    len = strlen(szOutput);
    wcOutput = (wchar_t*)_alloca(sizeof(wchar_t) * (len + 1));
    mbstowcs(wcOutput, szOutput, len + 1);
    curl_free(szOutput);
    return SysAllocString(wcOutput);    
}

BSTR __stdcall vbcurl_string_unescape(BSTR bsInput, int length)
{
    char *szInput, *szOutput;
    int len;
    wchar_t* wcOutput;

    len = wcslen((wchar_t*)bsInput);
    szInput = _alloca(len + 1);
    wcstombs(szInput, (wchar_t*)bsInput, len + 1);
    szOutput = curl_unescape(szInput, length);
    len = strlen(szOutput);
    wcOutput = (wchar_t*)_alloca(sizeof(wchar_t) * (len + 1));
    mbstowcs(wcOutput, szOutput, len + 1);
    curl_free(szOutput);
    return SysAllocString(wcOutput);    
}

BSTR __stdcall vbcurl_string_version()
{
    wchar_t* wcVersion;
    int len;
    char* szVersion = curl_version();
    len = strlen(szVersion);
    wcVersion = (wchar_t*)_alloca(sizeof(wchar_t) * (len + 1));
    mbstowcs(wcVersion, szVersion, len + 1);
    return SysAllocString(wcVersion);
}
