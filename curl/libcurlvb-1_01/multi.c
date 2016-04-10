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
 * $Id: multi.c,v 1.1 2005/03/01 00:06:25 jeffreyphillips Exp $
 **************************************************************************/

#include <windows.h>
#include <curl/curl.h>
#include "easy.h"

#define MULTI_MAGIC 0x7F7F0003

typedef struct tagMultiContext
{
    unsigned int    _magic;
    CURLM*          _multi;
    fd_set          _readSet;
    fd_set          _writeSet;
    fd_set          _excSet;
    int             _maxFD;
}   MULTI_CONTEXT;

MULTI_CONTEXT* __stdcall vbcurl_multi_init()
{
    MULTI_CONTEXT* pmc = NULL;
    CURLM* pm = curl_multi_init();
    if (!pm)
        return NULL;
    pmc = (MULTI_CONTEXT*)malloc(sizeof(MULTI_CONTEXT));
    pmc->_magic = MULTI_MAGIC;
    pmc->_multi = pm;
    FD_ZERO(&pmc->_readSet);
    FD_ZERO(&pmc->_writeSet);
    FD_ZERO(&pmc->_excSet);
    pmc->_maxFD = 0;
    return pmc;
}

CURLMcode __stdcall vbcurl_multi_cleanup(MULTI_CONTEXT* pmc)
{
    CURLMcode ret;
    if (pmc->_magic != MULTI_MAGIC)
        RaiseException(0xE0000000, EXCEPTION_NONCONTINUABLE, 0, NULL);
    ret = curl_multi_cleanup(pmc->_multi);
    free(pmc);
    return ret;
}

CURLMcode __stdcall vbcurl_multi_add_handle(MULTI_CONTEXT* pmc,
    void* pvEasyContext)
{
    void* pvEasy;
    if (pmc->_magic != MULTI_MAGIC)
        RaiseException(0xE0000000, EXCEPTION_NONCONTINUABLE, 0, NULL);
    pvEasy = easy_get_inner(pvEasyContext);
    if (!pvEasy)
        return CURLM_BAD_EASY_HANDLE;
    return curl_multi_add_handle(pmc->_multi, pvEasy);
}

CURLMcode __stdcall vbcurl_multi_remove_handle(MULTI_CONTEXT* pmc,
    void* pvEasyContext)
{
    void* pvEasy;
    if (pmc->_magic != MULTI_MAGIC)
        RaiseException(0xE0000000, EXCEPTION_NONCONTINUABLE, 0, NULL);
    pvEasy = easy_get_inner(pvEasyContext);
    if (!pvEasy)
        return CURLM_BAD_EASY_HANDLE;
    return curl_multi_remove_handle(pmc->_multi, pvEasy);
}

CURLMcode __stdcall vbcurl_multi_perform(MULTI_CONTEXT* pmc,
    int* runningHandles)
{
    CURLMcode cl;
    if (pmc->_magic != MULTI_MAGIC)
        RaiseException(0xE0000000, EXCEPTION_NONCONTINUABLE, 0, NULL);
    cl = curl_multi_perform(pmc->_multi, runningHandles);
    return cl;
}

CURLMcode __stdcall vbcurl_multi_fdset(MULTI_CONTEXT* pmc)
{
    CURLMcode cl;
    if (pmc->_magic != MULTI_MAGIC)
        RaiseException(0xE0000000, EXCEPTION_NONCONTINUABLE, 0, NULL);
    cl = curl_multi_fdset(pmc->_multi, &pmc->_readSet,
        &pmc->_writeSet, &pmc->_excSet, &pmc->_maxFD);
    return cl;
}

int __stdcall vbcurl_multi_select(MULTI_CONTEXT* pmc,
    int timeoutMillis)
{
    int ret;
    struct timeval timeout;
    if (pmc->_magic != MULTI_MAGIC)
        RaiseException(0xE0000000, EXCEPTION_NONCONTINUABLE, 0, NULL);

    timeout.tv_sec  = timeoutMillis / 1000;
    timeout.tv_usec = (timeoutMillis % 1000) * 1000;
    ret = select(pmc->_maxFD + 1, &pmc->_readSet,
        &pmc->_writeSet, &pmc->_excSet, &timeout);
    return ret;
}

BSTR __stdcall vbcurl_multi_strerror(CURLMcode errorNum)
{
    const char* szError;
    szError = curl_multi_strerror(errorNum);
    if (szError)
    {
        int len = strlen(szError);
        wchar_t* wcError = (wchar_t*)_alloca(sizeof(wchar_t)*(len + 1));
        mbstowcs(wcError, szError, len + 1);
        return SysAllocString(wcError);
    }
    else
        return SysAllocString(L"");
}

int __stdcall vbcurl_multi_info_read(MULTI_CONTEXT* pmc,
    CURLMSG* pMsg, void** ppvEasy, CURLcode* pResult)
{
    int msgsInQueue;
    CURLMsg* pcMsg;

    if (pmc->_magic != MULTI_MAGIC)
        RaiseException(0xE0000000, EXCEPTION_NONCONTINUABLE, 0, NULL);

    pcMsg = curl_multi_info_read(pmc->_multi,
        &msgsInQueue);
    if (pcMsg) {
        *pMsg = pcMsg->msg;
        // get EASY_CONTEXT value from easy handle!
        *ppvEasy = easy_get_outer(pcMsg->easy_handle);
        *pResult = pcMsg->data.result;
    }
    else {
        *pMsg = 0;
        *ppvEasy = 0;
        *pResult = 0;
    }
    return msgsInQueue;
}

