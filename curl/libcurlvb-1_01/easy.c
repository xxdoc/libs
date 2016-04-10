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
 * $Id: easy.c,v 1.2 2005/06/18 22:07:50 jeffreyphillips Exp $
 **************************************************************************/

#include <windows.h>
#include <curl/curl.h>
#include "form.h"
#include "seq.h"
#include "list.h"
#include "mem.h"
#include "table.h"
#include "slist.h"

// typedefs for callbacks into VB6
typedef size_t (__stdcall *VBCURL_WRITEFUNCTION)
    (char*, size_t, size_t, void*);
typedef int (__stdcall *VBCURL_PROGRESSFUNCTION)
    (void*, double, double, double, double);
typedef size_t (__stdcall *VBCURL_READFUNCTION)
    (char*, size_t, size_t, void*);
typedef size_t (__stdcall *VBCURL_HEADERFUNCTION)
    (char*, size_t, size_t, void*);
typedef int (__stdcall *VBCURL_DEBUGFUNCTION)
    (curl_infotype, char*, size_t, void*);
typedef int (__stdcall *VBCURL_SSL_CTX_FUNCTION)
    (void*, void*);
typedef int (__stdcall *VBCURL_IOCTLFUNCTION)
    (int, void*);

#define EASY_MAGIC  0x7F7F0000

typedef struct
{
    unsigned int            _magic;
    CURL*                   _curl;
    Seq_T                   _strings;
    VBCURL_PROGRESSFUNCTION _vbCurlProgressFunction;
    VBCURL_WRITEFUNCTION    _vbCurlWriteFunction;
    VBCURL_READFUNCTION     _vbCurlReadFunction;
    VBCURL_DEBUGFUNCTION    _vbCurlDebugFunction;
    VBCURL_HEADERFUNCTION   _vbCurlHeaderFunction;
    VBCURL_SSL_CTX_FUNCTION _vbCurlSslCtxFunction;
    VBCURL_IOCTLFUNCTION    _vbCurlIoCtlFunction;
    void*                   _vbCurlPrivateData;
    void*                   _vbCurlProgressData;
    void*                   _vbCurlWriteData;
    void*                   _vbCurlReadData;
    void*                   _vbCurlDebugData;
    void*                   _vbCurlHeaderData;
    void*                   _vbCurlSslCtxData;
    void*                   _vbCurlIoCtlData;
}   EASY_CONTEXT;

// table of easy handle -> EASY_CONTEXT pointers: not thread safe!
static Table_T g_contextTable;

void easy_create_context_table()
{
    g_contextTable = Table_new(16, NULL, NULL);
}

void easy_free_context_table()
{
    Table_free(&g_contextTable);
}

void* easy_get_inner(EASY_CONTEXT* context)
{
    if (context->_magic != EASY_MAGIC)
        return NULL;
    return context->_curl;
}

void* easy_get_outer(void* pvEasy)
{
    return Table_get(g_contextTable, pvEasy);
}

static size_t __cdecl write_function(char* szptr, size_t sz,
    size_t nmemb, void* pvData)
{
    EASY_CONTEXT* context = (EASY_CONTEXT*)pvData;
    if (context->_vbCurlWriteFunction)
    {
        return context->_vbCurlWriteFunction(szptr, sz,
            nmemb, context->_vbCurlWriteData);
    }
    else
        return sz * nmemb;
}

static size_t __cdecl read_function(void* szptr, size_t sz,
    size_t nmemb, void* pvData)
{
    EASY_CONTEXT* context = (EASY_CONTEXT*)pvData;
    if (context->_vbCurlReadFunction)
    {
        return context->_vbCurlReadFunction(szptr, sz,
            nmemb, context->_vbCurlReadData);
    }
    else
        return 0;
}

static int __cdecl progress_function(void* pvData, double dlTotal,
    double dlNow, double ulTotal, double ulNow)
{
    EASY_CONTEXT* context = (EASY_CONTEXT*)pvData;
    if (context->_vbCurlProgressFunction)
    {
        return context->_vbCurlProgressFunction(
            context->_vbCurlProgressData,
            dlTotal, dlNow, ulTotal, ulNow);
    }
    else
        return 0;
}

static size_t __cdecl header_function(char* szptr, size_t sz,
    size_t nmemb, void* pvData)
{
    EASY_CONTEXT* context = (EASY_CONTEXT*)pvData;
    if (context->_vbCurlHeaderFunction)
    {
        return context->_vbCurlHeaderFunction(szptr, sz,
            nmemb, context->_vbCurlHeaderData);
    }
    else
        return sz * nmemb;
}

static int debug_function(void* pvCurl, int infoType,
    char* szMsg, size_t msgSize, void* pvData)
{
    EASY_CONTEXT* context = (EASY_CONTEXT*)pvData;
    if (context->_vbCurlDebugFunction)
    {
        return context->_vbCurlDebugFunction(infoType,
            szMsg, msgSize, context->_vbCurlDebugData);
    }
    else
        return 0;
}

static int ssl_ctx_function(void* pvCurl, void* ctx, void* pvData)
{
    EASY_CONTEXT* context = (EASY_CONTEXT*)pvData;
    if (context->_vbCurlSslCtxFunction)
    {
        return context->_vbCurlSslCtxFunction(ctx,
            context->_vbCurlSslCtxData);
    }
    else
        return 0;
}

static int ioctl_function(void* pvCurl, int cmd, void* pvData)
{
    EASY_CONTEXT* context = (EASY_CONTEXT*)pvData;
    if (context->_vbCurlIoCtlFunction)
    {
        return context->_vbCurlIoCtlFunction(cmd,
            context->_vbCurlIoCtlData);
    }
    else
        return 0;
}

static EASY_CONTEXT* vbcurl_easy_init_impl(EASY_CONTEXT* ctxFrom)
{
    EASY_CONTEXT* context;
    CURL* curl;

    if (ctxFrom)
        curl = curl_easy_duphandle(ctxFrom->_curl);
    else
        curl = curl_easy_init();
    if (!curl)
        return NULL;
    context = (EASY_CONTEXT*)malloc(sizeof(EASY_CONTEXT));
    context->_magic = EASY_MAGIC;
    context->_curl = curl;

    // set our callback hooks now
    curl_easy_setopt(context->_curl, CURLOPT_WRITEFUNCTION,
        write_function);
    curl_easy_setopt(context->_curl, CURLOPT_WRITEDATA,
        context);
    curl_easy_setopt(context->_curl, CURLOPT_READFUNCTION,
        read_function);
    curl_easy_setopt(context->_curl, CURLOPT_READDATA,
        context);
    curl_easy_setopt(context->_curl, CURLOPT_PROGRESSFUNCTION,
        progress_function);
    curl_easy_setopt(context->_curl, CURLOPT_PROGRESSDATA,
        context);
    curl_easy_setopt(context->_curl, CURLOPT_HEADERFUNCTION,
        header_function);
    curl_easy_setopt(context->_curl, CURLOPT_HEADERDATA,
        context);
    curl_easy_setopt(context->_curl, CURLOPT_DEBUGFUNCTION,
        debug_function);
    curl_easy_setopt(context->_curl, CURLOPT_DEBUGDATA,
        context);
    curl_easy_setopt(context->_curl, CURLOPT_SSL_CTX_FUNCTION,
        ssl_ctx_function);
    curl_easy_setopt(context->_curl, CURLOPT_SSL_CTX_DATA,
        context);
    curl_easy_setopt(context->_curl, CURLOPT_IOCTLFUNCTION,
        ioctl_function);
    curl_easy_setopt(context->_curl, CURLOPT_IOCTLDATA,
        context);

    // storage for strings
    context->_strings = Seq_new(0);

    // initialize callbacks and data
    context->_vbCurlWriteFunction = NULL;
    context->_vbCurlProgressFunction = NULL;
    context->_vbCurlReadFunction = NULL;
    context->_vbCurlHeaderFunction = NULL;
    context->_vbCurlDebugFunction = NULL;
    context->_vbCurlIoCtlFunction = NULL;
    context->_vbCurlSslCtxFunction = NULL;
    context->_vbCurlWriteData = NULL;
    context->_vbCurlProgressData = NULL;
    context->_vbCurlReadData = NULL;
    context->_vbCurlHeaderData = NULL;
    context->_vbCurlDebugData = NULL;
    context->_vbCurlPrivateData = NULL;
    context->_vbCurlSslCtxData = NULL;
    context->_vbCurlIoCtlData = NULL;

    // update context table
    Table_put(g_contextTable, context->_curl, context);

    // return the context
    return context;
}

EASY_CONTEXT* __stdcall vbcurl_easy_init()
{
    return vbcurl_easy_init_impl(NULL);
}

static void vbcurl_easy_check_context(EASY_CONTEXT* context)
{
    if (context->_magic != EASY_MAGIC)
        RaiseException(0xE0000000, EXCEPTION_NONCONTINUABLE, 0, NULL);
}

void __stdcall vbcurl_easy_cleanup(EASY_CONTEXT* context)
{
    int i, count;
    
    // check the context
    vbcurl_easy_check_context(context);

    // remove context table entry
    Table_remove(g_contextTable, context->_curl);

    // and cleanup the easy handle
    curl_easy_cleanup(context->_curl);

    // and now free the strings
    count = Seq_length(context->_strings);
    for (i = 0; i < count; i++)
        free(Seq_get(context->_strings, i));
    Seq_free(&context->_strings);

    // and finally the context itself
    free(context);
}

int __stdcall vbcurl_easy_perform(EASY_CONTEXT* context)
{
    // check the context
    vbcurl_easy_check_context(context);
    return curl_easy_perform(context->_curl);
}

int __stdcall vbcurl_easy_setopt(EASY_CONTEXT* context,
    CURLoption option, VARIANT* pvValue)
{
    VARIANT vTemp;
    VariantInit(&vTemp);

    // check the context
    vbcurl_easy_check_context(context);

    // numeric cases
    if (option < CURLOPTTYPE_OBJECTPOINT)
    {
        // for now
        if (option == CURLOPT_TIMEVALUE)
            return CURLE_BAD_FUNCTION_ARGUMENT;
        VariantChangeType(&vTemp, pvValue, 0, VT_I4);
        return curl_easy_setopt(context->_curl, option,
            vTemp.lVal);
    }

    // object cases: the majority
    if (option < CURLOPTTYPE_FUNCTIONPOINT)
    {
        switch(option)
        {
            // various data items
            case CURLOPT_PRIVATE:
                VariantChangeType(&vTemp, pvValue, 0, VT_I4);
                context->_vbCurlPrivateData = (void*)vTemp.lVal;
                break;
            case CURLOPT_READDATA:
                VariantChangeType(&vTemp, pvValue, 0, VT_I4);
                context->_vbCurlReadData = (void*)vTemp.lVal;
                break;
            case CURLOPT_DEBUGDATA:
                VariantChangeType(&vTemp, pvValue, 0, VT_I4);
                context->_vbCurlDebugData = (void*)vTemp.lVal;
                break;
            case CURLOPT_HEADERDATA:
                VariantChangeType(&vTemp, pvValue, 0, VT_I4);
                context->_vbCurlHeaderData = (void*)vTemp.lVal;
                break;
            case CURLOPT_PROGRESSDATA:
                VariantChangeType(&vTemp, pvValue, 0, VT_I4);
                context->_vbCurlProgressData = (void*)vTemp.lVal;
                break;
            case CURLOPT_SSL_CTX_DATA:
                VariantChangeType(&vTemp, pvValue, 0, VT_I4);
                context->_vbCurlSslCtxData = (void*)vTemp.lVal;
                break;
            case CURLOPT_IOCTLDATA:
                VariantChangeType(&vTemp, pvValue, 0, VT_I4);
                context->_vbCurlIoCtlData = (void*)vTemp.lVal;
                break;
            case CURLOPT_WRITEDATA:
                VariantChangeType(&vTemp, pvValue, 0, VT_I4);
                context->_vbCurlWriteData = (void*)vTemp.lVal;
                break;

            // items that can't be set externally or are obsolete
            case CURLOPT_ERRORBUFFER:
            case CURLOPT_STDERR:
            case CURLOPT_SOURCE_HOST:
            case CURLOPT_SOURCE_PATH:
            case CURLOPT_PASV_HOST:
                return CURLE_BAD_FUNCTION_ARGUMENT;

            // singular case for share
            case CURLOPT_SHARE: 
                return CURLE_BAD_FUNCTION_ARGUMENT;

            // multipart HTTP post
            case CURLOPT_HTTPPOST:
            {
                void* pvPost;
                VariantChangeType(&vTemp, pvValue, 0, VT_I4);
                pvPost = form_get_post((void*)vTemp.lVal);
                if (pvPost)
                    return curl_easy_setopt(context->_curl, option, pvPost);
                else
                    return CURLE_BAD_FUNCTION_ARGUMENT;
            }

            // items requiring an Slist
            case CURLOPT_HTTPHEADER:
            case CURLOPT_PREQUOTE:
            case CURLOPT_QUOTE:
            case CURLOPT_POSTQUOTE:
            case CURLOPT_SOURCE_QUOTE:
            case CURLOPT_TELNETOPTIONS:
            case CURLOPT_HTTP200ALIASES:
            {
                void* pvSlist;
                VariantChangeType(&vTemp, pvValue, 0, VT_I4);
                pvSlist = slist_get_inner((void*)vTemp.lVal);
                if (pvSlist)
                {
                    return curl_easy_setopt(context->_curl,
                        option, pvSlist);
                }
                else
                    return CURLE_BAD_FUNCTION_ARGUMENT;
            }

            // string items
            default:
            {
                const wchar_t* pInStr;
                char* pOutStr;
                VariantChangeType(&vTemp, pvValue, 0, VT_BSTR);
                pInStr = (const wchar_t*)vTemp.bstrVal;
                pOutStr = (char*)malloc(wcslen(pInStr) + 1);
                wcstombs(pOutStr, pInStr, wcslen(pInStr) + 1);
                Seq_addhi(context->_strings, pOutStr);
                return curl_easy_setopt(context->_curl, option, pOutStr);
            }
        }
        return CURLE_OK;
    }

    // FUNCTIONPOINT args, for delegates
    if (option < CURLOPTTYPE_OFF_T)
    {
        VariantChangeType(&vTemp, pvValue, 0, VT_I4);
        switch(option)
        {
            case CURLOPT_PROGRESSFUNCTION:
                context->_vbCurlProgressFunction =
                    (VBCURL_PROGRESSFUNCTION)vTemp.lVal;
                break;

            case CURLOPT_WRITEFUNCTION:
                context->_vbCurlWriteFunction =
                    (VBCURL_WRITEFUNCTION)vTemp.lVal;
                break;

            case CURLOPT_READFUNCTION:
                context->_vbCurlReadFunction =
                    (VBCURL_READFUNCTION)vTemp.lVal;
                break;

            case CURLOPT_DEBUGFUNCTION:
                context->_vbCurlDebugFunction =
                    (VBCURL_DEBUGFUNCTION)vTemp.lVal;
                break;

            case CURLOPT_HEADERFUNCTION:
                context->_vbCurlHeaderFunction =
                    (VBCURL_HEADERFUNCTION)vTemp.lVal;
                break;

            case CURLOPT_SSL_CTX_FUNCTION:
                context->_vbCurlSslCtxFunction =
                    (VBCURL_SSL_CTX_FUNCTION)vTemp.lVal;
                break;

            case CURLOPT_IOCTLFUNCTION:
                context->_vbCurlIoCtlFunction =
                    (VBCURL_IOCTLFUNCTION)vTemp.lVal;
                break;

            default:
                return CURLE_BAD_FUNCTION_ARGUMENT;
        }

        return CURLE_OK;
    }

    // now we're into those 64-bit off_t dudes
    VariantChangeType(&vTemp, pvValue, 0, VT_I8);
    return curl_easy_setopt(context->_curl, option,
        vTemp.llVal);
}

CURLcode __stdcall vbcurl_easy_getinfo(EASY_CONTEXT* context,
    CURLINFO info, VARIANT* pv)
{
    VariantInit(pv);
    
    // check the context
    vbcurl_easy_check_context(context);

    if (info == CURLINFO_FILETIME)
    {
        int n = 0;
        struct tm* ptm;
        SYSTEMTIME st;
        curl_easy_getinfo(context->_curl, info, &n);
        if (n >= 0)
        {
            ptm = localtime(&n);
	        ZeroMemory(&st, sizeof(SYSTEMTIME));
            st.wYear = (WORD)(ptm->tm_year + 1900);
            st.wMonth = (WORD)(ptm->tm_mon + 1);
            st.wDay = (WORD)(ptm->tm_mday);
            st.wHour = (WORD)(ptm->tm_hour);
            st.wMinute = (WORD)(ptm->tm_min);
            st.wSecond = (WORD)(ptm->tm_sec);
            pv->vt = VT_DATE;
            SystemTimeToVariantTime(&st, &(pv->date));
        }
    }
    else if (info >= CURLINFO_STRING && info < CURLINFO_LONG)
    {
        char* p = NULL;
        curl_easy_getinfo(context->_curl, info, &p);
        pv->vt = VT_BSTR;
        if (p)
        {
            int n = (int)strlen(p);
            wchar_t* pwc = (wchar_t*)_alloca(sizeof(wchar_t) * (n + 1));
            mbstowcs(pwc, p, n + 1);
            pv->bstrVal = SysAllocString(pwc);
        }
        else
            pv->bstrVal = SysAllocString(L"");
    }
    else if (info >= CURLINFO_LONG && info < CURLINFO_DOUBLE)
    {
        int n = 0;
        curl_easy_getinfo(context->_curl, info, &n); 
        pv->vt = VT_I4;
        pv->lVal = n;
    }
    else if (info >= CURLINFO_DOUBLE && info < CURLINFO_SLIST)
    {
        double d = 0.0;
        curl_easy_getinfo(context->_curl, info, &d);
        pv->vt = VT_R8;
        pv->dblVal = d;
    }
    else if (info >= CURLINFO_SLIST)
    {
        struct curl_slist *psl, *ptemp;
        int ncount = 0;
        curl_easy_getinfo(context->_curl, info, &psl);
        ptemp = psl;
        while (ptemp != NULL) {
            ncount++;
            ptemp = ptemp->next;
        }
        pv->vt = VT_ARRAY | VT_BSTR;
        pv->parray = SafeArrayCreateVector(VT_BSTR, 0, ncount);
        if (ncount > 0)
        {
            LONG rgIndex = 0;
            ptemp = psl;
            for (rgIndex = 0; rgIndex < ncount; rgIndex++)
            {
                int nlen = (int)strlen(ptemp->data);
                wchar_t* pwc = (wchar_t*)_alloca(
                    sizeof(wchar_t)*(nlen + 1));
                mbstowcs(pwc, ptemp->data, nlen + 1);
                SafeArrayPutElement(pv->parray, &rgIndex,
                    SysAllocString(pwc));
                ptemp = ptemp->next;
            }
            curl_slist_free_all(psl);        
        }
    }
    return CURLE_OK;
}

void __stdcall vbcurl_easy_reset(EASY_CONTEXT* context)
{
    vbcurl_easy_check_context(context);
    curl_easy_reset(context->_curl);
}

EASY_CONTEXT* __stdcall vbcurl_easy_duphandle(EASY_CONTEXT* context)
{
    vbcurl_easy_check_context(context);
    return vbcurl_easy_init_impl(context);
}

BSTR __stdcall vbcurl_easy_strerror(CURLcode errorNum)
{
    const char* szError;
    szError = curl_easy_strerror(errorNum);
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

