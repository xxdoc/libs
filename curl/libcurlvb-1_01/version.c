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
 * $Id: version.c,v 1.1 2005/03/01 00:06:25 jeffreyphillips Exp $
 **************************************************************************/

#include <windows.h>
#include <curl/curl.h>

// ensure passed structures are CURL_VERSION pointers
#define VERSION_MAGIC           0x7F7F0005

// offset to various items in curl_version_info_data struct
#define OFFSET_AGE              0
#define OFFSET_VERSION          4
#define OFFSET_VERSION_NUM      8
#define OFFSET_HOST             12
#define OFFSET_FEATURES         16
#define OFFSET_SSL_VERSION      20
#define OFFSET_SSL_VERSION_NUM  24
#define OFFSET_LIBZ_VERSION     28
#define OFFSET_PROTOCOLS        32
#define OFFSET_ARES_VERSION     36
#define OFFSET_ARES_VERSION_NUM 40
#define OFFSET_LIBIDN_VERSION   44

// for not applicable items
#define NOT_APPLICABLE          L"n.a."

typedef struct tagCurlVersion
{
    unsigned int            _magic;
    curl_version_info_data* _pcvData;
}   CURL_VERSION;
static CURL_VERSION g_curlVersion;

CURL_VERSION* __stdcall vbcurl_version_info(CURLversion type)
{
    curl_version_info_data* pd = curl_version_info(type);
    if (pd) {
        g_curlVersion._magic = VERSION_MAGIC;
        g_curlVersion._pcvData = pd;
        return &g_curlVersion;
    }
    else
        return NULL;    
}

static int vbcurl_version_int(CURL_VERSION* pv, int offset)
{
    int *q = NULL;
    if (pv->_magic != VERSION_MAGIC)
        RaiseException(0xE0000000, EXCEPTION_NONCONTINUABLE, 0, NULL);
    q = (int*)pv->_pcvData;
    q += offset / sizeof(int);
    return *q;
}

static BSTR vbcurl_version_bstr(CURL_VERSION* pv, int offset)
{
    if (pv->_magic != VERSION_MAGIC)
        RaiseException(0xE0000000, EXCEPTION_NONCONTINUABLE, 0, NULL);
    // trick the compiler into not issuing a warning on the
    // (char**)q cast
    {
        char* q = &((char*)pv->_pcvData)[offset];
        char** qq = (char**)q;
        if (qq && *qq)
        {
            int len = (int)strlen(*qq);
            wchar_t* pwc = (wchar_t*)_alloca(sizeof(wchar_t) * (len + 1));
            mbstowcs(pwc, *qq, len + 1);
            return SysAllocString((OLECHAR*)pwc);
        }
        return SysAllocString((OLECHAR*)L"");
    }
}

static int vbcurl_num_protocols(CURL_VERSION* pv, int offset)
{
    int nProtocols = 0;
    char* q = &((char*)pv->_pcvData)[offset];
    char*** qq = (char***)q;
    char** rr = *qq;
    while(*rr++)
        nProtocols++;
    return nProtocols;
}

static BSTR vbcurl_protocol_string(CURL_VERSION* pv,
    int offset, int nProt)
{
    char* q = &((char*)pv->_pcvData)[offset];
    char*** qq = (char***)q;
    char** rr = *qq;
    if (rr[nProt])
    {
        int len = (int)strlen(rr[nProt]);
        wchar_t* pwc = (wchar_t*)_alloca(sizeof(wchar_t) * (len + 1));
        mbstowcs(pwc, rr[nProt], len + 1);
        return SysAllocString((OLECHAR*)pwc);
    }    
    return SysAllocString(L"");
}

CURLversion __stdcall vbcurl_version_age(CURL_VERSION* pv)
{
    return (CURLversion)vbcurl_version_int(pv, OFFSET_AGE);
}

BSTR __stdcall vbcurl_version_string(CURL_VERSION* pv)
{
    return vbcurl_version_bstr(pv, OFFSET_VERSION);
}

int __stdcall vbcurl_version_num(CURL_VERSION* pv)
{
    return vbcurl_version_int(pv, OFFSET_VERSION_NUM);
}

BSTR __stdcall vbcurl_version_host(CURL_VERSION* pv)
{
    return vbcurl_version_bstr(pv, OFFSET_HOST);
}

int __stdcall vbcurl_version_features(CURL_VERSION* pv)
{
    return vbcurl_version_int(pv, OFFSET_FEATURES);
}

BSTR __stdcall vbcurl_version_ssl(CURL_VERSION* pv)
{
    return vbcurl_version_bstr(pv, OFFSET_SSL_VERSION);
}

int __stdcall vbcurl_version_ssl_num(CURL_VERSION* pv)
{
    return vbcurl_version_int(pv, OFFSET_SSL_VERSION_NUM);
}

BSTR __stdcall vbcurl_version_libz(CURL_VERSION* pv)
{
    return vbcurl_version_bstr(pv, OFFSET_LIBZ_VERSION);
}

void __stdcall vbcurl_version_protocols(CURL_VERSION* pv,
    SAFEARRAY** ppsa)
{
    LONG rgIndex = 0;
    int nProtocols = 0;
    if (pv->_magic != VERSION_MAGIC)
        RaiseException(0xE0000000, EXCEPTION_NONCONTINUABLE, 0, NULL);
    nProtocols = vbcurl_num_protocols(pv, OFFSET_PROTOCOLS);
    *ppsa = SafeArrayCreateVector(VT_BSTR, 0, nProtocols);
    for (rgIndex = 0; rgIndex < nProtocols; rgIndex++)
    {
        SafeArrayPutElement(*ppsa, &rgIndex,
            vbcurl_protocol_string(pv, OFFSET_PROTOCOLS, rgIndex));
    }
}

BSTR __stdcall vbcurl_version_ares(CURL_VERSION* pv)
{
    if (pv->_magic != VERSION_MAGIC)
        RaiseException(0xE0000000, EXCEPTION_NONCONTINUABLE, 0, NULL);
    if (pv->_pcvData->age > CURLVERSION_FIRST)
        return vbcurl_version_bstr(pv, OFFSET_ARES_VERSION);
    return SysAllocString((OLECHAR*)NOT_APPLICABLE);
}

int __stdcall vbcurl_version_ares_num(CURL_VERSION* pv)
{
    if (pv->_magic != VERSION_MAGIC)
        RaiseException(0xE0000000, EXCEPTION_NONCONTINUABLE, 0, NULL);
    if (pv->_pcvData->age > CURLVERSION_FIRST)
        return vbcurl_version_int(pv, OFFSET_ARES_VERSION_NUM);
    return 0;
}

BSTR __stdcall vbcurl_version_libidn(CURL_VERSION* pv)
{
    if (pv->_magic != VERSION_MAGIC)
        RaiseException(0xE0000000, EXCEPTION_NONCONTINUABLE, 0, NULL);
    if (pv->_pcvData->age > CURLVERSION_SECOND)
        return vbcurl_version_bstr(pv, OFFSET_LIBIDN_VERSION);
    return SysAllocString((OLECHAR*)NOT_APPLICABLE);
}
