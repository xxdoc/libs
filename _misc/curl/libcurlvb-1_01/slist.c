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
 * $Id: slist.c,v 1.2 2005/06/18 22:07:50 jeffreyphillips Exp $
 **************************************************************************/

#include <windows.h>
#include <curl\curl.h>

#define SLIST_MAGIC 0x7F7F0004

typedef struct
{
    unsigned int _magic;
    struct curl_slist* _theList;    
}   SLIST;

void* slist_get_inner(void* pvSlist)
{
    SLIST* pl = (SLIST*)pvSlist;
    if (pl->_magic == SLIST_MAGIC)
        return (void*)pl->_theList;
    return NULL;
}

void* __stdcall vbcurl_slist_create()
{
    SLIST* pl = (SLIST*)malloc(sizeof(SLIST));
    pl->_magic = SLIST_MAGIC;
    pl->_theList = NULL;
    return pl;
}

void __stdcall vbcurl_slist_append(SLIST* pl, BSTR bs)
{
    int len = 0;
    char* pszStr = NULL;
    wchar_t* pwcStr = (wchar_t*)bs;

    if (pl->_magic != SLIST_MAGIC)
        RaiseException(0xE0000000, EXCEPTION_NONCONTINUABLE, 0, NULL);
    len = (int)wcslen(pwcStr);
    pszStr = (char*)_alloca(len + 1);        
    wcstombs(pszStr, pwcStr, len + 1);
    pl->_theList = curl_slist_append(pl->_theList, (void*)pszStr);
}

void __stdcall vbcurl_slist_free(SLIST* pl)
{
    void* pvStr = NULL;
    if (pl->_magic != SLIST_MAGIC)
        RaiseException(0xE0000000, EXCEPTION_NONCONTINUABLE, 0, NULL);
    curl_slist_free_all(pl->_theList);
    free(pl);
}
