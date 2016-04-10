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
 * $Id: form.c,v 1.1 2005/03/01 00:06:25 jeffreyphillips Exp $
 **************************************************************************/

#include <windows.h>
#include <curl/curl.h>
#include "seq.h"
#include "slist.h"

#define FORM_MAGIC  0x7F7F0001
#define PART_MAGIC  0x7F7F0002

typedef struct tagFormContext
{
    unsigned int _magic;
    struct curl_httppost* _post;
    struct curl_httppost* _last;
    Seq_T _parts;
}   FORM_CONTEXT;

typedef struct tagFormPart
{
    unsigned int _magic;
    Seq_T _part;
}   FORM_PART;

static HMODULE g_hModCurl;
static FARPROC g_fpFormAdd;

void* form_get_post(void* pvFormContext)
{
    FORM_CONTEXT* fc = (FORM_CONTEXT*)pvFormContext;
    if (fc->_magic != FORM_MAGIC)
        return NULL;
    return fc->_post;
}

void* __stdcall vbcurl_form_create()
{
    FORM_CONTEXT* fc = (FORM_CONTEXT*)malloc(sizeof(FORM_CONTEXT));
    fc->_magic = FORM_MAGIC;
    fc->_post = NULL;
    fc->_last = NULL;
    fc->_parts = Seq_new(0);
    return fc;
}

void __stdcall vbcurl_form_free(FORM_CONTEXT* fc)
{
    // need to free all the parts
    int i, j;
    int numThisPart, numParts;

    if (fc->_magic != FORM_MAGIC)
        RaiseException(0xE0000000, EXCEPTION_NONCONTINUABLE, 0, NULL);

    numParts = Seq_length(fc->_parts);
    for (i = 0; i < numParts; i++)
    {
        FORM_PART* fp = (FORM_PART*)Seq_get(fc->_parts, i);
        Seq_T part = fp->_part;
        numThisPart = Seq_length(part);
        for (j = 0; j < numThisPart; j += 2)
        {
            CURLformoption opt = (CURLformoption)Seq_get(part, j);
            switch(opt)
            {
                case CURLFORM_BUFFER:
                case CURLFORM_CONTENTTYPE:
                case CURLFORM_COPYNAME:
                case CURLFORM_COPYCONTENTS:
                case CURLFORM_FILE:
                case CURLFORM_FILECONTENT:
                case CURLFORM_FILENAME:
                    free(Seq_get(part, j + 1));
                    break;
                default:
                    break;
            }
        }
        Seq_free(&part);
        free(fp);
    }
    Seq_free(&fc->_parts);
    curl_formfree(fc->_post);
    free(fc);
}

FORM_PART* __stdcall vbcurl_form_create_part(FORM_CONTEXT* fc)
{
    FORM_PART* fp;

    if (fc->_magic != FORM_MAGIC)
        RaiseException(0xE0000000, EXCEPTION_NONCONTINUABLE, 0, NULL);

    fp = (FORM_PART*)malloc(sizeof(FORM_PART));
    fp->_magic = PART_MAGIC;
    fp->_part = Seq_new(0);
    Seq_addhi(fc->_parts, fp);
    return fp;
}

CURLFORMcode __stdcall vbcurl_form_add_pair_to_part(FORM_PART* fp,
    CURLformoption option, VARIANT* pvarValue)
{
    VARIANT vTemp;
    VariantInit(&vTemp);

    if (fp->_magic != PART_MAGIC)
        RaiseException(0xE0000000, EXCEPTION_NONCONTINUABLE, 0, NULL);

    switch(option)
    {
        // numeric cases: trivial
        case CURLFORM_BUFFERLENGTH:
        case CURLFORM_CONTENTSLENGTH:
        case CURLFORM_NAMELENGTH:
            VariantChangeType(&vTemp, pvarValue, 0, VT_I4);
            Seq_addhi(fp->_part, (void*)option);
            Seq_addhi(fp->_part, (void*)vTemp.lVal);
            break;

        // string cases
        case CURLFORM_BUFFER:
        case CURLFORM_COPYNAME:
        case CURLFORM_COPYCONTENTS:
        case CURLFORM_FILE:
        case CURLFORM_FILECONTENT:
        case CURLFORM_FILENAME:
        case CURLFORM_CONTENTTYPE:
        {
            wchar_t* pwcStr;
            int len;
            char* pszStr;

            VariantChangeType(&vTemp, pvarValue, 0, VT_BSTR);
            pwcStr = (wchar_t*)vTemp.bstrVal;
            len = (int)wcslen(pwcStr);
            pszStr = (char*)malloc(len + 1);
            
            // add the option
            Seq_addhi(fp->_part, (void*)option);

            // now deal with the string stuff
            wcstombs(pszStr, pwcStr, len + 1);
            Seq_addhi(fp->_part, pszStr);
            break;
        }

        // slist case
        case CURLFORM_CONTENTHEADER:
        {
            void* pvSlist;
            VariantChangeType(&vTemp, pvarValue, 0, VT_I4);
            pvSlist = slist_get_inner((void*)vTemp.lVal);
            if (pvSlist)
            {
                Seq_addhi(fp->_part, (void*)option);
                Seq_addhi(fp->_part, pvSlist);
                break;
            }
            else
                return CURL_FORMADD_UNKNOWN_OPTION;            
        }

        default:
            return CURL_FORMADD_UNKNOWN_OPTION;
    }
    return CURL_FORMADD_OK;
}

CURLFORMcode __stdcall vbcurl_form_add_four_to_part(FORM_PART* fp,
    CURLformoption opt1, VARIANT* pvarVal1,
    CURLformoption opt2, VARIANT* pvarVal2)
{
    CURLFORMcode code =
        vbcurl_form_add_pair_to_part(fp, opt1, pvarVal1);
    if (code != CURL_FORMADD_OK)
        return code;
    return vbcurl_form_add_pair_to_part(fp, opt2, pvarVal2);
}

CURLFORMcode __stdcall vbcurl_form_add_six_to_part(FORM_PART* fp,
    CURLformoption opt1, void* pvarVal1,
    CURLformoption opt2, void* pvarVal2,
    CURLformoption opt3, void* pvarVal3)
{
    CURLFORMcode code =
        vbcurl_form_add_four_to_part(fp, opt1, pvarVal1,
            opt2, pvarVal2);
    if (code != CURL_FORMADD_OK)
        return code;
    return vbcurl_form_add_pair_to_part(fp, opt3, pvarVal3);
}

CURLFORMcode __stdcall vbcurl_form_add_part(
    FORM_CONTEXT* fc, FORM_PART* fp)
{
    CURLFORMcode retVal;
    void*  ppLast;
    void*  ppPost;
    int    i;
    int    argsInPairs;
    int    stackFix;
    void** ppArgs;

    if (fc->_magic != FORM_MAGIC)
        RaiseException(0xE0000000, EXCEPTION_NONCONTINUABLE, 0, NULL);
    if (fp->_magic != PART_MAGIC)
        RaiseException(0xE0000000, EXCEPTION_NONCONTINUABLE, 0, NULL);

    ppLast = &fc->_last;
    ppPost = &fc->_post;
    argsInPairs = Seq_length(fp->_part);
    stackFix = sizeof(int) * (argsInPairs + 3);
    ppArgs = (void**)_alloca(argsInPairs * sizeof(void*));

    if (!g_hModCurl)
    {
        g_hModCurl = GetModuleHandle("libcurl.dll");
        g_fpFormAdd = GetProcAddress(g_hModCurl, "curl_formadd");
    }

    // load the void** array from first to last
    for (i = 0; i < argsInPairs; i++)
        ppArgs[i] = Seq_get(fp->_part, i);

    // now decrement before going into assembly
    argsInPairs--;

    __asm
    {
        // terminator in call to curl_formadd
        push CURLFORM_END

        // set up loop to push the option->value pairs
        // in reverse order
        mov  ebx, ppArgs
        mov  ecx, argsInPairs
Args:   mov  eax, [ebx + 4 * ecx]
        push eax
        dec  ecx
        jns  args

        // push the two struct curl_httpost** pointers
        push ppLast
        push ppPost

        // and now call curl_formadd
        call g_fpFormAdd

        // get the return value
        mov  retVal, eax

        // and finally, clean up the stack, since we called
        // a __cdecl function
        add  esp, stackFix
    }
    return retVal;
}
