#include <stdlib.h>
#include "minilzo.h"
#include <stdio.h>
#include <Windows.h>
#include <comutil.h>

#pragma comment(lib, "comsuppw.lib")

bool initilized = false;
char* lastError[500];

/* We want to compress the data block at 'in' with length 'IN_LEN' to
 * the block at 'out'. Because the input block may be incompressible,
 * we must provide a little more output space in case that compression
 * is not possible.
 */


#define IN_LEN      (128*1024ul)
#define OUT_LEN     (IN_LEN + IN_LEN / 16 + 64 + 3)

static unsigned char __LZO_MMODEL in  [ IN_LEN ];
static unsigned char __LZO_MMODEL out [ OUT_LEN ];


/* Work-memory needed for compression. Allocate memory in units
 * of 'lzo_align_t' (instead of 'char') to make sure it is properly aligned.
 */

#define HEAP_ALLOC(var,size) \
    lzo_align_t __LZO_MMODEL var [ ((size) + (sizeof(lzo_align_t) - 1)) / sizeof(lzo_align_t) ]

static HEAP_ALLOC(wrkmem, LZO1X_1_MEM_COMPRESS);


#define EXPORT comment(linker, "/EXPORT:"__FUNCTION__"="__FUNCDNAME__)


int __stdcall LZOGetMsg(char* buf, int bufsz, int msgID)
{
#pragma EXPORT
	
	char* b = (char*)lastError;

	if(msgID==0){
		sprintf(b,"LZO real-time data compression library (v%s, %s).\n", lzo_version_string(), lzo_version_date());
		strcat(b,"Copyright (C) 1996-2015 All Rights Reserved.\n");
		strcat(b,"Markus Franz Xaver Johannes Oberhumer");
	}
	
	int sz = strlen(b);
	if(sz==0) return 0;
	if(sz > bufsz-1) return 0;
	strcpy(buf,b);

}

int __stdcall Compress(unsigned char* buf, int bInSz , unsigned char* bOut, int bOutSz)
{
#pragma EXPORT

	int r;
	char* b = (char*)lastError;

	if(!initilized){
		if (lzo_init() != LZO_E_OK)
		{
			sprintf(b,"internal error - lzo_init() failed !!!\n");
			strcat(b,"(this usually indicates a compiler bug - try recompiling\nwithout optimizations, and enable '-DLZO_DEBUG' for diagnostics)\n");
			return -1;
		}
		initilized = true;
	}

	lzo_uint in_len = bInSz;
	lzo_uint out_len = bOutSz;
	
	if(out_len <= in_len ){
		sprintf(b,"Error Compress outbuffer (%d) must be larger than inbuffer(%d) just in case.",bOutSz,bInSz);
		return -2;
	}

	r = lzo1x_1_compress(buf,in_len,bOut,&out_len,wrkmem);
    if (r != LZO_E_OK)
    {
        /* this should NEVER happen */
        sprintf(b, "internal error - compression failed: %d\n", r);
        return -3;
    }
    /* check for an incompressible block */
    if (out_len >= in_len)
    {
        sprintf(b,"This block contains incompressible data.\n");
        return -4;
    }

	return out_len;

}

int __stdcall DeCompress(unsigned char* buf, int bInSz , unsigned char* bOut, int bOutSz)
{
#pragma EXPORT

	int r;
	char* b = (char*)lastError;

	if(!initilized){
		if (lzo_init() != LZO_E_OK)
		{
			sprintf(b,"internal error - lzo_init() failed !!!\n");
			strcat(b,"(this usually indicates a compiler bug - try recompiling\nwithout optimizations, and enable '-DLZO_DEBUG' for diagnostics)\n");
			return -1;
		}
		initilized = true;
	}

	lzo_uint in_len = bInSz;
	lzo_uint out_len = bOutSz;
	lzo_uint new_len;

	if(out_len <= in_len ){
		sprintf(b,"Error Compress outbuffer (%d) must be larger than inbuffer(%d)",bOutSz,bInSz);
		return -2;
	}

    r = lzo1x_decompress(buf,in_len,bOut,&out_len,NULL);
    if (r != LZO_E_OK)
    {
        sprintf(b,"internal error - decompression failed: %d\n", r);
        return -3;
    }

	return out_len;

}