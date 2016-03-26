/*	
   This DLL is an encapsulation of the GPL UCL Library designed
   so it could be called from both VB and C. 
   
   This project links to the included ucl.lib.
   
   You can download the full source to the compression library from
   the following URL:
   
   http://www.oberhumer.com/opensource/ucl/

   The UCL library is Copyright (C) 1996-2002 Markus Franz Xaver 
   Johannes Oberhumer All Rights Reserved.

   The UCL library is free software; you can redistribute it and/or
   modify it under the terms of the GNU General Public License as
   published by the Free Software Foundation; either version 2 of
   the License, or (at your option) any later version.

 */

#include "ucl.h"
#include "lutil.h"
#include "ucl_dll.h"

int level = 5;                
char ErrorString[50];


void __stdcall LastError(char* buf){ strcpy(buf, ErrorString); }
int __stdcall CalcBufSize(int inBufLen){ return inBufLen+(inBufLen/8)+256;}

void __stdcall SetLevel(int x){
	if(x > 10 || x < 1){
		strcpy(ErrorString, "Valid Compression Levels are 1-10");
	}
	else{ level = x; }
}

int __stdcall Init(void){

	if (ucl_init() != UCL_E_OK){
        strcpy(ErrorString, "ucl_init() failed");
		return -1;
    }

	return 1;

}

int __stdcall Compress(const unsigned char* inBuf, unsigned char* outBuf, unsigned int in_len,  unsigned int* out_len){
	
	int ret;
	
 
	ret = ucl_nrv2b_99_compress(inBuf, in_len, outBuf, out_len, NULL, level, NULL, NULL);

	if (ret != UCL_E_OK){ // this should NEVER happen 
		sprintf(ErrorString,"Compression failed: %d\n", ret);
		return -1;
	}

	if( in_len == *out_len) return 0;

	return 1;
 
}



int __stdcall DeCompress(const unsigned char* inBuf, unsigned char* outBuf,  unsigned int in_len,  unsigned int* out_len){
	
	int ret;
	
	try{

		ret = ucl_nrv2b_decompress_8(inBuf, in_len, outBuf, out_len,NULL);
	
		if (ret != UCL_E_OK){ // this should NEVER happen 
			sprintf(ErrorString,"Decompression failed: %d\n", ret);
			return -1;

		}

		if( in_len == *out_len) return 0;

		return 1;

	}
	catch(...){
		strcpy(ErrorString, "Not a valid Compressed Buffer");
		return -1;
	}

	

}




