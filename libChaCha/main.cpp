#include <stdio.h>
#include "stdint.h"
#include <string.h>
#include <stdlib.h>
#include <malloc.h>
#include <Windows.h>
#include <comutil.h>
#include "chacha20_simple.h" 

#pragma comment(lib, "comsuppw.lib")
#define EXPORT comment(linker, "/EXPORT:"__FUNCTION__"="__FUNCDNAME__)

uint64_t counter = 0;
uint8_t key[32];  
uint8_t nonce[8];
bool isInit = false;

void __stdcall chacha_init(char* _key, int kLen, int count)
{
#pragma EXPORT
	counter = count;
	memset(nonce,0,sizeof(nonce));
	memset(key,0,sizeof(key));
	if(kLen > 32) kLen = 32;
	memcpy(key,_key,kLen);
	isInit = true;
}



//vb true = -1, false = 0
SAFEARRAY* __stdcall chacha(SAFEARRAY** buf, short doEncrypt)
{
#pragma EXPORT

  if(!isInit) return 0;  
  if(*buf==0) return 0;
  if( (*buf)->cbElements != 1) return 0; //1 dimension  
  if( (*buf)->rgsabound[0].cElements < 1) return 0; //empty

  SafeArrayLock(*buf);
  uint32_t len = (*buf)->rgsabound[0].cElements; 

  uint8_t *plain = (uint8_t *)(*buf)->pvData;
  uint8_t *output = (uint8_t *)malloc(len);
  memset(output, 0, len);

  chacha20_ctx ctx;
  chacha20_setup(&ctx, key, sizeof(key), nonce);
  chacha20_counter_set(&ctx, counter);

  if(doEncrypt == -1) 
		chacha20_encrypt(&ctx, plain, output, len);
  else
		chacha20_decrypt(&ctx, plain, output, len);

  SAFEARRAYBOUND arrayBounds[1] = { {len, 0}};
  SAFEARRAY* psa = SafeArrayCreate(VT_I1, 1, arrayBounds);
  SafeArrayLock(psa);

  memcpy(psa->pvData,output,len);
  free(output);

  SafeArrayUnlock(psa);
  SafeArrayUnlock(*buf);
  isInit = false;

  return psa;
}

