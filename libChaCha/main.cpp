#include <stdio.h>
#include "stdint.h"
#include <string.h>
#include <stdlib.h>
#include <malloc.h>
#include <Windows.h>
#include <comutil.h>
#include "chacha20_simple.h" 

#pragma comment(lib, "comsuppw.lib")

uint64_t counter = 0;
uint8_t key[32];  
uint8_t nonce[8];
bool isInit = false;

//for api simplicity I didnt add access to setting nonce, you can if required.
void __stdcall chainit(char* _key, int kLen, int count)
{
	counter = count;
	memset(nonce,0,sizeof(nonce));
	memset(key,0,sizeof(key));
	if(kLen > 32) kLen = 32;
	memcpy(key,_key,kLen);
	isInit = true;
}


//decrypt is actually just a wrapper for encrypt..
//its symetric so no need for calling both or an isEncrypt flag..
SAFEARRAY* __stdcall chacha(SAFEARRAY** buf)
{

  if(!isInit) return 0; 
  if(buf==0 || *buf==0) return 0;
  if( (*buf)->cbElements != 1) return 0; //1 dimension  
  if( (*buf)->rgsabound[0].cElements < 1) return 0; //empty
  
  uint32_t len = (*buf)->rgsabound[0].cElements; 
  uint8_t *plain = (uint8_t *)(*buf)->pvData;

  SAFEARRAYBOUND arrayBounds[1] = { {len, 0}};
  SAFEARRAY* psa = SafeArrayCreate(VT_I1, 1, arrayBounds);
  if(psa==0) return 0;

  SafeArrayLock(*buf);
  SafeArrayLock(psa);
  memset(psa->pvData,0,len); //looks like its actually already zeroed out..
  
  chacha20_ctx ctx;
  chacha20_setup(&ctx, key, sizeof(key), nonce);
  chacha20_counter_set(&ctx, counter);
  chacha20_encrypt(&ctx, plain, (uint8_t *)psa->pvData, len);
   
  SafeArrayUnlock(psa);
  SafeArrayUnlock(*buf);
  isInit = false;

  return psa;
}

