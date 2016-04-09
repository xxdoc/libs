#include <stdio.h>
#include "stdint.h"
#include <string.h>
#include <stdlib.h>
#include <malloc.h>
#include <Windows.h>
#include <comutil.h>
#include <Atlconv.h>
#include "chacha20_simple.h" 

#pragma comment(lib, "comsuppw.lib")

uint64_t counter = 0;
uint8_t key[32];  
uint8_t nonce[8];
bool isInit = false;

//byval char* _key, or BSTR
void __stdcall test(char* _key)
{
	char buf[1024];
	if(_key != 0){
		uint32_t kLen = SysStringByteLen((BSTR)_key);
		if(kLen > 1000) return; //fu
		sprintf(buf,"BSTR '%s' Len: %d", _key, kLen);
		MessageBox(0,buf,"",0);
		_key[kLen-1] = '?';
	}
}

//byref char** or BSTR*
void __stdcall test2(BSTR* _key)
{
	char buf[1024];
	if(_key != 0){
		uint32_t kLen = SysStringByteLen(*_key);
		if(kLen > 1000) return; //fu
		sprintf(buf,"BSTR '%s' Len: %d", *_key, kLen);
		MessageBox(0,buf,"",0);
		char* b = (char*)*_key;
		b[kLen-1] = '?';
	}
}


//you can use this method to explicitly set the parameters, including binary keys..
void __stdcall chainit(BSTR* _key, BSTR* _nonce, uint32_t count)
{
	counter = count;
	memset(nonce,0,sizeof(nonce));
	memset(key,0,sizeof(key));

	if(_nonce != 0){
		uint32_t nLen = SysStringByteLen(*_nonce);
		if(nLen > 0){
			if(nLen > 8) nLen = 8;
			memcpy(nonce,(uint8_t*)*_nonce,nLen);
		}
	}

	if(_key != 0){
		uint32_t kLen = SysStringByteLen(*_key);
		if(kLen > 0){
			if(kLen > 32) kLen = 32;
			memcpy(key,(uint8_t*)*_key,kLen);
		}
	}

	isInit = true;
}


//decrypt is actually just a wrapper for encrypt..
//its symetric so no need for calling both or an isEncrypt flag..
SAFEARRAY* __stdcall chacha(SAFEARRAY** buf, BSTR* _key=0)
{

  //we will let key be optional parameter but include for convience. if you need 
  //to use a binary key or configure other params, use chainit directly.
  if(_key != 0){
	uint32_t kLen = SysStringByteLen(*_key);
	if(kLen > 0) chainit(_key,0,0);
  }

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

//in case you prefer to pass in a string..
//note: this is ok for binary strings on US systems only..
//for international software only use this as an encryption convience method.
SAFEARRAY* __stdcall chacha2(BSTR *buf, BSTR* _key=0)
{

  //we will let key be optional parameter but include for convience. if you need 
  //to use a binary key or configure other params, use chainit directly.
  if(_key != 0){
	uint32_t kLen = SysStringByteLen(*_key);
	if(kLen > 0) chainit(_key,0,0);
  }

  if(!isInit) return 0; 
  if(buf==0 || *buf==0) return 0;
  
  uint32_t bufLen = SysStringByteLen(*buf);
  if(bufLen==0) return 0;

  SAFEARRAYBOUND arrayBounds[1] = { {bufLen, 0}};
  SAFEARRAY* psa = SafeArrayCreate(VT_I1, 1, arrayBounds);
  if(psa==0) return 0;

  SafeArrayLock(psa);
  memset(psa->pvData,0,bufLen); //looks like its actually already zeroed out..
  
  chacha20_ctx ctx;
  chacha20_setup(&ctx, key, sizeof(key), nonce);
  chacha20_counter_set(&ctx, counter);
  chacha20_encrypt(&ctx, (uint8_t *)*buf, (uint8_t *)psa->pvData, bufLen);
   
  SafeArrayUnlock(psa);
  isInit = false;

  return psa;
}

BSTR __stdcall chacha4(BSTR *buf, BSTR* _key=0)
{

  //we will let key be optional parameter but include for convience. if you need 
  //to use a binary key or configure other params, use chainit directly.
  if(_key != 0){
	uint32_t kLen = SysStringByteLen(*_key);
	if(kLen > 0) chainit(_key,0,0);
  }

  if(!isInit) return 0;
  if(buf==0 || *buf==0) return 0;
  
  uint32_t bufLen = SysStringByteLen(*buf);
  if(bufLen==0) return 0;
   
  uint8_t * tmp = (uint8_t *)malloc(bufLen);
  if(tmp==0) return 0;
  
  chacha20_ctx ctx;
  chacha20_setup(&ctx, key, sizeof(key), nonce);
  chacha20_counter_set(&ctx, counter);
  chacha20_encrypt(&ctx, (uint8_t *)*buf, tmp, bufLen);

  //this can contain binary data ok if system uses english codepage 
  BSTR b = SysAllocStringByteLen((LPCSTR)tmp,bufLen);
  isInit = false;
  free(tmp);

  return b;
}
