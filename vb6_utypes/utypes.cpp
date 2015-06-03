#include <windows.h>
#include <stdio.h>

enum op{
	op_add = 0,
	op_sub = 1,
	op_div = 2,
	op_mul = 3,
	op_mod = 4,
	op_xor = 5,
	op_and = 6,
	op_or  = 7
};

enum modes{
	mUnsigned = 0,
	mSigned = 1,
	mHex = 2
};

struct x64{
	unsigned int lo;
	unsigned int hi;
};

unsigned int __stdcall ULong(unsigned int v1, unsigned int v2, int operation){

	switch(operation){
		case op_add: return v1 + v2;
		case op_sub: return v1 - v2;
		case op_div: return v1 / v2;
		case op_mul: return v1 * v2;
		case op_mod: return v1 % v2;
		case op_xor: return v1 ^ v2;
		case op_and: return v1 & v2;
		case op_or:  return v1 | v2;
	}

	return -1;

}

unsigned short __stdcall UInt(unsigned short v1, unsigned short v2, int operation){

	switch(operation){
		case op_add: return v1 + v2;
		case op_sub: return v1 - v2;
		case op_div: return v1 / v2;
		case op_mul: return v1 * v2;
		case op_mod: return v1 % v2;
		case op_xor: return v1 ^ v2;
		case op_and: return v1 & v2;
		case op_or:  return v1 | v2;
	}

	return -1;

}  

unsigned __int64 __stdcall U64(unsigned __int64 v1, unsigned __int64 v2, int operation){

	switch(operation){
		case op_add: return v1 + v2;
		case op_sub: return v1 - v2;
		case op_div: return v1 / v2;
		case op_mul: return v1 * v2;
		case op_mod: return v1 % v2;
		case op_xor: return v1 ^ v2;
		case op_and: return v1 & v2;
		case op_or:  return v1 | v2;
	}

	return -1;

}

__int64 __stdcall S64(__int64 v1, __int64 v2, int operation){

	switch(operation){
		case op_add: return v1 + v2;
		case op_sub: return v1 - v2;
		case op_div: return v1 / v2;
		case op_mul: return v1 * v2;
		case op_mod: return v1 % v2;
		case op_xor: return v1 ^ v2;
		case op_and: return v1 & v2;
		case op_or:  return v1 | v2;
	}

	return -1;

}


int __stdcall U642Str(unsigned __int64 v1, LPSTR pszString, LONG cSize, int mode){

	char buf[64]={0};
    int i;

	switch(mode){
		case mUnsigned: sprintf(buf, "%I64u", v1); break;
		case mSigned:   sprintf(buf, "%I64d", v1); break;
		case mHex:      sprintf(buf, "%I64x", v1); break;
	}
	
	//printf("%08X%08X", static_cast<UINT32>((u64>>32)&0xFFFFFFFF), static_cast<UINT32>(u64)&0xFFFFFFFF));

	i = strlen(buf);
	if (cSize > i ) strcpy(pszString, buf);
	return i;

}

int __stdcall U2Str(unsigned int v1, LPSTR pszString, LONG cSize){

	char buf[64]={0};
    int i;
    sprintf(buf, "%u", v1);
	i = strlen(buf);
	if (cSize > i ) strcpy(pszString, buf);
	return i;

}

unsigned __int64  __stdcall Str264(LPSTR pszString, int base){
	//supports - but doesnt seem to work with 0x prefix? so had to include base arg..
	unsigned __int64 ret = _strtoui64(pszString, NULL, base);
	return ret;
}

unsigned __int64 __stdcall toU64(unsigned int v1, unsigned int v2){

	unsigned __int64 ret = 0;
	x64 *x = (struct x64*)&ret;
	x->hi = v1;
	x->lo = v2;
	return ret;

}

void __stdcall fromU64(unsigned __int64 v0, unsigned int *v1, unsigned int *v2){

	x64 *x = (struct x64*)&v0;
	*v1 = x->hi;
	*v2 = x->lo;

}


/* 

was hoping doubles allowed you to do native + - in vb6 but seems wonky..
Private Declare Sub toU64d Lib "utypes.dll" (ByVal v1 As Long, ByVal v2 As Long, ByRef outVar As Double)

void __stdcall toU64d(unsigned int v1, unsigned int v2, unsigned __int64 *outVar){

	unsigned __int64 ret = 0;
	x64 *x = (struct x64*)&ret;
	x->hi = v1;
	x->lo = v2;
	*outVar = ret;

}*/