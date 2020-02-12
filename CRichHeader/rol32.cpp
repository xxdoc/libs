
#include <stdio.h>
#include <conio.h>

#define uint32_t unsigned int

 /*static inline uint32_t rol32_generic(const uint32_t& x, int i) {
       return (x << static_cast<uint32_t>(i & 31)) |
              (x >> static_cast<uint32_t>((32 - (i & 31)) & 31));
 }*/

 uint32_t __stdcall proto_CallWindowProc_rol32_generic(int dummy, const uint32_t x , int i, int arg4){
	return (x << static_cast<uint32_t>(i & 31)) |
              (x >> static_cast<uint32_t>((32 - (i & 31)) & 31));

 }

 uint32_t __stdcall _add32(int dummy, const uint32_t x , uint32_t i, int arg4){
	return x+i;
 }

/*
turn off optimizations required

_add32: 55 8B EC 8B 45 0C 03 45 10 5d C2 10 00

.text:00401030                         sub_401030      proc near               ; CODE XREF: _main+2E?p
.text:00401030
.text:00401030                         arg_4           = dword ptr  0Ch
.text:00401030                         arg_8           = dword ptr  10h
.text:00401030
.text:00401030 55                                      push    ebp
.text:00401031 8B EC                                   mov     ebp, esp
.text:00401033 8B 45 0C                                mov     eax, [ebp+arg_4]
.text:00401036 03 45 10                                add     eax, [ebp+arg_8]
.text:00401039 5D                                      pop     ebp
.text:0040103A C2 10 00                                retn    10h
.text:0040103A                         sub_401030      endp


rol32 : 55 8B EC 56 8B 4D 10 83 E1 1F 8B 45 0C D3 E0 8B 4D 10 83 E1 1F BA 20 00 00 00 2B D1 83 E2 1F 8B 75 0C 8B CA D3 EE 0B C6 5E 5D C2 10 00

.text:00401000                         sub_401000      proc near               ; CODE XREF: _main+C?p
.text:00401000
.text:00401000                         arg_4           = dword ptr  0Ch
.text:00401000                         arg_8           = dword ptr  10h
.text:00401000
.text:00401000 55                                      push    ebp
.text:00401001 8B EC                                   mov     ebp, esp
.text:00401003 56                                      push    esi
.text:00401004 8B 4D 10                                mov     ecx, [ebp+arg_8]
.text:00401007 83 E1 1F                                and     ecx, 1Fh
.text:0040100A 8B 45 0C                                mov     eax, [ebp+arg_4]
.text:0040100D D3 E0                                   shl     eax, cl
.text:0040100F 8B 4D 10                                mov     ecx, [ebp+arg_8]
.text:00401012 83 E1 1F                                and     ecx, 1Fh
.text:00401015 BA 20 00 00 00                          mov     edx, 20h ; ' '
.text:0040101A 2B D1                                   sub     edx, ecx
.text:0040101C 83 E2 1F                                and     edx, 1Fh
.text:0040101F 8B 75 0C                                mov     esi, [ebp+arg_4]
.text:00401022 8B CA                                   mov     ecx, edx
.text:00401024 D3 EE                                   shr     esi, cl
.text:00401026 0B C6                                   or      eax, esi
.text:00401028 5E                                      pop     esi
.text:00401029 5D                                      pop     ebp
.text:0040102A C2 10 00                                retn    10h
.text:0040102A                         sub_401000      endp





*/

void main(void){
	
	//printf("%x\n", rol32_generic(1,1));
	uint32_t x = proto_CallWindowProc_rol32_generic(0,1,1,0);
	printf("%x\n",x );
	printf("%x\n", _add32(0,-1,1,0) );
	getch();
}

