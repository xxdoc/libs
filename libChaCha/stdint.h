#include "inttypes.h"
#pragma warning( disable : 4244 ) //size conversion possible losss of data
#define UINT32_MAX  0xffffffff
#define UINT32_C(x)  ((x) + (UINT32_MAX - UINT32_MAX))

//#define BIG_ENDIAN 0
//#define LIL_ENDIAN 1
//#define BYTE_ORDER LIL_ENDIAN
