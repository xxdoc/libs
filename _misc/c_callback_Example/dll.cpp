
typedef int(__stdcall *vbCallback)(int arg);
#define EXPORT comment(linker, "/EXPORT:"__FUNCTION__"="__FUNCDNAME__)

int __stdcall myFunc(vbCallback lpfnCallBack, int arg1)
{
#pragma EXPORT

	return lpfnCallBack(arg1);

}