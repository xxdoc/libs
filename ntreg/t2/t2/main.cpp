#include <windows.h>
#include "stdstring.h"
#include "NtRegistry.h"
#include "msvbvm60.tlh"

CNtRegistry ntreg;
bool isInit = false;

const int VB_TRUE = -1;
const int VB_FALSE = 0;

BOOL APIENTRY DllMain( HANDLE hModule, DWORD  ul_reason_for_call,  LPVOID lpReserved){ 
	if(ul_reason_for_call==1){
		if(!isInit){
			isInit = true;
			ntreg.InitNtRegistry();
		}
	}
	return TRUE;
}

void addStr(_CollectionPtr p , char* str){
	_variant_t vv;
	vv.SetString(str);
	p->Add(&vv.GetVARIANT());
}

/*
Public hive As hKey
Function keyExists(path) As Boolean
Function DeleteValue(path, ValueName) As Boolean
    Function DeleteKey(path) As Boolean
Function CreateKey(path) As Boolean
Function SetValue(path, KeyName, Data, dType As dataType) As Boolean
Function ReadValue(path, ByVal KeyName)
Function EnumKeys(path) As String()
Function EnumValues(path) As String()
*/

/*
int __stdcall EnumMutex2(_CollectionPtr *pColl, void* doEventsCallback){
addStr(*pColl,buf);
*/

int __stdcall keyExists_(char* path){
	CStdString p = path;
	return ntreg.KeyExists(p) ? VB_TRUE : VB_FALSE;
}

int __stdcall deleteValue_(char* path, char* valName){
	CStdString p = path;
	CStdString v = valName;
	ntreg.SetPathVars(p);
	return ntreg.DeleteValue(v) ? VB_TRUE : VB_FALSE;
}

/*
int __stdcall createKey_(char* path){
	CStdString p = path;
	return ntreg.CreateKey(p) ? VB_TRUE : VB_FALSE;
}
*/

int __stdcall userSID(char* sid, int *size){
	
	CStdString s="";

	if( !ntreg.LookupSID(s) ){
		*size = 0;
		return VB_FALSE;
	}

	if(s.length() >= *size){
		*size = s.length() + 1;
		return VB_FALSE;
	}

	const char* ret = s.c_str();
	strncpy(sid, ret, s.length());
	*size = s.length();
	return VB_TRUE;
}


void main(void){


    

	//CStdString key = "\\Registry\\Machine\\SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Run";
	CStdString key = ntreg.GetRootPathFor(HKEY_LOCAL_MACHINE) + "\\SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Run";
	//CStdString key = "HKEY_LOCAL_MACHINE\\SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Run";
	CStdString name = "tvncontrol";
	CStdString def = "";

	CStdString ret = ntreg.ReadString(key,name,def); 

	printf("ret=%s\n",ret.c_str());
	printf("def=%s\n",def.c_str());


}

