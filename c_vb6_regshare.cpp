
#include <windows.h>  
#include <conio.h>
#include <stdio.h>

/* 
	vb6 code to set this registry key 

	Private Sub Form_Load()
		SaveSetting "FastBuild", "Settings", "DisplayAsHex", 22
		End
	End Sub
*/

#define ERROR_NO_KEY      0x11223344

int ReadRegInt(char* baseKey, char* name){

	 char tmp[20] = {0};
     unsigned long l = sizeof(tmp);
	 HKEY h;
	 
	 int rv = RegOpenKeyEx(HKEY_CURRENT_USER, baseKey, 0, KEY_READ, &h);
	 rv = RegQueryValueExA(h, name, 0,0, (unsigned char*)tmp, &l);
	 RegCloseKey(h);

	 if(rv != ERROR_SUCCESS) return ERROR_NO_KEY;
	 return atoi(tmp);
}


bool FileExists(char* szPath)
{
  DWORD dwAttrib = GetFileAttributes(szPath);
  bool rv = (dwAttrib != INVALID_FILE_ATTRIBUTES && !(dwAttrib & FILE_ATTRIBUTE_DIRECTORY)) ? true : false;
  return rv;
}

bool RegKeyExists(HKEY key , char* subPath){
	HKEY subKey = NULL;
	LONG result = RegOpenKeyEx(key, subPath, 0, KEY_READ, &subKey);
	RegCloseKey(subKey);
	return (result == ERROR_SUCCESS);
}

void main(void){
		
	char baseKey[200] = "Software\\VB and VBA Program Settings\\FastBuild3\\Settings";

	bool keyExists = RegKeyExists(HKEY_CURRENT_USER,baseKey);

	//if(keyExists){
		//printf("parent Key found\n");
		int v = ReadRegInt(baseKey, "DisplayAsHex");
		if(v == ERROR_NO_KEY){
			printf("value not set");
		}else{
			printf("%s\\DisplayAsHex = %d", baseKey, v);
		}
	//}else{
	//	printf("Key not found");
	//}

	getch();


}