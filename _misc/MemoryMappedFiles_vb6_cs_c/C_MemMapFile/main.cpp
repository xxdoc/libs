#include <windows.h>
#include <conio.h>
#include <stdio.h>

char* ErrorMessage = 0;
HANDLE hFile = 0;
void* gAddr = 0;
unsigned int MaxSize = 0;

void setErrMsg(char*msg){
	if(ErrorMessage != NULL) free(ErrorMessage);
	if(msg != NULL && strlen(msg) > 0) ErrorMessage = strdup(msg);
}

bool CreateMemMapFile(char* fName, int mSize){
    
	setErrMsg(0);

    if((int)hFile!=0){
        setErrMsg("Cannot open multiple virtural files with one class");
        return false;
    }

    MaxSize = 0;
    hFile = CreateFileMapping((HANDLE)-1, 0, PAGE_READWRITE, 0, mSize, fName);

    if( hFile == 0 ){
        setErrMsg("Unable to create virtual file");
        return false;
    }

     gAddr = MapViewOfFile(hFile, FILE_MAP_ALL_ACCESS, 0, 0, mSize);

     if (gAddr == 0) return false;

	 MaxSize = mSize;
     return true;
    
}

bool WriteMemFile(unsigned char* bData, int blen){

	setErrMsg(0);
    if(blen == 0) return false;

    if(blen > MaxSize){
        setErrMsg("Data is to large for buffer!");
        return false;
    }

    if(hFile==0){
        setErrMsg("Virtual File or Virtual File Interface not initialized");
        return false;
    }

	memcpy(gAddr,bData,blen);
    return true;
}

bool ReadMemFile(unsigned char* buf, int blen){

	setErrMsg(0);

    if(blen > MaxSize){
        setErrMsg("ReadLength to large for buffer!");
        return false;
    }

    if(hFile==0){
        setErrMsg("Virtual File or Virtual File Interface not initialized");
        return false;
    }

	memcpy(buf,gAddr, blen); 
    return true;
}


void main(void){

	if(!CreateMemMapFile("DAVES_VFILE", 20)){
        printf("Failed to create vfile");
        return;
    }

	unsigned char* b = (unsigned char*)malloc(20);
	memset(b,0,20);

    if (!ReadMemFile(b, 5))
    {
        printf("Failed to read");
    }

    for (int i = 0; i < 5; i++)
    {
        printf("%x ", b[i]);
        b[i] = (byte)(0x41 + i);
    }
    printf("\n");

    if (!WriteMemFile(b,5))
    {
        printf("Failed to write file\n");
    }

    
    if (!ReadMemFile(b, 5))
    {
        printf("Failed to read");
    }

    for (int i = 0; i < 5; i++)
    {
        printf("%x ", b[i]);
    }

    printf("\nDo a write from extern app then Press any key to do one more read now..\n");
    getch();

    if (!ReadMemFile(b, 5))
    {
        printf("Failed to read");
    }

    for (int i = 0; i < 5; i++)
    {
        printf("%x ", b[i]);
    }

    printf("\nPress any key to exit...\n");
	getch();

}