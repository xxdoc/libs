
#include <stdio.h>
#include <conio.h>
#include <string.h>
#include <windows.h>

HWND hServer;

int TransferData(char *Buffer, int dwBytes) 
{

 
  COPYDATASTRUCT cpStructData;
  memset(&cpStructData,0, sizeof(struct tagCOPYDATASTRUCT )) ;

  cpStructData.dwData = 3;
  cpStructData.cbData = dwBytes ;
  cpStructData.lpData = (void*)Buffer;
  
  printf("Transferring message, structure size = %d\n", sizeof(struct tagCOPYDATASTRUCT ));
  
  return SendMessage(hServer, WM_COPYDATA, (WPARAM)hServer,(LPARAM)&cpStructData);

}  

void main(int argc, char* argv[])
{
 
	char data[256];
	int lRet=0;
	
	char *vbIDEClassName = "ThunderFormDC" ;
	char *vbEXEClassName = "ThunderRT6FormDC" ;
	char *vbWindowCaption = "VB - C InterProcess Communications Using WM_COPYDATA" ;

	if(argc < 2){
		printf("Usage: wmcopy message\n\n");
		return;
	}

	//hServer = (HWND)atoi(argv[1]);
	strcpy(data,argv[1]);
	
	hServer = FindWindowA( vbIDEClassName, vbWindowCaption );

	if(hServer==0){
		hServer = FindWindowA( vbEXEClassName, vbWindowCaption );
	}


	if( IsWindow(hServer) == 0){
		printf("Could not find Window HServer= %x\n",hServer);
		printf("Press any key to continue");
		getch();
		return;
	} 
		
	lRet = TransferData( data , strlen(data) );
	
	printf("Found VB Server Window HWND: %d\n",hServer);
	printf("Message Sent Result: %d\n", lRet);
	printf("Press any key to continue");
	getch();

	return ;

}