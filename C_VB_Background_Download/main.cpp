#include <windows.h>
#include <urlmon.h>
#include <wininet.h>
#include <stdio.h>

#pragma comment(lib, "urlmon.lib")
#pragma comment (lib, "WinInet.lib")

#define EXPORT comment(linker, "/EXPORT:"__FUNCTION__"="__FUNCDNAME__)
#define WIN32_LEAN_AND_MEAN

//quick test of async download using another thread and updating vb variables directly by address
//supports multiple background downloads at once

//note xp test machine seems to fail for some https downloads, probably them not accepting tls 1.0

struct CONFIG{
	char* SERVER;
	char* WEBPATH;
	char* TOFILE;
	int   isSSL;
	int*  RETVAL;
	int*  ABORT;
	int*  STATUS_CODE;
	int*  PROGRESS;
	int*  CONTENT_LENGTH;
};

void releaseCFG(CONFIG* cfg){
	  free(cfg->SERVER);
	  free(cfg->WEBPATH);
	  free(cfg->TOFILE);
	  free(cfg);
}

DWORD WINAPI BackgroundWinInetDownload_Thread( LPVOID lpParam ) 
{
  int rv = 0 ;
  unsigned long sz;
  unsigned char buf[0x1001];
  char lpOutBuffer[256] = {0};
  DWORD dwSize = sizeof(lpOutBuffer);
  DWORD sz2 = 4;
  DWORD opts;
  DWORD timeout = 3 * 1000;	//3 seconds	

  HINTERNET hUrl = 0;
  HINTERNET hOpen = 0;
  HINTERNET hConnect = 0;
  HINTERNET hRequest = 0;
  FILE *hFile = 0;

  CONFIG *cfg = (CONFIG*)lpParam;

  DWORD ignoreCertErrors =  SECURITY_FLAG_IGNORE_CERT_DATE_INVALID |
							SECURITY_FLAG_IGNORE_CERT_CN_INVALID | 
							SECURITY_FLAG_IGNORE_UNKNOWN_CA |
							SECURITY_FLAG_IGNORE_REVOCATION |
							SECURITY_FLAG_IGNORE_WRONG_USAGE; 
  

  *cfg->RETVAL = -1;
  hOpen = InternetOpen("WininetDl", INTERNET_OPEN_TYPE_PRECONFIG, NULL,NULL, 0 );
  if(hOpen == NULL) goto errOut;

  InternetSetOption(hOpen, INTERNET_OPTION_CONNECT_TIMEOUT, &timeout, sizeof(timeout));

  if(cfg->isSSL){
	  hConnect = InternetConnect(hOpen, cfg->SERVER, INTERNET_DEFAULT_HTTPS_PORT, NULL,NULL, INTERNET_SERVICE_HTTP, INTERNET_FLAG_SECURE,0);
  }else{
	  hConnect = InternetConnect(hOpen, cfg->SERVER, INTERNET_DEFAULT_HTTP_PORT, NULL,NULL,INTERNET_SERVICE_HTTP,0,0);
  }

  if(hConnect == NULL) goto errOut;

  InternetSetOption(hConnect, INTERNET_OPTION_RECEIVE_TIMEOUT, &timeout, sizeof(timeout));
  InternetSetOption(hConnect, INTERNET_OPTION_SEND_TIMEOUT, &timeout, sizeof(timeout));
  
  hRequest = HttpOpenRequest(hConnect,
                                "GET",
                                cfg->WEBPATH,
                                "HTTP/1.1", NULL, NULL, INTERNET_FLAG_RELOAD | INTERNET_FLAG_EXISTING_CONNECT, 0); 

  if(hRequest == NULL) goto errOut;
  
  /*
  rv = InternetQueryOption(hRequest, INTERNET_OPTION_SECURITY_FLAGS, &opts, &sz2);
  if(rv==0){
	  rv = GetLastError();
  }else{
	opts = opts | ignoreCertErrors;
	rv = InternetSetOption(hRequest, INTERNET_OPTION_SECURITY_FLAGS, &opts, sz2);
    if(rv==0) rv = GetLastError();
  } */

  rv = HttpSendRequest(hRequest,0,0,0,0);
  if(rv==0) goto errOut;

  rv = HttpQueryInfo(hRequest, HTTP_QUERY_STATUS_CODE, (LPVOID)lpOutBuffer, &dwSize, NULL);
  if(rv) *cfg->STATUS_CODE = atoi(lpOutBuffer); 

  rv = HttpQueryInfo(hRequest, HTTP_QUERY_CONTENT_LENGTH , (LPVOID)lpOutBuffer, &dwSize, NULL);
  if(rv) *cfg->CONTENT_LENGTH = atoi(lpOutBuffer); 

  hFile = fopen(cfg->TOFILE, "wb");
  if(hFile == NULL) goto errOut;

  while(InternetReadFile(hRequest, buf, sizeof(buf)-1, &sz) && sz !=0)
  {
	 if(*cfg->ABORT) goto errOut;
	 fwrite(buf, 1, sz, hFile);
	 *cfg->PROGRESS += sz;
  }
  
  
  *cfg->RETVAL = 1;
  goto cleanup;

errOut:
   if(*cfg->ABORT == 0) *cfg->RETVAL = 3; else *cfg->RETVAL = 2;
   rv = GetLastError();
   //12152       ERROR_HTTP_INVALID_SERVER_RESPONSE
   //12031       ERROR_INTERNET_CONNECTION_RESET
   //https://support.microsoft.com/en-us/kb/193625

cleanup:
  if(hFile != NULL) fclose(hFile);
  if(hRequest!=NULL) InternetCloseHandle(hRequest);
  if(hConnect!=NULL) InternetCloseHandle(hConnect);
  if(hOpen != NULL) InternetCloseHandle(hOpen);

  releaseCFG(cfg);
  return 0;

}

int __stdcall BackgroundDownload(char* server, char* webpath, int ssl, char* toFile, int* retVal, int* abort, int* statusCode, int* content_length, int* progress)
{
#pragma EXPORT
	
	if(retVal==0) return 0;
	if(server==0) return  0;
	if(webpath==0) return  0;
	if(toFile==0) return  0;
	if(abort==0) return  0;
	if(statusCode==0) return 0;
	if(content_length==0) return 0;
	if(progress==0) return 0;

	*retVal = -2;
	*abort = 0;
	*statusCode = 0; 
	*content_length = 0;
	*progress = 0;

	CONFIG *cfg = (CONFIG*)malloc(sizeof(CONFIG));
	if(cfg==NULL) return 0;

	memset(cfg,0,sizeof(CONFIG));

	cfg->isSSL = ssl;
	cfg->ABORT = abort;
	cfg->RETVAL = retVal;
	cfg->STATUS_CODE = statusCode;
	cfg->CONTENT_LENGTH = content_length;
	cfg->PROGRESS = progress;
	cfg->SERVER = strdup(server);
	cfg->WEBPATH = strdup(webpath);
	cfg->TOFILE = strdup(toFile);

	HANDLE hThread = CreateThread(NULL, 0, BackgroundWinInetDownload_Thread, (LPVOID)cfg, 0, 0); 
	
	if(hThread==NULL){
		releaseCFG(cfg);
		return 0;
	}

	return (int)hThread; 
}


