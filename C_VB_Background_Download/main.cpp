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
	char* lpOutBuffer;
	char* buf;
	int bufSz;
	int outBufSz;
	HINTERNET hOpen;
	HINTERNET hConnect;
	HINTERNET hRequest;
	FILE *hFile ;
};

void releaseCFG(CONFIG* cfg){
	  free(cfg->SERVER);
	  free(cfg->WEBPATH);
	  free(cfg->TOFILE);
	  if(cfg->lpOutBuffer != NULL) free(cfg->lpOutBuffer);
	  if(cfg->buf != NULL) free(cfg->buf);
	  free(cfg);
}

DWORD WINAPI BackgroundWinInetDownload_Thread( LPVOID lpParam ) 
{
  unsigned long sz;
  int rv = 0 ;

  CONFIG *cfg = (CONFIG*)lpParam;

  DWORD ignoreCertErrors =  SECURITY_FLAG_IGNORE_CERT_DATE_INVALID |
							SECURITY_FLAG_IGNORE_CERT_CN_INVALID | 
							SECURITY_FLAG_IGNORE_UNKNOWN_CA |
							SECURITY_FLAG_IGNORE_REVOCATION |
							SECURITY_FLAG_IGNORE_WRONG_USAGE; 
  
  DWORD sz2 = 4;
  DWORD opts;
  DWORD dwSize = cfg->outBufSz;

  *cfg->RETVAL = -1;
  cfg->hOpen = InternetOpen("WininetDl", INTERNET_OPEN_TYPE_PRECONFIG, NULL,NULL, 0 );
  if(cfg->hOpen == NULL) goto errOut;

  if(cfg->isSSL){
	  cfg->hConnect = InternetConnect(cfg->hOpen, cfg->SERVER, INTERNET_DEFAULT_HTTPS_PORT, NULL,NULL, INTERNET_SERVICE_HTTP, INTERNET_FLAG_SECURE,0);
  }else{
	  cfg->hConnect = InternetConnect(cfg->hOpen, cfg->SERVER, INTERNET_DEFAULT_HTTP_PORT, NULL,NULL,INTERNET_SERVICE_HTTP,0,0);
  }

  if(cfg->hConnect == NULL) goto errOut;

  cfg->hRequest = HttpOpenRequest(cfg->hConnect,
                                "GET",
                                cfg->WEBPATH,
                                "HTTP/1.1", NULL, NULL, INTERNET_FLAG_RELOAD | INTERNET_FLAG_EXISTING_CONNECT, 0); 

  if(cfg->hRequest == NULL) goto errOut;
  
  /*
  rv = InternetQueryOption(hRequest, INTERNET_OPTION_SECURITY_FLAGS, &opts, &sz2);
  if(rv==0){
	  rv = GetLastError();
  }else{
	opts = opts | ignoreCertErrors;
	rv = InternetSetOption(hRequest, INTERNET_OPTION_SECURITY_FLAGS, &opts, sz2);
    if(rv==0) rv = GetLastError();
  } */

  rv = HttpSendRequest(cfg->hRequest,0,0,0,0);
  if(rv==0) goto errOut;

  rv = HttpQueryInfo(cfg->hRequest, HTTP_QUERY_STATUS_CODE, (LPVOID)cfg->lpOutBuffer, &dwSize, NULL);
  if(rv) *cfg->STATUS_CODE = atoi(cfg->lpOutBuffer); 

  rv = HttpQueryInfo(cfg->hRequest, HTTP_QUERY_CONTENT_LENGTH , (LPVOID)cfg->lpOutBuffer, &dwSize, NULL);
  if(rv) *cfg->CONTENT_LENGTH = atoi(cfg->lpOutBuffer); 

  cfg->hFile = fopen(cfg->TOFILE, "wb");
  if(cfg->hFile == NULL) goto errOut;

  while(InternetReadFile(cfg->hRequest, cfg->buf, cfg->bufSz-1, &sz) && sz !=0)
  {
	 if(*cfg->ABORT) goto errOut;
	 fwrite(cfg->buf, 1, sz, cfg->hFile);
	 //cfg->buf[sz] = '\0';
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
  if(cfg->hFile != NULL) fclose(cfg->hFile);
  if(cfg->hOpen != NULL) InternetCloseHandle(cfg->hOpen);

  releaseCFG(cfg);

  return 0;

}

int __stdcall BackgroundDownload(char* server, char* webpath, int ssl, char* toFile, int* retVal, int* abort,	int* statusCode, int* content_length, int* progress)
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
	
	cfg->bufSz = 0x1001;
	cfg->buf = (char*)malloc(cfg->bufSz);

	cfg->outBufSz = 256;
	cfg->lpOutBuffer = (char*)malloc(cfg->outBufSz);
	
	if(cfg->buf == NULL || cfg->lpOutBuffer == NULL){
		releaseCFG(cfg);
		return 0;
	}

	HANDLE hThread = CreateThread(NULL, 0, BackgroundWinInetDownload_Thread, (LPVOID)cfg, 0, 0); 
	
	if(hThread==NULL){
		releaseCFG(cfg);
		return 0;
	}

	return (int)hThread; 
}


