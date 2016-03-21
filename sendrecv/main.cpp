#include <stdio.h>
#include <stdlib.h>
#include <winsock.h>

#pragma comment(lib, "ws2_32.lib")

char* server=0;
int port=0;
int ms_timeout=0;

SOCKET sockfd; 
char *buffer;//[4096]; //client supplied buffer now...
int bufSz=0;

#define SD_SEND 1
//bool debug = false;
bool initilized = false;
bool allowPartial = true;
char errBuf[2048];

//cant use this since I also have a c client which needs a correct lib, this will wack the lib file..
//#define EXPORT comment(linker, "/EXPORT:"__FUNCTION__"="__FUNCDNAME__)

//most of this was utilized from a post by arcomber Aug4 2012
//http://codereview.stackexchange.com/questions/14340/how-to-fix-retrieval-of-receive-data-very-basic-socket-library-in-c

/* helper function to connect to server */
SOCKET establish_connection(u_long nRemoteAddr, u_short nPort)
{
    /* Create a stream socket */
   struct sockaddr_in sinRemote;

   sockfd = socket(AF_INET, SOCK_STREAM, 0);
   if (sockfd != INVALID_SOCKET) {
       sinRemote.sin_family = AF_INET;
       sinRemote.sin_addr.s_addr = nRemoteAddr;
       sinRemote.sin_port = nPort;
       if (connect(sockfd, (struct sockaddr*)&sinRemote, sizeof(struct sockaddr_in)) == SOCKET_ERROR) {
           sockfd = INVALID_SOCKET;
       }
   }

   return sockfd;
}


bool init() {
   WORD ver = MAKEWORD( 1, 1 );
   WSADATA wsa = {0};
   if(WSAStartup(ver, &wsa)==0) initilized = true;
   return initilized;
}
 

u_long LookupAddress(const char* pcHost)
{
    u_long nRemoteAddr = inet_addr(pcHost);
    if (nRemoteAddr == INADDR_NONE) {
        /* pcHost isn't a dotted IP, so resolve it through DNS */
        struct hostent* pHE = gethostbyname(pcHost);
        if (pHE == 0) {
            return INADDR_NONE;
        }
        nRemoteAddr = *((u_long*)pHE->h_addr_list[0]);
    }
    return nRemoteAddr;
}

int tcp_connect(const char* pcHost, int nPort) {
    struct in_addr Address; 

    u_long nRemoteAddress = LookupAddress(pcHost);
    if (nRemoteAddress == INADDR_NONE) {
        strcpy(errBuf,"lookup address error");
        return 3;
    }
    memcpy(&Address, &nRemoteAddress, sizeof(u_long)); 

    sockfd = establish_connection(nRemoteAddress, htons(nPort));
    if (sockfd == INVALID_SOCKET) {
		sprintf(errBuf,"Unable to establish a connection with server %s:%d", pcHost, nPort);
        return 3;
    }
    //if(debug) printf("successfully connected to server %s\n", pcHost);
    return 0;
}

int sendData(const char* data, int len) {
   int rv = 0, rv2=0;
   rv = send(sockfd, data, len, 0);
   if(rv == SOCKET_ERROR) {
        sprintf(errBuf,"send failed with error: %d", WSAGetLastError());
        return -1;
   }
   return rv;
}

int dorecv(int ms_timeout) {
   int n = 0, offset=0;
   unsigned int startTime = GetTickCount();
   do{
		n = recv(sockfd, &buffer[offset], bufSz-offset-2 , 0); 
		offset+=n;
		if(n==0) break;
		if((bufSz-offset) <= 2){
			strcpy(errBuf,"Buffer to small");
			if(allowPartial) break; //leave this one as a soft error because I may use it as a feature..
			return -5;
		}
		if(GetTickCount() - startTime > ms_timeout){
			strcpy(errBuf,"Timeout");
			//break;
			return -6; //this one we will consider an error, buffer may be partially full though..
		}
   }while(n > 0);
   //if(debug && offset == 0) printf("no data received?\n");	
   if(n < 0) sprintf(errBuf,"recv failed with error: %d", WSAGetLastError());	
   buffer[offset+1]= 0;
   return offset;
}

int sendRecv(char* server, int port, char* msg, int msgLen, int ms_timeout){

	if(tcp_connect(server,port) != 0) return -1;
	if(sendData(msg,msgLen) == -1) return -2;
	
	int sz = dorecv(ms_timeout);
	if(sz < 1) return -3;

	closesocket(sockfd);
	sockfd = 0;
	return sz;

}

void __stdcall qsConfig(char* _server, int _port, int _timeout, short partialRespOk){
//#pragma EXPORT
	
	if(_server==0) return;
	if(server!=0) free(server);
	
	server= strdup(_server);
	port = _port;
	ms_timeout = _timeout;
	allowPartial = partialRespOk == 0 ? false : true;
}

int __stdcall QuickSend(char* request, int reqLen, char* response_buffer, int response_buflen){
//#pragma EXPORT

	if(server==0){
		strcpy(errBuf,"You must call qsConfig to set server first");
		return -8;
	}

	if(response_buffer ==0) return -6;
	if(response_buflen < 1) return -7;

	errBuf[0] = 0;
	buffer = response_buffer;
	bufSz = response_buflen;
	
	if(!initilized) init();
	if(!initilized){
		strcpy(errBuf,"Failed to initilize winsock");
		return -4;
	}

	return sendRecv(server, port, request, reqLen, ms_timeout); 
}

int __stdcall LastError(char* buffer, int buflen){
//#pragma EXPORT
	
	int eLen = strlen(errBuf);
	if(buflen < 1) return 0;
	if(eLen > buflen) eLen = buflen-1;

	if(eLen==0){
		buffer[0]=0;
	}else{
		strcpy(buffer,errBuf);
		buffer[eLen] = 0;
	}

	return eLen;

}