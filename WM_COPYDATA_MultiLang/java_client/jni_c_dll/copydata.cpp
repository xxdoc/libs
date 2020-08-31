
#include <stdio.h>
#include <string.h>
#include <windows.h>
#include "JNITest.h"
#include <jni.h>

const int WM_DISPLAY_TEXT = 3;

//this is a super quick and dirty demo of a C client..

typedef struct{
    int dwFlag;
    int cbSize;
    int lpData;
} cpyData;


HWND hServer;
WNDPROC oldProc;
cpyData CopyData;

DWORD uThreadId =0;
int  threadInitilized = 0;
int  IDA_HWND=0;
int  m_debug = 0;
int  m_ServerHwnd = 0;
char* IDA_SERVER_NAME = "IDA_SERVER";
char m_msg[2020];
bool received_response = false;
HANDLE hConOut = 0;

static JNIEnv *g_env = NULL;
static jobject g_obj = NULL;
static JavaVM *g_jvm = NULL;
bool async_active = true;

#pragma warning (disable:4996)

void end_color(void){ SetConsoleTextAttribute(hConOut,7); }
void start_color(){ SetConsoleTextAttribute(hConOut, 14); }

bool SendTextMessage(int hwnd, char *Buffer, int blen) 
{
		  char* nullString = "NULL";
		  if(blen==0){ //in case they are waiting on a message with data len..
				Buffer = nullString;
				blen=4;
		  }
		  if(m_debug){
			  printf("Trying to send message to %x size:%d\n", hwnd, blen);
			  printf(" msg> %s\n", Buffer);
		  }
		  if(Buffer[blen] != 0) Buffer[blen]=0; ;
		  cpyData cpStructData;  
		  cpStructData.cbSize = blen;
		  cpStructData.lpData = (int)Buffer;
		  cpStructData.dwFlag = 3;
		  SendMessage((HWND)hwnd, WM_COPYDATA, (WPARAM)hwnd,(LPARAM)&cpStructData);  
		  return true;
}  

bool SendIntMessage(int hwnd, int resp){
	char tmp[30]={0};
	sprintf(tmp, "%d", resp);
	if(m_debug) printf("SendIntMsg(%d, %s)", hwnd, tmp);
	return SendTextMessage(hwnd,tmp, sizeof(tmp));
} 

//these next 2 are a very simple implementation..SendMessage automatically blocks so this works...
int ReceiveInt(char* command, int hwnd){
	memset(m_msg,0,2020);
	received_response = false;
	SendTextMessage(hwnd,command,strlen(command)+1);
	return atoi(m_msg);
}

char* ReceiveText(char* command, int hwnd){
	memset(m_msg,0,2020);
	received_response = false;
	async_active = false;
	SendTextMessage(hwnd, command,strlen(command)+1);
	async_active = true;
	return m_msg;
}

void AsyncCallBack(){

	if(g_obj == 0 || g_env == 0) return;
    if(strlen(m_msg) == 0) return;

	start_color();
	jclass cls = g_env->GetObjectClass(g_obj);
	
	if (cls == 0){
		printf("JNI: Failed to GetObjectClass\n");
		end_color();
		return;
	}

	jmethodID mid = g_env->GetMethodID(cls,"AsyncMsg", "(Ljava/lang/String;)V");

	if (mid == 0){
		printf("JNI: Failed to find AsyncMsg method id\n");
		end_color();
		return;
	}

	end_color();
	jstring js_ret = g_env->NewStringUTF(m_msg);
	g_env->CallVoidMethod(g_obj, mid, js_ret);
}


LRESULT CALLBACK WindowProc(HWND hwnd,UINT uMsg,WPARAM wParam,LPARAM lParam){
		
		if( uMsg != WM_COPYDATA) return 0;
		if( lParam == 0) return 0;
		
		memcpy((void*)&CopyData, (void*)lParam, sizeof(cpyData));
    
		if( CopyData.dwFlag == 3 ){
			if( CopyData.cbSize >= sizeof(m_msg) ) CopyData.cbSize = sizeof(m_msg)-1;
			memcpy((void*)&m_msg[0], (void*)CopyData.lpData, CopyData.cbSize);
			if(m_debug){
				start_color();
				printf("JNI: WindowProc Message Received: %s \n", m_msg); 
				end_color();
			}
			if(async_active) AsyncCallBack();
			received_response = true;
			return 0;
		}
			
    return CallWindowProc(oldProc, hwnd, uMsg, wParam, lParam);

}

JNIEXPORT jint JNICALL Java_JNITest_SendCMDRecvInt(JNIEnv *env, jobject obj, jstring javaString)
{
	 char* cmd = (char*)env->GetStringUTFChars(javaString, 0);
     
	 start_color();
	 printf("JNI: SendCMDRecvInt(%s)\n", cmd);
	 end_color();

	 return ReceiveInt(cmd,IDA_HWND);
}

JNIEXPORT void JNICALL Java_JNITest_SendCMD(JNIEnv *env, jobject obj, jstring javaString)
{
	char* cmd = (char*)env->GetStringUTFChars(javaString, 0);

	start_color();
	printf("JNI: SendCMD(%s)\n", cmd); 
	end_color();

	SendTextMessage(IDA_HWND, cmd, strlen(cmd) );
	
	env->ReleaseStringUTFChars(javaString, cmd);
}


JNIEXPORT jstring JNICALL Java_JNITest_SendCMDRecvText(JNIEnv *env, jobject obj, jstring javaString)
{
	 char* cmd = (char*)env->GetStringUTFChars(javaString, 0);
     
	 start_color();
	 printf("JNI: SendCMDRecvText(%s)\n", cmd);
	 
	 char* resp = ReceiveText(cmd,IDA_HWND);

	 printf("JNI: received: %s\n", resp);
	 end_color();

	 jstring js_ret = env->NewStringUTF(resp);
	 env->ReleaseStringUTFChars(javaString, cmd);

	 return js_ret;


}

int FindVBWindow(){

	char *vbIDEClassName = "ThunderFormDC" ;
	char *vbEXEClassName = "ThunderRT6FormDC" ;
	char *vbWindowCaption = "VB - C InterProcess Communications Using WM_COPYDATA" ;

	HWND hServer = FindWindowA( vbIDEClassName, vbWindowCaption );

	if(hServer==0){
		hServer = FindWindowA( vbEXEClassName, vbWindowCaption );
	}

	if( IsWindow(hServer) == 0){
		start_color();
		printf("JNI: Could not find VB6 command Window HServer= %x\n",hServer);
		end_color();
		return 0;
	} 

	return (int)hServer;
}

DWORD __stdcall CreateWndThread(LPVOID pThreadParam) 
{
	m_ServerHwnd = (int)CreateWindowA("EDIT","MESSAGE_WINDOW", 0, 0, 0, 0, 0, 0, 0, 0, 0);
	oldProc = (WNDPROC)SetWindowLongA((HWND)m_ServerHwnd, GWL_WNDPROC, (LONG)WindowProc);

	jint nStatus = g_jvm->AttachCurrentThread(reinterpret_cast<void**>(&g_env), NULL);
		 
	threadInitilized = 1;

	MSG Msg;
	while(GetMessage(&Msg, 0, 0, 0)) {
		TranslateMessage(&Msg);
		DispatchMessage(&Msg);
	}
	return Msg.wParam;

}

JNIEXPORT jint JNICALL Java_JNITest_InitHwnd(JNIEnv *env, jobject obj)
{
	
	g_env = env;
	env->GetJavaVM(&g_jvm);
	
	hConOut = GetStdHandle( STD_OUTPUT_HANDLE );
	start_color();

	//this has to be in a new thread or you will get threadlock..
	HANDLE hThread = (HANDLE)CreateThread(NULL, 0, &CreateWndThread, NULL, 0, &uThreadId);
	if(!hThread) 
	{
		printf("JNI: Fail creating thread");
		return 0;
	}
	
	g_obj = env->NewGlobalRef(obj); 

	IDA_HWND = FindVBWindow();
	if(!IsWindow((HWND)IDA_HWND)) IDA_HWND = 0;

	while(threadInitilized==0){ Sleep(1); }; //cheesy delay tactic do better..

	if( m_ServerHwnd == 0){
		printf("JNI: Could not create listener window to receive data on\n");
		end_color();
		return 0;
	}

	if( IDA_HWND==0 ){
		printf("JNI: Target vb6 hwnd window not found\n");
		end_color();
		return 0;
	}
	
	printf("JNI: Listening for responses on hwnd: %d\n", m_ServerHwnd); 
	printf("JNI: Target hwnd to send messages to: %d\n", IDA_HWND);
	end_color();

	return m_ServerHwnd;	

}