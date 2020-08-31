#include <Windows.h>
#include <stdio.h>
#include <conio.h>

typedef struct {
	ULONG_PTR dwFlag; // dwData;
	DWORD     cbSize; // cbData;
	PVOID     lpData;
} cpyData;

HWND ServerHwnd=0;
WNDPROC oldProc=0;
char m_msg[2020];
cpyData CopyData;
CRITICAL_SECTION m_cs;
char* IPC_NAME = "C_SERVER";

//we can only assume these args/ret val to be 32bit because we must support a 32 bit sendmessage caller (vb6)
//The integral types WPARAM , LPARAM , and LRESULT are 32 bits wide on 32-bit systems and 64 bits wide on 64-bit systems

LRESULT CALLBACK WindowProc(HWND hwnd, UINT uMsg, WPARAM wParam, LPARAM lParam){
		

		if( uMsg != WM_COPYDATA) return DefWindowProc(hwnd, uMsg, wParam, lParam);
		if( lParam == 0)         return DefWindowProc(hwnd, uMsg, wParam, lParam);
		
		int retVal = 0;
		EnterCriticalSection(&m_cs);
		memcpy((void*)&CopyData, (void*)lParam, sizeof(cpyData));
    
		if (CopyData.dwFlag == 3 && CopyData.cbSize > 0) {

			if (CopyData.cbSize >= sizeof(m_msg) - 2) CopyData.cbSize = sizeof(m_msg) - 2;

			memcpy((void*)&m_msg[0], (void*)CopyData.lpData, CopyData.cbSize);
			m_msg[CopyData.cbSize] = 0; //always null terminate..

			printf("Message Received: %s \n", m_msg);
		}
			
		LeaveCriticalSection(&m_cs);
		return retVal;
}



int CreateServerWindow()
{
	WNDCLASSEX wc = { };
	MSG msg;
	HWND hwnd;

	wc.cbSize = sizeof(wc);
	wc.style = 0;
	wc.lpfnWndProc = WindowProc;
	wc.cbClsExtra = 0;
	wc.cbWndExtra = 0;
	wc.hInstance = GetModuleHandle(NULL);
	wc.hIcon = NULL;
	wc.hCursor = NULL;
	wc.hbrBackground = NULL;
	wc.lpszMenuName = NULL;
	wc.lpszClassName = IPC_NAME;
	wc.hIconSm = NULL;

	if (!RegisterClassEx(&wc)) {
		printf("Could not register window class\n");
		return 0;
	}

	ServerHwnd = CreateWindowEx(WS_EX_LEFT,
		IPC_NAME,
		NULL,
		WS_OVERLAPPEDWINDOW,
		CW_USEDEFAULT,
		CW_USEDEFAULT,
		CW_USEDEFAULT,
		CW_USEDEFAULT,
		NULL,
		NULL,
		GetModuleHandle(NULL),
		NULL);

	if (!ServerHwnd) {
		printf("Could not create window\n");
		return 0;
	}

	printf("ServerHwnd = %d\n", ServerHwnd);
	return 1;

}


void main(void){

	InitializeCriticalSection(&m_cs);
	if( CreateServerWindow() )
	{
		printf("Server running, try command: vb6_client \"%d,test\"\n", ServerHwnd);
		MSG msg;
		while (PeekMessage(&msg, 0, 0, 0, PM_NOREMOVE)) {
			TranslateMessage(&msg);
			DispatchMessage(&msg);
		}
	}

	printf("Press any key to exit...");
	getch();

}
