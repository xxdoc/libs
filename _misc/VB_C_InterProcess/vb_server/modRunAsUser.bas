Attribute VB_Name = "modRunAsUser"
Option Explicit

'RunAsDesktopUser:
'   Start a process as the currently logged in user from an elevated process
'
'ported to vb6 from C code here:
'    https://blogs.msdn.microsoft.com/aaron_margosis/2009/06/06/faq-how-do-i-start-a-program-as-the-desktop-user-from-an-elevated-app/

Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessID As Long) As Long
Private Declare Function OpenProcessToken Lib "advapi32.dll" (ByVal ProcessHandle As Long, ByVal DesiredAccess As Long, TokenHandle As Long) As Long
Private Declare Function GetCurrentProcess Lib "kernel32.dll" () As Long
Private Declare Function LookupPrivilegeValue Lib "advapi32.dll" Alias "LookupPrivilegeValueA" (ByVal lpSystemName As String, ByVal lpName As String, lpLuid As LUID) As Long
Private Declare Function AdjustTokenPrivileges Lib "advapi32.dll" (ByVal TokenHandle As Long, ByVal DisableAllPrivileges As Long, ByRef NewState As TOKEN_PRIVILEGES, ByVal BufferLength As Long, ByRef PreviousState As TOKEN_PRIVILEGES, ByRef ReturnLength As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function GetWindowThreadProcessId Lib "user32.dll" (ByVal hwnd As Long, ByRef hINst As Long) As Long

Private Type SECURITY_ATTRIBUTES
    nLength As Long
    lpSecurityDescriptor As Long
    bInheritHandle As Long
End Type

Private Enum SECURITY_IMPERSONATION_LEVEL
    SecurityAnonymous
    SecurityIdentification
    SecurityImpersonation
    SecurityDelegation
End Enum

Enum TOKEN_TYPE
        TokenPrimary = 1
        TokenImpersonation = 2
End Enum
    
Private Declare Function DuplicateTokenEx Lib "advapi32.dll" ( _
    ByVal hExistingToken As Long, _
    ByVal dwDesiredAccess As Long, _
    ByRef lpTokenAttributes As Any, _
    ByVal ImpersonationLevel As Long, _
    ByVal TokenType As Long, _
    ByRef phNewToken As Long _
) As Long

Private Type LUID
    LowPart As Long
    HighPart As Long
End Type

Private Type LUID_AND_ATTRIBUTES
   pLuid As LUID
   Attributes As Long
End Type

Private Const ANYSIZE_ARRAY = 1

'Private Type TOKEN_PRIVILEGES
'   PrivilegeCount As Long
'   Privileges(ANYSIZE_ARRAY) As LUID_AND_ATTRIBUTES
'End Type

Private Type TOKEN_PRIVILEGES
    PrivilegeCount As Long
    TheLuid As LUID
    Attributes As Long
End Type

Private Enum ProcessAccessTypes
  PROCESS_TERMINATE = (&H1)
  PROCESS_CREATE_THREAD = (&H2)
  PROCESS_SET_SESSIONID = (&H4)
  PROCESS_VM_OPERATION = (&H8)
  PROCESS_VM_READ = (&H10)
  PROCESS_VM_WRITE = (&H20)
  PROCESS_DUP_HANDLE = (&H40)
  PROCESS_CREATE_PROCESS = (&H80)
  PROCESS_SET_QUOTA = (&H100)
  PROCESS_SET_INFORMATION = (&H200)
  PROCESS_QUERY_INFORMATION = (&H400)
'  STANDARD_RIGHTS_REQUIRED = &HF0000
  SYNCHRONIZE = &H100000
  'PROCESS_ALL_ACCESS = (STANDARD_RIGHTS_REQUIRED Or SYNCHRONIZE Or &HFFF)
End Enum

'HWND WINAPI GetShellWindow(void) User32.dll
Private Declare Function GetShellWindow Lib "user32.dll" () As Long

' TOKEN_QUERY | TOKEN_ASSIGN_PRIMARY | TOKEN_DUPLICATE | TOKEN_ADJUST_DEFAULT | TOKEN_ADJUST_SESSIONID
Public Const TOKEN_QUERY As Long = &H8
Public Const TOKEN_ASSIGN_PRIMARY As Long = 1
Public Const TOKEN_DUPLICATE   As Long = 2
Public Const TOKEN_ADJUST_DEFAULT As Long = &H80
Public Const TOKEN_ADJUST_SESSIONID As Long = &H100
Public Const TOKEN_ADJUST_PRIVILEGES As Long = &H20
       
'vista+ Advapi32.lib
'https://msdn.microsoft.com/en-us/library/windows/desktop/ms682434(v=vs.85).aspx
'
'BOOL WINAPI CreateProcessWithTokenW(
'  _In_        HANDLE                hToken,
'  _In_        DWORD                 dwLogonFlags,
'  _In_opt_    LPCWSTR               lpApplicationName,
'  _Inout_opt_ LPWSTR                lpCommandLine,
'  _In_        DWORD                 dwCreationFlags,
'  _In_opt_    LPVOID                lpEnvironment,
'  _In_opt_    LPCWSTR               lpCurrentDirectory,
'  _In_        LPSTARTUPINFOW        lpStartupInfo,
'  _Out_       LPPROCESS_INFORMATION lpProcessInfo
')

Private Declare Function CreateProcessWithTokenW Lib "advapi32.dll" ( _
    ByVal hExistingToken As Long, _
    ByVal dwLogonFlags As Long, _
    ByVal lpApplicationName As Long, _
    ByVal lpCommandLine As Long, _
    ByVal dwCreationFlags As Long, _
    ByVal lpEnvironment As Long, _
    ByVal lpCurrentDirectory As Long, _
    ByRef lpStartupInfo As STARTUPINFO, _
    ByRef lpProcessInfo As PROCESS_INFORMATION) As Long

Public Type STARTUPINFO
    cb                  As Long
    lpReserved          As String
    lpDesktop           As String
    lpTitle             As String
    dwX                 As Long
    dwY                 As Long
    dwXSize             As Long
    dwYSize             As Long
    dwXCountChars       As Long
    dwYCountChars       As Long
    dwFillAttribute     As Long
    dwFlags             As Long
    wShowWindow         As Integer
    cbReserved2         As Integer
    lpReserved2         As Long
    hStdInput           As Long
    hStdOutput          As Long
    hStdError           As Long
End Type

Public Type PROCESS_INFORMATION
    hProcess            As Long
    hThread             As Long
    dwProcessID         As Long
    dwThreadID          As Long
End Type

Public Const SE_INCREASE_QUOTA_NAME As String = "SeIncreaseQuotaPrivilege"
Public Const SE_PRIVILEGE_ENABLED As Long = 2
Public Const ERROR_SUCCESS = 0

Private Declare Function GetLastError Lib "kernel32.dll" () As Long

'// Definition of the function this sample is all about.
'// The szApp, szCmdLine, szCurrDir, si, and pi parameters are passed directly to CreateProcessWithTokenW.
'// sErrorInfo returns text describing any error that occurs.
'// Returns "true" on success, "false" on any error.
'// It is up to the caller to close the HANDLEs returned in the PROCESS_INFORMATION structure.

Public si As STARTUPINFO
Public pi As PROCESS_INFORMATION
Public ErrInfo As String

Function RunAsDesktopUser(Optional szApp As String, Optional szCmdLine As String, Optional szCurrDir As String) As Boolean

    Dim hShellProcess As Long, hShellProcessToken As Long, hPrimaryToken  As Long, dwTokenRights As Long
    Dim hwnd As Long, dwPID As Long, ret As Boolean, dwLastErr As Long, retval As Boolean, lret As Long
    Dim tkp As TOKEN_PRIVILEGES, prev As TOKEN_PRIVILEGES
    
    ErrInfo = Empty
    
    '// Enable SeIncreaseQuotaPrivilege in this process.  (This won't work if current process is not elevated.)
    Dim hProcessToken As Long
    If OpenProcessToken(GetCurrentProcess(), TOKEN_ADJUST_PRIVILEGES, hProcessToken) = 0 Then
        dwLastErr = Err.LastDllError
        ErrInfo = "OpenProcessToken failed:  " & dwLastErr
        Exit Function
    End If
        
    tkp.PrivilegeCount = 1
    LookupPrivilegeValue "", SE_INCREASE_QUOTA_NAME, tkp.TheLuid
    tkp.Attributes = SE_PRIVILEGE_ENABLED
    
    AdjustTokenPrivileges hProcessToken, False, tkp, Len(tkp), prev, lret
    dwLastErr = GetLastError()
    CloseHandle hProcessToken
    
    If ERROR_SUCCESS <> dwLastErr Then
        ErrInfo = "AdjustTokenPrivileges failed:  " & dwLastErr
        Exit Function
    End If
    
    '// Get an HWND representing the desktop shell.
    '// CAVEATS:  This will fail if the shell is not running (crashed or terminated), or the default shell has been
    '// replaced with a custom shell.  This also won't return what you probably want if Explorer has been terminated and
    '// restarted elevated.
    hwnd = GetShellWindow()
    If 0 = hwnd Then
        ErrInfo = "No desktop shell is present"
        Exit Function
    End If

    '// Get the PID of the desktop shell process.
    GetWindowThreadProcessId hwnd, dwPID
    If 0 = dwPID Then
        ErrInfo = "Unable to get PID of desktop shell."
        Exit Function
    End If

    '// Open the desktop shell process in order to query it (get the token)
    hShellProcess = OpenProcess(PROCESS_QUERY_INFORMATION, False, dwPID)
    If hShellProcess = 0 Then
        dwLastErr = GetLastError()
        ErrInfo = "Can't open desktop shell process:  " & dwLastErr
        Exit Function
    End If

    '// From this point down, we have handles to close, so make sure to clean up.
 
    '// Get the process token of the desktop shell.
    ret = OpenProcessToken(hShellProcess, TOKEN_DUPLICATE, hShellProcessToken)
    If ret = 0 Then
        dwLastErr = GetLastError()
        ErrInfo = "Can't get process token of desktop shell:  " & dwLastErr
        GoTo cleanup
    End If

    '// Duplicate the shell's process token to get a primary token.
    '// Based on experimentation, this is the minimal set of rights required for CreateProcessWithTokenW (contrary to current documentation).
    dwTokenRights = TOKEN_QUERY Or TOKEN_ASSIGN_PRIMARY Or TOKEN_DUPLICATE Or TOKEN_ADJUST_DEFAULT Or TOKEN_ADJUST_SESSIONID
    ret = DuplicateTokenEx(hShellProcessToken, dwTokenRights, 0, SecurityImpersonation, TokenPrimary, hPrimaryToken)
    If ret = 0 Then
        dwLastErr = GetLastError()
        ErrInfo = "Can't get primary token:  " & dwLastErr
        GoTo cleanup
    End If

    '// Start the target process with the new token.
    Dim appPtr As Long, cmdLinePtr As Long, curDirPtr As Long
    If Len(szApp) > 0 Then appPtr = StrPtr(szApp)
    If Len(szCmdLine) > 0 Then cmdLinePtr = StrPtr(szCmdLine)
    If Len(szCurrDir) > 0 Then curDirPtr = StrPtr(szCurrDir)
    
    ret = CreateProcessWithTokenW(hPrimaryToken, 0, appPtr, cmdLinePtr, 0, 0, curDirPtr, si, pi)
        
    If ret = 0 Then
        dwLastErr = GetLastError()
        ErrInfo = "CreateProcessWithTokenW failed:  " & dwLastErr
        GoTo cleanup
    End If

    retval = True

cleanup:
    '// Clean up resources
    CloseHandle hShellProcessToken
    CloseHandle hPrimaryToken
    CloseHandle hShellProcess
    RunAsDesktopUser = retval
End Function






