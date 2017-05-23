Attribute VB_Name = "mdlDeclares"
'*********************************************************************************************
'
' Debugging processes with VB
'
' API Declarations
'
'*********************************************************************************************
'
' Author: Eduardo A. Morcillo
' E-Mail: e_morcillo@yahoo.com
' Web Page: http://www.domaindlx.com/e_morcillo
'
' Distribution: You can freely use this code in your own applications but you
'               can't publish this code in a web site, online service, or any
'               other media, without my express permission.
'
' Usage: at your own risk.
'
' Tested on: Windows 98 + VB5
'
' History:
'           03/09/2000 - This code was released
'
'*********************************************************************************************
Option Explicit

Type SECURITY_ATTRIBUTES
   nLength As Long
   lpSecurityDescriptor As Long
   bInheritHandle As Long
End Type

Type PROCESS_INFORMATION
   hProcess As Long
   hThread As Long
   dwProcessId As Long
   dwThreadId As Long
End Type

Type STARTUPINFO
   cb As Long
   lpReserved As String
   lpDesktop As String
   lpTitle As String
   dwX As Long
   dwY As Long
   dwXSize As Long
   dwYSize As Long
   dwXCountChars As Long
   dwYCountChars As Long
   dwFillAttribute As Long
   dwFlags As Long
   wShowWindow As Integer
   cbReserved2 As Integer
   lpReserved2 As Byte
   hStdInput As Long
   hStdOutput As Long
   hStdError As Long
End Type

Public Const DEBUG_PROCESS = &H1
Public Const DEBUG_ONLY_THIS_PROCESS = &H2

Public Const NORMAL_PRIORITY_CLASS = &H20
Public Const IDLE_PRIORITY_CLASS = &H40
Public Const HIGH_PRIORITY_CLASS = &H80
Public Const REALTIME_PRIORITY_CLASS = &H100

Declare Function CreateProcess Lib "kernel32" Alias "CreateProcessA" (ByVal lpApplicationName As String, ByVal lpCommandLine As String, lpProcessAttributes As Any, lpThreadAttributes As Any, ByVal bInheritHandles As Long, ByVal dwCreationFlags As Long, lpEnvironment As Any, ByVal lpCurrentDriectory As String, lpStartupInfo As STARTUPINFO, lpProcessInformation As PROCESS_INFORMATION) As Long
Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long

Public Const EXCEPTION_DEBUG_EVENT = 1
Public Const CREATE_THREAD_DEBUG_EVENT = 2
Public Const CREATE_PROCESS_DEBUG_EVENT = 3
Public Const EXIT_THREAD_DEBUG_EVENT = 4
Public Const EXIT_PROCESS_DEBUG_EVENT = 5
Public Const LOAD_DLL_DEBUG_EVENT = 6
Public Const UNLOAD_DLL_DEBUG_EVENT = 7
Public Const OUTPUT_DEBUG_STRING_EVENT = 8
Public Const RIP_EVENT = 9

Public Const EXCEPTION_NONCONTINUABLE = 1 ' Noncontinuable exception
Public Const EXCEPTION_MAXIMUM_PARAMETERS = 15

Type EXCEPTION_RECORD
   ExceptionCode As Long
   ExceptionFlags As Long
   pExceptionRecord As Long  ' Pointer to an EXCEPTION_RECORD structure
   ExceptionAddress As Long
   NumberParameters As Long
   ExceptionInformation(EXCEPTION_MAXIMUM_PARAMETERS) As Long
End Type

Type EXCEPTION_DEBUG_INFO
   pExceptionRecord As EXCEPTION_RECORD
   dwFirstChance As Long
End Type

Type CREATE_THREAD_DEBUG_INFO
   hThread As Long
   lpThreadLocalBase As Long
   lpStartAddress As Long
End Type

Type CREATE_PROCESS_DEBUG_INFO
   hFile As Long
   hProcess As Long
   hThread As Long
   lpBaseOfImage As Long
   dwDebugInfoFileOffset As Long
   nDebugInfoSize As Long
   lpThreadLocalBase As Long
   lpStartAddress As Long
   lpImageName As Long
   fUnicode As Integer
End Type

Type EXIT_THREAD_DEBUG_INFO
   dwExitCode As Long
End Type

Type EXIT_PROCESS_DEBUG_INFO
   dwExitCode As Long
End Type

Type LOAD_DLL_DEBUG_INFO
   hFile As Long
   lpBaseOfDll As Long
   dwDebugInfoFileOffset As Long
   nDebugInfoSize As Long
   lpImageName As Long
   fUnicode As Integer
End Type

Type UNLOAD_DLL_DEBUG_INFO
   lpBaseOfDll As Long
End Type

Type OUTPUT_DEBUG_STRING_INFO
   lpDebugStringData As Long
   fUnicode As Integer
   nDebugStringLength As Integer
End Type

Type RIP_INFO
   dwError As Long
   dwType As Long
End Type

Type DEBUG_EVENT
   dwDebugEventCode As Long
   dwProcessId As Long
   dwThreadId As Long
   DEBUG_INFO(0 To 187) As Byte
End Type

Declare Function WaitForDebugEvent Lib "kernel32" (lpDebugEvent As Any, ByVal dwMilliseconds As Long) As Long
Declare Function ContinueDebugEvent Lib "kernel32" (ByVal dwProcessId As Long, ByVal dwThreadId As Long, ByVal dwContinueStatus As Long) As Long

Declare Sub MoveMemory Lib "kernel32" Alias "RtlMoveMemory" (Dest As Any, Src As Any, ByVal bl As Long)

Declare Function ReadProcessMemory Lib "kernel32" (ByVal hProcess As Long, ByVal lpBaseAddress As Long, lpBuffer As Any, ByVal nSize As Long, lpNumberOfBytesWritten As Long) As Long

Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long

Public Const PROCESS_TERMINATE = (&H1)
Public Const PROCESS_CREATE_THREAD = (&H2)
Public Const PROCESS_SET_SESSIONID = (&H4)
Public Const PROCESS_VM_OPERATION = (&H8)
Public Const PROCESS_VM_READ = (&H10)
Public Const PROCESS_VM_WRITE = (&H20)
Public Const PROCESS_DUP_HANDLE = (&H40)
Public Const PROCESS_CREATE_PROCESS = (&H80)
Public Const PROCESS_SET_QUOTA = (&H100)
Public Const PROCESS_SET_INFORMATION = (&H200)
Public Const PROCESS_QUERY_INFORMATION = (&H400)
Public Const DELETE = &H10000
Public Const READ_CONTROL = &H20000
Public Const WRITE_DAC = &H40000
Public Const WRITE_OWNER = &H80000
Public Const SYNCHRONIZE = &H100000
Public Const STANDARD_RIGHTS_REQUIRED = &HF0000
Public Const STANDARD_RIGHTS_READ = READ_CONTROL
Public Const STANDARD_RIGHTS_WRITE = READ_CONTROL
Public Const STANDARD_RIGHTS_EXECUTE = READ_CONTROL
Public Const STANDARD_RIGHTS_ALL = &H1F0000
Public Const SPECIFIC_RIGHTS_ALL = &HFFFF
Public Const PROCESS_ALL_ACCESS = STANDARD_RIGHTS_REQUIRED Or SYNCHRONIZE Or &HFFF

Public Const STATUS_WAIT_0 = &H0
Public Const STATUS_ABANDONED_WAIT_0 = &H80
Public Const STATUS_USER_APC = &HC0
Public Const STATUS_TIMEOUT = &H102
Public Const STATUS_PENDING = &H103
Public Const DBG_CONTINUE = &H10002
Public Const STATUS_SEGMENT_NOTIFICATION = &H40000005
Public Const DBG_TERMINATE_THREAD = &H40010003
Public Const DBG_TERMINATE_PROCESS = &H40010004
Public Const DBG_CONTROL_C = &H40010005
Public Const DBG_CONTROL_BREAK = &H40010008
Public Const STATUS_GUARD_PAGE_VIOLATION = &H80000001
Public Const STATUS_DATATYPE_MISALIGNMENT = &H80000002
Public Const STATUS_BREAKPOINT = &H80000003
Public Const STATUS_SINGLE_STEP = &H80000004
Public Const DBG_EXCEPTION_NOT_HANDLED = &H80010001
Public Const STATUS_ACCESS_VIOLATION = &HC0000005
Public Const STATUS_IN_PAGE_ERROR = &HC0000006
Public Const STATUS_INVALID_HANDLE = &HC0000008
Public Const STATUS_NO_MEMORY = &HC0000017
Public Const STATUS_ILLEGAL_INSTRUCTION = &HC000001D
Public Const STATUS_NONCONTINUABLE_EXCEPTION = &HC0000025
Public Const STATUS_INVALID_DISPOSITION = &HC0000026
Public Const STATUS_ARRAY_BOUNDS_EXCEEDED = &HC000008C
Public Const STATUS_FLOAT_DENORMAL_OPERAND = &HC000008D
Public Const STATUS_FLOAT_DIVIDE_BY_ZERO = &HC000008E
Public Const STATUS_FLOAT_INEXACT_RESULT = &HC000008F
Public Const STATUS_FLOAT_INVALID_OPERATION = &HC0000090
Public Const STATUS_FLOAT_OVERFLOW = &HC0000091
Public Const STATUS_FLOAT_STACK_CHECK = &HC0000092
Public Const STATUS_FLOAT_UNDERFLOW = &HC0000093
Public Const STATUS_INTEGER_DIVIDE_BY_ZERO = &HC0000094
Public Const STATUS_INTEGER_OVERFLOW = &HC0000095
Public Const STATUS_PRIVILEGED_INSTRUCTION = &HC0000096
Public Const STATUS_STACK_OVERFLOW = &HC00000FD
Public Const STATUS_CONTROL_C_EXIT = &HC000013A
Public Const STATUS_FLOAT_MULTIPLE_FAULTS = &HC00002B4
Public Const STATUS_FLOAT_MULTIPLE_TRAPS = &HC00002B5
Public Const STATUS_REG_NAT_CONSUMPTION = &HC00002C9

Public Const EXCEPTION_ACCESS_VIOLATION = STATUS_ACCESS_VIOLATION
Public Const EXCEPTION_DATATYPE_MISALIGNMENT = STATUS_DATATYPE_MISALIGNMENT
Public Const EXCEPTION_BREAKPOINT = STATUS_BREAKPOINT
Public Const EXCEPTION_SINGLE_STEP = STATUS_SINGLE_STEP
Public Const EXCEPTION_ARRAY_BOUNDS_EXCEEDED = STATUS_ARRAY_BOUNDS_EXCEEDED
Public Const EXCEPTION_FLT_DENORMAL_OPERAND = STATUS_FLOAT_DENORMAL_OPERAND
Public Const EXCEPTION_FLT_DIVIDE_BY_ZERO = STATUS_FLOAT_DIVIDE_BY_ZERO
Public Const EXCEPTION_FLT_INEXACT_RESULT = STATUS_FLOAT_INEXACT_RESULT
Public Const EXCEPTION_FLT_INVALID_OPERATION = STATUS_FLOAT_INVALID_OPERATION
Public Const EXCEPTION_FLT_OVERFLOW = STATUS_FLOAT_OVERFLOW
Public Const EXCEPTION_FLT_STACK_CHECK = STATUS_FLOAT_STACK_CHECK
Public Const EXCEPTION_FLT_UNDERFLOW = STATUS_FLOAT_UNDERFLOW
Public Const EXCEPTION_INT_DIVIDE_BY_ZERO = STATUS_INTEGER_DIVIDE_BY_ZERO
Public Const EXCEPTION_INT_OVERFLOW = STATUS_INTEGER_OVERFLOW
Public Const EXCEPTION_PRIV_INSTRUCTION = STATUS_PRIVILEGED_INSTRUCTION
Public Const EXCEPTION_IN_PAGE_ERROR = STATUS_IN_PAGE_ERROR
Public Const EXCEPTION_ILLEGAL_INSTRUCTION = STATUS_ILLEGAL_INSTRUCTION
Public Const EXCEPTION_NONCONTINUABLE_EXCEPTION = STATUS_NONCONTINUABLE_EXCEPTION
Public Const EXCEPTION_STACK_OVERFLOW = STATUS_STACK_OVERFLOW
Public Const EXCEPTION_INVALID_DISPOSITION = STATUS_INVALID_DISPOSITION
Public Const EXCEPTION_GUARD_PAGE = STATUS_GUARD_PAGE_VIOLATION
Public Const EXCEPTION_INVALID_HANDLE = STATUS_INVALID_HANDLE

Declare Function lstrlenA Lib "kernel32" (ByVal Str As Any) As Long
Declare Function lstrcpyA Lib "kernel32" (ByVal Dest As Any, ByVal Src As Any) As Long

Public Const MAXIMUM_SUPPORTED_EXTENSION = 512
Public Const SIZE_OF_80387_REGISTERS = 80

Type FLOATING_SAVE_AREA
   ControlWord As Long
   StatusWord As Long
   TagWord As Long
   ErrorOffset As Long
   ErrorSelector As Long
   DataOffset As Long
   DataSelector As Long
   RegisterArea(1 To SIZE_OF_80387_REGISTERS) As Byte
   Cr0NpxState As Long
End Type

'
' Context Frame
'
'  This frame has a several purposes: 1) it is used as an argument to
'  NtContinue, 2) is is used to constuct a call frame for APC delivery,
'  and 3) it is used in the user level thread creation routines.
'
'  The layout of the record conforms to a standard call frame.
'

Type CONTEXT

    '
    ' The flags values within this flag control the contents of
    ' a CONTEXT record.
    '
    ' If the context record is used as an input parameter, then
    ' for each portion of the context record controlled by a flag
    ' whose value is set, it is assumed that that portion of the
    ' context record contains valid context. If the context record
    ' is being used to modify a threads context, then only that
    ' portion of the threads context will be modified.
    '
    ' If the context record is used as an IN OUT parameter to capture
    ' the context of a thread, then only those portions of the thread's
    ' context corresponding to set flags will be returned.
    '
    ' The context record is never used as an OUT only parameter.
    '

    ContextFlags As Long

    '
    ' This section is specified/returned if CONTEXT_DEBUG_REGISTERS is
    ' set in ContextFlags.  Note that CONTEXT_DEBUG_REGISTERS is NOT
    ' included in CONTEXT_FULL.
    '

    Dr0 As Long
    Dr1 As Long
    Dr2 As Long
    Dr3 As Long
    Dr6 As Long
    Dr7 As Long

    '
    ' This section is specified/returned if the
    ' ContextFlags word contians the flag CONTEXT_FLOATING_POINT.
    '

    FloatSave As FLOATING_SAVE_AREA

    '
    ' This section is specified/returned if the
    ' ContextFlags word contians the flag CONTEXT_SEGMENTS.
    '

    SegGs As Long
    SegFs As Long
    SegEs As Long
    SegDs As Long

    '
    ' This section is specified/returned if the
    ' ContextFlags word contians the flag CONTEXT_INTEGER.
    '

    Edi As Long
    Esi As Long
    Ebx As Long
    Edx As Long
    Ecx As Long
    Eax As Long

    '
    ' This section is specified/returned if the
    ' ContextFlags word contians the flag CONTEXT_CONTROL.
    '

    Ebp As Long
    Eip As Long
    SegCs As Long              ' MUST BE SANITIZED
    EFlags As Long             ' MUST BE SANITIZED
    Esp As Long
    SegSs As Long

    '
    ' This section is specified/returned if the ContextFlags word
    ' contains the flag CONTEXT_EXTENDED_REGISTERS.
    ' The format and contexts are processor specific
    '

    ExtendedRegisters(1 To MAXIMUM_SUPPORTED_EXTENSION) As Byte

End Type

Declare Function GetThreadContext Lib "kernel32" ( _
   ByVal hThread As Long, _
   lpContext As CONTEXT) As Long

Declare Function SetThreadContext Lib "kernel32" ( _
   ByVal hThread As Long, _
   lpContext As CONTEXT) As Long

'
' StrFromPtrPtrPID
'
' Returns a string given a pointer to a pointer
' to it and the process ID that owns it.
'
' Parameters:
'
' lPtr      Pointer to pointer to the string in the given process.
' ProcID    Process ID.
' bUnicode  If is set to True the string is Unicode.
' lStrLen   Optional. Length of the string.
'
Public Function StrFromPtrPtrPID(ByVal lPtr As Long, ByVal ProcID As Long, ByVal bUnicode As Boolean, Optional ByVal lStrLen As Long) As String
Dim lWritten As Long, lpStr As Long, hProc As Long
   
   On Error Resume Next
   
   If lPtr <> 0 Then
   
      ' Open the process
      hProc = OpenProcess(PROCESS_ALL_ACCESS, False, ProcID)
      
      ' Read the pointer to the string
      ' from the process memory
      ReadProcessMemory hProc, lPtr, lpStr, Len(lpStr), lWritten
      
      ' Read the string from the
      ' process memory
      If lStrLen Then
         StrFromPtrPtrPID = Space$(lStrLen + 2)
      Else
         StrFromPtrPtrPID = Space$(512)
      End If
      
      If lpStr <> 0 Then
      
         If bUnicode Then
            ReadProcessMemory hProc, lpStr, ByVal StrPtr(StrFromPtrPtrPID), LenB(StrFromPtrPtrPID), lWritten
         Else
            ReadProcessMemory hProc, lpStr, ByVal StrFromPtrPtrPID, Len(StrFromPtrPtrPID), lWritten
         End If
         
         ' Trim anything after the first
         ' null character
         StrFromPtrPtrPID = Left$(StrFromPtrPtrPID, InStr(StrFromPtrPtrPID, vbNullChar) - 1)
      
      End If
   
      ' Close the process handle
      CloseHandle hProc
      
   End If
   
End Function



'
' StrFromPtrPID
'
' Returns a string given a pointer it
' and the process ID that owns it.
'
' Parameters:
'
' lPtr      Pointer to pointer to the string in the given process
' ProcID    Process ID
' bUnicode  If is set to True the string is Unicode.
' lStrLen   Optional. Length of the string.
'
Public Function StrFromPtrPID(ByVal lPtr As Long, ByVal ProcID As Long, ByVal bUnicode As Boolean, Optional ByVal lStrLen As Long) As String
Dim lWritten As Long
Dim hProc As Long

   On Error Resume Next
   
   If lPtr <> 0 Then
   
      hProc = OpenProcess(PROCESS_ALL_ACCESS, False, ProcID)
      
      ' Read the string from the
      ' process memory
      If lStrLen Then
         StrFromPtrPID = Space$(lStrLen + 2)
      Else
         StrFromPtrPID = Space$(512)
      End If
      
      If bUnicode Then
         ReadProcessMemory hProc, lPtr, ByVal StrPtr(StrFromPtrPID), LenB(StrFromPtrPID), lWritten
      Else
         ReadProcessMemory hProc, lPtr, ByVal StrFromPtrPID, Len(StrFromPtrPID), lWritten
      End If
         
      ' Trim anything after the first
      ' null character
      StrFromPtrPID = Left$(StrFromPtrPID, InStr(StrFromPtrPID, vbNullChar) - 1)
   
      CloseHandle hProc
      
   End If
   
End Function




