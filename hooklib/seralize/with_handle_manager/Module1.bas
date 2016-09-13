Attribute VB_Name = "Module1"
Option Explicit

Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal length As Long)
Declare Function lstrlen Lib "kernel32" Alias "lstrlenA" (ByVal lpString As Long) As Long
Public Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Public Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Public Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long
Public Declare Sub DebugBreak Lib "kernel32" ()

Enum hookType
    ht_jmp = 0
    ht_pushret = 1
    ht_jmp5safe = 2
    ht_jmpderef = 3
    ht_micro = 4
End Enum

Enum hookErrors
    he_None = 0
    he_cantDisasm
    he_cantHook
    he_maxHooks
    he_UnknownHookType
    he_Other
End Enum

'BOOL __stdcall HookFunction(ULONG_PTR OriginalFunction, ULONG_PTR NewFunction, char *name, hookType ht)
Public Declare Function HookFunction Lib "hooklib.dll" (ByVal lpOrgFunc As Long, ByVal lpNewFunc As Long, ByVal name As String, ByVal ht As hookType) As Long

'char* __stdcall GetHookError(void)
Public Declare Function GetHookError Lib "hooklib.dll" () As Long

'void __stdcall SetDebugHandler(ULONG_PTR lpfn); --> callback prototype: void  (__stdcall *debugMsgHandler)(char* msg);
Public Declare Sub SetDebugHandler Lib "hooklib.dll" (ByVal lpCallBack As Long, Optional ByVal logLevel As Long = 0)

'VOID __stdcall DisableHook(ULONG_PTR Function)
Public Declare Function DisableHook Lib "hooklib.dll" (ByVal lpOrgFunc As Long) As Long

'VOID __stdcall EnableHook(ULONG_PTR Function)
Public Declare Function EnableHook Lib "hooklib.dll" (ByVal lpOrgFunc As Long) As Long

'VOID __stdcall RemoveHook(ULONG_PTR Function)
Public Declare Function RemoveHook Lib "hooklib.dll" (ByVal lpOrgFunc As Long) As Long

Public Declare Function CallOriginal Lib "hooklib.dll" ( _
    ByVal lpOrgFunc As Long, _
    Optional ByVal arg1 As Long = &HDEADBEEF, _
    Optional ByVal arg2 As Long = &HDEADBEEF, _
    Optional ByVal arg3 As Long = &HDEADBEEF, _
    Optional ByVal arg4 As Long = &HDEADBEEF, _
    Optional ByVal arg5 As Long = &HDEADBEEF, _
    Optional ByVal arg6 As Long = &HDEADBEEF, _
    Optional ByVal arg7 As Long = &HDEADBEEF, _
    Optional ByVal arg8 As Long = &HDEADBEEF, _
    Optional ByVal arg9 As Long = &HDEADBEEF, _
    Optional ByVal arg10 As Long = &HDEADBEEF _
) As Long

Public Declare Sub RemoveAllHooks Lib "hooklib.dll" ()
Public Declare Sub UnInitilizeHookLib Lib "hooklib.dll" Alias "UnInitilize" ()

Const LANG_US = &H409
Global hHookLib As Long
Global lpWriteFile As Long
Global lpCloseHandle As Long
Global lpCreateFileA As Long
Global lpReadFile As Long
Global lpGetFileType As Long

Global Const trigger = "c:\HOOKME"

Public Type OVERLAPPED
    ternal As Long
    ternalHigh As Long
    offset As Long
    OffsetHigh As Long
    hEvent As Long
End Type

Public Type SECURITY_ATTRIBUTES
    nLength As Long
    lpSecurityDescriptor As Long
    bInheritHandle As Long
End Type

'Public Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" ( _
'    ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, _
'    lpSecurityAttributes As SECURITY_ATTRIBUTES, ByVal dwCreationDisposition As Long, _
'    ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long _
') As Long

'we have modified some args since we are using it as a pass through from the hook...
Private Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" ( _
    ByVal lpFileName As Long, _
    ByVal dwDesiredAccess As Long, _
    ByVal dwShareMode As Long, _
    ByVal lpSecurityAttributes As Long, _
    ByVal dwCreationDisposition As Long, _
    ByVal dwFlagsAndAttributes As Long, _
    ByVal hTemplateFile As Long _
) As Long

Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long

Private Declare Function WriteFile Lib "kernel32" ( _
    ByVal hFile As Long, _
    ByVal lpBuffer As Long, _
    ByVal nNumberOfBytesToWrite As Long, _
    lpNumberOfBytesWritten As Long, _
    ByVal lpOverlapped As Long _
) As Long

 

Public manager As New CHandleManager

Function logit(msg)
    MsgBox msg
    Form1.List1.AddItem msg
End Function

'void  (__stdcall *debugMsgHandler)(char* msg);
Public Sub DebugMsgHandler(ByVal lpMsg As Long)

    Dim msg As String
    Dim tmp() As String
    Dim x
    
    msg = CStringToVBString(lpMsg)
    
    If InStr(msg, vbTab) > 0 Then msg = Replace(msg, vbTab, "    ")
    
    If InStr(msg, vbLf) Then
        msg = Replace(msg, vbCr, Empty)
        tmp() = Split(msg, vbLf)
        For Each x In tmp
            If Len(Trim(x)) > 0 Then Form1.List1.AddItem "DebugMsg: " & x
        Next
    Else
        Form1.List1.AddItem "DebugMsg: " & msg
    End If
    
End Sub

'HANDLE WINAPI CreateFile(
'  _In_     LPCTSTR               lpFileName,
'  _In_     DWORD                 dwDesiredAccess,
'  _In_     DWORD                 dwShareMode,
'  _In_opt_ LPSECURITY_ATTRIBUTES lpSecurityAttributes,
'  _In_     DWORD                 dwCreationDisposition,
'  _In_     DWORD                 dwFlagsAndAttributes,
'  _In_opt_ HANDLE                hTemplateFile
');

Function RandomNum()
    Randomize
    Dim tmp
    tmp = Round(Timer * Now * Rnd(), 0)
    RandomNum = tmp
End Function

'we hook createfile so we can track what system file handle belongs to which file by name.
'we passthrough the args unchanged to the windows api using a modified declaration
Function my_CreateFile(ByVal lpFileName As Long, ByVal access As Long, ByVal shareMode As Long, ByVal sec As Long, ByVal cdisp As Long, ByVal flags As Long, ByVal template As Long) As Long
    
    Dim h As Long
    Dim file As String
    
    file = CStringToVBString(lpFileName)
    logit "my_CreateFile called for: " & file
    
    'runtime needs the file to actually exist we must hook more i guess...
    If file = trigger Then
        Do
            h = RandomNum
        Loop While manager.HandleExists(h)
        manager.Add file, h
        my_CreateFile = h
    Else
        my_CreateFile = CallOriginal(lpCreateFileA, lpFileName, access, shareMode, sec, cdisp, flags, template)
    End If
    
    'h = CallOriginal(lpCreateFileA, lpFileName, access, shareMode, sec, cdisp, flags, template)
    'manager.Add file, h
    'my_CreateFile = h
    
    
    'DebugBreak
    
End Function

'DWORD WINAPI GetFileType(
'  _In_ HANDLE hFile
');
Function my_GetFileType(ByVal hObject As Long) As Long
    
    logit "GetFileType " & hObject
    
    If manager.HandleExists(hObject) Then
        my_GetFileType = 1 'disk file
    Else
        my_GetFileType = CallOriginal(lpGetFileType, hObject)
    End If
    
End Function

'BOOL WINAPI CloseHandle(
'  _In_ HANDLE hObject
');

Function my_CloseHandle(ByVal hObject As Long) As Long
    logit "CloseHandle " & hObject
    my_CloseHandle = CallOriginal(lpCloseHandle, hObject)
    manager.Remove hObject
End Function
 
'If the function succeeds, the return value is nonzero (TRUE).
'BOOL WINAPI WriteFile(
'  _In_        HANDLE       hFile,
'  _In_        LPCVOID      lpBuffer,
'  _In_        DWORD        nNumberOfBytesToWrite,
'  _Out_opt_   LPDWORD      lpNumberOfBytesWritten,
'  _Inout_opt_ LPOVERLAPPED lpOverlapped
');

Function my_WriteFile(ByVal hFile As Long, ByVal lpBuf As Long, ByVal nWriteLength As Long, nBytesWritten As Long, lpOverlapped As Long) As Long
        
    Dim h As CFileHandle
    
    'logit "WriteFile " & hFile & " buf: " & lpBuf & " size: " & nWriteLength
    
    If manager.HandleExists(hFile) Then
        Set h = manager.GetHandle(hFile)
        If h.isRedirected Then
            my_WriteFile = HandleRedirection(h, hFile, lpBuf, nWriteLength, nBytesWritten)
            Exit Function
        End If
    End If
           
    my_WriteFile = CallOriginal(lpWriteFile, hFile, lpBuf, nWriteLength, nBytesWritten, lpOverlapped)
    
End Function

Function HandleRedirection(h As CFileHandle, hFile As Long, lpBuf As Long, nLength As Long, ByRef nBytesHandled As Long, Optional isWrite As Boolean = True) As Long
    
    HandleRedirection = 1
    nBytesHandled = nLength
    
    If h.RedirectTo = rt_memory Then
        If isWrite Then
            h.AddData lpBuf, nLength
        Else
            h.ReadData lpBuf, nLength
        End If
    End If
    
    
End Function

'If the function succeeds, the return value is nonzero (TRUE).
'Public Declare Function ReadFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToRead As Long, lpNumberOfBytesRead As Long, lpOverlapped As OVERLAPPED) As Long

Public Function my_ReadFile(ByVal hFile As Long, ByVal lpBuffer As Long, ByVal nNumberOfBytesToRead As Long, lpNumberOfBytesRead As Long, lpOverlapped As Long) As Long
    
     Dim h As CFileHandle
    
    logit "ReadFile " & hFile & " buf: " & Hex(lpBuffer) & " size: " & nNumberOfBytesToRead
    
    If manager.HandleExists(hFile) Then
        Set h = manager.GetHandle(hFile)
        If h.isRedirected Then
            my_ReadFile = HandleRedirection(h, hFile, lpBuffer, nNumberOfBytesToRead, lpNumberOfBytesRead, False)
            Exit Function
        End If
    End If
           
    my_ReadFile = CallOriginal(lpReadFile, hFile, lpBuffer, nNumberOfBytesToRead, lpNumberOfBytesRead, lpOverlapped)

    
End Function




Function CStringToVBString(lpCstr As Long) As String

    Dim x As Long
    Dim sBuffer As String
    Dim lpBuffer As Long
    Dim b() As Byte
    
    If lpCstr <> 0 Then
        x = lstrlen(lpCstr)
        If x > 0 Then
            ReDim b(x)
            CopyMemory ByVal VarPtr(b(0)), ByVal lpCstr, x
            CStringToVBString = StrConv(b, vbUnicode, LANG_US)
        End If
    End If
    
    CStringToVBString = Replace(CStringToVBString, Chr(0), Empty) 'just in case..
    
End Function




Function AryIsEmpty(ary) As Boolean
  On Error GoTo oops
  Dim i As Long
  
    i = UBound(ary)  '<- throws error if not initalized
    AryIsEmpty = False
  Exit Function
oops: AryIsEmpty = True
End Function




Function HexDump(ByVal str, Optional hexOnly = 0) As String
    Dim s() As String, chars As String, tmp As String
    On Error Resume Next
    Dim ary() As Byte
    Dim offset As Long
    Const LANG_US = &H409
    Dim i As Long, tt, h, x

    offset = 0
    str = " " & str
    ary = StrConv(str, vbFromUnicode, LANG_US)
    
    chars = "   "
    For i = 1 To UBound(ary)
        tt = Hex(ary(i))
        If Len(tt) = 1 Then tt = "0" & tt
        tmp = tmp & tt & " "
        x = ary(i)
        'chars = chars & IIf((x > 32 And x < 127) Or x > 191, Chr(x), ".") 'x > 191 causes \x0 problems on non us systems... asc(chr(x)) = 0
        chars = chars & IIf((x > 32 And x < 127), Chr(x), ".")
        If i > 1 And i Mod 16 = 0 Then
            h = Hex(offset)
            While Len(h) < 6: h = "0" & h: Wend
            If hexOnly = 0 Then
                push s, h & "   " & tmp & chars
            Else
                push s, tmp
            End If
            offset = offset + 16
            tmp = Empty
            chars = "   "
        End If
    Next
    'if read length was not mod 16=0 then
    'we have part of line to account for
    If tmp <> Empty Then
        If hexOnly = 0 Then
            h = Hex(offset)
            While Len(h) < 6: h = "0" & h: Wend
            h = h & "   " & tmp
            While Len(h) <= 56: h = h & " ": Wend
            push s, h & chars
        Else
            push s, tmp
        End If
    End If
    
    HexDump = Join(s, vbCrLf)
    
    If hexOnly <> 0 Then
        HexDump = Replace(HexDump, " ", "")
        HexDump = Replace(HexDump, vbCrLf, "")
    End If
    
End Function



Sub push(ary, value) 'this modifies parent ary object
    On Error GoTo init
    Dim x
    
    x = UBound(ary) '<-throws Error If Not initalized
    ReDim Preserve ary(UBound(ary) + 1)
    ary(UBound(ary)) = value
    Exit Sub
init:     ReDim ary(0): ary(0) = value
End Sub
