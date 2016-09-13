Attribute VB_Name = "mSer"
Option Explicit

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal length As Long)
Private Declare Function lstrlen Lib "kernel32" Alias "lstrlenA" (ByVal lpString As Long) As Long
Private Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long

Private Enum hookType
    ht_jmp = 0
    ht_pushret = 1
    ht_jmp5safe = 2
    ht_jmpderef = 3
    ht_micro = 4
End Enum

Private Enum hookErrors
    he_None = 0
    he_cantDisasm
    he_cantHook
    he_maxHooks
    he_UnknownHookType
    he_Other
End Enum

Private Declare Function HookFunction Lib "hooklib.dll" (ByVal lpOrgFunc As Long, ByVal lpNewFunc As Long, ByVal name As String, ByVal ht As hookType) As Long
Private Declare Function GetHookError Lib "hooklib.dll" () As Long
Private Declare Sub SetDebugHandler Lib "hooklib.dll" (ByVal lpCallBack As Long, Optional ByVal logLevel As Long = 0)
Private Declare Function DisableHook Lib "hooklib.dll" (ByVal lpOrgFunc As Long) As Long
Private Declare Function EnableHook Lib "hooklib.dll" (ByVal lpOrgFunc As Long) As Long
Private Declare Function RemoveHook Lib "hooklib.dll" (ByVal lpOrgFunc As Long) As Long

Private Declare Function CallOriginal Lib "hooklib.dll" ( _
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
Private Declare Sub UnInitilizeHookLib Lib "hooklib.dll" Alias "UnInitilize" ()

Private hHookLib As Long
Private lpWriteFile As Long
Private lpCloseHandle As Long
Private lpCreateFileA As Long
Private lpReadFile As Long
Private lpGetFileType As Long

Private Const trigger = "c:\HOOKME" 'never created we intercept it..
Private buf() As Byte               'read/write buffer
Private pointer As Long             'data write pointer \_ could combine...
Private readPointer As Long         'data read pointer  /
Private Const hFile = &HCAFEBABE    'our fake file handle used as a marker
Private vb_hFile As Long            'vb freefile used internally...

Public isInit As Boolean

Function Init(List1 As ListBox) As Boolean

    Dim h As Long
    Dim ret As Long
    Dim lpMsg As Long
    
    If hHookLib <> 0 Then Exit Function 'already called...
    
    hHookLib = LoadLibrary("hooklib.dll")
    If hHookLib = 0 Then hHookLib = LoadLibrary(App.Path & "\hooklib.dll")
    If hHookLib = 0 Then hHookLib = LoadLibrary(App.Path & "\..\hooklib.dll")
    
    If hHookLib = 0 Then
        List1.AddItem "Could not find hooklib.dll compile or download from github."
        Exit Function
    End If
    
    List1.AddItem "Hooklib base address: 0x" & Hex(hHookLib)
    
    'this is optional but were debugging the library so..
    SetDebugHandler AddressOf DebugMsgHandler, 1
    
    h = LoadLibrary("kernel32.dll")
    lpCloseHandle = GetProcAddress(h, "CloseHandle")
    lpCreateFileA = GetProcAddress(h, "CreateFileA")
    lpWriteFile = GetProcAddress(h, "WriteFile")
    lpReadFile = GetProcAddress(h, "ReadFile")
    lpGetFileType = GetProcAddress(h, "GetFileType")
    
    If lpCloseHandle = 0 Or lpCreateFileA = 0 Or lpWriteFile = 0 Or lpReadFile = 0 Or lpGetFileType = 0 Then
        List1.AddItem "GetProcAddress failed for one of the functions??"
        Exit Function
    End If
        
    ret = HookFunction(lpCloseHandle, AddressOf my_CloseHandle, "CloseHandle", ht_jmp)
    If ret = 0 Then
        lpMsg = GetHookError()
        List1.AddItem "Hook Function failed msg: " & CStringToVBString(lpMsg)
        Exit Function
    Else
        DisableHook lpCloseHandle
    End If
    
    ret = HookFunction(lpWriteFile, AddressOf my_WriteFile, "WriteFile", ht_jmp)
    If ret = 0 Then
        lpMsg = GetHookError()
        List1.AddItem "Hook Function failed msg: " & CStringToVBString(lpMsg)
        Exit Function
    Else
        DisableHook lpWriteFile
    End If
    
    ret = HookFunction(lpCreateFileA, AddressOf my_CreateFile, "CreateFileA", ht_jmp)
    If ret = 0 Then
        lpMsg = GetHookError()
        List1.AddItem "Hook Function failed msg: " & CStringToVBString(lpMsg)
        Exit Function
    Else
        DisableHook lpCreateFileA
    End If
    
    ret = HookFunction(lpReadFile, AddressOf my_ReadFile, "ReadFile", ht_jmp)
    If ret = 0 Then
        lpMsg = GetHookError()
        List1.AddItem "Hook Function failed msg: " & CStringToVBString(lpMsg)
        Exit Function
    Else
        DisableHook lpReadFile
    End If
    
    ret = HookFunction(lpGetFileType, AddressOf my_GetFileType, "GetFileType", ht_jmp)
    If ret = 0 Then
        lpMsg = GetHookError()
        List1.AddItem "Hook Function failed msg: " & CStringToVBString(lpMsg)
        Exit Function
    Else
        DisableHook lpGetFileType
    End If
    
    List1.AddItem "Api Successfully Hooked !"
    Init = True
    isInit = True
    
End Function

Function StartSeralize() As Long

    If Not isInit Then
        logit "StartSeralize: error hook lib not initilized!"
        Exit Function
    End If
    
    EnableHook lpWriteFile
    EnableHook lpReadFile
    EnableHook lpCreateFileA
    EnableHook lpGetFileType
    EnableHook lpCloseHandle
    ResetData
    
    vb_hFile = FreeFile
    Open trigger For Binary As vb_hFile
    StartSeralize = vb_hFile
    
End Function

Sub SetReadBuffer(readBuffer() As Byte)
    buf() = readBuffer()
End Sub

Function EndSeralize() As Byte()
    
    If Not isInit Then
        logit "EndSeralize: error hook lib not initilized!"
        Exit Function
    End If
    
    Close vb_hFile
    DisableHook lpWriteFile
    DisableHook lpCreateFileA
    DisableHook lpGetFileType
    DisableHook lpCloseHandle
    DisableHook lpReadFile
    EndSeralize = Data()
    ResetData
    
End Function



'----------------------------------------------------------------------
' data buffer functions below here
'----------------------------------------------------------------------


Private Function Data() As Byte()
    If pointer < 1 Then Exit Function
    ReDim Preserve buf(pointer - 1)
    Data = buf
End Function

Private Sub AddData(lpBuf As Long, bufSize As Long)
    Dim size As Long
    
    If bufSize = 0 Then Exit Sub
    
    size = bufSize
    If size < &H1000 Then size = &H1000
    If size > &H1000 Then size = size + &H1000
    If AryIsEmpty(buf) Then ReDim buf(size)
    
    If pointer + bufSize > UBound(buf) Then
        size = UBound(buf) + bufSize + &H1000
        ReDim Preserve buf(size)
    End If
    
    CopyMemory ByVal VarPtr(buf(pointer)), ByVal lpBuf, bufSize
    pointer = pointer + bufSize
    
End Sub

Private Sub ResetData()
    readPointer = 0
    pointer = 0
    Erase buf
End Sub


Private Function ReadData(lpBuf As Long, length As Long) As Boolean
    
    If AryIsEmpty(buf) Then Exit Function
    If readPointer + length > UBound(buf) + 1 Then Exit Function
    
    CopyMemory ByVal lpBuf, ByVal VarPtr(buf(readPointer)), length
    readPointer = readPointer + length
    ReadData = True
    
End Function



'----------------------------------------------------------------------
' hook implementations functions below here (must be public for addressof)
'----------------------------------------------------------------------

'we hook createfile so we can track what system file handle belongs to which file by name.
'we passthrough the args unchanged to the windows api using a modified declaration
Private Function my_CreateFile(ByVal lpFileName As Long, ByVal access As Long, ByVal shareMode As Long, ByVal sec As Long, ByVal cdisp As Long, ByVal flags As Long, ByVal template As Long) As Long
    
    Dim file As String
    
    file = CStringToVBString(lpFileName)
    'logit "my_CreateFile called for: " & file
    
    If file = trigger Then
        my_CreateFile = hFile
    Else
        'just in case something else calls this while enabled...
        my_CreateFile = CallOriginal(lpCreateFileA, lpFileName, access, shareMode, sec, cdisp, flags, template)
    End If
   
End Function

Private Function my_GetFileType(ByVal hObject As Long) As Long
    
    'logit "GetFileType " & Hex(hObject)
    
    If hObject = hFile Then
        my_GetFileType = 1 'disk file we fake out vbruntime...or it says permission denied...
    Else
        my_GetFileType = CallOriginal(lpGetFileType, hObject)
    End If
    
End Function

'If the function succeeds, the return value is nonzero.
Private Function my_CloseHandle(ByVal hObject As Long) As Long
'    logit "CloseHandle " & hex(hObject)
   
    If hObject = hFile Then
        my_CloseHandle = 1
    Else
        my_CloseHandle = CallOriginal(lpCloseHandle, hObject)
    End If

End Function

Private Function my_WriteFile(ByVal fileHandle As Long, ByVal lpBuf As Long, ByVal nWriteLength As Long, nBytesWritten As Long, lpOverlapped As Long) As Long
    
    'logit "WriteFile " & fileHandle & " buf: " & lpBuf & " size: " & nWriteLength
    
    If fileHandle = hFile Then
         my_WriteFile = HandleRedirection(lpBuf, nWriteLength, nBytesWritten)
    Else
         my_WriteFile = CallOriginal(lpWriteFile, fileHandle, lpBuf, nWriteLength, nBytesWritten, lpOverlapped)
    End If
    
End Function

'If the function succeeds, the return value is nonzero (TRUE).
Private Function my_ReadFile(ByVal fileHandle As Long, ByVal lpBuffer As Long, ByVal nNumberOfBytesToRead As Long, lpNumberOfBytesRead As Long, lpOverlapped As Long) As Long
        
    'logit "ReadFile " & fileHandle & " buf: " & Hex(lpBuffer) & " size: " & nNumberOfBytesToRead
    
    If fileHandle = hFile Then
        my_ReadFile = HandleRedirection(lpBuffer, nNumberOfBytesToRead, lpNumberOfBytesRead, False)
    Else
        my_ReadFile = CallOriginal(lpReadFile, fileHandle, lpBuffer, nNumberOfBytesToRead, lpNumberOfBytesRead, lpOverlapped)
    End If
    
End Function


Private Function HandleRedirection(lpBuf As Long, nLength As Long, ByRef nBytesHandled As Long, Optional isWrite As Boolean = True) As Long
    
    HandleRedirection = 1
    nBytesHandled = nLength
    
    If isWrite Then
        AddData lpBuf, nLength
    Else
        ReadData lpBuf, nLength
    End If
     
End Function






'----------------------------------------------------------------------
' library functions below here
'----------------------------------------------------------------------

'void  (__stdcall *debugMsgHandler)(char* msg);
Private Sub DebugMsgHandler(ByVal lpMsg As Long)

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

Private Function CStringToVBString(lpCstr As Long) As String

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

