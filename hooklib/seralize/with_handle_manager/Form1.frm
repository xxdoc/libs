VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   7605
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9345
   LinkTopic       =   "Form1"
   ScaleHeight     =   7605
   ScaleWidth      =   9345
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Back to UDT"
      Height          =   495
      Left            =   5880
      TabIndex        =   4
      Top             =   6480
      Width           =   2175
   End
   Begin VB.CommandButton cmdSetHooks 
      Caption         =   "Set Hooks"
      Height          =   495
      Left            =   120
      TabIndex        =   3
      Top             =   6480
      Width           =   2295
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Test serialize UDT to memory"
      Height          =   495
      Left            =   2760
      TabIndex        =   2
      Top             =   6480
      Width           =   2415
   End
   Begin VB.TextBox Text1 
      Height          =   3015
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   1
      Top             =   3360
      Width           =   9135
   End
   Begin VB.ListBox List1 
      Height          =   3180
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   9135
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Type test
    signature As Long
    buf As String
    signature2 As Long
    v As Variant
End Type

Dim seralized() As Byte

Private Sub Command1_Click()
    
    Dim f As Long
    Dim path As String
    Dim h As CFileHandle
    Dim udtTest As test
    
    path = trigger
    
    EnableHook lpWriteFile
    EnableHook lpCreateFileA
    EnableHook lpGetFileType
    
    f = FreeFile
    Open path For Binary As f
    
    If Not manager.HandleExists(path) Then
        logit "We did not find our handle?? vbHandle=" & f
        Exit Sub
    End If
    
    Set h = manager.GetHandle(path)
    h.vbHandle = f
    'logit "We found it! vbHandle=" & h.vbHandle & " systemHandle=" & Hex(h.sysHandle)
    h.RedirectTo = rt_memory
    
    With udtTest
        .signature = &H11223344
        .buf = "test string"
        .signature2 = &H99887766
        .v = Array(1, "blah!", 3.14)
    End With

    'logit "testing redirected object serialization"
    Put f, , udtTest
    'Close f 'do not close it now since we didnt really open it!

    'Put f, , CStr(String(16, "A"))
    'Put f, , CStr(String(16, "B"))
    'Close f
    
    DisableHook lpWriteFile
    DisableHook lpCreateFileA
    DisableHook lpGetFileType
    
    seralized() = h.Data
    manager.Remove h.sysHandle
    Text1 = HexDump(StrConv(seralized, vbUnicode))
    
End Sub

Private Sub cmdSetHooks_Click()
    
    Dim h As Long
    Dim ret As Long
    Dim lpMsg As Long
    
    If hHookLib <> 0 Then Exit Sub 'already called...
    
    hHookLib = LoadLibrary("hooklib.dll")
    If hHookLib = 0 Then hHookLib = LoadLibrary(App.path & "\hooklib.dll")
    If hHookLib = 0 Then hHookLib = LoadLibrary(App.path & "\..\hooklib.dll")
    If hHookLib = 0 Then hHookLib = LoadLibrary(App.path & "\..\..\hooklib.dll")
    
    If hHookLib = 0 Then
        List1.AddItem "Could not find hooklib.dll compile or download from github."
        Exit Sub
    End If
    
    List1.AddItem "Hooklib base address: 0x" & Hex(h)
    
    'this is optional but were debugging the library so..
    SetDebugHandler AddressOf DebugMsgHandler, 1
    
    h = LoadLibrary("kernel32.dll")
    lpCloseHandle = GetProcAddress(h, "CloseHandle")
    lpCreateFileA = GetProcAddress(h, "CreateFileA")
    lpWriteFile = GetProcAddress(h, "WriteFile")
    lpReadFile = GetProcAddress(h, "ReadFile")
    lpGetFileType = GetProcAddress(h, "GetFileType")
    
    If lpCloseHandle = 0 Or lpCreateFileA = 0 Or lpWriteFile = 0 Or lpReadFile = 0 Then
        List1.AddItem "GetProcAddress failed for one of the functions??"
        Exit Sub
    End If
        
'    ret = HookFunction(lpCloseHandle, AddressOf my_CloseHandle, "CloseHandle", ht_jmp)
'    If ret = 0 Then
'        lpMsg = GetHookError()
'        List1.AddItem "Hook Function failed msg: " & CStringToVBString(lpMsg)
'        Exit Sub
'    Else
'        DisableHook lpCloseHandle
'    End If
    
    ret = HookFunction(lpWriteFile, AddressOf my_WriteFile, "WriteFile", ht_jmp)
    If ret = 0 Then
        lpMsg = GetHookError()
        List1.AddItem "Hook Function failed msg: " & CStringToVBString(lpMsg)
        Exit Sub
    Else
        DisableHook lpWriteFile
    End If
    
    ret = HookFunction(lpCreateFileA, AddressOf my_CreateFile, "CreateFileA", ht_jmp)
    If ret = 0 Then
        lpMsg = GetHookError()
        List1.AddItem "Hook Function failed msg: " & CStringToVBString(lpMsg)
        Exit Sub
    Else
        DisableHook lpCreateFileA
    End If
    
    ret = HookFunction(lpReadFile, AddressOf my_ReadFile, "ReadFile", ht_jmp)
    If ret = 0 Then
        lpMsg = GetHookError()
        List1.AddItem "Hook Function failed msg: " & CStringToVBString(lpMsg)
        Exit Sub
    Else
        DisableHook lpReadFile
    End If
    
    ret = HookFunction(lpGetFileType, AddressOf my_GetFileType, "GetFileType", ht_jmp)
    If ret = 0 Then
        lpMsg = GetHookError()
        List1.AddItem "Hook Function failed msg: " & CStringToVBString(lpMsg)
        Exit Sub
    Else
        DisableHook lpGetFileType
    End If
    
    List1.AddItem "Api Successfully Hooked !"
    
    
End Sub

Private Sub Command2_Click()

    Dim f As Long
    Dim path As String
    Dim h As CFileHandle
    Dim udtTest As test
    
    path = trigger
    
    EnableHook lpReadFile
    EnableHook lpCreateFileA
    EnableHook lpGetFileType
    
    f = FreeFile
    Open path For Binary As f
    
    If Not manager.HandleExists(path) Then
        logit "We did not find our handle?? vbHandle=" & f
        Exit Sub
    End If
    
    Set h = manager.GetHandle(path)
    h.vbHandle = f
    'logit "We found it! vbHandle=" & h.vbHandle & " systemHandle=" & Hex(h.sysHandle)
    h.RedirectTo = rt_memory
    h.target = seralized()
    
    'With udtTest
    '    .signature = &H11223344
    '    .buf = "test string"
    '    .signature2 = &H99887766
    '    .v = Array(1, "blah!", 3.14)
    'End With

    'logit "testing redirected object serialization"
    Get f, , udtTest
    'Close f

    DisableHook lpReadFile
    DisableHook lpCreateFileA
    DisableHook lpGetFileType
    
    On Error Resume Next
    Dim tmp() As String
    
    push tmp, "Signature: " & Hex(udtTest.signature)
    push tmp, "Buf: " & udtTest.buf
    push tmp, "Signature2: " & Hex(udtTest.signature2)
    
    If IsArray(udtTest.v) Then
         push tmp, "v: " & Join(udtTest.v, ", ")
    Else
        push tmp, "error v is not array: " & TypeName(udtTest.v)
    End If
    
    Text1 = Join(tmp, vbCrLf)
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    RemoveAllHooks
    'UnInitilizeHookLib
    'FreeLibrary hHookLib
End Sub
