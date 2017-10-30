VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "VB - C InterProcess Communications Using WM_COPYDATA"
   ClientHeight    =   6375
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10875
   LinkTopic       =   "Form1"
   ScaleHeight     =   6375
   ScaleWidth      =   10875
   StartUpPosition =   2  'CenterScreen
   Begin VB.OptionButton Option7 
      Caption         =   "Python"
      Height          =   285
      Left            =   6030
      TabIndex        =   16
      Top             =   3690
      Width           =   1050
   End
   Begin VB.Frame Frame1 
      Caption         =   "Vista+ Elevated Process Options"
      Height          =   1050
      Left            =   8235
      TabIndex        =   13
      Top             =   3285
      Width           =   2580
      Begin VB.CheckBox chkLowPriv 
         Caption         =   "start child low priv"
         Enabled         =   0   'False
         Height          =   285
         Left            =   135
         TabIndex        =   15
         Top             =   630
         Width           =   1950
      End
      Begin VB.CheckBox chkAllowCopyDataMsg 
         Caption         =   "Allow WM_COPYDATA"
         Enabled         =   0   'False
         Height          =   285
         Left            =   135
         TabIndex        =   14
         Top             =   270
         Width           =   2085
      End
   End
   Begin VB.OptionButton Option6 
      Caption         =   "Delphi"
      Height          =   375
      Left            =   5130
      TabIndex        =   12
      Top             =   3630
      Width           =   795
   End
   Begin VB.OptionButton Option5 
      Caption         =   "VB6"
      Height          =   405
      Left            =   4350
      TabIndex        =   11
      Top             =   3600
      Width           =   675
   End
   Begin VB.CommandButton cmdAsync 
      Caption         =   "Async Test"
      Enabled         =   0   'False
      Height          =   345
      Left            =   6840
      TabIndex        =   10
      Top             =   3240
      Width           =   1305
   End
   Begin VB.OptionButton Option4 
      Caption         =   "java"
      Height          =   375
      Left            =   3600
      TabIndex        =   9
      Top             =   3630
      Width           =   705
   End
   Begin VB.OptionButton Option3 
      Caption         =   "CSharp"
      Height          =   345
      Left            =   2580
      TabIndex        =   7
      Top             =   3630
      Width           =   915
   End
   Begin VB.OptionButton Option2 
      Caption         =   "C x64"
      Height          =   285
      Left            =   1710
      TabIndex        =   6
      Top             =   3660
      Width           =   795
   End
   Begin VB.OptionButton Option1 
      Caption         =   "C"
      Height          =   285
      Left            =   1020
      TabIndex        =   5
      Top             =   3660
      Value           =   -1  'True
      Width           =   585
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   90
      TabIndex        =   3
      Top             =   4410
      Width           =   10695
   End
   Begin VB.TextBox Text2 
      Height          =   315
      Left            =   960
      TabIndex        =   0
      Text            =   "this is my test message"
      Top             =   3240
      Width           =   4395
   End
   Begin VB.ListBox List1 
      Height          =   3180
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   10755
   End
   Begin VB.CommandButton cmdTest 
      Caption         =   "Start and send"
      Height          =   345
      Left            =   5490
      TabIndex        =   1
      Top             =   3210
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Client: "
      Height          =   255
      Left            =   90
      TabIndex        =   8
      Top             =   3690
      Width           =   795
   End
   Begin VB.Label Label1 
      Caption         =   "Raw Bytes Received"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   4050
      Width           =   1575
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit

'I used my subclass library for simplicity, you can use whatever sub
'class technique or inline code you desire...
'
'see the following for inline code:
'   http://support.microsoft.com/kb/176058

'So in Vista+ operating systems they implemented UIPI - user interface privledge escalation
'If UAC is on, elevated processes can not receive many window messages from lower privledge level processes
'to prevent shatter attack type scenarios.
'
'There may be times we need to use this technique across privledge levels so I have included a couple
'of routines specificaly for testing this scenario

Dim WithEvents sc As clsSubClass
Attribute sc.VB_VarHelpID = -1

Private Declare Function GetShortPathName Lib "kernel32" Alias "GetShortPathNameA" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal cchBuffer As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
         
Const WM_COPYDATA = &H4A
Const WM_DISPLAY_TEXT = 3

Private Type COPYDATASTRUCT 'this works for both 32 and 64 bit ignore structure below it was testing..
    dwFlag As Long
    cbSize As Long
    lpData As Long
End Type

Private RemoteHWND As Long

Public Function GetShortName(sFile As String) As String
    Dim sShortFile As String * 67
    Dim lResult As Long
    
    'the path must actually exist to get the short path name !!
    'If Not FileExists(sFile) Then writeFile sFile, ""
    
    'Make a call to the GetShortPathName API
    lResult = GetShortPathName(sFile, sShortFile, _
    Len(sShortFile))

    'Trim out unused characters from the string.
    GetShortName = Left$(sShortFile, lResult)
    
    If Len(GetShortName) = 0 Then GetShortName = sFile

End Function


'Private Type COPYDATASTRUCT64
'    dwFlag As Long
'    dwFlagHigh As Long
'    cbSize As Long
'    cbSizeHigh As Long
'    lpData As Long
'    lpDataHigh As Long
'End Type



Function IsIde() As Boolean
    On Error GoTo out
    Debug.Print 1 / 0
out: IsIde = Err
End Function

Sub RunJavaExample()
    Dim cmdline As String
    
    cmdline = "cmd /k cd """ & App.path & "\java_client"" && java JNITest"
    If IsIde() Then cmdline = Replace(cmdline, "\vb_server", Empty)
    
    On Error Resume Next
    If chkLowPriv.Value = 1 Then
        List1.AddItem "Running as low priv not implemented for java sample..."
    End If
    List1.AddItem "Executing Cmdline: " & cmdline
    Err.Clear
    Shell cmdline, vbNormalFocus
    If Err.Number <> 0 Then MsgBox Err.Description
    
End Sub


Private Sub chkAllowCopyDataMsg_Click()
    AllowCopyDataAcrossUIPI (chkAllowCopyDataMsg.Value = 1)
End Sub

Private Sub cmdTest_Click()
    
    Dim exe As String, cmdline As String
    
    List1.Clear
    List1.AddItem Now
    RemoteHWND = 0
    cmdAsync.Enabled = False
    
    exe = App.path & "\wmcopy.exe"
    If Not FileExists(exe) Then exe = App.path & "\..\wmcopy.exe"
    
    If Option1.Value Then
        'standard C client, default already done..
    ElseIf Option2.Value Then
        'x64
        exe = Replace(exe, ".exe", ".x64.exe")
    ElseIf Option3.Value Then
        'csharp client...
        exe = Replace(exe, "wmcopy.exe", "csharp_client.exe")
    ElseIf Option5.Value Then
        'vb6 client
        exe = Replace(exe, "wmcopy.exe", "vb6_client.exe")
    ElseIf Option6.Value Then
        'delphi client
        exe = Replace(exe, "wmcopy.exe", "delphi6_client.exe")
    ElseIf Option7.Value Then
        'python client
        If InStr(1, Environ("path"), "python", vbTextCompare) < 1 Then
            List1.AddItem "python must be in your path variable to run this example"
            List1.AddItem "also make sure to pip install pypiwin32"
            Exit Sub
        End If
         exe = Replace(exe, "wmcopy.exe", "python_example.py")
    ElseIf Option4.Value Then
        RunJavaExample 'java sample does not use the test message from this UI, uses hardcoded message..
        Exit Sub
    End If
    
    If Option7.Value Then
        cmdline = "cmd /k python " & exe & " " & Me.hwnd & " """ & Text2 & """"
    Else
        cmdline = exe & " """ & Me.hwnd & "," & Text2 & """"
    End If
    
    If Not FileExists(exe) Then
        MsgBox "Could not locate executable " & exe, vbExclamation
    Else
        On Error Resume Next
        List1.AddItem "Executing Cmdline: " & Replace(cmdline, App.path & "\", Empty)
        Err.Clear
        If chkLowPriv.Value = 1 Then
            If Not RunAsDesktopUser(, cmdline) Then
                MsgBox modRunAsUser.ErrInfo
            End If
        Else
            Shell cmdline, vbNormalFocus
        End If
        If Err.Number <> 0 Then MsgBox Err.Description
    End If
    
End Sub
    
Private Sub cmdAsync_Click()
    
    If RemoteHWND = 0 Then
        MsgBox "The remote app didnt register a callback hwnd with us yet"
        Exit Sub
    End If
    
    SendData Text2, RemoteHWND
    
End Sub

Private Sub Form_Load()
    
    If IsVistaPlus() Then
        List1.AddItem "Vista+ detected"
        chkLowPriv.Enabled = True
        If isUACEnabled() Then
            List1.AddItem "UAC is enabled"
            If IsProcessElevated() Then
                List1.AddItem "Elevated Process detected allowing WM_CopyData for process..."
                chkAllowCopyDataMsg.Enabled = True
                chkAllowCopyDataMsg.Value = 1
                AllowCopyDataAcrossUIPI
            Else
                List1.AddItem "Process is not elevated no special actions necessary."
            End If
        Else
             List1.AddItem "UAC is not enabled"
        End If
    End If
    
    Set sc = New clsSubClass
    sc.AttachMessage Me.hwnd, WM_COPYDATA
    
    List1.AddItem "Watching for WM_COPYDATA for main Window Hwnd: " & Me.hwnd
    
End Sub

Private Sub sc_MessageReceived(hwnd As Long, wMsg As Long, wParam As Long, lParam As Long, Cancel As Boolean)
    
    If wMsg = WM_COPYDATA Then RecieveTextMessage lParam
    
End Sub


Private Sub RecieveTextMessage(lParam As Long)
   
    Dim CopyData As COPYDATASTRUCT
    'Dim CopyData64 As COPYDATASTRUCT64
    Dim Buffer(1 To 2048) As Byte
    Dim Temp As String
    Dim lpData As Long
    Dim sz As Long
    Dim tmp() As Byte
    ReDim tmp(30)
    
    CopyMemory CopyData, ByVal lParam, Len(CopyData)
    
    If CopyData.dwFlag = 3 Then
    
        CopyMemory tmp(0), ByVal lParam, Len(CopyData)
        Text1 = HexDump(tmp, Len(CopyData))
        
        lpData = CopyData.lpData
        sz = CopyData.cbSize
        
    'ElseIf CopyData.dwFlag = 64 Then
   '
   '     CopyMemory tmp(0), ByVal lParam, Len(CopyData64)
   '     Text1 = HexDump(tmp, Len(CopyData64))
   '
   '     CopyMemory CopyData64, ByVal lParam, Len(CopyData64)
   '     lpData = CopyData64.lpData
   '     sz = CopyData64.cbSize
   '
   '     List1.AddItem "Sizeof(CopyData64) = " & Len(CopyData64)
   '     List1.AddItem "HighSize: " & Hex(CopyData64.cbSizeHigh)
   '     List1.AddItem "HighData: " & Hex(CopyData64.lpDataHigh)
   '     List1.AddItem "HighFlag: " & Hex(CopyData64.dwFlagHigh)
   '
    Else
        List1.AddItem "Unknown flag: " & CopyData.dwFlag
        Exit Sub
    End If
    
    If sz > 2048 Then
        List1.AddItem "Size(" & Hex(sz) & " ) > buffer (2048)"
        Exit Sub
    End If
    
    'List1.AddItem "lpData: " & Hex(lpData)
    'List1.AddItem "Size: " & Hex(sz)
    
    CopyMemory Buffer(1), ByVal lpData, sz
    Temp = StrConv(Buffer, vbUnicode)
    Temp = Left$(Temp, InStr(1, Temp, Chr$(0)) - 1)
    'heres where we work with the intercepted message
    List1.AddItem "***** Message Received: " & Temp
     
    On Error Resume Next
    If InStr(Temp, "PINGME=") > 0 Then
        x = Split(Temp, "=")
        RemoteHWND = CLng(x(1))
        cmdAsync.Enabled = True
        SendData Text2, RemoteHWND
    End If
    
    
    
End Sub

Sub SendData(msg As String, hwnd As Long)
    Dim cds As COPYDATASTRUCT
    Dim ThWnd As Long
    Dim buf(1 To 255) As Byte

    List1.AddItem "SendingData to " & hwnd & " msg: " & msg
    
    Call CopyMemory(buf(1), ByVal msg, Len(msg))
    cds.dwFlag = 3
    cds.cbSize = Len(msg) + 1
    cds.lpData = VarPtr(buf(1))
    i = SendMessage(hwnd, WM_COPYDATA, Me.hwnd, cds)
    
End Sub

Function FileExists(path) As Boolean
  If Dir(path, vbHidden Or vbNormal Or vbReadOnly Or vbSystem) <> "" Then FileExists = True
End Function

Sub push(ary, Value) 'this modifies parent ary object
    On Error GoTo init
    x = UBound(ary) '<-throws Error If Not initalized
    ReDim Preserve ary(UBound(ary) + 1)
    ary(UBound(ary)) = Value
    Exit Sub
init: ReDim ary(0): ary(0) = Value
End Sub

Function HexDump(byteArray() As Byte, Optional max As Long = -1) As String
    Dim s() As String, chars As String, tmp As String
    On Error Resume Next
    Dim ary() As Byte
    Dim offset As Long
    Const LANG_US = &H409
    Dim endAt As Long
    Const hexOnly = 0
    
    offset = 0
    'str = " " & str
    'ary = StrConv(str, vbFromUnicode, LANG_US)
    
    ary = byteArray
    
    endAt = max
    If endAt = -1 Then endAt = UBound(ary)
    If endAt > UBound(ary) Then endAt = UBound(ary)
    
    chars = "   "
    For i = 0 To UBound(ary)
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

