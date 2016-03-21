VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5415
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8115
   LinkTopic       =   "Form1"
   ScaleHeight     =   5415
   ScaleWidth      =   8115
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   5190
      Left            =   135
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Text            =   "Form1.frx":0000
      Top             =   90
      Width           =   7890
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'this is a test for a small syncronous socket send/recv function.
'you can call this to send raw data to another pc over the network via tcp
'and it will not return until you have the full response from the server.
'errors are handled simply, C dll is 25k compressed.
'
'this is free for any use, open source
'
'note buffer full, and recv timeout errors do not make quicksend return false
'you will receive partial data, you can double check the lastError to see if they
'hit
'
'the reason i coded this is because sometimes you want an inline data send/recv
'without having to force syncronous behavior on top of a mswinsck.ocx control.
'which I hate.
'
'this is easier, smaller (25k vrs 122k) and does not require installation (regsvr32)

Private Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Private Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long

'void __stdcall qsCfg(char* _server, int _port, int _timeout, short partialRespOk){
Private Declare Sub qsConfig Lib "sendrecv.dll" ( _
            ByVal server As String, _
            ByVal port As Long, _
            Optional ByVal msTimeout As Long = 12000, _
            Optional ByVal allowPartial As Boolean = True _
        )

'int __stdcall LastError(char* buffer, int buflen){
Private Declare Function CLastError Lib "sendrecv.dll" Alias "LastError" ( _
            ByVal buffer As String, _
            ByVal bufLen As Long) As Long


'int __stdcall QuickSend(char* request,int reqLen, char* response_buffer, int response_buflen ){

'allowPartial=true lets you specify a small buffer and have it return ok
'without an error. You can still determine if the buffer was full based on
'the size returned, or if no error was returned by lastErr has data in it.
'alternativly to reject any non-complete response set allowPartial = false
'and CQuickSend will return false unless it gets the full reply back.
'not to over complicate it, but its sufficiently useful both ways..

Private Declare Function CQuickSend Lib "sendrecv.dll" Alias "QuickSend" ( _
            ByVal request As String, _
            ByVal reqLen As Long, _
            ByVal response_buffer As String, _
            ByVal respBufLen As Long _
        ) As Long

Dim hLib As Long

'only needed in teh ide or if dll is in different directory..
Function ide_ensure_dll_loaded(Optional ByRef msg, Optional dllName = "sendrecv.dll") As Boolean

    If hLib = 0 Then hLib = LoadLibrary(dllName)
    If hLib = 0 Then hLib = LoadLibrary(App.Path & "\" & dllName)
    If hLib = 0 Then hLib = LoadLibrary(App.Path & "\Release\" & dllName)
    If hLib = 0 Then hLib = LoadLibrary(App.Path & "\Debug\" & dllName)
    
    If hLib <> 0 Then
        ide_ensure_dll_loaded = True
    Else
        msg = "Could not find library " & dllName
    End If

End Function

Function QuickSend(msg, Optional ByRef response, Optional maxSize As Long = 4096) As Boolean
    
    Dim buf As String
    Dim sz As Long
    
    buf = String(maxSize, Chr(0))
    sz = CQuickSend(msg, Len(msg), buf, Len(buf))
    
    If sz < 1 Then 'we had an error
        sz = CLastError(buf, Len(buf))
        If sz < 1 Then
            response = "Unknown error"
        Else
            response = Mid(buf, 1, sz)
        End If
    Else
        response = Mid(buf, 1, sz)
        QuickSend = True
    End If
    
End Function

Private Sub Form_Load()
    
    Dim buf As String
    Dim ok As Boolean
    Dim server As String
    
    Const http = "GET /tools.php HTTP/1.0" & vbCrLf & _
                "Host: sandsprite.com" & vbCrLf & _
                "User-Agent: Mozilla/5.0 (Windows NT 5.1; rv:45.0)" & vbCrLf & _
                "Accept-Encoding: none" & vbCrLf & _
                "Connection: close" & vbCrLf & _
                "" & vbCrLf & _
                "" & vbCrLf
    
    Const maxSz = 300
    
    server = "sandsprite.com"
    'server = "192.168.0.10"
    
    If Not ide_ensure_dll_loaded(buf) Then
        MsgBox buf
        Exit Sub
    End If
    
    qsConfig server, 80
    ok = QuickSend(http, buf, maxSz)
    Me.Caption = IIf(ok, "Success!", "Failed!")
    
    If ok Then
        If Len(buf) = maxSz - 2 Then
            Me.Caption = Me.Caption & " - Partial content - Buffer full"
        Else
            Me.Caption = Me.Caption & " Size: " & Len(buf)
        End If
        Text1 = buf
    Else
        Text1 = "Error: " & buf
    End If
    

End Sub

Private Sub Form_Unload(Cancel As Integer)
    'this is just for testing in the ide, so ide doesnt hang onto
    'the dll and we can recompile it without closing ide down..
    If hLib <> 0 Then FreeLibrary (hLib)
End Sub
