VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6525
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13230
   LinkTopic       =   "Form1"
   ScaleHeight     =   6525
   ScaleWidth      =   13230
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "end"
      Height          =   465
      Left            =   1800
      TabIndex        =   1
      Top             =   4410
      Width           =   1320
   End
   Begin VB.Timer Timer1 
      Interval        =   200
      Left            =   810
      Top             =   4455
   End
   Begin VB.ListBox List1 
      Height          =   4155
      Left            =   45
      TabIndex        =   0
      Top             =   45
      Width           =   13020
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub d(x)
    List1.AddItem x
    List1.Refresh
    DoEvents
    Me.Refresh
End Sub

Private Sub Command1_Click()
    End
End Sub

Private Sub Form_Load()
    Dim connect_to As String, message As String, message_size As Long
    Dim ctx As Long, s As Long, rc As Long, i As Long
    Dim hLib As Long
    Dim msgBuf() As Byte

    Me.Visible = True
    
    'in case we hit stop in ide..i dont want dll stuck in memory
    'requires close ide to recompile dll..this is only for dev..
    hLib = GetModuleHandle("libzmq.dll")
    If hLib <> 0 Then FreeLibrary (hLib)
    
    hLib = LoadLibrary(App.Path & "\libzmq.dll")
    If hLib = 0 Then
        d "failed to find dll? " & Err.LastDllError
        Exit Sub
    End If
    
    connect_to = "tcp://localhost:5546"
           
    ReDim msgBuf(255)

    
    ctx = zmq_ctx_new()
    If ctx = 0 Then
        d "error in zmq_init"
        Exit Sub
    End If
    
    s = zmq_socket(ctx, ZMQ_REQ)
    If s = 0 Then
        d "error in zmq_socket"
        Exit Sub
    End If

    ' Add your socket options here.
    ' For example ZMQ_RATE, ZMQ_RECOVERY_IVL and ZMQ_MCAST_LOOP for PGM.

    rc = zmq_connect(s, connect_to)
    
    If rc <> 0 Then
        d "error in zmq_connect: " & zError(rc)
        Exit Sub
    End If

    rc = zmq_setsockopt(s, ZMQ_LINGER, 0, 4)
    
    poller = zmq_poller_new()
    rc = zmq_poller_add(poller, s, s, ZMQ_POLLIN)
    d "poller: " & poller & " add:" & rc
    
    Dim evt As zmq_poller_event
    
    For i = 0 To 3
        message = "Hello " & Now
        rc = zmq_send(s, message, Len(message), 0)
        d i & " send:" & rc
        rc = zmq_poller_wait(poller, evt, 0)
        d "poller: " & rc
        rc = zmq_recv(s, msgBuf(0), UBound(msgBuf), 0)
        d i & " recv:" & rc & " > " & StrConv(msgBuf, vbUnicode)
    Next

    rc = zmq_close(s)
    If rc <> 0 Then
        d "error in zmq_close: " & zError(rc)
        Exit Sub
    End If

    rc = zmq_ctx_term(ctx)
    If rc <> 0 Then
        d "error in zmq_ctx_term: " & zError(rc)
        Exit Sub
    End If

    d "done"
    FreeLibrary hLib

End Sub











'Private Sub Form_Load()
'    Dim connect_to As String, message As String, message_size As Long
'    Dim ctx As Long, s As Long, rc As Long, i As Long, msg As z_msg_t
'    Dim hLib As Long
'    Dim msgBuf() As Byte
'
'    'in case we hit stop in ide..i dont want dll stuck in memory
'    'requires close ide to recompile dll..this is only for dev..
'    hLib = GetModuleHandle("libzmq.dll")
'    If hLib <> 0 Then FreeLibrary (hLib)
'
'    hLib = LoadLibrary("C:\Documents and Settings\david\Desktop\lib\libzmq.dll")
'    If hLib = 0 Then
'        d "failed to find dll? " & Err.LastDllError
'        Exit Sub
'    End If
'
'    connect_to = "tcp://localhost:5546"
'    message = "Hello!"
'
'    msgBuf = StrConv(message, vbFromUnicode)
'    message_size = UBound(msgBuf)
'    'msg = VarPtr(msgBuf(0))
'
'    ctx = zmq_init(1)
'    If ctx = 0 Then
'        d "error in zmq_init"
'        Exit Sub
'    End If
'
'    s = zmq_socket(ctx, ZMQ_REQ)
'    If s = 0 Then
'        d "error in zmq_socket"
'        Exit Sub
'    End If
'
'    ' Add your socket options here.
'    ' For example ZMQ_RATE, ZMQ_RECOVERY_IVL and ZMQ_MCAST_LOOP for PGM.
'
'    rc = zmq_connect(s, connect_to)
'
'    If rc <> 0 Then
'        d "error in zmq_connect: " & zError(rc)
'        Exit Sub
'    End If
'
'    For i = 0 To 10
'        rc = zmq_msg_init_size(msg, message_size)
'        If rc <> 0 Then
'            d ("error in zmq_msg_init_size: " & zError(rc))
'            Exit Sub
'        End If
'        rc = zmq_sendmsg(s, msg, 0)
'        If rc <> 0 Then
'            d ("error in zmq_sendmsg: " & zError(rc))
'            Exit Sub
'        End If
'        rc = zmq_msg_close(msg)
'        If rc <> 0 Then
'            d ("error in zmq_msg_close: " & zError(rc))
'            Exit Sub
'        End If
'    Next
'
'    rc = zmq_close(s)
'    If rc <> 0 Then
'        d "error in zmq_close: " & zError(rc)
'        Exit Sub
'    End If
'
'    rc = zmq_ctx_term(ctx)
'    If rc <> 0 Then
'        d "error in zmq_ctx_term: " & zError(rc)
'        Exit Sub
'    End If
'
'
'    FreeLibrary hLib
'
'End Sub

Private Sub Timer1_Timer()
    DoEvents
End Sub
