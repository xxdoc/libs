VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5790
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9120
   LinkTopic       =   "Form1"
   ScaleHeight     =   5790
   ScaleWidth      =   9120
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox List2 
      Height          =   2400
      Left            =   945
      TabIndex        =   7
      Top             =   2700
      Width           =   7935
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Abort 2"
      Height          =   375
      Left            =   7875
      TabIndex        =   6
      Top             =   90
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Abort 1"
      Height          =   375
      Left            =   6705
      TabIndex        =   3
      Top             =   90
      Width           =   1095
   End
   Begin VB.ListBox List1 
      Height          =   2010
      Left            =   945
      TabIndex        =   2
      Top             =   630
      Width           =   7935
   End
   Begin VB.Timer Timer1 
      Left            =   45
      Top             =   1035
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   945
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   135
      Width           =   2850
   End
   Begin VB.Label lblProgress 
      Height          =   285
      Left            =   4770
      TabIndex        =   5
      Top             =   180
      Width           =   1680
   End
   Begin VB.Label Label3 
      Caption         =   "Progress"
      Height          =   240
      Left            =   3915
      TabIndex        =   4
      Top             =   180
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "Cur time:"
      Height          =   240
      Left            =   135
      TabIndex        =   1
      Top             =   135
      Width           =   690
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim hLib As Long
Dim d1 As CAsyncDownload
Dim d2 As CAsyncDownload

Dim d1Done As Boolean 'just for debug message filtering and since one test sets a class = nothing to force an abort for crash test
Dim d2Done As Boolean

Private Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Private Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long


Private Sub Command1_Click()
    lst "setting d1 = nothing", 2
    Set d1 = Nothing 'force auto abort to see if we can crash it
End Sub

Private Sub Command2_Click()
    lst "aborting d2", 2
    d2.AbortDownload
End Sub

Function lst(x, Optional index As Long = 0)
    Dim l As ListBox
    
    Set l = List1
    If index = 2 Then Set l = List2
    
    l.AddItem x
    l.ListIndex = l.ListCount - 1
    
End Function

Private Sub Form_Unload(Cancel As Integer)
    Timer1.Enabled = False
    Set d1 = Nothing
    Set d2 = Nothing
    FreeLibrary hLib 'so ide doesnt hang onto it and we can recompile if we want without closing ide
End Sub

Function LoadDLL() As Boolean
    
    Dim dll As String
    
    If hLib <> 0 Then
        lst "Dll already loaded"
    Else
        dll = App.path & "\Debug\bgdl.dll"
        hLib = LoadLibrary(dll)
        
        If hLib = 0 Then
            lst "Failed to find lib: " & dll
            dll = App.path & "\Release\bgdl.dll"
            hLib = LoadLibrary(dll)
        End If
        
        If hLib = 0 Then
            lst "Failed to find lib: " & dll
            dll = App.path & "\bgdl.dll"
            hLib = LoadLibrary(dll)
        End If
        
        If hLib = 0 Then
            lst "Fatal - Failed to find lib: " & dll
            Exit Function
        End If
    End If
    
    LoadDLL = True
    
End Function

Private Sub Form_Load()
    
    On Error Resume Next
        
    Dim url1, url2
    Dim f1 As String, f2 As String
    
    Text1 = Now
    Me.Visible = True
    d1Done = False
    d2Done = False
    List1.Clear
    List2.Clear
    
    If Not LoadDLL() Then Exit Sub
    
    url1 = "http://speedtest.ftp.otenet.gr/files/test10Mb.db"
    url2 = "http://www.x-ways.net/winhex.zip"
    
    f1 = App.path & "\x1.bin" 'note api cant handle accidental double slashs
    f2 = App.path & "\x2.bin"
    Kill f1
    Kill f2

    Set d1 = New CAsyncDownload
    Set d2 = New CAsyncDownload
    
    If Not d1.StartDownload(url1, f1) Then
        lst "Failed to create download thread for d1!"
        Exit Sub
    Else
        lst "d1 started: " & url1
    End If
        
    If Not d2.StartDownload(url2, f2) Then
        lst "Failed to create download thread for d2!"
        Exit Sub
    Else
        lst "d2 started: " & url2
    End If
    
    lst "Downloads started!", 2
    
    Timer1.Interval = 500
    Timer1.Enabled = True 'to show we arent frozen and poll download status from classes

End Sub

Private Sub Timer1_Timer()

    On Error Resume Next 'lazy - note for test d1 may be set to nothing on abort
    Dim tmp As String
    
    Text1 = Now
    
    If Not d1 Is Nothing Then
        If Not d1Done Then
            tmp = "d1: " & Join(Array(d1.StatusCode, d1.Progress, d1.DownloadSize, d1.DownloadStatus), ",")
            If d1.isRunning Then
                lst tmp
            Else
                lst "d1 complete retval=" & d1.DownloadStatusStr, 2
                d1Done = True
            End If
        End If
    Else
        d1Done = True
    End If
    
    If d2.isRunning Then
        tmp = "d2: " & Join(Array(d2.StatusCode, d2.Progress, d2.DownloadSize, d2.DownloadStatus), ",")
        If Not d2Done Then lst tmp
    Else
        If Not d2Done Then lst "d2 complete retval=" & d2.DownloadStatusStr, 2
        d2Done = True
    End If

    If d1Done And d2.DownloadStatus <> ds_Downloading Then
        Timer1.Enabled = False
        d2Done = True
        d1Done = True
        lst "downloads complete", 2
    End If
    
End Sub
