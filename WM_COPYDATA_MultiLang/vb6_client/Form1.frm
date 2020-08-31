VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "VB6 IDASrvr Example"
   ClientHeight    =   5385
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10620
   LinkTopic       =   "Form1"
   ScaleHeight     =   5385
   ScaleWidth      =   10620
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox List2 
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2700
      Left            =   60
      TabIndex        =   1
      Top             =   2520
      Width           =   10455
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2220
      Left            =   30
      TabIndex        =   0
      Top             =   30
      Width           =   10515
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'http://support.microsoft.com/kb/176058
'this uses inline subclassing code, I would recommend using a library such as
'my spSubclass or VBaccelerator's subclass lib for stability when running in the IDE.
'ps dont hit end from within IDE or it will crash as subclass isnt cleaned up.


Private Sub Form_Load()
    
    On Error Resume Next
    'expected command line format: "hwnd,message with spaces possible"
    
    Dim tmp As String, msg As String
    tmp = Command
    
    a = InStr(tmp, ",")
    If a > 0 Then
        List2.AddItem "Command line: " & tmp
        msg = Mid(tmp, a + 1)
        msg = Replace(msg, """", Empty)
        List2.AddItem "parsed msg: " & msg
        b = InStrRev(tmp, """", a)
        If b > 0 Then
            tmp = Mid(tmp, b, a - b)
            tmp = Replace(tmp, """", Empty)
            tmp = Replace(tmp, ",", Empty)
            List2.AddItem "parsed hwnd: " & tmp
            Server_Hwnd = CLng(tmp)
        End If
    End If
    
    If Len(msg) = 0 Then
        List1.AddItem "Failed to extract message from command line..."
        Exit Sub
    End If
    
    If Server_Hwnd = 0 Then
        List1.AddItem "Failed to extract server hwnd from command line..."
        Exit Sub
    End If
    
    Hook Me.hwnd
     
    List1.AddItem "Sending " & msg
    SendCMD msg
    
    List1.AddItem "Sending pingback"
    resp = SendCmdRecvText("PINGME=" & Me.hwnd)
    
    List1.AddItem "SendCmdRecvText response was: " & resp
    
    
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Unhook
End Sub
 

