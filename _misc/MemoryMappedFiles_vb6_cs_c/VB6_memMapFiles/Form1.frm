VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   2070
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10680
   LinkTopic       =   "Form1"
   ScaleHeight     =   2070
   ScaleWidth      =   10680
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   1800
      TabIndex        =   3
      Text            =   "Text2"
      Top             =   135
      Width           =   4515
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Read"
      Height          =   255
      Left            =   8760
      TabIndex        =   2
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Write"
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   120
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   1455
      Left            =   240
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   480
      Width           =   9855
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mmf As New CMemMapFile

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, source As Any, ByVal Length As Long)

Private Declare Sub CopyToMem Lib "kernel32" Alias "RtlMoveMemory" (ByVal Destination As Long, source As Any, ByVal Length As Long)
Private Declare Sub CopyFromMem Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, ByVal source As Long, ByVal Length As Long)


Private Sub Command1_Click()
    
    Dim b(5) As Byte
    Dim tmp() As Byte
    ReDim tmp(5)
    
    If App.PrevInstance Then
        tmp(0) = Asc("A")
        tmp(1) = Asc("B")
        tmp(2) = Asc("C")
        tmp(3) = Asc("D")
        tmp(4) = Asc("E")
        Text2 = "Wrote ABCDE"
    Else
        tmp(0) = &H80
        tmp(1) = &H90
        tmp(2) = 100
        tmp(3) = 110
        tmp(4) = 120
        Text2 = "Wrote 0x80 0x90 100 110 120"
    End If
    
    Dim x As String
    x = StrConv(tmp, vbUnicode)
    If Not mmf.WriteFile(x) Then
        MsgBox mmf.ErrorMessage, vbInformation
    End If

    
End Sub

Private Sub Command2_Click()
    
    Dim x As String
    Dim b() As Byte
    If Not mmf.ReadLength(x, 5) Then MsgBox mmf.ErrorMessage
    b() = StrConv(x, vbFromUnicode)
    For i = 0 To UBound(b)
        tmp = tmp & Hex(b(i)) & " "
    Next
    Text1 = tmp

    

End Sub

Private Sub Form_Load()
    
    If Not mmf.CreateMemMapFile("DAVES_VFILE", 20) Then
        MsgBox mmf.ErrorMessage, vbInformation
    End If
    
    Me.Caption = IIf(App.PrevInstance, "Second instance", "First Instance")
    
    If Not App.PrevInstance Then Call Command1_Click  'do a write to fill buff
    Command2_Click 'now display the buffer
    
End Sub
