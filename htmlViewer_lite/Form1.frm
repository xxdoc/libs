VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   7980
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10440
   LinkTopic       =   "Form1"
   ScaleHeight     =   7980
   ScaleWidth      =   10440
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Load"
      Height          =   330
      Left            =   7380
      TabIndex        =   2
      Top             =   45
      Width           =   1365
   End
   Begin VB.TextBox Text1 
      Height          =   330
      Left            =   270
      OLEDropMode     =   1  'Manual
      TabIndex        =   1
      Text            =   "drag and drop here"
      Top             =   45
      Width           =   6900
   End
   Begin HtmlViewer.htmlControl htmlControl 
      Height          =   5415
      Left            =   45
      TabIndex        =   0
      Top             =   405
      Width           =   9195
      _ExtentX        =   16219
      _ExtentY        =   9551
      editor          =   0
      ww              =   -1  'True
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function GetTickCount Lib "kernel32" () As Long

Public Enum BYTEVALUES
  KiloByte = 1024
  MegaByte = 1048576
  GigaByte = 107374182
End Enum

Private Sub Command1_Click()
    On Error Resume Next
    Dim a As Long, b As Long, sz As Long, c As Long, ret As String
    a = GetTickCount
    htmlControl.LoadFile Text1
    b = GetTickCount
    sz = Len(htmlControl.text)
    c = b - a
    
    If sz > 1024 Then
        ret = "DataSize: " & ConvertKiloBytes(sz)
    Else
        ret = "DataSize: " & sz & " bytes "
    End If
    
    ret = ret & "   LoadTime: "
    If c > 2000 Then
        ret = ret & Round(CDbl(c / 1000), 3) & " secs"
    Else
        ret = ret & c & " ms"
    End If
    
    Me.Caption = ret
    
End Sub

Private Sub Form_Load()
    With htmlControl
        .WordWrap = True
        .EditorDisplayed = Highlight
        Text1 = App.path & "\test.html"
        Command1_Click
    End With
End Sub

Private Sub Form_Resize()
    htmlControl.MatchSize Me, 100
End Sub

Private Sub Text1_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    Text1 = Data.Files(1)
End Sub



Function ConvertKiloBytes(Bytes As Long, Optional NumDigitsAfterDecmal As Long = 0) As String
  ConvertKiloBytes = FormatNumber(Bytes / BYTEVALUES.KiloByte, NumDigitsAfterDecmal) & "kb "
End Function

