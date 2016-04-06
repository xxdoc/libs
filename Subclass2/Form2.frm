VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text2 
      Height          =   975
      Left            =   60
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   2160
      Width           =   4515
   End
   Begin VB.TextBox Text1 
      Height          =   1995
      Left            =   60
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Text            =   "Form2.frx":0000
      Top             =   60
      Width           =   4515
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents SubClassMessage As clsSubClass
Attribute SubClassMessage.VB_VarHelpID = -1
Private Const WM_VScroll = &H115
Private Const WM_CHAR = &H102

Private Sub Form_Load()
    
    Set SubClassMessage = New clsSubClass
    
    If Not SubClassMessage.AttachMessage(Text1, WM_VScroll) Then
        MsgBox "WM_SCROLL: " & SubClassMessage.ErrorMessage
    End If
    
    If Not SubClassMessage.AttachMessage(Text1, WM_CHAR) Then
        MsgBox "WM_CHAR: " & SubClassMessage.ErrorMessage
    End If
        
    Dim tmp, i
    For i = 0 To 10
        tmp = tmp & String(20, Chr(65 + i)) & vbCrLf
    Next
    
    Text1.Text = tmp
        
End Sub

Private Sub SubClassMessage_MessageReceived(hwnd As Long, wMsg As Long, wParam As Long, lParam As Long, Cancel As Boolean)
    
    Text2.SelText = wMsg & vbCrLf
    Text2.SelLength = 0
    
    If wMsg = WM_VScroll Then Cancel = True
    
End Sub
