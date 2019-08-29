VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5205
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   2955
   LinkTopic       =   "Form1"
   ScaleHeight     =   5205
   ScaleWidth      =   2955
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Caption         =   "Frame2"
      Height          =   2955
      Left            =   1200
      TabIndex        =   3
      Top             =   2160
      Width           =   1575
      Begin VB.CommandButton Command1 
         Caption         =   "Command1"
         Height          =   435
         Left            =   180
         TabIndex        =   5
         Top             =   240
         Width           =   1155
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   3015
      Left            =   60
      TabIndex        =   2
      Top             =   1380
      Width           =   2295
      Begin VB.TextBox Text1 
         Height          =   315
         Left            =   240
         TabIndex        =   4
         Text            =   "Text1"
         Top             =   420
         Width           =   1695
      End
   End
   Begin VB.PictureBox pictStack 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   390
      Left            =   60
      Picture         =   "Form1.frx":0000
      ScaleHeight     =   390
      ScaleWidth      =   2100
      TabIndex        =   1
      Top             =   600
      Width           =   2100
   End
   Begin VB.PictureBox pictCode 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   390
      Left            =   0
      Picture         =   "Form1.frx":2AEC
      ScaleHeight     =   390
      ScaleWidth      =   2100
      TabIndex        =   0
      Top             =   60
      Width           =   2100
   End
   Begin VB.Image pictTabLine 
      Height          =   435
      Left            =   1800
      Picture         =   "Form1.frx":55D8
      Stretch         =   -1  'True
      Top             =   900
      Width           =   675
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    
    With pictCode
        pictStack.Move .Left, .Top, .Width, .Height
        pictTabLine.Move .Left + .Width, .Top, pictTabLine.Width, .Height
        pictStack.Visible = False
    End With
    
    With Frame1
        .Top = pictStack.Top + pictStack.Height + 20
        Frame2.Move .Left, .Top, .Width, .Height
        Frame2.Visible = False
    End With
    
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    pictTabLine.Width = Me.Width - pictTabLine.Left - 10
    Frame1.Width = Me.Width - 200
    Frame2.Width = Frame1.Width
    Frame1.Height = Me.Height - Frame1.Top - 500
    Frame2.Height = Frame1.Height
End Sub

Private Sub pictCode_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Debug.Print X
End Sub

Private Sub pictCode_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If X >= 990 And X <= 1620 Then
        pictStack.Visible = True
        pictCode.Visible = False
        Frame2.Visible = True
        Frame1.Visible = False
    End If
End Sub
 
Private Sub pictStack_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If X < 990 Then
        pictStack.Visible = False
        pictCode.Visible = True
        Frame2.Visible = False
        Frame1.Visible = True
    End If
End Sub

