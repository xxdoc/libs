VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5625
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10200
   LinkTopic       =   "Form1"
   ScaleHeight     =   5625
   ScaleWidth      =   10200
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox Check1 
      Caption         =   "Check1"
      Height          =   375
      Left            =   8235
      TabIndex        =   4
      Top             =   1305
      Width           =   1320
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   510
      Left            =   8220
      TabIndex        =   3
      Top             =   225
      Width           =   1230
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   510
      Left            =   8235
      TabIndex        =   2
      Top             =   3330
      Width           =   735
   End
   Begin VB.ListBox List1 
      Height          =   2205
      Left            =   315
      TabIndex        =   1
      Top             =   3195
      Width           =   7845
   End
   Begin RichTextLib.RichTextBox rtf 
      Height          =   2805
      Left            =   270
      TabIndex        =   0
      Top             =   135
      Width           =   7800
      _ExtentX        =   13758
      _ExtentY        =   4948
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   3
      TextRTF         =   $"Form1.frx":0000
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim WithEvents re As clsRtfExtender
Attribute re.VB_VarHelpID = -1

 
Private Sub Command1_Click()
    List1.Clear
End Sub

Private Sub Command2_Click()
    Dim f As Integer
    MsgBox re.WordBeforeCursor(f, ".", " ", ":")
End Sub

Private Sub Form_Load()
    Set re = New clsRtfExtender
    re.InitRtf rtf.hwnd
    re.AutoIndent = True
    rtf.LoadFile App.Path & "\..\ips.txt"
    
End Sub

Private Sub re_ArrowDownLine(prevlineIndex As Long)
    List1.AddItem "ArrowDownLine"
End Sub

Private Sub re_ArrowUpLine(prevlineIndex As Long)
    List1.AddItem "ArrowupLine"
End Sub

Private Sub re_ClickedToNewLine(PrevLine As Long, curLine As Long)
    List1.AddItem "ClickedToNewLine"
End Sub

Private Sub re_newLine(lineIndex As Long)
    List1.AddItem "newLine"
End Sub

Private Sub re_Scrolled()
    List1.AddItem "Scrolled curline:" & re.TopLineIndex
End Sub

Private Sub rtf_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Me.Caption = re.WordUnderMouse
End Sub
