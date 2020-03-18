VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6000
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   6000
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkSmooth 
      Caption         =   "Smooth"
      Height          =   495
      Left            =   1740
      TabIndex        =   2
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Index           =   1
      Left            =   3480
      TabIndex        =   1
      Top             =   60
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Index           =   0
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'http://www.vbforums.com/showthread.php?884089-Progress-Bar-Using-API
'Feb 27th, 2020, 06:05 PM #1
'ChildOfTheKing
'Join Date  Jul 2016
'Posts 25

Public Sub Command1_Click(index As Integer)

  Call InitComctl32(ICC_PROGRESS_CLASS)
  
  ' PBS_SMOOTH: look of progressbar is smooth
  ' Without PBS_SMOOTH, Progress bar has a "standard" look
  ' WS_VISIBLE: Progress bar starts out visible
  
  Dim flags As Long
  flags = WS_CHILD Or WS_VISIBLE
  If chkSmooth.Value Then flags = WS_CHILD Or WS_VISIBLE Or PBS_SMOOTH
  
  If index = 0 Then ' Horizontal Progress Bar
    qwProgressBar = CreateWindowEx(0, PROGRESS_CLASS, vbNullString, flags, 15, 80, 240, 20, hWnd, 0, App.hInstance, ByVal 0)
  Else
    flags = flags Or PBS_VERTICAL
    qwProgressBar = CreateWindowEx(0, PROGRESS_CLASS, vbNullString, flags, 15, 80, 20, 225, hWnd, 0, App.hInstance, ByVal 0)
  End If

  
  ' Set Progress Bar Color
  Call SendMessage(qwProgressBar, PBM_SETBARCOLOR, 0, ByVal RGB(216, 43, 51)) ' Red
  ' Set Progress Bar Background Color
  Call SendMessage(qwProgressBar, PBM_SETBKCOLOR, 0, ByVal RGB(246, 243, 11)) ' Yellow
  ' Set Progress Bar Position
  Call SendMessage(qwProgressBar, PBM_SETPOS, 50&, 0&) ' Set Progress Bar to 50%
'  Call SendMessage(qwProgressBar, PBM_SETPOS, 0&, 0&) ' Reset Progressbar

  ' Set Progress Bar range and step interval
  Dim qwIterations As Long
  qwIterations = 25000
  Dim dwRange As Long
  dwRange = MAKELPARAM(0, qwIterations)
  Call SendMessage(qwProgressBar, PBM_SETRANGE, 0&, ByVal dwRange) ' Range
  Call SendMessage(qwProgressBar, PBM_SETSTEP, ByVal 1, 0&) ' Step interval
  
  '  Set the bar's parent; Used if Progress Bar is in a different dialog box than the Main.
  '  Call SetParent(qwProgressBarUpload, frmUploadFiles.hwnd)
  
  ' The following is a sample of showing the Progress Bar moving.
  Dim qwSpin As Long
  
  For qwSpin = 1 To qwIterations
  
    Call SendMessage(qwProgressBar, PBM_STEPIT, 0&, ByVal 0&)
    
  Next

End Sub

