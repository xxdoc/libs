VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "VB Developers Kit -  http://sandsprite.com"
   ClientHeight    =   4590
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8010
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4590
   ScaleWidth      =   8010
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.Timer Timer1 
      Interval        =   2000
      Left            =   120
      Top             =   3180
   End
   Begin VB.Label Label2 
      Caption         =   $"Form1.frx":0000
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Index           =   0
      Left            =   240
      TabIndex        =   4
      Top             =   2400
      Width           =   7395
   End
   Begin VB.Label Label2 
      Caption         =   $"Form1.frx":0091
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Index           =   1
      Left            =   240
      TabIndex        =   3
      Top             =   3780
      Width           =   7395
   End
   Begin VB.Label lblLink 
      Caption         =   "http://sandsprite.com/vbdevkit/"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   1620
      MousePointer    =   1  'Arrow
      TabIndex        =   2
      Top             =   3180
      Width           =   4455
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " VB Developers Kit"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   1155
      Left            =   180
      TabIndex        =   1
      Top             =   60
      Width           =   7695
   End
   Begin VB.Label Label1 
      Caption         =   $"Form1.frx":012A
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   915
      Left            =   240
      TabIndex        =   0
      Top             =   1380
      Width           =   7215
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Type POINTAPI
    X As Long
    Y As Long
End Type
Private Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Sub Form_Load()

End Sub

Private Sub lblLink_Click()
    On Error Resume Next
    ShellExecute Me.hWnd, vbNullString, "http://sandsprite.com/vbdevkit/", vbNullString, "C:\", 1
End Sub


Private Sub Timer1_Timer()
    Timer1.Enabled = False
    isInitalized = True
    
    Dim pt As POINTAPI, hWnd As Long
    
    'GetCursorPos pt
    'hWnd = WindowFromPoint(pt.X, pt.Y)
    'If hWnd <> Me.hWnd Then
        Unload Me
    'End If
    
End Sub
