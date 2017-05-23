VERSION 5.00
Begin VB.Form frmPreview 
   Caption         =   "Menu Preview"
   ClientHeight    =   3240
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5895
   Icon            =   "frmPreview.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3240
   ScaleWidth      =   5895
   StartUpPosition =   1  '소유자 가운데
   Begin vbMenuBuilder.CodeEdit txtResult 
      Height          =   2535
      Left            =   120
      TabIndex        =   1
      Top             =   240
      Width           =   5655
      _extentx        =   9975
      _extenty        =   4471
      font            =   "frmPreview.frx":0E42
      backstyle       =   1
      linenumbers     =   0   'False
   End
   Begin VB.CommandButton cmdOK 
      Appearance      =   0  '평면
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   315
      Left            =   2040
      TabIndex        =   0
      Top             =   2880
      Width           =   1575
   End
   Begin VB.Menu mnuBlank 
      Caption         =   "mnuBlank"
      Visible         =   0   'False
   End
End
Attribute VB_Name = "frmPreview"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'
' Created & released by KSY, 06/14/2003
'
Private Sub cmdOK_Click()
   Unload Me
End Sub

Private Sub Form_Resize()
   On Error GoTo Bye
   With cmdOK
      .Move (ScaleWidth - .Width) / 2, ScaleHeight - .Height
   End With
   txtResult.Move 0, 0, ScaleWidth, cmdOK.Top - 50
Bye:
End Sub

