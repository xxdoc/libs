VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   645
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   645
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdDiv 
      Caption         =   "Divider"
      Height          =   330
      Left            =   3600
      TabIndex        =   3
      Top             =   180
      Width           =   1005
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Send printf"
      Height          =   330
      Left            =   1305
      TabIndex        =   2
      Top             =   180
      Width           =   1095
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear"
      Height          =   330
      Left            =   2520
      TabIndex        =   1
      Top             =   180
      Width           =   960
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Send"
      Height          =   330
      Left            =   135
      TabIndex        =   0
      Top             =   180
      Width           =   1050
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdClear_Click()
    debugClear
End Sub

Private Sub cmdDiv_Click()
    debugDiv
End Sub

Private Sub Command1_Click()
    debugPrint "v1", 2, 3
End Sub

Private Sub Command2_Click()
    debugPrintf "%s( %d, 0x%X)", "func", 21, &HDD
End Sub
