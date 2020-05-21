VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   12255
   ClientLeft      =   2985
   ClientTop       =   2100
   ClientWidth     =   6585
   LinkTopic       =   "Form1"
   ScaleHeight     =   12255
   ScaleWidth      =   6585
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Click()
    DebugPrint "Two Random Numbers:", Rnd, Rnd
End Sub

