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
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()

    Const F = "C:\Documents and Settings\david\Desktop\courses\kav2012_12.0.0.374aEN_2777.exe"
    
    Dim h As New CWinHash
    
    hh = h.HashFile(F)
    If Len(hh) = 0 Then
        MsgBox h.error_message
    End If
    
    
    
End Sub
