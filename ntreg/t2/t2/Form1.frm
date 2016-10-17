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
    Dim r As New CNtRegistry
    r.hive = HKEY_LOCAL_MACHINE
    
    'MsgBox r.keyExists("\Registry\Machine\SOFTWARE\Microsoft\Windows\CurrentVersion\Run")
    'MsgBox r.keyExists("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Run")
    'MsgBox r.keyExists("hklm\SOFTWARE\Microsoft\Windows\CurrentVersion\Run")
    
    MsgBox r.CreateKey("hklm\SOFTWARE\Microsoft\Windows\CurrentVersion\Run\test\")
    
End Sub
