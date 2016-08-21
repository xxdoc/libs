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
    
    'testing the auto drop of the dependancy dll/exe from resource file...works ok
    
    Dim p As New CProcessLib
    Dim c As Collection
    
    'Set c = p.EnumMutexes
    '
    'Dim x As New Cx64
    'x.isExe_x64 App.Path & "\" & App.EXEName
    
    Me.Visible = True
    
    Dim cmd As New CCmdOutput2
    cmd.CfgOpts False, 10
    a = cmd.LaunchProcess("c:\windows\system32\notepad.exe")
    Me.Caption = a & "Exit code: " & cmd.exitCode
    
    
    
End Sub
