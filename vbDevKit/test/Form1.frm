VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4815
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6825
   LinkTopic       =   "Form1"
   ScaleHeight     =   4815
   ScaleWidth      =   6825
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   2820
      TabIndex        =   0
      Top             =   2160
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim dlg As New CCmnDlg



Private Sub Command1_Click()
 
    MsgBox dlg.FolderDialog2

End Sub

Private Sub Form_Load()

    MsgBox dlg.FolderDialog2
    
'Dim c As New clsIniFile
'
'c.LoadFile App.Path & "\blah.ini"
'
'
'MsgBox c.AddKey("fart", "smart", "guy")
'c.AddKey "fart", "smart", "guy2"
'
'MsgBox c.SectionExists("fart")
'
'
'c.Save
'
'Shell "notepad """ & App.Path & "\blah.ini" & """", vbNormalFocus
'End

End Sub



'create from scratch
'c.AddSection "fart"
'c.AddKey "fart", "who", "you"
'
'c.AddSection "smell"
'c.AddKey "smell", "who", "me"
'
'c.AddSection "fart2"
'c.AddKey "fart2", "who2", "you2"
