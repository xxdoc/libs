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
   Begin VB.CommandButton Command4 
      Caption         =   "Save File"
      Height          =   510
      Left            =   1665
      TabIndex        =   3
      Top             =   810
      Width           =   1455
   End
   Begin VB.CommandButton Command3 
      Caption         =   "ColorDlg"
      Height          =   510
      Left            =   270
      TabIndex        =   2
      Top             =   810
      Width           =   1185
   End
   Begin VB.CommandButton Command2 
      Caption         =   "FolderDlg"
      Height          =   510
      Left            =   1710
      TabIndex        =   1
      Top             =   135
      Width           =   1410
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Folder Dlg 2"
      Height          =   510
      Left            =   270
      TabIndex        =   0
      Top             =   135
      Width           =   1185
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim dlg As New CCmnDlg

Private Sub Command1_Click()
    
    MsgBox dlg.FolderDialog2()
    MsgBox dlg.FolderDialog2(App.path)
    
    
End Sub

Private Sub Command2_Click()
    MsgBox dlg.FolderDialog()
End Sub

Private Sub Command3_Click()
    MsgBox Hex(dlg.ColorDialog())
    
End Sub

Private Sub Command4_Click()
    MsgBox dlg.SaveDialog("test.txt", , "Saveit")
    MsgBox dlg.SaveDialog("test.txt", App.path)
End Sub
