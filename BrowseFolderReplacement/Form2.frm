VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   660
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7905
   LinkTopic       =   "Form2"
   ScaleHeight     =   660
   ScaleWidth      =   7905
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Height          =   315
      Left            =   120
      TabIndex        =   1
      Top             =   150
      Width           =   6225
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Browse Folder"
      Height          =   315
      Left            =   6510
      TabIndex        =   0
      Top             =   150
      Width           =   1245
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Dim f As New frmDlg
    Text1 = f.BrowseForFolder()
End Sub
