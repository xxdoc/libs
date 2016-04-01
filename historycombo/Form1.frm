VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "History Combo Box User control"
   ClientHeight    =   1995
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   1995
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Clear History"
      Height          =   285
      Left            =   1125
      TabIndex        =   3
      Top             =   720
      Width           =   1005
   End
   Begin VB.CommandButton cmdDoit 
      Caption         =   "Save Item"
      Height          =   285
      Left            =   3555
      TabIndex        =   1
      Top             =   720
      Width           =   915
   End
   Begin Project1.HistoryCombo hc 
      Height          =   330
      Left            =   1080
      TabIndex        =   0
      Top             =   225
      Width           =   3345
      _extentx        =   5900
      _extenty        =   582
   End
   Begin VB.Label Label1 
      Caption         =   "Enter some text in combo, and hit save item. Close the form and restart, your history will be saved across sessions. "
      Height          =   510
      Left            =   225
      TabIndex        =   2
      Top             =   1170
      Width           =   4200
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    hc.LoadHistory App.path & "\hc.dat"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    hc.SaveHistory
End Sub

Private Sub cmdDoit_Click()
    hc.RecordIfNew
End Sub

Private Sub Command1_Click()
    hc.ClearHistory
End Sub
