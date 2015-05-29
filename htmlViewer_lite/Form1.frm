VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   7980
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10440
   LinkTopic       =   "Form1"
   ScaleHeight     =   7980
   ScaleWidth      =   10440
   StartUpPosition =   2  'CenterScreen
   Begin HtmlViewer.htmlControl htmlControl 
      Height          =   5415
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9195
      _ExtentX        =   16219
      _ExtentY        =   9551
      editor          =   0
      ww              =   -1  'True
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    With htmlControl
        .WordWrap = True
        .EditorDisplayed = Highlight
        .LoadFile App.path & "\test.html"
    End With
End Sub

Private Sub Form_Resize()
    htmlControl.MatchSize Me, 100
End Sub
