VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "MiniZipDir demo"
   ClientHeight    =   4965
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   5310
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   4965
   ScaleWidth      =   5310
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2475
      Left            =   180
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   240
      Width           =   2955
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'
'Be sure to set "Break on Unhandled Errors" in the
'Tools|Options... dialog when testing in the IDE!
'
Private Sub Form_Load()
    Dim Path As String

    ChDir App.Path
    ChDrive App.Path
    With New MiniZipDir
        .OpenZip App.Path & "\sample.zip"
        Path = .FirstFile()
        Do While Len(Path)
            Text1.SelText = Path & vbNewLine
            Path = .NextFile()
        Loop
        .CloseZip
    End With
End Sub

Private Sub Form_Resize()
    If WindowState <> vbMinimized Then
        Text1.Move 0, 0, ScaleWidth, ScaleHeight
    End If
End Sub
