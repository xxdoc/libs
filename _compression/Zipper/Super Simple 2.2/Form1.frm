VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Super Simple Zip Demo"
   ClientHeight    =   3420
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   4575
   BeginProperty Font 
      Name            =   "Segoe UI"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3420
   ScaleWidth      =   4575
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdZipIt 
      Caption         =   "Zip It"
      Enabled         =   0   'False
      Height          =   495
      Left            =   1680
      TabIndex        =   1
      Top             =   2460
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   2295
      Left            =   60
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   60
      Width           =   4455
   End
   Begin VB.Label lblStatus 
      BorderStyle     =   1  'Fixed Single
      Height          =   300
      Left            =   0
      TabIndex        =   2
      Top             =   3120
      Width           =   4575
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdZipIt_Click()
    Dim TextBLOB() As Byte

    Text1.Enabled = False
    cmdZipIt.Enabled = False

    'Zip Text1.Text as "Text.txt" in "Sample.zip" archive:
    With New ZipperSync
        TextBLOB = StrConv(Text1.Text, vbFromUnicode) 'Store as ANSI.
        .AddBLOB TextBLOB, "Text.txt"
        .Zip App.Path & "\Sample.zip"
    End With

    lblStatus.Caption = "Complete"
End Sub

Private Sub Text1_Change()
    If Len(Text1.Text) > 0 Then
        cmdZipIt.Enabled = True
        lblStatus.Caption = "Ready to zip"
    Else
        cmdZipIt.Enabled = False
        lblStatus.Caption = ""
    End If
End Sub
