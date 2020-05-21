VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Form1"
   ClientHeight    =   7845
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   15810
   LinkTopic       =   "Form2"
   ScaleHeight     =   7845
   ScaleWidth      =   15810
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   15
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6675
      Left            =   360
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   300
      Width           =   15075
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()

    Dim rh As New CRichHeader 'todo: clear_data no counts for hashing?
    Dim hash As New CWinHash  'dzrt reference (CRichHeader has no dependancies)
    Dim ret() As String
    
    If rh.Load("D:\_code\libs\pe_lib2\_sppe2.dll") Then
        'yara workbench: pe.dbg(hash.md5(pe.rich_signature.clear_data)) = 869487aa2b9d84eb86208f66ccbe7dd0
        push ret, "hash.HashBytes(rh.clearData)     = " & hash.HashBytes(rh.clearData)           ' = 869487AA2B9D84EB86208F66CCBE7DD0
        push ret, "hash.HashString(rh.strClearData) = " & hash.HashString(rh.strClearData)       ' = 869487AA2B9D84EB86208F66CCBE7DD0
        push ret, "version(7299,rhMasm613) = " & rh.version(7299, rhMasm613) 'yara pe.rich_signature.version(verion,[toolId]) compatiable
        Text1 = Join(ret, vbCrLf) & vbCrLf
    End If
    
    Text1 = Text1 & rh.dump 'ok if failed
   
   
End Sub


