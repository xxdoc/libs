VERSION 5.00
Begin VB.Form frmMenuItem 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Menu Item was Clicked"
   ClientHeight    =   2625
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5910
   Icon            =   "MenuItem.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2625
   ScaleWidth      =   5910
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdOk 
      Caption         =   "Ok"
      Height          =   435
      Left            =   120
      TabIndex        =   4
      Top             =   2100
      Width           =   1155
   End
   Begin VB.Frame fraAction 
      Caption         =   "Menu item was clicked."
      Height          =   1935
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   5775
      Begin VB.OptionButton optAction 
         Caption         =   "Insert Sub-menu After this Point."
         Height          =   315
         Index           =   4
         Left            =   2640
         TabIndex        =   7
         Top             =   780
         Width           =   3015
      End
      Begin VB.OptionButton optAction 
         Caption         =   "Insert Item After this Point."
         Height          =   315
         Index           =   2
         Left            =   180
         TabIndex        =   6
         Top             =   1200
         Width           =   2235
      End
      Begin VB.OptionButton optAction 
         Caption         =   "Delete Item."
         Height          =   315
         Index           =   5
         Left            =   2640
         TabIndex        =   5
         Top             =   1200
         Width           =   1695
      End
      Begin VB.OptionButton optAction 
         Caption         =   "Insert Sub-menu Before this Point."
         Height          =   315
         Index           =   3
         Left            =   2640
         TabIndex        =   3
         Top             =   360
         Width           =   3015
      End
      Begin VB.OptionButton optAction 
         Caption         =   "Continue."
         Height          =   315
         Index           =   0
         Left            =   180
         TabIndex        =   2
         Top             =   360
         Value           =   -1  'True
         Width           =   1695
      End
      Begin VB.OptionButton optAction 
         Caption         =   "Insert Item Before this Point."
         Height          =   315
         Index           =   1
         Left            =   180
         TabIndex        =   1
         Top             =   780
         Width           =   2355
      End
   End
End
Attribute VB_Name = "frmMenuItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Using the Menu APIs to Grow or Shrink a Menu During Run-time
'(c) Jon Vote, 2003
'
'Idioma Software Inc.
'jon@idioma-software.com
'www.idioma-software.com
'www.skycoder.com

Option Explicit

Private m_maAction As MenuAction

Private Sub Form_Load()

  m_maAction = ACTION_CONTINUE
  
End Sub

'ProcessMenuClick: Public function - shows user item just clicked.
'Prompts for selection, returns choice.
Public Function ProcessMenuClick(ByVal strMenuItemCaption As String) As MenuAction

  Me.Caption = strMenuItemCaption
  fraAction.Caption = strMenuItemCaption & " was clicked."
  Me.Show vbModal
  ProcessMenuClick = m_maAction
  
End Function

'Converts option button selection to MenuAction
Private Function GetSelectedItem() As MenuAction
  
  Dim i As Integer
  
  For i = 0 To optAction.Count - 1
    If optAction(i).Value Then
      GetSelectedItem = i
      Exit For
    End If
  Next i
  
End Function

'Ok button
Private Sub cmdOk_Click()

  m_maAction = GetSelectedItem
  Unload Me
  
End Sub
