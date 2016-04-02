VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmDebugger 
   Caption         =   "VB Debugger"
   ClientHeight    =   3060
   ClientLeft      =   3150
   ClientTop       =   4140
   ClientWidth     =   6675
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   204
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   445
   Begin ComctlLib.ListView lvwThreads 
      Height          =   1110
      Left            =   0
      TabIndex        =   5
      Top             =   285
      Width           =   3090
      _ExtentX        =   5450
      _ExtentY        =   1958
      View            =   3
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   327682
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   2
      BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Thread ^"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   1
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Start Address "
         Object.Width           =   2646
      EndProperty
   End
   Begin ComctlLib.ListView lvwDLLs 
      Height          =   1110
      Left            =   3210
      TabIndex        =   4
      Top             =   270
      Width           =   3090
      _ExtentX        =   5450
      _ExtentY        =   1958
      View            =   3
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   327682
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   4
      BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "DLL ^"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   1
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Base Address "
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(3) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   2
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Thread "
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(4) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   3
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Path "
         Object.Width           =   5292
      EndProperty
   End
   Begin VB.TextBox txtOutput 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1110
      Left            =   0
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   1650
      Width           =   6270
   End
   Begin VB.Label lblThreads 
      AutoSize        =   -1  'True
      Caption         =   "&Threads:"
      Height          =   195
      Left            =   15
      TabIndex        =   3
      Top             =   45
      Width           =   630
   End
   Begin VB.Label lblDLLs 
      AutoSize        =   -1  'True
      Caption         =   "&Loaded DLLs:"
      Height          =   195
      Left            =   3195
      TabIndex        =   2
      Top             =   60
      Width           =   1020
   End
   Begin VB.Label lblOutput 
      AutoSize        =   -1  'True
      Caption         =   "&Debug output:"
      Height          =   195
      Left            =   15
      TabIndex        =   1
      Top             =   1410
      Width           =   1020
   End
End
Attribute VB_Name = "frmDebugger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*********************************************************************************************
'
' Debugging processes with VB
'
'*********************************************************************************************
'
' Author: Eduardo A. Morcillo
' E-Mail: e_morcillo@yahoo.com
' Web Page: http://www.domaindlx.com/e_morcillo
'
' Distribution: You can freely use this code in your own applications but you
'               can't publish this code in a web site, online service, or any
'               other media, without my express permission.
'
' Usage: at your own risk.
'
' Tested on: Windows 98 + VB5
'
' History:
'           03/09/2000 - This code was released
'
' Notes: Thanks to Stuart McBane for giving me this idea in
'        a newsgroup message.
'
'*********************************************************************************************
Option Explicit

Public lPID As Long
Public bRunning As Boolean

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   
   If bRunning Then
      MsgBox "The application is still running. Please close the application before closing this window.", vbInformation, Me.Caption
      Cancel = True
   End If

End Sub

Private Sub Form_Resize()
Dim H As Long

   On Error Resume Next
   
   H = ScaleHeight - lblThreads.Top - lblThreads.Height
   
   lvwThreads.Width = ScaleWidth / 2 - 1
   lvwThreads.Height = H / 2
   
   With lvwDLLs
      .Left = lvwThreads.Width + 1
      .Width = lvwThreads.Width
      .Height = lvwThreads.Height
   End With
   
   lblDLLs.Left = lvwThreads.Width + 1
   
   lblOutput.Move 2, lvwDLLs.Top + lvwDLLs.Height + 2
   txtOutput.Move 0, lblOutput.Top + lblOutput.Height + 2, ScaleWidth, ScaleHeight - lblOutput.Top - lblOutput.Height - 2
   
End Sub

Private Sub SortLV(ByVal LV As ListView, Column As ColumnHeader)
   
   With LV
      If .SortKey = Column.Index - 1 Then
         .SortOrder = 1 - .SortOrder
         If .SortOrder = lvwAscending Then
            Column.Text = Left$(Column.Text, Len(Column.Text) - 1) & "^"
         Else
            Column.Text = Left$(Column.Text, Len(Column.Text) - 1) & "v"
         End If
      Else
         With .ColumnHeaders(.SortKey + 1)
            .Text = Left$(.Text, Len(.Text) - 1)
         End With
         
         .SortKey = Column.Index - 1
         .SortOrder = lvwAscending
         Column.Text = Column.Text & "^"

      End If
   End With

End Sub

Private Sub lvwDLLs_ColumnClick(ByVal ColumnHeader As ComctlLib.ColumnHeader)

   SortLV lvwDLLs, ColumnHeader
   
End Sub

Private Sub lvwThreads_ColumnClick(ByVal ColumnHeader As ComctlLib.ColumnHeader)

   SortLV lvwThreads, ColumnHeader
   
End Sub
