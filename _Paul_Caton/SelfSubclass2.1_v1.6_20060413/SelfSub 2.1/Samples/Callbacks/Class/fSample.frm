VERSION 5.00
Begin VB.Form fMain 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   3345
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   7530
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3345
   ScaleWidth      =   7530
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frame 
      Caption         =   "Comm2"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1380
      Index           =   1
      Left            =   165
      TabIndex        =   4
      Top             =   1800
      Width           =   7155
      Begin VB.TextBox txtSend 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   1
         Left            =   2700
         TabIndex        =   6
         Top             =   360
         Width           =   4230
      End
      Begin VB.CommandButton cmdSend 
         Caption         =   "Send to Comm1"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   1
         Left            =   210
         TabIndex        =   5
         Top             =   360
         Width           =   2325
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Received:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   1740
         TabIndex        =   8
         Top             =   870
         Width           =   795
      End
      Begin VB.Label lblComm2 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   360
         Left            =   2700
         TabIndex        =   7
         Top             =   795
         Width           =   4230
      End
   End
   Begin VB.Frame frame 
      Caption         =   "Comm1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1380
      Index           =   0
      Left            =   165
      TabIndex        =   0
      Top             =   195
      Width           =   7155
      Begin VB.CommandButton cmdSend 
         Caption         =   "Send to Comm2"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   0
         Left            =   210
         TabIndex        =   2
         Top             =   360
         Width           =   2325
      End
      Begin VB.TextBox txtSend 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   0
         Left            =   2700
         TabIndex        =   1
         Text            =   "Hello Comm2, how are you?"
         Top             =   360
         Width           =   4230
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Received:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   1740
         TabIndex        =   9
         Top             =   870
         Width           =   795
      End
      Begin VB.Label lblComm1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   360
         Left            =   2700
         TabIndex        =   3
         Top             =   795
         Width           =   4230
      End
   End
End
Attribute VB_Name = "fMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents c1 As cComm1
Attribute c1.VB_VarHelpID = -1
Private WithEvents c2 As cComm2

Private Sub Form_Load()
  Set c1 = New cComm1
  Set c2 = New cComm2
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Set c1 = Nothing
  Set c2 = Nothing
End Sub

Private Sub cmdSend_Click(Index As Integer)
  Select Case Index
  Case 0
    c1.SendMessage c2, txtSend(0).Text
    
  Case 1
    c2.SendMessage c1, txtSend(1).Text
  
  End Select
End Sub

Private Sub txtSend_Change(Index As Integer)
  'cmdSend(Index).Enabled = Len(txtSend(Index)) > 0
End Sub

Private Sub c1_Comm1Msg(ByVal sText As Variant)
  lblComm1.Caption = " " & sText
End Sub

Private Sub c2_Comm2Msg(ByVal sText As Variant)
  lblComm2.Caption = " " & sText
End Sub
