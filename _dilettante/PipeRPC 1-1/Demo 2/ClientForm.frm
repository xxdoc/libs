VERSION 5.00
Begin VB.Form ClientForm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Calculation Client"
   ClientHeight    =   3015
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   3675
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "ClientForm"
   MaxButton       =   0   'False
   ScaleHeight     =   3015
   ScaleWidth      =   3675
   StartUpPosition =   2  'CenterScreen
   Begin Client.PipeRPC pipeCalculate 
      Left            =   2760
      Top             =   1980
      _ExtentX        =   847
      _ExtentY        =   847
      PipeName        =   "Calc Server Pipe"
   End
   Begin VB.OptionButton optOperation 
      Caption         =   "÷"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   3060
      TabIndex        =   5
      Top             =   1140
      Width           =   555
   End
   Begin VB.OptionButton optOperation 
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   2340
      TabIndex        =   4
      Top             =   1140
      Width           =   555
   End
   Begin VB.OptionButton optOperation 
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   1620
      TabIndex        =   3
      Top             =   1140
      Width           =   555
   End
   Begin VB.OptionButton optOperation 
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   900
      TabIndex        =   2
      Top             =   1140
      Value           =   -1  'True
      Width           =   555
   End
   Begin VB.TextBox txtResults 
      Height          =   315
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   2580
      Width           =   2235
   End
   Begin VB.CommandButton cmdCalculate 
      Caption         =   "="
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   2160
      TabIndex        =   7
      Top             =   2040
      Width           =   495
   End
   Begin VB.TextBox txtB 
      Height          =   315
      Left            =   1320
      TabIndex        =   6
      Text            =   "0"
      Top             =   1620
      Width           =   2235
   End
   Begin VB.TextBox txtA 
      Height          =   315
      Left            =   1320
      TabIndex        =   1
      Text            =   "0"
      Top             =   660
      Width           =   2235
   End
   Begin VB.TextBox txtServerName 
      Height          =   315
      Left            =   1320
      TabIndex        =   0
      Text            =   "."
      Top             =   120
      Width           =   2235
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   "C"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   60
      TabIndex        =   11
      Top             =   2610
      Width           =   1215
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "B"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   60
      TabIndex        =   10
      Top             =   1650
      Width           =   1215
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   60
      TabIndex        =   9
      Top             =   690
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Server Name"
      Height          =   255
      Left            =   60
      TabIndex        =   8
      Top             =   150
      Width           =   1215
   End
End
Attribute VB_Name = "ClientForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Request() As Byte
Private Response() As Byte
Private Operation As String

Private Sub cmdCalculate_Click()
    Request = txtA.Text & "|" _
            & Operation & "|" _
            & txtB.Text
    ReDim Response(199)
    pipeCalculate.PipeCall Request, Response
    txtResults.Text = Response
    txtA.SetFocus
End Sub

Private Sub Form_Load()
    Operation = "+"
End Sub

Private Sub optOperation_Click(Index As Integer)
    Operation = Choose(Index + 1, "+", "-", "*", "÷")
End Sub
