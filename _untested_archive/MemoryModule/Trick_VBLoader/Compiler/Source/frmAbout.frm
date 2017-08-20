VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About MyApp"
   ClientHeight    =   3750
   ClientLeft      =   2340
   ClientTop       =   1935
   ClientWidth     =   4800
   ClipControls    =   0   'False
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   204
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2588.317
   ScaleMode       =   0  'User
   ScaleWidth      =   4507.448
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   405
      Left            =   3420
      TabIndex        =   0
      Top             =   2970
      Width           =   1260
   End
   Begin VB.Image Image1 
      Height          =   960
      Left            =   60
      Picture         =   "frmAbout.frx":000C
      Top             =   240
      Width           =   960
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000010&
      X1              =   1014.176
      X2              =   4338.419
      Y1              =   372.718
      Y2              =   372.718
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Index           =   1
      X1              =   98.6
      X2              =   4394.762
      Y1              =   1832.528
      Y2              =   1832.528
   End
   Begin VB.Label lblDescription 
      ForeColor       =   &H00000000&
      Height          =   1680
      Left            =   1140
      TabIndex        =   1
      Top             =   810
      Width           =   3495
   End
   Begin VB.Label lblTitle 
      Caption         =   "Application Title"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   285
      Left            =   1080
      TabIndex        =   3
      Top             =   195
      Width           =   3600
   End
   Begin VB.Label lblVersion 
      Alignment       =   1  'Right Justify
      Caption         =   "Version"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000010&
      Height          =   180
      Left            =   1080
      TabIndex        =   4
      Top             =   615
      Width           =   3600
   End
   Begin VB.Label lblDisclaimer 
      Caption         =   "Warning: This program works correctly with programs compiled in vb6. With any other programs correct operation can't guaranteed."
      ForeColor       =   &H008080FF&
      Height          =   885
      Left            =   255
      TabIndex        =   2
      Top             =   2760
      Width           =   3030
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' // frmAbout.frm - "About" form of TrickVB6Installer application
' // © Krivous Anatoly Anatolevich (The trick), 2014

Option Explicit

Private Sub cmdOK_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    ' // Set icon
    SetWindowIcon Me.hWnd
    
    Me.Caption = "About " & App.Title
    lblVersion.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision & " by The Trick"
    lblTitle.Caption = App.Title
    lblDescription.Caption = "Keywords path:" & vbNewLine & vbNewLine & _
                           "<app> - application installed path;" & vbNewLine & _
                           "<win> - system windows directory;" & vbNewLine & _
                           "<sys> - System32 directory;" & vbNewLine & _
                           "<drv> - system drive;" & vbNewLine & _
                           "<tmp> - temporary directory;" & vbNewLine & _
                           "<dtp> - user desktop"
End Sub
