VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form fMain 
   Caption         =   "Shadow"
   ClientHeight    =   1860
   ClientLeft      =   60
   ClientTop       =   360
   ClientWidth     =   3510
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   1860
   ScaleWidth      =   3510
   StartUpPosition =   3  'Windows Default
   Begin MSComCtl2.UpDown udTrans 
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   795
      Width           =   270
      _ExtentX        =   476
      _ExtentY        =   661
      _Version        =   393216
      Increment       =   5
      Max             =   255
      Enabled         =   -1  'True
   End
   Begin MSComCtl2.UpDown udDepth 
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   240
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   661
      _Version        =   393216
      Max             =   32
      Enabled         =   -1  'True
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Another form..."
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   1350
      Width           =   1455
   End
   Begin MSComCtl2.UpDown udR 
      Height          =   375
      Left            =   2160
      TabIndex        =   5
      Top             =   240
      Width           =   270
      _ExtentX        =   476
      _ExtentY        =   661
      _Version        =   393216
      Increment       =   5
      Max             =   255
      Enabled         =   -1  'True
   End
   Begin MSComCtl2.UpDown udG 
      Height          =   375
      Left            =   2160
      TabIndex        =   6
      Top             =   795
      Width           =   270
      _ExtentX        =   476
      _ExtentY        =   661
      _Version        =   393216
      Increment       =   5
      Max             =   255
      Enabled         =   -1  'True
   End
   Begin MSComCtl2.UpDown udB 
      Height          =   375
      Left            =   2160
      TabIndex        =   7
      Top             =   1350
      Width           =   270
      _ExtentX        =   476
      _ExtentY        =   661
      _Version        =   393216
      Increment       =   5
      Max             =   255
      Enabled         =   -1  'True
   End
   Begin VB.Label lblB 
      AutoSize        =   -1  'True
      Caption         =   "Blue:"
      Height          =   195
      Left            =   2505
      TabIndex        =   10
      Top             =   1440
      Width           =   360
   End
   Begin VB.Label lblG 
      AutoSize        =   -1  'True
      Caption         =   "Green:"
      Height          =   195
      Left            =   2505
      TabIndex        =   9
      Top             =   885
      Width           =   495
   End
   Begin VB.Label lblR 
      AutoSize        =   -1  'True
      Caption         =   "Red:"
      Height          =   195
      Left            =   2505
      TabIndex        =   8
      Top             =   330
      Width           =   345
   End
   Begin VB.Label lblTrans 
      AutoSize        =   -1  'True
      Caption         =   "Transparency:"
      Height          =   195
      Left            =   480
      TabIndex        =   4
      Top             =   885
      Width           =   1050
   End
   Begin VB.Label lblDepth 
      AutoSize        =   -1  'True
      Caption         =   "Depth:"
      Height          =   195
      Left            =   480
      TabIndex        =   2
      Top             =   330
      Width           =   495
   End
End
Attribute VB_Name = "fMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*************************************************************************************************
'* Self-subclassing class sample - demonstrates form shadows.
'*
'* Paul_Caton@hotmail.com
'* Copyright free, use and abuse as you see fit.
'*
'* v1.0 Original version................................................................. 20060322
'*************************************************************************************************

Option Explicit

Private oShadow As cShadow

Private Sub Form_Load()
  Set oShadow = New cShadow

  With oShadow
    If .Shadow(Me) Then
      udDepth.Value = .Depth
      udTrans.Value = .Transparency
      udR = 0
      udG = 0
      udB = 0
    Else
      udDepth.Enabled = False
      udTrans.Enabled = False
      MsgBox "The cShadow class requires Windows 2000 or better, and no less than 16 bit color to display form shadows.", vbInformation
    End If
  End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Set oShadow = Nothing
End Sub

Private Sub Command1_Click()
  Dim frm As New fMain
  
  On Error Resume Next
  Load frm

  With frm
    .udDepth.Value = Me.udDepth.Value
    .udTrans.Value = Me.udTrans.Value
    .Show vbModeless
  End With
  
  On Error GoTo 0
End Sub

Private Sub udB_Change()
  lblB.Caption = "Blue: " & udB
  oShadow.Color = RGB(udR, udG, udB)
End Sub

Private Sub udG_Change()
  lblG.Caption = "Green: " & udG
  oShadow.Color = RGB(udR, udG, udB)
End Sub

Private Sub udR_Change()
  lblR.Caption = "Red: " & udR
  oShadow.Color = RGB(udR, udG, udB)
End Sub

Private Sub udDepth_Change()
  lblDepth.Caption = "Depth: " & udDepth.Value
  oShadow.Depth = udDepth.Value
End Sub

Private Sub udTrans_Change()
  lblTrans.Caption = "Transparency: " & udTrans.Value
  oShadow.Transparency = udTrans.Value
End Sub
