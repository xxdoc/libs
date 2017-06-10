VERSION 5.00
Object = "*\AHyperLbl.vbp"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   Caption         =   "Label on steroids, but still light weight"
   ClientHeight    =   6915
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7710
   LinkTopic       =   "Form1"
   ScaleHeight     =   6915
   ScaleWidth      =   7710
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Properties"
      Height          =   4455
      Left            =   5370
      TabIndex        =   1
      Top             =   90
      Width           =   2265
      Begin VB.CommandButton cmdRestoreSize 
         Caption         =   "&Restore Size"
         Enabled         =   0   'False
         Height          =   420
         Left            =   120
         TabIndex        =   15
         Top             =   2055
         Width           =   1890
      End
      Begin VB.CheckBox chkAutoSize 
         Caption         =   "Auto&Size"
         Height          =   240
         Left            =   165
         TabIndex        =   6
         Top             =   1740
         Width           =   1725
      End
      Begin VB.CommandButton cmdPlainCaption 
         Caption         =   "&PlainCaption..."
         Height          =   435
         Left            =   135
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   3885
         Width           =   1890
      End
      Begin VB.CommandButton cmdForeColor 
         Height          =   285
         Left            =   135
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   3450
         Width           =   1890
      End
      Begin VB.CommandButton cmdBackColor 
         Height          =   285
         Left            =   135
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   2835
         Width           =   1890
      End
      Begin VB.CheckBox chkWordWrap 
         Caption         =   "&Word wrap"
         Height          =   240
         Left            =   165
         TabIndex        =   5
         Top             =   1410
         Width           =   1725
      End
      Begin VB.CheckBox chkAutoNavigate 
         Caption         =   "&AutoNavigate"
         Height          =   240
         Left            =   165
         TabIndex        =   4
         Top             =   1080
         Width           =   1725
      End
      Begin VB.ComboBox cmbBorder 
         Height          =   315
         ItemData        =   "Form1.frx":0000
         Left            =   135
         List            =   "Form1.frx":000A
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   600
         Width           =   1965
      End
      Begin VB.Label Label3 
         Caption         =   "&Fore Color:"
         Height          =   225
         Left            =   150
         TabIndex        =   9
         Top             =   3225
         Width           =   1500
      End
      Begin VB.Label Label2 
         Caption         =   "Back &Color:"
         Height          =   225
         Left            =   150
         TabIndex        =   7
         Top             =   2610
         Width           =   1500
      End
      Begin VB.Label Label1 
         Caption         =   "&Border style:"
         Height          =   240
         Left            =   150
         TabIndex        =   2
         Top             =   345
         Width           =   1875
      End
   End
   Begin HyperLbl.HyperLabel HyperLabel1 
      Height          =   4365
      Left            =   105
      TabIndex        =   0
      Top             =   135
      Width           =   5085
      Visible         =   0   'False
      _ExtentX        =   8969
      _ExtentY        =   7699
      BorderStyle     =   1
      Caption         =   $"Form1.frx":0022
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "&Update Caption"
      Height          =   360
      Left            =   90
      TabIndex        =   14
      Top             =   6525
      Width           =   1695
   End
   Begin VB.TextBox txtCaption 
      Height          =   1575
      Left            =   60
      MultiLine       =   -1  'True
      TabIndex        =   13
      Text            =   "Form1.frx":03FC
      Top             =   4845
      Width           =   7560
   End
   Begin MSComDlg.CommonDialog cdlColor 
      Left            =   4935
      Top             =   4275
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin VB.Label Label4 
      Caption         =   "C&aption:"
      Height          =   210
      Left            =   90
      TabIndex        =   12
      Top             =   4635
      Width           =   1230
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : Form1
' DateTime  : 04 jan 2006 03:19
' Author    : Joacim Andersson, Brixoft Software, http://www.brixoft.net
' Purpose   : Simple demo application for the HyperLabel control
'---------------------------------------------------------------------------------------
Option Explicit

Private Function ShowColorDialog(ByVal nColor As Long) As Long
    On Error Resume Next
    cdlColor.Flags = cdlCCRGBInit
    cdlColor.Color = nColor
    cdlColor.ShowColor
    If Err.Number <> cdlCancel Then
        ShowColorDialog = cdlColor.Color
    Else
        ShowColorDialog = -1
    End If
End Function

Private Sub chkAutoNavigate_Click()
    HyperLabel1.AutoNavigate = (chkAutoNavigate.Value = vbChecked)
End Sub

Private Sub chkAutoSize_Click()
    HyperLabel1.AutoSize = (chkAutoSize.Value = vbChecked)
    cmdRestoreSize.Enabled = Not HyperLabel1.AutoSize
End Sub

Private Sub chkWordWrap_Click()
    HyperLabel1.WordWrap = (chkWordWrap.Value = vbChecked)
End Sub

Private Sub cmbBorder_Click()
    HyperLabel1.BorderStyle = cmbBorder.ListIndex
End Sub

Private Sub cmdBackColor_Click()
    Dim nColor As Long
    
    nColor = ShowColorDialog(cmdBackColor.BackColor)
    If nColor <> -1 Then
        cmdBackColor.BackColor = nColor
        HyperLabel1.BackColor = nColor
    End If
End Sub

Private Sub cmdForeColor_Click()
    Dim nColor As Long
    
    nColor = ShowColorDialog(cmdForeColor.BackColor)
    If nColor <> -1 Then
        cmdForeColor.BackColor = nColor
        HyperLabel1.ForeColor = nColor
    End If
End Sub

Private Sub cmdPlainCaption_Click()
    MsgBox HyperLabel1.PlainCaption
End Sub

Private Sub cmdRestoreSize_Click()
    HyperLabel1.Move txtCaption.Left, 90, 5085, 4365
    cmdRestoreSize.Enabled = False
End Sub

Private Sub cmdUpdate_Click()
    HyperLabel1.Caption = txtCaption.Text
End Sub

Private Sub Form_Load()
    With HyperLabel1
        cmbBorder.ListIndex = .BorderStyle
        chkAutoNavigate.Value = Abs(.AutoNavigate)
        chkWordWrap.Value = Abs(.WordWrap)
        chkAutoSize.Value = Abs(.AutoSize)
        cmdBackColor.BackColor = .BackColor
        cmdForeColor.BackColor = .ForeColor
        txtCaption.Text = .Caption
        .RenderToContainer = True
    End With
End Sub

Private Sub Form_Resize()
    With txtCaption
        .Move .Left, Me.ScaleHeight - .Height - cmdUpdate.Height - .Left * 2, Me.ScaleWidth - .Left * 2
        Label4.Top = .Top - Label4.Height
        cmdUpdate.Move .Left, .Top + .Height + .Left
        HyperLabel1.Left = .Left
        Frame1.Move Me.ScaleWidth - Frame1.Width - .Left, Label4.Top - Frame1.Height
    End With
End Sub

Private Sub HyperLabel1_Click()
    Debug.Print "Click event fired"
End Sub

Private Sub HyperLabel1_DblClick()
    Debug.Print "DblClick event fired"
End Sub

Private Sub HyperLabel1_HyperlinkClick(ByVal URL As String)
    'This event is only raised if the
    'AutoNavigate property is set to False
    MsgBox "HyperlinkClick event fired for the following link:" & vbCrLf & URL
End Sub
