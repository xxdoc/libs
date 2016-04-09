VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmBinCode 
   Caption         =   "BinCode"
   ClientHeight    =   6615
   ClientLeft      =   165
   ClientTop       =   480
   ClientWidth     =   9660
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   441
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   644
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.Toolbar tb 
      Align           =   1  'Align Top
      Height          =   540
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   9660
      _ExtentX        =   17039
      _ExtentY        =   953
      ButtonWidth     =   820
      ButtonHeight    =   794
      Appearance      =   1
      ImageList       =   "il"
      DisabledImageList=   "ild"
      HotImageList    =   "ilh"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   8
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Open (Ctrl + O)"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.ToolTipText     =   "Copy (Ctrl + C)"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.ToolTipText     =   "8 bit (String)"
            ImageIndex      =   3
            Style           =   2
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.ToolTipText     =   "32 bit (Long)"
            ImageIndex      =   4
            Style           =   2
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.ToolTipText     =   "64 bit (Currency)"
            ImageIndex      =   5
            Style           =   2
            Value           =   1
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
            Object.Width           =   100
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.ToolTipText     =   "Refresh"
            ImageIndex      =   6
         EndProperty
      EndProperty
      BorderStyle     =   1
      Begin VB.TextBox txtVarName 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   2610
         TabIndex        =   2
         Text            =   "z_Code"
         ToolTipText     =   "Array name"
         Top             =   90
         Width           =   1305
      End
   End
   Begin MSComctlLib.ImageList il 
      Left            =   8685
      Top             =   5745
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   16711935
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "BinCode.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "BinCode.frx":0712
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "BinCode.frx":0B68
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "BinCode.frx":127A
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "BinCode.frx":198C
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "BinCode.frx":209E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.TextBox txtCode 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4665
      Left            =   315
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   1620
      Width           =   7575
   End
   Begin MSComDlg.CommonDialog cd 
      Left            =   8040
      Top             =   5835
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList ild 
      Left            =   8685
      Top             =   5010
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   16711935
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "BinCode.frx":255B
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "BinCode.frx":2C6D
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "BinCode.frx":3381
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "BinCode.frx":3A93
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "BinCode.frx":41A5
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "BinCode.frx":48B7
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ilh 
      Left            =   8685
      Top             =   4275
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   16711935
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "BinCode.frx":4D68
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "BinCode.frx":547A
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "BinCode.frx":5B8E
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "BinCode.frx":62A0
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "BinCode.frx":69B2
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "BinCode.frx":70C4
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmBinCode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private bData()     As Byte
Private nLen        As Long

Private Declare Sub RtlMoveMemory Lib "kernel32" (ByVal Destination As Long, ByVal Source As Long, ByVal Length As Long)

Private Sub Form_Load()
  Show
  tb_ButtonClick tb.Buttons(1)
End Sub

Private Sub Form_Resize()
  If WindowState <> vbMinimized Then
    txtCode.Move 0, tb.Height, Me.ScaleWidth, ScaleHeight - tb.Height
  End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  If Shift = 2 Then
    If KeyCode = 67 Then
      Clipboard.Clear
      Clipboard.SetText txtCode.Text
    ElseIf KeyCode = 79 Then
      tb_ButtonClick tb.Buttons(1)
    End If
  End If
End Sub

Private Sub tb_ButtonClick(ByVal Button As MSComctlLib.Button)
  Select Case Button.Index
  Case 1
    If OpenFile Then
      txtCode.Text = vbNullString
      
      With tb.Buttons
        .Item(2).Enabled = True
        .Item(4).Enabled = True
        .Item(5).Enabled = True
        .Item(6).Enabled = True
      End With
      
      Select Case True
      Case tb.Buttons(4).Value = tbrPressed
        txtVarName.Enabled = False
        tb.Buttons(8).Enabled = False
        ProcessString

      Case tb.Buttons(5).Value = tbrPressed
        txtVarName.Enabled = True
        tb.Buttons(8).Enabled = True
        ProcessLong

      Case tb.Buttons(6).Value = tbrPressed
        txtVarName.Enabled = True
        tb.Buttons(8).Enabled = True
        ProcessCurr

      End Select
    End If

  Case 2
    Clipboard.Clear
    Clipboard.SetText txtCode.Text
    
  Case 4
    txtVarName.Enabled = False
    tb.Buttons(8).Enabled = False
    ProcessString
    
  Case 5
    txtVarName.Enabled = True
    tb.Buttons(8).Enabled = True
    ProcessLong
    
  Case 6
    txtVarName.Enabled = True
    tb.Buttons(8).Enabled = True
    ProcessCurr
    
  Case 8
    Select Case True
    Case tb.Buttons(4).Value = tbrPressed
      ProcessString
    Case tb.Buttons(5).Value = tbrPressed
      ProcessLong
    Case tb.Buttons(6).Value = tbrPressed
      ProcessCurr
    End Select
  End Select
End Sub

Private Function OpenFile() As Boolean
  With cd
    On Error GoTo Cancel

    .CancelError = True
    .Filter = "Bin files|*.bin"
    .DialogTitle = "Open .bin file"
    .Flags = &H4
    .ShowOpen

    Open .FileName For Binary As #1

    nLen = LOF(1)
    If nLen Mod 8 Then
      nLen = nLen + (8 - (nLen Mod 8))
    End If
      
    ReDim bData(0 To nLen - 1)

    Get #1, , bData
    Close #1
    
    OpenFile = True
Cancel:
  End With
End Function

Private Sub ProcessCurr()
  Dim c As Currency
  Dim i As Long
  Dim s As String

  Caption = App.Title & " - " & cd.FileTitle & " As Currency"
  txtCode.Text = vbNullString

  For i = 0 To (nLen \ 8) - 1
    c = Get8(i)

    If c <> 0@ Then
      s = s & txtVarName.Text & "(" & i & ") = " & c & "@"

      If Len(s) > 900 Then
        txtCode.Text = txtCode.Text & s & vbNewLine
        s = vbNullString
      Else
        s = s & ": "
      End If
    End If
  Next i

  If Right$(s, 1) = ":" Then
    s = Left$(s, Len(s) - 1)
  Else
    If Right$(s, 2) = ": " Then
      s = Left$(s, Len(s) - 2)
    End If
  End If
  txtCode.Text = txtCode.Text & s
End Sub

Private Sub ProcessLong()
  Dim i As Long
  Dim l As Long
  Dim s As String
  
  Caption = App.Title & " - " & cd.FileTitle & " As Long"
  txtCode.Text = vbNullString

  For i = 0 To (nLen \ 4) - 1
    l = Get4(i)

    If l <> 0 Then
      s = s & txtVarName.Text & "(" & i & ") = &H" & Hex$(l) & "&"

      If Len(s) > 900 Then
        txtCode.Text = txtCode.Text & s & vbNewLine
        s = vbNullString
      Else
        s = s & ": "
      End If
    End If
  Next i

  If Right$(s, 1) = ":" Then
    s = Left$(s, Len(s) - 1)
  Else
    If Right$(s, 2) = ": " Then
      s = Left$(s, Len(s) - 2)
    End If
  End If

  txtCode.Text = txtCode.Text & s
End Sub

Private Sub ProcessString()
  Dim i As Long

  Caption = App.Title & " - " & cd.FileTitle & " As String"
  txtCode.Text = """"

  For i = 0 To nLen - 1
    txtCode.Text = txtCode.Text & HexFmt(Get1(i))
  Next i

  txtCode.Text = txtCode.Text & """"
End Sub

Private Function HexFmt(ByVal b As Byte) As String
  HexFmt = Right$("0" & Hex$(b), 2)
End Function

Private Function Get1(ByVal nIndex As Long) As Byte
  Get1 = bData(nIndex)
End Function

Private Function Get4(ByVal nIndex As Long) As Long
  RtlMoveMemory VarPtr(Get4), VarPtr(bData(nIndex * 4)), 4
End Function

Private Function Get8(ByVal nIndex As Long) As Currency
  RtlMoveMemory VarPtr(Get8), VarPtr(bData(nIndex * 8)), 8
End Function

Private Sub txtVarName_KeyPress(KeyAscii As Integer)
  If KeyAscii = 32 Then
    KeyAscii = 0
    Beep
  End If
End Sub
