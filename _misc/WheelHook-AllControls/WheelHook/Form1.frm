VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   7140
   ClientLeft      =   750
   ClientTop       =   1155
   ClientWidth     =   10725
   LinkTopic       =   "Form1"
   ScaleHeight     =   7140
   ScaleWidth      =   10725
   Begin VB.ListBox List1 
      Height          =   1230
      Left            =   120
      TabIndex        =   6
      Top             =   5760
      Width           =   5175
   End
   Begin VB.PictureBox Picture2 
      Height          =   975
      Left            =   2640
      ScaleHeight     =   915
      ScaleWidth      =   2475
      TabIndex        =   5
      Top             =   3240
      Width           =   2535
   End
   Begin VB.PictureBox Picture1 
      Height          =   975
      Left            =   120
      ScaleHeight     =   915
      ScaleWidth      =   2355
      TabIndex        =   4
      Top             =   3240
      Width           =   2415
   End
   Begin VB.TextBox Text1 
      Height          =   1215
      Left            =   5400
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Text            =   "Form1.frx":0000
      Top             =   5760
      Width           =   5175
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   120
      TabIndex        =   2
      Text            =   "Combo1"
      Top             =   5280
      Width           =   5175
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid2 
      Height          =   5535
      Left            =   5400
      TabIndex        =   1
      Top             =   120
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   9763
      _Version        =   393216
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   3015
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   5318
      _Version        =   393216
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
  Dim N As Integer
  Dim I As Integer
  
  ' Dummy values
  With MSFlexGrid1
    .Rows = 100
    .Cols = 5
    For N = .FixedRows To .Rows - 1
      .TextMatrix(N, 0) = "Row " & N
    Next N
  End With
  
  With MSFlexGrid2
    .Rows = 100
    .Cols = 5
    For N = .FixedRows To .Rows - 1
      .TextMatrix(N, 0) = "Row " & N
    Next N
  End With
  
  For N = 0 To 20
    Combo1.AddItem "Test " & N
    List1.AddItem "Test " & N
  Next N

  ' Hook Form
  Call WheelHook(Me.hWnd)
  
  ' Hook Controls to be ignored
  Call WheelHook(Combo1.hWnd)
  Call WheelHook(List1.hWnd)
  Call WheelHook(Text1.hWnd)
  
  Form2.Show
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Call WheelUnHook(Me.hWnd)
  Call WheelUnHook(Combo1.hWnd)
  Call WheelUnHook(Text1.hWnd)
  Unload Form2
End Sub

' Here you can add scrolling support to controls that don't normally respond.
' This Sub could always be moved to a module to make scrollwheel behaviour
' generic across forms.
' ===========================================================================
Public Sub MouseWheel(ByVal MouseKeys As Long, ByVal Rotation As Long, ByVal Xpos As Long, ByVal Ypos As Long)
  Dim ctl As Control
  Dim bHandled As Boolean
  Dim bOver As Boolean
  
  For Each ctl In Controls
    ' Is the mouse over the control
    On Error Resume Next
    bOver = (ctl.Visible And IsOver(ctl.hWnd, Xpos, Ypos))
    On Error GoTo 0
    
    If bOver Then
      ' If so, respond accordingly
      bHandled = True
      Select Case True
      
        Case TypeOf ctl Is MSFlexGrid
          FlexGridScroll ctl, MouseKeys, Rotation, Xpos, Ypos
          
        Case TypeOf ctl Is PictureBox
          PictureBoxZoom ctl, MouseKeys, Rotation, Xpos, Ypos
          
        Case TypeOf ctl Is ListBox, TypeOf ctl Is TextBox, TypeOf ctl Is ComboBox
          ' These controls already handle the mousewheel themselves, so allow them to:
          If ctl.Enabled Then ctl.SetFocus
          
        Case Else
          bHandled = False

      End Select
      If bHandled Then Exit Sub
    End If
    bOver = False
  Next ctl
  
  ' Scroll was not handled by any controls, so treat as a general message send to the form
  Me.Caption = "Form Scroll " & IIf(Rotation < 0, "Down", "Up")
End Sub
