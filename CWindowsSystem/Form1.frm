VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   9345
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13410
   LinkTopic       =   "Form1"
   ScaleHeight     =   9345
   ScaleWidth      =   13410
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      Height          =   345
      Left            =   12420
      TabIndex        =   9
      Top             =   3210
      Width           =   945
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   6990
      TabIndex        =   8
      Text            =   "Combo1"
      Top             =   3240
      Width           =   2115
   End
   Begin VB.TextBox Text3 
      Height          =   315
      Left            =   9510
      TabIndex        =   7
      Top             =   3240
      Width           =   1365
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   315
      Left            =   11010
      TabIndex        =   6
      Top             =   3210
      Width           =   1245
   End
   Begin VB.ListBox List2 
      Height          =   3570
      Left            =   7200
      TabIndex        =   5
      Top             =   3600
      Width           =   5295
   End
   Begin VB.ListBox List1 
      Height          =   2985
      Left            =   7260
      TabIndex        =   4
      Top             =   180
      Width           =   5205
   End
   Begin VB.TextBox Text2 
      Height          =   3615
      Left            =   180
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   3
      Top             =   4590
      Width           =   6585
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   525
      Left            =   3510
      TabIndex        =   2
      Top             =   3960
      Width           =   1725
   End
   Begin VB.TextBox Text1 
      Height          =   405
      Left            =   1020
      TabIndex        =   1
      Text            =   "11995156"
      Top             =   4020
      Width           =   2295
   End
   Begin MSComctlLib.TreeView TreeView1 
      Height          =   3765
      Left            =   150
      TabIndex        =   0
      Top             =   120
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   6641
      _Version        =   393217
      Style           =   7
      Appearance      =   1
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 Dim ws As New CWindowsSystem
 
 Private Sub Command1_Click()
    Dim c As Collection
    Dim w As New Cwindow
    
    Dim tmp() As String
    
    w.hWnd = CLng(Text1)
    Set c = w.CopyRemoteTv(TreeView1)
    Text2 = ColToStr(c)
    
End Sub

Private Sub Command2_Click()
    Dim w As New Cwindow
    Dim c As Collection
    
    If Len(Text3) = 0 Then
        w.hWnd = List1.hWnd
    Else
        w.hWnd = CLng(Text3)
        Me.Caption = w.ClassName
    End If
    
    Set c = w.CopyListBox(List2)
    
End Sub

Private Sub Command3_Click()
    Dim w As New Cwindow
    Dim c As Collection
    
    If Len(Text3) = 0 Then
        w.hWnd = Combo1.hWnd
    Else
        w.hWnd = CLng(Text3)
        Me.Caption = w.ClassName
    End If
    
    Set c = w.CopyComboBox()
    Text2 = ColToStr(c)
End Sub

Private Sub Form_Load()
    Me.Caption = ws.GetWindowsVersion(True) & " - " & ws.GetWindowsVersionName
    
    For i = 0 To 10
        List1.AddItem "test " & i
    Next
    
    For i = 0 To 10
        Combo1.AddItem "combo test " & i
    Next
    
End Sub
