VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   9345
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   18615
   LinkTopic       =   "Form1"
   ScaleHeight     =   9345
   ScaleWidth      =   18615
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command6 
      Caption         =   "clone Kanal treeview"
      Height          =   405
      Left            =   1710
      TabIndex        =   15
      Top             =   5070
      Width           =   2235
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   120
      Top             =   4470
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Get Text from remote IE window"
      Height          =   435
      Left            =   780
      TabIndex        =   14
      Top             =   4560
      Width           =   2595
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Copy remote or local listview"
      Height          =   345
      Left            =   13440
      TabIndex        =   10
      Top             =   3270
      Width           =   4365
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   2955
      Left            =   12990
      TabIndex        =   9
      Top             =   270
      Width           =   5265
      _ExtentX        =   9287
      _ExtentY        =   5212
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "col1"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "col2"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "col3"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Copy remote or local combobox"
      Height          =   345
      Left            =   3990
      TabIndex        =   8
      Top             =   5040
      Width           =   2775
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   4020
      TabIndex        =   7
      Text            =   "Combo1"
      Top             =   4650
      Width           =   2745
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Copy remote or local listbox"
      Height          =   315
      Left            =   7710
      TabIndex        =   6
      Top             =   3240
      Width           =   4215
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
      Left            =   210
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   3
      Top             =   5580
      Width           =   6585
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Copy Remote Treeview"
      Height          =   525
      Left            =   5010
      TabIndex        =   2
      Top             =   3960
      Width           =   1725
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Left            =   2640
      TabIndex        =   1
      Text            =   "11995156"
      Top             =   3990
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
   Begin MSComctlLib.ListView ListView2 
      Height          =   3255
      Left            =   13020
      TabIndex        =   11
      Top             =   3720
      Width           =   5265
      _ExtentX        =   9287
      _ExtentY        =   5741
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "col1"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "col2"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "col3"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Image imgSniper 
      Height          =   480
      Left            =   420
      Picture         =   "Form1.frx":0000
      ToolTipText     =   "Drag & Drop over External IE Window"
      Top             =   3930
      Width           =   480
   End
   Begin VB.Label Label2 
      Caption         =   "Remote HWND"
      Height          =   255
      Left            =   1170
      TabIndex        =   13
      Top             =   4020
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "Text Output Pane"
      Height          =   225
      Left            =   270
      TabIndex        =   12
      Top             =   5310
      Width           =   1275
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
    
    w.hwnd = CLng(Text1)
    
    If Not w.isValid Then
        Text1 = Empty
        Exit Sub
    End If
        
    Set c = w.CopyRemoteTreeView(TreeView1)
    Text2 = ColToStr(c)
    
End Sub

Private Sub Command2_Click()
    Dim w As New Cwindow
    Dim c As Collection
    
    If Len(Text1) > 0 Then
        w.hwnd = CLng(Text1)
        Me.Caption = w.className
        If Not w.isValid Then
            Text1 = Empty
            Exit Sub
        End If
    Else
        w.hwnd = List1.hwnd
    End If
    
    Set c = w.CopyListBox(List2)
    
End Sub

Private Sub Command3_Click()
    Dim w As New Cwindow
    Dim c As Collection
    
    If Len(Text1) > 0 Then
        w.hwnd = CLng(Text1)
        Me.Caption = w.className
        If Not w.isValid Then
            Text1 = Empty
            Exit Sub
        End If
    Else
        w.hwnd = Combo1.hwnd
    End If
    
    Set c = w.CopyComboBox()
    Text2 = ColToStr(c)
End Sub

Private Sub Command4_Click()
    Dim w As New Cwindow
    Dim c As Collection
    
    If Len(Text1) > 0 Then
        w.hwnd = CLng(Text1)
        Me.Caption = w.className
        If Not w.isValid Then
            Text1 = Empty
            Exit Sub
        End If
    Else
        w.hwnd = ListView1.hwnd
    End If
    
    Set c = w.CopyRemoteListView(ListView2)
    Text2 = ColToStr(c)
    
End Sub
 

Private Sub Command5_Click()
    Dim w As New Cwindow
    Dim d As HTMLDocument
    Dim url As String
    Dim body As String
    
    If Len(Text1) > 0 Then
        w.hwnd = CLng(Text1)
        Me.Caption = w.className
        If Not w.isValid Then
            Text1 = "Invalid Hwnd"
            Exit Sub
        End If
    Else
        MsgBox "You must fill out a valid IE hwnd"
        Exit Sub
    End If
    
    Set d = w.IEDOMFromhWnd()
    If d Is Nothing Then
        Text2 = "Failed..."
    Else
        url = d.location.href
        body = d.body.innerHTML
        Text2 = url & vbCrLf & vbCrLf & body
    End If
    
End Sub

Private Sub Command6_Click()
    Dim c As Collection, c2 As Collection
    Dim w As Cwindow, wTv As Cwindow
    
    Set c = ws.ChildWindows()
    For Each w In c
        If VBA.Left(w.Caption, 5) = "KANAL" Then
            Set wTv = w.FindChild("SysTreeView32")
            If wTv.isValid Then
                 Set c2 = wTv.CopyRemoteTreeView(TreeView1)
                 Text2 = ColToStr(c2)
                 w.CloseWindow
            End If
            Exit Sub
        End If
    Next
    
    Text2 = "failed to find kanal window?"
    
End Sub

Private Sub Form_Load()
    Me.Caption = ws.GetWindowsVersion(True) & " - " & ws.GetWindowsVersionName
    
    For i = 0 To 10
        List1.AddItem Now & " test " & i
    Next
    
    For i = 0 To 10
        Combo1.AddItem Now & " combo test " & i
    Next
    
    Dim li As ListItem
    For i = 0 To 6
        Set li = ListView1.ListItems.Add(, , Now & " li text " & i)
        For j = 1 To ListView1.ColumnHeaders.count - 1
            li.SubItems(j) = "row " & i & " col " & j
        Next
    Next
    
    
End Sub

Private Sub imgSniper_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    Screen.MousePointer = 99 'custom
    Screen.MouseIcon = LoadResPicture("sniper.ico", vbResIcon)
    Timer1.Enabled = True
End Sub

Private Sub imgSniper_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    On Error Resume Next
    Timer1.Enabled = False
    Screen.MousePointer = vbDefault
    DoEvents
    Dim w As New Cwindow
    w.hwnd = CLng(Text1)
    Me.Caption = w.hwnd & " - " & w.className
End Sub

Private Sub Timer1_Timer()
   Text1 = WindowUnderCursor()
End Sub


