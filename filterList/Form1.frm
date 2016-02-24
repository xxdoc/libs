VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   8730
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   9540
   LinkTopic       =   "Form1"
   ScaleHeight     =   8730
   ScaleWidth      =   9540
   StartUpPosition =   3  'Windows Default
   Begin Project1.ucFilterList lvFilter 
      Height          =   4650
      Left            =   180
      TabIndex        =   0
      Top             =   135
      Width           =   8520
      _extentx        =   15028
      _extenty        =   8202
   End
   Begin VB.Menu mnuPopup 
      Caption         =   "mnuPopup"
      Begin VB.Menu mnuPopupTest 
         Caption         =   "Test"
      End
      Begin VB.Menu mnuCopyTable 
         Caption         =   "Copy Table"
      End
      Begin VB.Menu mnuCopySecondSel 
         Caption         =   "Copy 2nd Column Selected"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'author:  David Zimmer <dzzie@yahoo.com>
'site:    http://sandsprite.com
'License: free for any use

Private Sub Form_Load()
    
    mnuPopup.Visible = False
    lvFilter.MultiSelect = True
    lvFilter.FilterColumn = 1
    lvFilter.SetColumnHeaders "test1,test2,test3,test4"
    
    Dim li As ListItem
    For i = 0 To 5
        Set li = lvFilter.AddItem("text" & i)
        li.SubItems(1) = "taco1 " & i
        li.SubItems(2) = "test3 " & i
        li.SubItems(3) = "test4 " & i
        
        Set li = lvFilter.AddItem("item " & i)
        li.SubItems(1) = "item taco2  " & i
        li.SubItems(2) = "item 2 test " & i
        li.SubItems(3) = "item 2 test " & i
    Next
    
    
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    lvFilter.Width = Me.Width - lvFilter.Left - 300
End Sub

Private Sub lvFilter_Click()
    Me.Caption = "lvFilter_Click"
End Sub

Private Sub lvFilter_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    Me.Caption = "lvFilter_ColumnClick(" & ColumnHeader.Text & ")"
End Sub

Private Sub lvFilter_DblClick()
    Me.Caption = "lvFilter_DblClick"
End Sub

Private Sub lvFilter_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then PopupMenu mnuPopup
End Sub

Private Sub mnuCopySecondSel_Click()
    Clipboard.Clear
    Clipboard.SetText lvFilter.GetAllText(2, True)
End Sub

Private Sub mnuCopyTable_Click()
    Clipboard.Clear
    Clipboard.SetText lvFilter.GetAllElements()
End Sub

Private Sub mnuPopupTest_Click()
    Dim li As ListItem
    On Error Resume Next
    Set li = lvFilter.selItem
    If li Is Nothing Then Exit Sub
    MsgBox li.Text
End Sub
