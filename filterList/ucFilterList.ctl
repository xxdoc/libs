VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.UserControl ucFilterList 
   ClientHeight    =   6315
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7605
   ScaleHeight     =   6315
   ScaleWidth      =   7605
   Begin VB.Timer tmrFilter 
      Enabled         =   0   'False
      Interval        =   400
      Left            =   5880
      Top             =   4440
   End
   Begin VB.TextBox txtFilter 
      Height          =   330
      Left            =   540
      TabIndex        =   3
      Top             =   4320
      Width           =   1995
   End
   Begin MSComctlLib.ListView lvFilter 
      Height          =   3300
      Left            =   360
      TabIndex        =   1
      Top             =   360
      Visible         =   0   'False
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   5821
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
      NumItems        =   0
   End
   Begin MSComctlLib.ListView lv 
      Height          =   4155
      Left            =   45
      TabIndex        =   0
      Top             =   90
      Width           =   7170
      _ExtentX        =   12647
      _ExtentY        =   7329
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
      NumItems        =   0
   End
   Begin VB.Image imgX 
      Height          =   225
      Left            =   2580
      Picture         =   "ucFilterList.ctx":0000
      Stretch         =   -1  'True
      Top             =   4380
      Width           =   165
   End
   Begin VB.Label Label1 
      Caption         =   "Filter"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   45
      TabIndex        =   2
      Top             =   4320
      Width           =   420
   End
   Begin VB.Menu mnuTools 
      Caption         =   "mnuTools"
      Begin VB.Menu mnuCopyAll 
         Caption         =   "Copy All"
      End
      Begin VB.Menu mnuCopySel 
         Caption         =   "Copy Sel"
      End
      Begin VB.Menu mnuCopyColumn 
         Caption         =   "Copy Column"
      End
      Begin VB.Menu mnuTotalCol 
         Caption         =   "Total Column"
      End
      Begin VB.Menu mnuspacer4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFilterHelp 
         Caption         =   "Filter Help"
      End
      Begin VB.Menu mnuSetFilterCol 
         Caption         =   "Set Filter Column"
      End
      Begin VB.Menu mnuResults 
         Caption         =   "Results:"
      End
      Begin VB.Menu mnuspacer 
         Caption         =   "-"
      End
      Begin VB.Menu mnuToggleMulti 
         Caption         =   "MultiSelect"
      End
      Begin VB.Menu mnuHideSel 
         Caption         =   "Hide Selection"
      End
      Begin VB.Menu mnuSelectInverse 
         Caption         =   "Inverse Selection"
      End
      Begin VB.Menu mnuAlertColWidths 
         Caption         =   "Alert Column Widths (IDE Only)"
      End
   End
End
Attribute VB_Name = "ucFilterList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'author:  David Zimmer <dzzie@yahoo.com>
'site:    http://sandsprite.com
'License: free for any use
'
'note this version is not synced with the newest!

Public AllowDelete As Boolean

Private m_Locked As Boolean
Private m_FilterColumn As Long
Private m_FilterColumnPreset As Long

'we need to track the index map between listviews in case they delete from lvFilter..
Private indexMapping As Collection

Event Click()
'Event ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
Event DblClick()
Event ItemClick(ByVal Item As MSComctlLib.ListItem)
Event MouseUp(Button As Integer, Shift As Integer, X As Single, y As Single)

Const LVM_FIRST = &H1000
Const LVM_GETSELECTEDCOUNT = (LVM_FIRST + 50)

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long


#If 0 Then
    Dim X, y, Column, nextone 'force lowercase so ide doesnt switch around on its own whim...
#End If

'note when locked you wont receive events, and can not add items..
Property Get Locked() As Boolean
    Locked = m_Locked
End Property

Property Let Locked(X As Boolean)
    m_Locked = X
    txtFilter.BackColor = IIf(X, &HC0C0C0, vbWhite)
    txtFilter.Enabled = Not X
End Property
    
Property Get SelCount() As Long
    Dim v As ListView
    Dim li As ListItem
    Dim cnt As Long
    
    Set v = currentLV
    cnt = SendMessage(v.hwnd, LVM_GETSELECTEDCOUNT, 0, 0)
    SelCount = cnt
    
End Property

Property Get selItems() As Collection

    Dim c As New Collection
    Dim li As ListItem
    Dim cnt As Long
    
    Set selItems = c
 
    For Each li In currentLV.ListItems
        If li.selected Then c.Add li
    Next
    
End Property
    
Property Get FilterColumn() As Long
    FilterColumn = m_FilterColumn
End Property

Property Let FilterColumn(X As Long)
    On Error Resume Next
    Dim tmp As String
    Dim ch As ColumnHeader
    
    If lv.ColumnHeaders.Count = 0 Then
        m_FilterColumnPreset = X
        Exit Property
    End If
    
    If X <= 0 Then X = 1
    
    If X > lv.ColumnHeaders.Count Then
        X = lv.ColumnHeaders.Count
    End If
    
    'remove the visual marker that this is the filter column
    Set ch = lv.ColumnHeaders(m_FilterColumn)
    ch.Text = Trim(Replace(ch.Text, "*", Empty))
    
    Set ch = lvFilter.ColumnHeaders(m_FilterColumn)
    ch.Text = Trim(Replace(ch.Text, "*", Empty))

    'add the visual marker to the new column
    Set ch = lv.ColumnHeaders(X)
    ch.Text = ch.Text & " *"
    
    Set ch = lvFilter.ColumnHeaders(X)
    ch.Text = ch.Text & " *"

    m_FilterColumn = X
    
End Property

'doesnt seem to work as intended in all cases?
'note this only hands out a ref to the main listview not filtered
'this is only for compatability with existing code to make integration easier..
Property Get ListItems() As ListItems
    Set ListItems = lv.ListItems
End Property

Property Get MultiSelect() As Boolean
    MultiSelect = lv.MultiSelect
End Property

Property Let MultiSelect(X As Boolean)
    lv.MultiSelect = X
    lvFilter.MultiSelect = X
    mnuToggleMulti.Checked = X
End Property

Property Get HideSelection() As Boolean
    HideSelection = lv.MultiSelect
End Property

Property Let HideSelection(X As Boolean)
    lv.HideSelection = X
    lvFilter.HideSelection = X
    mnuHideSel.Checked = X
End Property

Property Get GridLines() As Boolean
    GridLines = lv.GridLines
End Property

Property Let GridLines(X As Boolean)
    lv.GridLines = X
    lvFilter.GridLines = X
End Property

'which ever one is currently displayed
Property Get currentLV() As ListView
    On Error Resume Next
    If lvFilter.Visible Then
        Set currentLV = lvFilter
    Else
        Set currentLV = lv
    End If
End Property

Property Get mainLV() As ListView
    Set mainLV = lv
End Property


'compatability with normal listview
Property Get SelectedItem() As ListItem
    Set SelectedItem = selItem
End Property

Property Get selItem() As ListItem
    On Error Resume Next
    If lvFilter.Visible Then
        Set selItem = lvFilter.SelectedItem
    Else
        Set selItem = lv.SelectedItem
    End If
End Property

Property Get Filter() As String
    Filter = txtFilter
End Property

Property Let Filter(txt As String)
     txtFilter = txt
End Property

Function AddItem(txt, ParamArray subItems()) As ListItem
    On Error Resume Next
    
    Dim i As Integer
    
    If m_Locked Then Exit Function
    
    Set AddItem = lv.ListItems.Add(, , CStr(txt))
    
    For Each si In subItems
        AddItem.subItems(i + 1) = si
        i = i + 1
    Next
    
    ApplyFilter
    
End Function

Sub Clear()

    If m_Locked Then Exit Sub
    
    Dim li As ListItem
    For Each li In lv.ListItems
        If IsObject(li.Tag) Then Set li.Tag = Nothing
    Next
    
    For Each li In lvFilter.ListItems
        If IsObject(li.Tag) Then Set li.Tag = Nothing
    Next
    
    lv.ListItems.Clear
    lvFilter.ListItems.Clear
End Sub

Sub SetFont(Name As String, Size As Long)
    lv.Font.Name = Name
    lv.Font.Size = Size
    lvFilter.Font.Name = Name
    lvFilter.Font.Size = Size
    txtFilter.Font.Name = Name
    txtFilter.Font.Size = Size
End Sub

Sub SetColumnHeaders(csvList As String, Optional csvWidths As String)
    
    On Error Resume Next
    Dim i As Long, fc As Long, ch As ColumnHeader
    
    fc = -1
    lv.ColumnHeaders.Clear
    lvFilter.ColumnHeaders.Clear
    
    tmp = Split(csvList, ",")
    For Each t In tmp
        i = i + 1
        If InStr(t, "*") > 0 Then
            fc = i
            t = Trim(Replace(t, "*", Empty))
        End If
        lv.ColumnHeaders.Add , , Trim(t)
        lvFilter.ColumnHeaders.Add , , Trim(t)
    Next
    
    If fc <> -1 Then FilterColumn = fc  'this sets the visual marker on the column if they specified it..
    If m_FilterColumnPreset <> -1 Then FilterColumn = m_FilterColumnPreset 'they called FilterColumn manually first, now apply..
    If m_FilterColumn = -1 Then FilterColumn = 1 'they never specified it so default to first column
    
    If Len(csvWidths) > 0 Then
        tmp = Split(csvWidths, ",")
        For i = 0 To UBound(tmp)
            If Len(tmp(i)) > 0 Then
                lv.ColumnHeaders(i + 1).Width = CLng(tmp(i))
                lvFilter.ColumnHeaders(i + 1).Width = CLng(tmp(i))
            End If
        Next
    End If
    
End Sub

Sub SelectAll(Optional selected As Boolean = True)
    Dim v As ListView, li As ListItem
    
    If Not Me.MultiSelect Then Exit Sub
    If lv.Visible Then Set v = lv Else Set v = lvFilter
    
    For Each li In v
        li.selected = selected
    Next
        
End Sub

Private Sub imgX_Click()
    txtFilter.Text = Empty
End Sub

Private Sub lv_KeyDown(KeyCode As Integer, Shift As Integer)

    Dim i As Long
    Dim li As ListItem
    
    On Error Resume Next
    
    If m_Locked Then Exit Sub
    
    If KeyCode = vbKeyDelete And AllowDelete Then
        For i = lv.ListItems.Count To 1 Step -1
            If lv.ListItems(i).selected Then lv.ListItems.Remove i
        Next
    End If
    
    If KeyCode = vbKeyA And Shift = 2 Then SelectAll 'ctrl-a
    
    If KeyCode = vbKeyC And Shift = 2 Then 'ctrl-C
        Clipboard.Clear
        Clipboard.SetText GetAllElements()
    End If
    
             
End Sub

Private Sub lvFilter_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim i As Long
    Dim liMain As ListItem
    
    On Error Resume Next
    
    If m_Locked Then Exit Sub
    
    If KeyCode = vbKeyDelete And AllowDelete Then
        For i = lvFilter.ListItems.Count To 1 Step -1
            If lvFilter.ListItems(i).selected Then
                Set liMain = getMainListItemFor(lvFilter.ListItems(i))
                If Not liMain Is Nothing Then lv.ListItems.Remove liMain.Index
                lvFilter.ListItems.Remove i
            End If
        Next
    End If
    
    If KeyCode = vbKeyA And Shift = 2 Then SelectAll 'ctrl-a
             
    If KeyCode = vbKeyC And Shift = 2 Then 'ctrl-C
        Clipboard.Clear
        Clipboard.SetText GetAllElements()
    End If
    
End Sub

Function totalColumn(ByVal colIndex As Long, Optional ByRef hadErr As Boolean) As Long
    On Error Resume Next
    Dim i As Long, tot As Long, li As ListItem
    
    colIndex = colIndex - 1 'we expect a 1 based index to be consistant
    
    If colIndex < 0 Or colIndex > currentLV.ColumnHeaders.Count Then
        hadErr = True
        Exit Function
    End If
    
    hadErr = False
    For Each li In currentLV.ListItems
        If colIndex = 0 Then
           If Len(li.Text) > 0 Then i = CLng(li.Text)
        Else
            If Len(li.subItems(colIndex)) > 0 Then i = CLng(li.subItems(colIndex))
        End If
        tot = tot + i
    Next
        
    hadErr = Not (Err.Number = 0)
    totalColumn = tot

End Function

Private Sub mnuTotalCol_Click()
    Dim i As Long, tmp As String, b As Boolean, tot As Long
    On Error Resume Next
    tmp = InputBox("Enter column index to total (1-" & (lv.ColumnHeaders.Count) & ")")
    If Len(tmp) = 0 Then Exit Sub
    i = CLng(tmp)
    If Err.Number <> 0 Then Exit Sub
    tot = totalColumn(i, b)
    MsgBox "Total for " & lv.ColumnHeaders(i).Text & " = " & tot & IIf(b, " An error was generated", ""), vbInformation
End Sub

Private Sub mnuAlertColWidths_Click()
    Dim tmp(), c As ColumnHeader
    For Each c In lv.ColumnHeaders
        push tmp, Round(c.Width)
    Next
    InputBox "Column Widths are: ", , Join(tmp, ",")
End Sub

Private Sub Label1_Click()
    If m_Locked Then Exit Sub
    mnuResults.Caption = "Results: " & Me.currentLV.ListItems.Count
    PopupMenu mnuTools
End Sub

Private Sub mnuCopyAll_Click()
    Clipboard.Clear
    Clipboard.SetText Me.GetAllElements()
End Sub

Private Sub mnuCopyColumn_Click()
    On Error Resume Next
    Dim X As Long
    X = InputBox("Enter column index to copy", , 1)
    If Len(X) = 0 Then Exit Sub
    X = CLng(X) - 1 'we are 0 based internally..
    Clipboard.Clear
    Clipboard.SetText Me.GetAllText(X)
End Sub

Private Sub mnuCopySel_Click()
    Clipboard.Clear
    Clipboard.SetText Me.GetAllElements(True)
End Sub

Private Sub mnuFilterHelp_Click()
    
    Const msg = "You can enter multiple criteria to filter \n" & _
                "on by seperating with commas. You can also\n" & _
                "utilize a subtractive filter if the first \n" & _
                "character in the textbox is a minus sign\n" & _
                "Filter also understands: bold,selected, color:red|blue|etc\n\n" & _
                "The FilterColumn is marked with an * this is \n" & _
                "the column that is being searched. You can \n" & _
                "modify it on the filter menu, or by entering\n" & _
                "/[index] in the filter textbox and hitting return\n" & _
                "/t [1-ColCount] | or colName to total column is also supported\n" & _
                "/d will filter for unique entries only\n" & _
                "numeric columns support > [index] or < [index] filters\n\n" & _
                "Pressing escape in the filter textbox will clear it.\n\n" & _
                "If the AllowDelete property has been set, you can\n" & _
                "select list items and press the delete key to remove\n" & _
                "them."
                
                
    MsgBox Replace(msg, "\n", vbCrLf), vbInformation
                
End Sub

Private Sub mnuHideSel_Click()
    Me.HideSelection = Not lv.HideSelection
End Sub

Private Sub mnuSelectInverse_Click()
    InvertSelection
End Sub

Public Sub InvertSelection()
    If Not MultiSelect Then Exit Sub
    Dim li As ListItem
    For Each li In Me.currentLV.ListItems
        li.selected = Not li.selected
    Next
End Sub

Private Sub mnuSetFilterCol_Click()
    On Error Resume Next
    Dim X As Long
    X = InputBox("Enter column that filter searches", , FilterColumn)
    If Len(X) = 0 Then Exit Sub
    X = CLng(X)
    FilterColumn = X
End Sub

Private Sub mnuToggleMulti_Click()
    Me.MultiSelect = Not lv.MultiSelect
End Sub

Function ColorConstantsToLong(ByVal s As String) As Long
    
    Dim c As ColorConstants
    s = LCase(s)
    
    c = -1
    If InStr(s, "black") > 0 Then c = vbBlack
    If InStr(s, "blue") > 0 Then c = vbBlue
    If InStr(s, "cyan") > 0 Then c = vbCyan
    If InStr(s, "green") > 0 Then c = vbGreen
    If InStr(s, "magenta") > 0 Then c = vbMagenta
    If InStr(s, "red") > 0 Then c = vbRed
    If InStr(s, "white") > 0 Then c = vbWhite
    If InStr(s, "yellow") > 0 Then c = vbYellow
    
    ColorConstantsToLong = c
    
End Function

Private Sub tmrFilter_Timer()
    tmrFilter.Enabled = False
    Call ApplyFilter
End Sub

'on huge lists it can take a while so let them finish typing first
Private Sub txtFilter_Change()
    If lv.ListItems.Count > 100 Then
        tmrFilter.Enabled = False 'reset the timer it will apply once they pause and wait
        tmrFilter.Enabled = True
    Else
        ApplyFilter
    End If
End Sub

Function distinctFilter()
    Dim vals As New Collection, v As String
    Dim li As ListItem
    
    On Error Resume Next
    
    If m_FilterColumn = -1 Then Exit Function
    
    lvFilter.Visible = True
    lvFilter.ListItems.Clear
    
    For Each li In lv.ListItems
        If m_FilterColumn = 1 Then
            v = li.Text
        Else
            v = li.subItems(m_FilterColumn - 1)
        End If
        If Not KeyExistsInCollection(vals, v) Then
            vals.Add v, v
            CloneListItemTo li, lvFilter
            'Debug.Print "unique: " & v
        End If
    Next
    
End Function

Private Function myIsNumeric(ByVal v, ByRef outV As Long) As Boolean
    On Error GoTo hell
    If IsNumeric(v) Then
        outV = CLng(v)
    Else
        v = Replace(v, "0x", Empty)
        outV = CLng("&h" & v)
    End If
    myIsNumeric = True
hell:
End Function

Sub ApplyFilter()
    Dim li As ListItem
    Dim t As String
    Dim useSubtractiveFilter As Boolean
    Dim tmp() As String, addit As Boolean, X
    Dim gtMode As Boolean, ltMode As Boolean
    Dim uv As Long, v As Long
    
    On Error Resume Next
    
    If m_Locked Then Exit Sub
    
    If Len(txtFilter) = 0 Then GoTo hideExit
    
    If Len(txtFilter) = 1 Then
        If VBA.Left(txtFilter, 1) = "/" Then GoTo hideExit
    End If
        
    If VBA.Left(txtFilter, 1) = "/" Then
        t = Replace(txtFilter, "/", Empty)
        If IsNumeric(t) Then GoTo hideExit 'they are going to change the FilterColumn on "cmdline"
    End If
    
    If txtFilter = "/distinct" Or txtFilter = "/d" Then
         distinctFilter
         Exit Sub
    End If
    
    If VBA.Left(txtFilter, 1) = ">" Then
        gtMode = True
        If Len(Trim(txtFilter)) = 1 Then GoTo hideExit
    End If
    
    If VBA.Left(txtFilter, 1) = "<" Then
        ltMode = True
        If Len(Trim(txtFilter)) = 1 Then GoTo hideExit
    End If
    
    If VBA.Left(txtFilter, 1) = "-" Then 'they are typing a subtractive filter..give them time to formulate it..
        If Len(txtFilter) = 1 Then GoTo hideExit
        If VBA.Right(txtFilter, 1) = "," Then Exit Sub 'they are adding more criteria
    End If

    If Left(txtFilter, 2) = "/t" Then
         t = Trim(Mid(txtFilter, 3))
         If myIsNumeric(t, uv) Then  'we will use converted UserVal (uv) value below..
            If uv > 0 Or uv <= lv.ColumnHeaders.Count Then
                v = totalColumn(uv, addit)
                MsgBox "Total for " & lv.ColumnHeaders(uv).Text & " = " & v & IIf(addit, " - An error was generated", ""), vbInformation
                txtFilter = Empty
            End If
         Else
            uv = ColIndexForName(t) + 1 '0 based , -1 on error
            If uv > 0 Then
                tot = totalColumn(uv, addit)
                MsgBox "Total for " & lv.ColumnHeaders(uv).Text & " = " & tot & IIf(addit, vbCrLf & vbCrLf & " - An error was generated", ""), vbInformation
                txtFilter = Empty
            End If
         End If
         GoTo hideExit
    End If
    
    'should multiple (csv) filters only apply on hitting return?
    'so you can see full list to work off of?
    
    lvFilter.Visible = True
    lvFilter.ListItems.Clear
    Set indexMapping = New Collection
    
    Dim sMatch As String
    Dim isColor As Boolean
    Dim lColor As Long
    
    If VBA.Left(txtFilter, 1) = "-" Then
        useSubtractiveFilter = True
        sMatch = Mid(txtFilter, 2)
    ElseIf VBA.Left(txtFilter, 6) = "color:" Then
        isColor = True
        sMatch = Replace(txtFilter, "color:", Empty)
        If Len(sMatch) = 0 Then Exit Sub 'they are still entering it...
        Err.Clear
        lColor = CLng(sMatch)
        If Err.Number <> 0 Then lColor = ColorConstantsToLong(sMatch)
        If lColor = -1 Then Exit Sub
    Else
        sMatch = txtFilter
    End If
    
    If ltMode Or gtMode Then
        t = Mid(txtFilter, 2)
        If Not myIsNumeric(t, uv) Then 'we will use converted UserVal (uv) value below..
            ltMode = False
            gtMode = False
        End If
    End If
    
    'we allow for csv multiple criteria, also
    'you can use a subtractive filter like -mnu,cmd,lv
     For Each li In lv.ListItems
        
         If FilterColumn = 1 Then
            t = li.Text
         Else
            t = li.subItems(m_FilterColumn - 1)
         End If
         
         addit = False
         
         If gtMode Or ltMode Then
            If myIsNumeric(t, v) Then
                If gtMode Then If v > uv Then addit = True
                If ltMode Then If v < uv Then addit = True
            End If
         ElseIf txtFilter = "bold" Then
            If li.Bold = True Then addit = True
         ElseIf txtFilter = "selected" Then
            If li.selected = True Then addit = True
         ElseIf isColor Then
            If li.ForeColor = lColor Then addit = True
         Else
         
            addit = useSubtractiveFilter
            If InStr(txtFilter, ",") Then
               tmp = Split(sMatch, ",")
            Else
               push tmp, sMatch
            End If
            
            For Each X In tmp
                If Len(X) > 0 Then
                    If InStr(1, t, X, vbTextCompare) > 0 Then
                        addit = Not addit
                        Exit For
                    End If
                End If
            Next
            
         End If
         
         If addit Then
             CloneListItemTo li, lvFilter
         End If
      
     Next

     
Exit Sub

hideExit:
            lvFilter.Visible = False
            Exit Sub
            
    
End Sub

Sub CloneListItemTo(li As ListItem, lv As ListView)
    Dim li2 As ListItem, i As Integer
    
    Set li2 = lv.ListItems.Add(, , li.Text)
    
    For i = 1 To lv.ColumnHeaders.Count - 1
        li2.subItems(i) = li.subItems(i)
    Next
    
    If li.ForeColor <> vbBlack Then SetLiColor li2, li.ForeColor
    If li.selected Then li2.selected = True
    
    On Error Resume Next
    If IsObject(li.Tag) Then
        Set li2.Tag = li.Tag
    Else
        li2.Tag = li.Tag
    End If
    
    indexMapping.Add li, "fObj:" & ObjPtr(li2)  'filter list item obj to lvFilter objPtr map
    
End Sub

'we had to switch from index mapping to object mapping to account for column click sorts..
Private Function getMainListItemFor(liFilt As ListItem) As ListItem
    On Error Resume Next
    Set getMainListItemFor = indexMapping("fObj:" & ObjPtr(liFilt))
End Function

Private Sub lv_Click()
    If m_Locked Then Exit Sub
    RaiseEvent Click
End Sub

Private Sub lv_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    If m_Locked Then Exit Sub
    Me.ColumnSort ColumnHeader
    'RaiseEvent ColumnClick(ColumnHeader)
End Sub

Private Sub lv_DblClick()
    If m_Locked Then Exit Sub
    RaiseEvent DblClick
End Sub

Private Sub lv_ItemClick(ByVal Item As MSComctlLib.ListItem)
    If m_Locked Then Exit Sub
    If Me.SelCount > 1 Then Exit Sub 'uses sendmessage its ok..
    RaiseEvent ItemClick(Item)
End Sub

Private Sub lv_MouseUp(Button As Integer, Shift As Integer, X As Single, y As Single)
    If m_Locked Then Exit Sub
    RaiseEvent MouseUp(Button, Shift, X, y)
End Sub

Private Sub lvFilter_Click()
    If m_Locked Then Exit Sub
    RaiseEvent Click
End Sub

Private Sub lvFilter_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    If m_Locked Then Exit Sub
    Me.ColumnSort ColumnHeader
    'RaiseEvent ColumnClick(ColumnHeader)
End Sub

Private Sub lvFilter_DblClick()
    If m_Locked Then Exit Sub
    RaiseEvent DblClick
End Sub

Private Sub lvFilter_ItemClick(ByVal Item As MSComctlLib.ListItem)
    If m_Locked Then Exit Sub
    If Me.SelCount > 1 Then Exit Sub 'uses sendmessage its ok..we dont want client to do processing on select all
    RaiseEvent ItemClick(Item)
End Sub

Private Sub lvFilter_MouseUp(Button As Integer, Shift As Integer, X As Single, y As Single)
    If m_Locked Then Exit Sub
    RaiseEvent MouseUp(Button, Shift, X, y)
End Sub

Private Sub txtFilter_KeyPress(KeyAscii As Integer)
    'MsgBox KeyAscii
    
    On Error Resume Next
    Dim t As String, b As Boolean, tot As Long, i As Long
    
    If m_Locked Then Exit Sub
    
    If KeyAscii = vbKeyEscape Then
        KeyAscii = 0
        Filter = Empty
    End If
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        If Len(txtFilter) > 0 Then
            If Left(txtFilter, 1) = "/" Then
                t = Replace(txtFilter, "/", Empty)
                If IsNumeric(t) Then
                    FilterColumn = CLng(t)
                    Filter = Empty
                End If
'                ElseIf Left(t, 3) = "tot" Then
'                    t = Trim(Mid(t, 4))
'                    If IsNumeric(t) Then
'                        tot = totalColumn(CLng(t), b)
'                        MsgBox "Total for " & lv.ColumnHeaders(CLng(t) + 1).Text & " = " & tot & IIf(b, vbCrLf & vbCrLf & " An error was generated", ""), vbInformation
'                    Else
'                        i = ColIndexForName(t)
'                        If i >= 0 Then
'                            tot = totalColumn(i, b)
'                            MsgBox "Total for " & lv.ColumnHeaders(i + 1).Text & " = " & tot & IIf(b, vbCrLf & vbCrLf & " An error was generated", ""), vbInformation
'                        End If
'                    End If
'                    Filter = Empty
'                End If
                
            End If
            
        End If
    End If
            
End Sub

Private Function ColIndexForName(n) As Long
    On Error Resume Next
    Dim i As Long
    If Len(n) > 0 Then
        For i = 1 To lv.ColumnHeaders.Count
            If Left(LCase(lv.ColumnHeaders(i).Text), Len(n)) = LCase(n) Then
                ColIndexForName = i - 1
                Exit Function
            End If
        Next
    End If
    ColIndexForName = -1
End Function
    

Private Sub UserControl_Initialize()
    m_FilterColumn = -1
    m_FilterColumnPreset = -1
    mnuAlertColWidths.Visible = isIde()
End Sub

Private Sub UserControl_Resize()
    On Error Resume Next
    With UserControl
        lv.Top = 0
        lv.Left = 0
        lv.Width = .Width
        lv.Height = .Height - txtFilter.Height - 300
        txtFilter.Top = .Height - txtFilter.Height - 150
        txtFilter.Width = .Width - txtFilter.Left - imgX.Width '- lblTools.Width - 100
        imgX.Top = txtFilter.Top
        imgX.Left = txtFilter.Width + txtFilter.Left ' 100
        'lblTools.Left = .Width - lblTools.Width
        Label1.Top = txtFilter.Top + 30
        'lblTools.Top = txtFilter.Top + 30
    End With
    lvFilter.Move lv.Left, lv.Top, lv.Width, lv.Height
    lv.ColumnHeaders(lv.ColumnHeaders.Count).Width = lv.Width - lv.ColumnHeaders(lv.ColumnHeaders.Count).Left - 200
    lvFilter.ColumnHeaders(lvFilter.ColumnHeaders.Count).Width = lv.ColumnHeaders(lv.ColumnHeaders.Count).Width
End Sub


Public Sub SetLiColor(li As ListItem, newcolor As Long)
    Dim f As ListSubItem
'    On Error Resume Next
    li.ForeColor = newcolor
    For Each f In li.ListSubItems
        f.ForeColor = newcolor
    Next
End Sub

Public Sub ColumnSort(Column As ColumnHeader)
    Dim ListViewControl As ListView
    On Error Resume Next
    
    Set ListViewControl = lv
    If lvFilter.Visible Then Set ListViewControl = lvFilter
        
    With ListViewControl
       If .SortKey <> Column.Index - 1 Then
             .SortKey = Column.Index - 1
             .SortOrder = lvwAscending
       Else
             If .SortOrder = lvwAscending Then
              .SortOrder = lvwDescending
             Else
              .SortOrder = lvwAscending
             End If
       End If
       .Sorted = -1
    End With
    
End Sub

Private Function pad(v, Optional L As Long = 8, Optional char As String = " ", Optional padRight As Boolean = True)
    On Error GoTo hell
    Dim X As Long
    X = Len(v)
    If X < L Then
        If padRight Then
             pad = v & String(L - X, char)
        Else
             pad = String(L - X, char) & v
        End If
    Else
hell:
        pad = v
    End If
End Function

Public Function GetAllElements(Optional selectedOnly As Boolean = False, Optional incHeader As Boolean = True, Optional autoPad As Boolean = True) As String
    Dim ret() As String, i As Integer, tmp As String
    Dim li As ListItem
    Dim ListViewControl As ListView
    Dim include  As Boolean
    Dim maxLen() As Long, sz As Long
    
    On Error Resume Next
    
    Set ListViewControl = lv
    If lvFilter.Visible Then Set ListViewControl = lvFilter
    
    'calculate min field width (including col header text)
    '------------------------------------------------------
    ReDim maxLen(ListViewControl.ColumnHeaders.Count) 'default 0 pad ok with that..
         
    If autoPad Then
    
         If incHeader Then
            For i = 1 To ListViewControl.ColumnHeaders.Count
                sz = Len(ListViewControl.ColumnHeaders(i).Text)
                If sz > maxLen(i - 1) Then maxLen(i - 1) = sz
            Next
         End If
        
         For Each li In ListViewControl.ListItems
             If selectedOnly Then
                If Not li.selected Then GoTo nextSize
             End If
             sz = Len(li.Text)
             If sz > maxLen(0) Then maxLen(0) = sz
             For i = 1 To ListViewControl.ColumnHeaders.Count - 1
                  sz = Len(li.subItems(i))
                  If sz > maxLen(i) Then maxLen(i) = sz
             Next
nextSize:
         Next
         
    End If
    '------------------------------------------------------
    
    If incHeader Then
        For i = 1 To ListViewControl.ColumnHeaders.Count
            tmp = tmp & pad(ListViewControl.ColumnHeaders(i).Text, maxLen(i - 1) + 2) & vbTab
        Next
        
        push ret, tmp
        push ret, String(50, "-")
    End If
 
    For Each li In ListViewControl.ListItems
    
        If selectedOnly Then
            If Not li.selected Then GoTo nextone
        End If
            
        tmp = pad(li.Text, maxLen(0) + 2) & vbTab
        For i = 1 To ListViewControl.ColumnHeaders.Count - 1
            tmp = tmp & pad(li.subItems(i), maxLen(i) + 2) & vbTab
        Next
        push ret, tmp
        
nextone:
    Next

    GetAllElements = Join(ret, vbCrLf)

End Function

Function GetAllText(Optional subItemRow As Long = 0, Optional selectedOnly As Boolean = False) As String
    Dim i As Long
    Dim tmp() As String, X As String
    Dim ListViewControl As ListView
    
    On Error Resume Next
    
    Set ListViewControl = lv
    If lvFilter.Visible Then Set ListViewControl = lvFilter
    
    For i = 1 To ListViewControl.ListItems.Count
        If subItemRow = 0 Then
            X = ListViewControl.ListItems(i).Text
            If selectedOnly And Not ListViewControl.ListItems(i).selected Then X = Empty
            If Len(X) > 0 Then
                push tmp, X
            End If
        Else
            X = ListViewControl.ListItems(i).subItems(subItemRow)
            If selectedOnly And Not ListViewControl.ListItems(i).selected Then X = Empty
            If Len(X) > 0 Then
                push tmp, X
            End If
        End If
    Next
    
    GetAllText = Join(tmp, vbCrLf)
End Function

Private Sub push(ary, value) 'this modifies parent ary object
    On Error GoTo init
    Dim X As Integer
    X = UBound(ary) '<-throws Error If Not initalized
    ReDim Preserve ary(UBound(ary) + 1)
    ary(UBound(ary)) = value
    Exit Sub
init:     ReDim ary(0): ary(0) = value
End Sub

Private Function isIde() As Boolean
' Brad Martinez  http://www.mvps.org/ccrp
    On Error GoTo out
    Debug.Print 1 / 0
out: isIde = Err
End Function

Private Sub UserControl_Terminate()
    m_Locked = False
    Me.Clear
End Sub

Private Function KeyExistsInCollection(c As Collection, val As String) As Boolean
    On Error GoTo nope
    Dim t
    t = c(val)
    KeyExistsInCollection = True
 Exit Function
nope: KeyExistsInCollection = False
End Function

