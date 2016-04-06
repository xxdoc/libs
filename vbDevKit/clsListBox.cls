VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsListBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Info:    This class is a wrapper to access listbox functions
'         making your code smaller and easier to read.
'
'         it also encapsulates a bunch of new features like
'         being able to filter lists, save to file, load to array
'         load from array etc.
'
'License: You are free to use this library in both private and commercial
'         products at no cost so long as this header stays intact.
'
'Author:  dzzie@yahoo.com
'
'Site:    http://sandsprite.com
'---------------------------------------------------------------

Private Declare Function SendMessage Lib "User32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Enum LB_MSGS
    'general
    LB_ERR = -1
    LB_SETTOPINDEX = &H197  'insure that a particular item is visible
    LB_GETCOUNT = &H18B     'retrieve the number of items in a list box
    LB_GETSEL = &H187       'retrieve the selection state of an item
    LB_ADDSTRING = &H180    'add a string and re-sort a sorted listbox
    LB_INSERTSTRING = &H181 'add a string but do not re-sort a sorted listbox
    LB_DELETESTRING = &H182 'delete a single string
    LB_FINDSTRING = &H18F   'find the first string containing the string specified
    LB_FINDSTRINGEXACT = &H1A2 'find the first string that matches the string specified
    LB_SETTABSTOPS = &H192  'set the tab-stop positions
    LB_RESETCONTENT = &H184 'remove all items
    LB_SETHORIZONTALEXTENT = &H194
    
    'Multi-Select Listboxes
    LB_SETSEL = &H185       'select a string
    LB_GETSELCOUNT = &H190  'retrieve the total number of selected items
    LB_GETSELITEMS = &H191

    'Single Selection Listboxes
    LB_SELECTSTRING = &H18C 'find the first string containing the string specified
    LB_GETCURSEL = &H188    'retrieve the index of the currently selected item
    LB_SETCURSEL = &H186    'select a string and scroll it into view
End Enum

Private lst As ListBox
Attribute lst.VB_VarHelpID = -1

Public noSelectionWarningMessage As String

Property Let ListObject(l As Object)
    Set lst = l
End Property

Property Get Count() As Long
  Count = lst.ListCount
End Property
 
Property Get value(index As Long)
    value = lst.List(index)
End Property
 
Property Get SelectedText() As String
    SelectedText = lst.List(lst.ListIndex)
End Property

Property Let SelectedText(newTxt As String)
    lst.List(lst.ListIndex) = newTxt
End Property

Property Get SelectedIndex() As Long
    SelectedIndex = lst.ListIndex
End Property

Property Get HasSelection() As Boolean
    If lst.ListIndex >= 0 Then HasSelection = True
End Property

Function EnsureSelection() As Boolean
    If Not HasSelection Then
        MsgBox noSelectionWarningMessage, vbInformation
    Else
        EnsureSelection = True
    End If
End Function

Sub Remove(index)
    lst.RemoveItem index
End Sub

Sub Clear()
    lst.Clear
End Sub

Sub AddItem(it)
    lst.AddItem it
End Sub

Sub UpdateValue(newVal, index)
    lst.List(index) = newVal
End Sub

Sub FilterList(filt, Optional Likeit As Boolean = True)
    If filt = "*" Or Trim(filt) = Empty Then Exit Sub
    Dim tmp()
    tmp() = GetListToArray
    tmp() = filterArray(tmp(), filt, Likeit)
    LoadArray tmp()
End Sub

Function GetListToArray() As Variant()
    Dim tmp(), i As Long
    For i = 0 To lst.ListCount - 1
        push tmp, lst.List(i)
    Next
    GetListToArray = tmp()
End Function

Function GetListContents(Optional JoinWith = vbCrLf) As String
    Dim tmp As String, i As Long
    For i = 0 To lst.ListCount - 1
        tmp = tmp & lst.List(i) & JoinWith
    Next
    GetListContents = tmp
End Function

Sub LoadFile(fpath As String, Optional delimiter = vbCrLf, Optional AppendIt As Boolean = False)
    Dim tmp
    If FileExists(fpath) Then
        tmp = Split(ReadFile(fpath), delimiter)
        LoadArray tmp, AppendIt
    End If
End Sub

Sub LoadArray(ary, Optional AppendIt As Boolean = False)
    If Not AppendIt Then lst.Clear
    If AryIsEmpty(ary) Then: lst.AddItem "[Empty Set]": Exit Sub
    Dim i As Long
    For i = LBound(ary) To UBound(ary)
        lst.AddItem ary(i)
    Next
End Sub

Sub LoadDelimitedString(dStr, delimiter, Optional AppendIt As Boolean = False)
    Dim tmp
    tmp = Split(dStr, delimiter)
    LoadArray tmp, AppendIt
End Sub

Function filterArray(ary, filtStr, Optional Likeit As Boolean = True) As Variant()
    If AryIsEmpty(ary) Then Exit Function
    
    Dim tmp(), i As Long
    filtStr = filtStr
    'if you use lcase() on somthing not expliticly defined string
    'it returns nothign! wildcard expression always second
    For i = LBound(ary) To UBound(ary)
        If Likeit Then
            If ary(i) Like filtStr Then push tmp, ary(i)
        Else
            If Not ary(i) Like filtStr Then push tmp, ary(i)
        End If
    Next
    
    filterArray = tmp()
End Function

Private Sub Class_Initialize()
    'default message, you can change it at anytime
    noSelectionWarningMessage = "You must first select an item"
End Sub

Sub MatchSize(it As Object)
    lst.width = it.width
    lst.Height = it.Height
End Sub


Private Function FileExists(path) As Boolean
  If Len(path) = 0 Then Exit Function
  If Dir(path, vbHidden Or vbNormal Or vbReadOnly Or vbSystem) <> "" Then FileExists = True
End Function

Private Function ReadFile(filename)
  Dim f As Long, temp
  f = FreeFile
  temp = ""
   Open filename For Binary As #f        ' Open file.(can be text or image)
     temp = Input(FileLen(filename), #f) ' Get entire Files data
   Close #f
   ReadFile = temp
End Function

Private Sub WriteFile(path, it)
    Dim f As Long
    f = FreeFile
    Open path For Output As #f
    Print #f, it
    Close f
End Sub

Private Sub AppendFile(path, it)
    Dim f As Long
    f = FreeFile
    Open path For Append As #f
    Print #f, it
    Close f
End Sub

Private Sub push(ary, value) 'this modifies parent ary object
    On Error GoTo init
    Dim x As Long
    x = UBound(ary) '<-throws Error If Not initalized
    ReDim Preserve ary(UBound(ary) + 1)
    ary(UBound(ary)) = value
    Exit Sub
init: ReDim ary(0): ary(0) = value
End Sub

Private Function AryIsEmpty(ary) As Boolean
  On Error GoTo oops
    Dim x As Long
    x = UBound(ary)
    AryIsEmpty = False
  Exit Function
oops: AryIsEmpty = True
End Function

'API Functions below
'------------------------------------------------------------------
Private Function SendMsg(msg As LB_MSGS, Optional wParam As Long = 0, Optional lParam As Long = 0) As Long
   SendMsg = SendMessage(lst.hWnd, msg, wParam, lParam)
End Function

Sub SetHorzontalScrollWidth(Optional width As Long = 600)
    If width > 600 Then width = 600
    SendMsg LB_SETHORIZONTALEXTENT, width, 0
End Sub

Sub ScrollToTop()
    SendMsg LB_SETTOPINDEX, 0, 0&
End Sub

Sub DeselectAll()
    SendMsg LB_SETCURSEL, -1, 0&
End Sub

Public Function SelectedItems() As String()
    Dim ItemIndexes() As Long, x As Integer, iNumItems As Integer
    Dim ret() As String
    
    If Not CBool(lst.MultiSelect) Then
        MsgBox "Listbox must be set to MultiSelect for this function.", vbInformation
        Exit Function
    End If
        
    iNumItems = lst.SelCount
    
    If iNumItems Then
        ReDim ItemIndexes(iNumItems - 1)
        SendMessage lst.hWnd, LB_GETSELITEMS, iNumItems, ItemIndexes(0)
    End If
    
    For x = 0 To iNumItems - 1
        push ret(), lst.List(ItemIndexes(x))
    Next x
    
    SelectedItems = ret()

End Function

Function isUnique(strText As String) As Boolean
    If FindEntryIndex(strText, True) >= 0 Then isUnique = False
End Function

Function FindEntryIndex(sSearch As String, Optional exactMatch As Boolean = False, Optional startAt As Long = -1) As Long
    ' Wparam indicates the index to start searching from. -1 = start at top.
    ' lParam is address of the string to search for.
    If exactMatch Then
        FindEntryIndex = SendMsg(LB_FINDSTRINGEXACT, startAt, ByVal sSearch)
    Else
        FindEntryIndex = SendMsg(LB_FINDSTRING, startAt, ByVal sSearch)
    End If
End Function

Function AddUniqueString(strAdd, Optional lIndex = -1) As Boolean
    If isUnique(strAdd) Then
        SendMsg LB_INSERTSTRING, lIndex, ByVal strAdd
        AddUniqueString = True
    End If
End Function

Sub InsertStringAtIndex(strAdd, index)
    SendMsg LB_INSERTSTRING, index, ByVal strAdd
End Sub

