VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "OrgCollectionEx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' A class wrapper for the VB6 Collection.
' It has the advantages of still using string keys, but they're handled in a completely binary fashion.
' In other words, they're case sensitive, and not restricted to the valid VB6 Unicode range of characters.
'
' There are also a few extra methods and properties:
'
'       KeyExists           Just a boolean check if a key exists in the collection.
'       Keys()              Returns a string array with all the Collection's keys.
'       ItemKey             Based on an Index value, returns or sets the Key of an existing item.
'       ItemIndex           Based on a Key value, returns or sets the Index of an existing item.
'       ChangeKey       Change old key to new key.
'       ChangeIndex     Change old index to new index.
'
Option Explicit
'
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSource As Any, ByVal ByteLen As Long)
'
Dim c As Collection
'

'
'*************************************************************
'*************************************************************
'
' The four methods and property of a typical VB6 Collection.
'
'*************************************************************
'*************************************************************
'
Friend Sub Add(vData As Variant, Optional Key As String, Optional Before As Variant, Optional After As Variant)
    ' If not specified, Before and/or After come is as IsError() = True, which the collection.Add method can handle.
    ' It still raises an error if both are specified, and that behavior is not changed.
    ' Also errors when trying to add duplicate keys, as expected.
    If Len(Key) Then
        c.Add vData, Base16Encode(Key), Before, After
    Else
        c.Add vData, , Before, After
    End If
End Sub

Public Function Item(ByVal IndexOrKey As Variant, Optional bForceInterpretAsKey As Boolean = False) As Variant
Attribute Item.VB_UserMemId = 0
    'Attribute Item.VB_UserMemId = 0 ' This must be placed in with Notepad.  Makes this method the default.
    ' Had to leave this one public so the "Default" method could be used.
    Select Case VarType(IndexOrKey)
    Case 2 To 7 ' Integer, Long, Single, Double, Currency, & Date.  (A bit weird to pass in a date, but who knows.)
        If bForceInterpretAsKey Then
            LetSetItem c.Item(Base16Encode(CStr(IndexOrKey))), Item
        Else
            LetSetItem c.Item(CLng(IndexOrKey)), Item
        End If
    Case 8
        LetSetItem c.Item(Base16Encode(CStr(IndexOrKey))), Item
    End Select
    ' Returns Item = Empty if not set, because of bad IndexOrKey type (such as Null, an object or other nonsense).
End Function

Friend Function Count()
    Count = c.Count
End Function

Friend Sub Remove(ByVal IndexOrKey As Variant, Optional bForceInterpretAsKey As Boolean = False)
    Select Case VarType(IndexOrKey)
    Case 2 To 7 ' Integer, Long, Single, Double, Currency, & Date.  (A bit weird to pass in a date, but who knows.)
        If bForceInterpretAsKey Then
            c.Remove Base16Encode(CStr(IndexOrKey))
        Else
            c.Remove CLng(IndexOrKey)
        End If
    Case 8
        c.Remove Base16Encode(CStr(IndexOrKey))
    End Select
    ' It does nothing if bad IndexOrKey type (such as Null, an object or other nonsense).
End Sub

'
'*************************************************************
'*************************************************************
'
' Some extra handy methods.
'
'*************************************************************
'*************************************************************
'
Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    'Attribute NewEnum.VB_UserMemId = -4 ' This must be placed in with Notepad.
    'Attribute NewEnum.VB_MemberFlags = "40" ' This is also a good idea, as it hides this method.
    '
    ' This allows use of the "For Each" syntax.
    ' Just enumerate with a variant for the item, using this class as the series.
    '
   Set NewEnum = c.[_NewEnum]
End Function

Friend Function KeyExists(Key As String) As Boolean
    On Error GoTo DoesntExist
    c.Item Base16Encode(Key)
    KeyExists = True
DoesntExist:
End Function

Friend Function Keys() As String()
    ' Returns a string array of all the keys.
    Dim sKeys() As String
    Dim j As Long
    Dim iHold As Long
    Dim ptr As Long
    Dim sKeyTemp  As String
    '
    If c.Count = 0 Then Exit Function
    '
    ReDim sKeys(1 To c.Count)
    j = 1
    CopyMemory ptr, ByVal ObjPtr(c) + &H18, 4                   ' First item pointer of collection header.
    GoSub MoveKeyToArray
    For j = 2 To c.Count
        CopyMemory ptr, ByVal ptr + &H18, 4                     ' Next item pointer of collection item.
        GoSub MoveKeyToArray
    Next j
    Keys = sKeys
    Exit Function
    '
MoveKeyToArray: ' j and ptr must be set to call this.
    iHold = StrPtr(sKeyTemp)                                    ' Save string pointer because we're going to borrow the string.
    CopyMemory ByVal VarPtr(sKeyTemp), ByVal ptr + &H10, 4      ' Key string of collection item.
    sKeys(j) = Base16Decode(sKeyTemp)                           ' Move key into array.
    CopyMemory ByVal VarPtr(sKeyTemp), iHold, 4                 ' Put string pointer back to keep memory straight.
    Return
End Function

Friend Property Let ChangeKey(ByVal OldKey As String, ByVal NewKey As String)
    ' Usage: c.ChangeKey(OldKey) = NewKey
    '
    ' OldKey must exist or error.
    Dim ptr   As Long
    Dim tKey  As String
    Dim iHold  As Long
    Dim Index As Long
    Dim vData As Variant
    '
    If c.Count Then
        OldKey = Base16Encode(OldKey)
        iHold = StrPtr(tKey)                                    ' Tuck away for later restore.  We're going to borrow this string.
        CopyMemory ptr, ByVal ObjPtr(c) + &H18, 4               ' First item pointer of collection header.
        Index = 1
        Do
            CopyMemory ByVal VarPtr(tKey), ByVal ptr + &H10, 4  ' Key string of collection item.
            If OldKey = tKey Then Exit Do
            Index = Index + 1
            CopyMemory ptr, ByVal ptr + &H18, 4                 ' Next item pointer of collection item.
        Loop Until ptr = 0
        CopyMemory ByVal VarPtr(tKey), iHold, 4                 ' Put string pointer back to keep memory straight.
    End If
    If ptr = 0 Then                                             ' We didn't find the old key.
        Err.Raise 5
        Exit Property
    End If
    '
    ' We've found the old key if we got to here.
    LetSetItem c.Item(Index), vData
    c.Remove Index
    '
    Select Case True
    Case c.Count = 0: c.Add vData, Base16Encode(NewKey)                 ' Add first item.
    Case Index > c.Count: c.Add vData, Base16Encode(NewKey), , c.Count  ' Add to end.
    Case Else: c.Add vData, Base16Encode(NewKey), Index                 ' Add prior to one specified, so it takes on that index number.
    End Select
End Property

Friend Property Let ChangeIndex(OldIndex As Long, NewIndex As Long)
    ' Usage: c.ChangeIndex(OldIndex) = NewIndex
    '
    ' Item with key must exist or error.
    ' Indexes must be in current range of the Collection, or error.
    '
    Dim vData As Variant
    Dim sKey As String
    Dim tKey As String
    Dim ptr As Long
    Dim i As Long
    '
    If OldIndex < 1 Or OldIndex > c.Count Or NewIndex < 1 Or NewIndex > c.Count Then
        Err.Raise 9
        Exit Property
    End If
    '
    LetSetItem c.Item(OldIndex), vData
    If c.Count = 1 Then Exit Property ' Nothing to do.  We wait until here to check to make sure the Key exists.
    '
    If OldIndex <= c.Count / 2 Then                             ' Start from front.
        CopyMemory ptr, ByVal ObjPtr(c) + &H18, 4               ' First item pointer of collection header.
        For i = 2 To OldIndex
            CopyMemory ptr, ByVal ptr + &H18, 4                 ' Next item pointer of collection item.
        Next i
    Else                                                        ' Start from end and go back.
        CopyMemory ptr, ByVal ObjPtr(c) + &H1C, 4               ' Last item pointer of collection header.
        For i = c.Count - 1 To OldIndex Step -1
            CopyMemory ptr, ByVal ptr + &H14, 4                 ' Previous item pointer of collection item.
        Next i
    End If
    '
    i = StrPtr(tKey)                                            ' Save string pointer because we're going to borrow the string.
    CopyMemory ByVal VarPtr(tKey), ByVal ptr + &H10, 4          ' Key string of collection item.
    sKey = tKey                                                 ' Move key into better spot.
    CopyMemory ByVal VarPtr(tKey), i, 4                         ' Put string pointer back to keep memory straight.
    '
    ' Now that we've got the data and key, we can delete and re-add.
    ' Key is still encoded.  No need to decode.
    c.Remove OldIndex
    '
    Select Case True
    Case c.Count = 0: c.Add vData, sKey                         ' Add first item.
    Case NewIndex > c.Count: c.Add vData, sKey, , c.Count       ' Add to end.
    Case Else: c.Add vData, sKey, NewIndex                      ' Add prior to one specified, so it takes on that index number.
    End Select
End Property

Friend Property Let ItemKey(Index As Long, ByVal Key As String)
    ' Usage: ItemKey(Index) = Key
    '
    ' Change an item key based on its index value.
    ' Index must be in current range of the Collection, or error.
    '
    Dim vData As Variant
    '
    If Index < 1 Or Index > c.Count Then
        Err.Raise 9
        Exit Property
    End If
    '
    LetSetItem c.Item(Index), vData
    c.Remove Index                                      ' No need to save old key, as it's being thrown out.
    '
    Select Case True
    Case c.Count = 0: c.Add vData, Base16Encode(Key)                    ' Add first item.
    Case Index > c.Count: c.Add vData, Base16Encode(Key), , c.Count     ' Add to end.
    Case Else: c.Add vData, Base16Encode(Key), Index                    ' Add prior to one specified, so it takes on that index number.
    End Select
End Property

Friend Property Get ItemKey(Index As Long) As String
    ' Get a key based on its index value.  Must be in range, or error.
    Dim i     As Long
    Dim ptr   As Long
    Dim sKey  As String
    '
    If Index < 1 Or Index > c.Count Then
        Err.Raise 9
        Exit Property
    End If
    '
    If Index <= c.Count / 2 Then                                ' Start from front.
        CopyMemory ptr, ByVal ObjPtr(c) + &H18, 4               ' First item pointer of collection header.
        For i = 2 To Index
            CopyMemory ptr, ByVal ptr + &H18, 4                 ' Next item pointer of collection item.
        Next i
    Else                                                        ' Start from end and go back.
        CopyMemory ptr, ByVal ObjPtr(c) + &H1C, 4               ' Last item pointer of collection header.
        For i = c.Count - 1 To Index Step -1
            CopyMemory ptr, ByVal ptr + &H14, 4                 ' Previous item pointer of collection item.
        Next i
    End If
    '
    i = StrPtr(sKey)                                            ' Save string pointer because we're going to borrow the string.
    CopyMemory ByVal VarPtr(sKey), ByVal ptr + &H10, 4          ' Key string of collection item.
    ItemKey = Base16Decode(sKey)                                ' Move key into property's return.
    CopyMemory ByVal VarPtr(sKey), i, 4                         ' Put string pointer back to keep memory straight.
End Property

Friend Property Let ItemIndex(ByVal Key As String, Index As Long)
    ' Usage: ItemIndex(Key) = Index
    '
    ' Change an item index based on its key value.
    ' Item with key must exist or error.
    ' Index must be in current range of the Collection, or error.
    '
    Dim vData As Variant
    '
    If Index < 1 Or Index > c.Count Then
        Err.Raise 9
        Exit Property
    End If
    '
    Key = Base16Encode(Key)
    LetSetItem c.Item(Key), vData
    If c.Count = 1 Then Exit Property ' Nothing to do.  We wait until here to check to make sure the Key exists.
    c.Remove Key
    '
    Select Case True
    Case c.Count = 0: c.Add vData, Key                      ' Add first item.
    Case Index > c.Count:   c.Add vData, Key, , c.Count     ' Add to end.
    Case Else: c.Add vData, Key, Index                      ' Add prior to one specified, so it takes on that index number.
    End Select
End Property

Friend Property Get ItemIndex(ByVal Key As String) As Long
    ' Get an index based on its key value.
    Dim ptr   As Long
    Dim tKey  As String
    Dim iHold  As Long
    '
    If c.Count Then
        Key = Base16Encode(Key)
        iHold = StrPtr(tKey)                                    ' Tuck away for later restore.  We're going to borrow this string.
        CopyMemory ptr, ByVal ObjPtr(c) + &H18, 4               ' First item pointer of collection header.
        ItemIndex = 1
        Do
            CopyMemory ByVal VarPtr(tKey), ByVal ptr + &H10, 4  ' Key string of collection item.
            If Key = tKey Then Exit Do
            ItemIndex = ItemIndex + 1
            CopyMemory ptr, ByVal ptr + &H18, 4                 ' Next item pointer of collection item.
        Loop Until ptr = 0
        CopyMemory ByVal VarPtr(tKey), iHold, 4                 ' Put string pointer back to keep memory straight.
    End If
    If ptr = 0 Then ItemIndex = 0
End Property

'
'*************************************************************
'*************************************************************
'
' Private procedures used above.
'
'*************************************************************
'*************************************************************
'
Private Sub Class_Initialize()
    Set c = New Collection
End Sub

Private Function Base16Encode(s As String) As String
    Dim i As Long
    Base16Encode = Space$(Len(s) * 4)
    For i = 0 To Len(s) - 1
        Mid$(Base16Encode, i * 4 + 1, 4) = Right$("0000" & Hex$(AscW(Mid$(s, i + 1, 1))), 4)
    Next i
End Function

Private Function Base16Decode(s As String) As String
    Dim i As Long
    Base16Decode = Space$(Len(s) \ 4)
    For i = 0 To Len(s) - 1 Step 4
        Mid$(Base16Decode, i \ 4 + 1, 1) = ChrW$(val("&h" & Mid$(s, i + 1, 4)))
    Next i
End Function

Private Sub LetSetItem(ItemIn As Variant, ItemOut As Variant)
    ' This is necessary because we don't know if the item is an object or not.
    ' If the item is an object, just using Let will cause an error.
    ' We must also use ItemOut passed ByRef, so we don't have to Let/Set it again from a Function.
    If IsObject(ItemIn) Then
        Set ItemOut = ItemIn
    Else
        ItemOut = ItemIn
    End If
End Sub

