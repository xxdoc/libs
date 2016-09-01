VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CCollectionExtender"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' the CCollectionExtender class allows you to do some more handy stuff not traditionally possible
'  author: elroy
'  link:   http://www.vbforums.com/showthread.php?834579-Wrapper-for-VB6-Collections

'we duplicate a few methods from CollectionEx so this class can be used independantly..
'(i dont want to force an extra bas into the mix)


Option Explicit
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSource As Any, ByVal ByteLen As Long)
Private c As Collection

Private Type dump
    Key As String
    data As Variant
End Type

Public Sub setTarget(ByRef cc As Collection)
    Set c = cc 'obj ref, not copy
End Sub

Public Function keyForIndex(index As Long) As String
    ' Get a key based on its index value.  Must be in range, or error.
    Dim i     As Long
    Dim ptr   As Long
    Dim sKey  As String
    '
    If index < 1 Or index > c.Count Then
        Err.Raise 9
        Exit Function
    End If
    '
    If index <= c.Count / 2 Then                                ' Start from front.
        CopyMemory ptr, ByVal ObjPtr(c) + &H18, 4               ' First item pointer of collection header.
        For i = 2 To index
            CopyMemory ptr, ByVal ptr + &H18, 4                 ' Next item pointer of collection item.
        Next i
    Else                                                        ' Start from end and go back.
        CopyMemory ptr, ByVal ObjPtr(c) + &H1C, 4               ' Last item pointer of collection header.
        For i = c.Count - 1 To index Step -1
            CopyMemory ptr, ByVal ptr + &H14, 4                 ' Previous item pointer of collection item.
        Next i
    End If
    '
    i = StrPtr(sKey)                                            ' Save string pointer because we're going to borrow the string.
    CopyMemory ByVal VarPtr(sKey), ByVal ptr + &H10, 4          ' Key string of collection item.
    keyForIndex = Base16Decode(sKey)                                ' Move key into property's return.
    CopyMemory ByVal VarPtr(sKey), i, 4                         ' Put string pointer back to keep memory straight.
End Function

Public Function indexForKey(ByVal Key As String) As Long
    ' Get an index based on its key value.
    Dim ptr   As Long
    Dim tKey  As String
    Dim iHold  As Long
    '
    If c.Count Then
        Key = Base16Encode(Key)
        iHold = StrPtr(tKey)                                    ' Tuck away for later restore.  We're going to borrow this string.
        CopyMemory ptr, ByVal ObjPtr(c) + &H18, 4               ' First item pointer of collection header.
        indexForKey = 1
        Do
            CopyMemory ByVal VarPtr(tKey), ByVal ptr + &H10, 4  ' Key string of collection item.
            If Key = tKey Then Exit Do
            indexForKey = indexForKey + 1
            CopyMemory ptr, ByVal ptr + &H18, 4                 ' Next item pointer of collection item.
        Loop Until ptr = 0
        CopyMemory ByVal VarPtr(tKey), iHold, 4                 ' Put string pointer back to keep memory straight.
    End If
    If ptr = 0 Then indexForKey = 0
End Function

Public Function Keys() As String()
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

Public Function changeKeyByIndex(ByVal index As Long, ByVal NewKey As String) As Boolean
    Dim curKey As String
    On Error GoTo hell
    curKey = keyForIndex(index)
    changeKeyByIndex = changeKey(curKey, NewKey)
    Exit Function
hell:
End Function

Public Function changeIndexByKey(ByVal Key As String, ByVal newIndex As Long) As Boolean
    Dim curIndex As Long
    On Error GoTo hell
    curIndex = indexForKey(Key)
    changeIndexByKey = changeIndex(curIndex, newIndex)
    Exit Function
hell:
End Function

Public Function changeKey(ByVal OldKey As String, ByVal NewKey As String) As Boolean
    Dim ptr   As Long
    Dim tKey  As String
    Dim iHold  As Long
    Dim index As Long
    Dim vData As Variant
    
    On Error GoTo hell
    
    If c.Count Then
        OldKey = Base16Encode(OldKey)
        iHold = StrPtr(tKey)                                    ' Tuck away for later restore.  We're going to borrow this string.
        CopyMemory ptr, ByVal ObjPtr(c) + &H18, 4               ' First item pointer of collection header.
        index = 1
        Do
            CopyMemory ByVal VarPtr(tKey), ByVal ptr + &H10, 4  ' Key string of collection item.
            If OldKey = tKey Then Exit Do
            index = index + 1
            CopyMemory ptr, ByVal ptr + &H18, 4                 ' Next item pointer of collection item.
        Loop Until ptr = 0
        CopyMemory ByVal VarPtr(tKey), iHold, 4                 ' Put string pointer back to keep memory straight.
    End If
    If ptr = 0 Then                                             ' We didn't find the old key.
        Err.Raise 5
        Exit Function
    End If
    '
    ' We've found the old key if we got to here.
    getValue CVar(index), vData 'vData = c.Item(Index)
    c.Remove index
    '
    Select Case True
    Case c.Count = 0: c.Add vData, Base16Encode(NewKey)                 ' Add first item.
    Case index > c.Count: c.Add vData, Base16Encode(NewKey), , c.Count  ' Add to end.
    Case Else: c.Add vData, Base16Encode(NewKey), index                 ' Add prior to one specified, so it takes on that index number.
    End Select
    
    changeKey = True
    Exit Function
hell:
End Function

Public Function changeIndex(OldIndex As Long, newIndex As Long) As Boolean
    Dim vData As Variant
    Dim sKey As String
    Dim tKey As String
    Dim ptr As Long
    Dim i As Long
    
    On Error GoTo hell
    
    If OldIndex < 1 Or OldIndex > c.Count Or newIndex < 1 Or newIndex > c.Count Then
        Err.Raise 9
        Exit Function
    End If
    '
    getValue CVar(OldIndex), vData 'vData = c.Item(OldIndex)
    If c.Count = 1 Then Exit Function ' Nothing to do.  We wait until here to check to make sure the Key exists.
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
    Case newIndex > c.Count: c.Add vData, sKey, , c.Count       ' Add to end.
    Case Else: c.Add vData, sKey, newIndex                      ' Add prior to one specified, so it takes on that index number.
    End Select
    
    changeIndex = True
    Exit Function
hell:
End Function

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
        Mid$(Base16Decode, i \ 4 + 1, 1) = ChrW$(Val("&h" & Mid$(s, i + 1, 4)))
    Next i
End Function

Private Sub getValue(IndexOrKey As Variant, ValueOut As Variant)
    ' This is necessary because we don't know if the item is an object or not.
    ' If the item is an object, just using Let will cause an error.
    ' We must also use ValueOut passed ByRef, so we don't have to Let/Set it again from a Function.
    '
    ' We may still error if IndexOrKey isn't valid, but this is the way it should be.
    ' the function above this one handles determining if its an index or encoded key
    
    If IsObject(c.Item(IndexOrKey)) Then
        Set ValueOut = c.Item(IndexOrKey)
    Else
        ValueOut = c.Item(IndexOrKey)
    End If
    
End Sub

Public Function toArray() As Variant()
    Dim tmp() As Variant
    Dim v, i As Long
    Dim o As Object
    
    If c.Count = 0 Then GoTo done
    
    ReDim tmp(c.Count - 1)
    
    For i = 1 To c.Count
        If IsObject(c(i)) Then
            Set tmp(i - 1) = c(i)
        Else
            tmp(i - 1) = c(i)
        End If
    Next
    
done:
    toArray = tmp()
End Function

Public Sub fromArray(ary As Variant, Optional append As Boolean = True)
    Dim i As Long
    
    If AryIsEmpty(ary) Then Exit Sub
    If Not append And c.Count > 0 Then Set c = New Collection
    
    For i = LBound(ary) To UBound(ary)
        c.Add ary(i)
    Next

End Sub

Public Function toFile(ByVal fPath As String) As Boolean
        
    On Error GoTo hell
    
    If FileExists(fPath) Then Kill fPath
    If FileExists(fPath) Then GoTo hell
    
    Dim tmp() As dump, d As dump
    Dim v, i As Long, f As Long
    
    If c.Count = 0 Then GoTo hell
    
    ReDim tmp(c.Count - 1)
    
    For i = 1 To c.Count
        tmp(i - 1).Key = keyForIndex(i)
        If IsObject(c(i)) Then
            tmp(i - 1).data = "[ invalid can not save " & TypeName(c(i)) & " object ]"
        Else
            tmp(i - 1).data = c(i)
        End If
    Next
    
    f = FreeFile
    Open fPath For Binary As f
    Put f, , CLng(UBound(tmp))
    Put f, , tmp()
    Close f
    
    toFile = True
    
hell:
End Function

Public Function fromFile(ByVal fPath As String, Optional append As Boolean) As Long
    
    On Error GoTo hell
    
    If Not FileExists(fPath) Then GoTo hell
    
    Dim tmp() As dump, d As dump
    Dim v, i As Long, f As Long, cnt As Long
    
    If c.Count > 0 And append = False Then Set c = New Collection
    
    f = FreeFile
    Open fPath For Binary As f
    Get f, , cnt
    ReDim tmp(cnt)
    Get f, , tmp()
    Close f
    
    For i = 0 To UBound(tmp)
        If Len(tmp(i).Key) = 0 Then
            c.Add tmp(i).data
        Else
            c.Add tmp(i).data, Base16Encode(uniqueKey(tmp(i).Key))
        End If
    Next
        
    fromFile = UBound(tmp)
    Exit Function
hell:
    fromFile = -1
End Function

Public Function KeyExists(Key As String) As Boolean
    On Error GoTo DoesntExist
    IsObject c(Base16Encode(Key))
    KeyExists = True
DoesntExist:
End Function

Public Function uniqueKey(ByVal suggested As String) As String
    Dim i As Long, tmp As String
    
    If Len(suggested) = 0 Then suggested = Base16Encode(Now)
    
    If KeyExists(suggested) Then
        Do
            i = i + 1
            tmp = suggested & "_" & i
        Loop While KeyExists(tmp)
        suggested = tmp
    End If
    
    uniqueKey = suggested
    
End Function

Private Function FileExists(path As String) As Boolean
  On Error GoTo hell
    
  If Len(path) = 0 Then Exit Function
  If Right(path, 1) = "\" Then Exit Function
  If Dir(path, vbHidden Or vbNormal Or vbReadOnly Or vbSystem) <> "" Then FileExists = True
  
  Exit Function
hell: FileExists = False
End Function

Private Function AryIsEmpty(ary) As Boolean
  On Error GoTo oops
  Dim i As Long
    i = UBound(ary)  '<- throws error if not initalized
    AryIsEmpty = False
  Exit Function
oops: AryIsEmpty = True
End Function

