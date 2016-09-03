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
'
' so it turns out when you boil it all down..the only method you really need to do funky copymemory stuff
' is keyforIndex, all of the other methods can be achieved with normal programming plus that one extra primitive

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
    Dim i As Long
    indexForKey = 0
    For i = 1 To c.Count
        If keyForIndex(i) = Key Then
            indexForKey = i
            Exit Function
        End If
    Next
End Function

Public Function Keys() As String()
    Dim tmp() As String, i As Long
    
    If c.Count = 0 Then Exit Function
    
    push tmp(), ""
    For i = 1 To c.Count
        push tmp(), keyForIndex(i)
    Next
    
    Keys = tmp()
    
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
    
    index = indexForKey(OldKey)
   
    If IsObject(c.Item(index)) Then
        Set vData = c.Item(index)
    Else
       vData = c.Item(index)
    End If
    
    c.Remove index
    
    If c.Count = 0 Then
        c.Add vData, Base16Encode(NewKey)             ' Add first item.
    ElseIf index > c.Count Then
        c.Add vData, Base16Encode(NewKey), , c.Count  ' Add to end.
    Else
        c.Add vData, Base16Encode(NewKey), index      ' Add prior to one specified, so it takes on that index number.
    End If
    
    changeKey = True
    Exit Function
hell:
End Function

Public Function changeIndex(OldIndex As Long, newIndex As Long) As Boolean
    Dim vData As Variant
    Dim sKey As String
    Dim i As Long
    
    On Error GoTo hell
    
    If OldIndex = newIndex Then GoTo ok
    If c.Count = 1 Then GoTo ok
    
    If IsObject(c.Item(OldIndex)) Then
        Set vData = c.Item(OldIndex)
    Else
        vData = c.Item(OldIndex)
    End If

    sKey = Base16Encode(keyForIndex(OldIndex))
    c.Remove OldIndex
    
    If c.Count = 0 Then
        c.Add vData, sKey             ' Add first item.
    ElseIf newIndex > c.Count Then
        c.Add vData, sKey, , c.Count  ' Add to end.
    Else
        c.Add vData, sKey, newIndex   ' Add prior to one specified, so it takes on that index number.
    End If
    
ok:
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

Public Function toString(Optional delimiter As String = vbCrLf, Optional includeKeys As Boolean = False) As String
    Dim tmp() As Variant
    Dim k, i As Long
    
    If c.Count = 0 Then GoTo done
    
    ReDim tmp(c.Count - 1)
    
    For i = 1 To c.Count
        k = Empty
        If includeKeys Then k = keyForIndex(i)
        If Len(k) > 0 Then k = k & " = "
        If IsObject(c(i)) Then
            tmp(i - 1) = k & "[ " & TypeName(c(i)) & " object ]"
        Else
            tmp(i - 1) = k & c(i)
        End If
    Next
    
done:
    toString = Join(tmp, delimiter)
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

Private Sub push(ary, value) 'this modifies parent ary object
    On Error GoTo init
    Dim x
    x = UBound(ary) '<-throws Error If Not initalized
    ReDim Preserve ary(UBound(ary) + 1)
    ary(UBound(ary)) = value
    Exit Sub
init:     ReDim ary(0): ary(0) = value
End Sub
