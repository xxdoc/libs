VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CollectionEx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Wrapper for VB6 collection object with some extended capabilities
' add now allows keys to be numeric only and is case sensitive
' if you use a numeric only key, make sure to use forceAsKey =1 when accessing by key
' the CCollectionExtender class allows you to do some more handy stuff not traditionally possible
'  author: elroy
'  link:   http://www.vbforums.com/showthread.php?834579-Wrapper-for-VB6-Collections

Option Explicit

Dim c As New Collection
Dim extras As CCollectionExtender

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSource As Any, ByVal ByteLen As Long)

Private Type dump
    Key As String
    data As Variant
End Type

'initilize the class on demand (memory saver)
Property Get ext() As CCollectionExtender
    If extras Is Nothing Then
        Set extras = New CCollectionExtender
        extras.setTarget c
    End If
    Set ext = extras
End Property


Public Sub Add(vData As Variant, Optional ByVal Key As String, Optional Before As Variant, Optional After As Variant)
    ' If not specified, Before and/or After come is as IsError() = True, which the collection.Add method can handle.
    ' It still raises an error if both are specified, and that behavior is not changed.
    ' Also errors when trying to add duplicate keys, as expected.
    If Len(Key) Then
        c.Add vData, Base16Encode(Key), Before, After
    Else
        c.Add vData, , Before, After
    End If
End Sub

Property Let Item(ByVal IndexOrKey As Variant, Optional ByVal forceAsKey As Byte = 0, val As Variant)

'this works as default as well since set in get
    Dim i As Integer
    Dim k As String
    Dim tmp As Variant
    
    i = VarType(IndexOrKey)
    If forceAsKey <> 0 Then i = 8
    
    Select Case i
        Case 2 To 7 ' Integer, Long, Single, Double, Currency, & Date.  (A bit weird to pass in a date, but who knows.)
            
            k = Me.keyForIndex(CLng(IndexOrKey))
            c.Remove CLng(IndexOrKey)
            c.Add val, k
            
        Case 8
            k = Base16Encode(CStr(IndexOrKey))
            Me.Remove IndexOrKey, forceAsKey
            Add val, IndexOrKey
    
    End Select
    

End Property

Property get  Item(ByVal IndexOrKey As Variant, Optional ByVal forceAsKey As Byte = 0) As Variant
Attribute Item.VB_UserMemId = 0
    'Attribute Item.VB_UserMemId = 0 ' This must be placed in with Notepad.  Makes this method the default.
    ' Had to leave this one public so the "Default" method could be used.
    Dim i As Integer
    Dim k As String
    
    i = VarType(IndexOrKey)
    If forceAsKey <> 0 Then i = 8
    
    Select Case i
        Case 2 To 7 ' Integer, Long, Single, Double, Currency, & Date.  (A bit weird to pass in a date, but who knows.)
       
            If IsObject(c.Item(IndexOrKey)) Then
                Set Item = c.Item(IndexOrKey)
            Else
                Item = c.Item(IndexOrKey)
            End If
    
        Case 8
            k = Base16Encode(CStr(IndexOrKey))
            
            If IsObject(c.Item(k)) Then
                Set Item = c.Item(k)
            Else
                Item = c.Item(k)
            End If
    
    End Select
    
End property

'property so we can mouse over it in ide to see value..
Property Get Count()
    Count = c.Count
End Property

Public Sub Remove(ByVal IndexOrKey As Variant, Optional ByVal forceAsKey As Byte = 0)
    Dim i As Integer
    
    i = VarType(IndexOrKey)
    If forceAsKey <> 0 Then i = 8
    
    Select Case i
        Case 2 To 7 ' Integer, Long, Single, Double, Currency, & Date.  (A bit weird to pass in a date, but who knows.)
            c.Remove CLng(IndexOrKey)
        Case 8
            c.Remove Base16Encode(CStr(IndexOrKey))
    End Select
    
    ' It does nothing if bad IndexOrKey type (such as Null, an object or other nonsense).
End Sub

Property Get isEmpty() As Boolean
    isEmpty = (c.Count = 0)
End Property

Public Sub Clear()
    Set c = New Collection
End Sub


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

Public Function KeyExists(Key As String) As Boolean
    On Error GoTo DoesntExist
    IsObject c(Base16Encode(Key))
    KeyExists = True
DoesntExist:
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

Private Function FileExists(path As String) As Boolean
  On Error GoTo hell
    
  If Len(path) = 0 Then Exit Function
  If Right(path, 1) = "\" Then Exit Function
  If Dir(path, vbHidden Or vbNormal Or vbReadOnly Or vbSystem) <> "" Then FileExists = True
  
  Exit Function
hell: FileExists = False
End Function


Function Sort(Optional descending As Boolean = True) As Boolean
    
    Dim tmp() As Variant
    Dim keys() As String
    
    Dim v, i As Long
    Dim o As Object
    
    If c.Count = 0 Then Exit Function
    
    ReDim tmp(c.Count - 1)
    ReDim keys(c.Count - 1)
    
    For i = 1 To c.Count
        If IsObject(c(i)) Then
            Exit Function
        Else
            keys(i - 1) = Me.keyForIndex(i)
            tmp(i - 1) = c(i)
        End If
    Next
    
    If Not SortArray(descending, tmp, keys) Then Exit Function
    
    Set c = New Collection
    
    For i = 0 To UBound(tmp)
        Add tmp(i), keys(i)
    Next
    
    Sort = True

End Function

Private Function SortArray(descending As Boolean, ByRef TheArray As Variant, ByRef keys As Variant) As Boolean
    Dim sorted As Boolean
    Dim temp, temp_key, x As Long
    Dim doit As Boolean
    
    On Error GoTo hell
    
    sorted = False
    Do While Not sorted
        sorted = True
        For x = 0 To UBound(TheArray) - 1
        
            doit = False
            
            If descending Then
                If TheArray(x) < TheArray(x + 1) Then doit = True
            Else
                If TheArray(x) > TheArray(x + 1) Then doit = True
            End If
            
            If doit Then
                temp = TheArray(x + 1)
                temp_key = keys(x + 1)
                TheArray(x + 1) = TheArray(x)
                keys(x + 1) = keys(x)
                TheArray(x) = temp
                keys(x) = temp_key
                sorted = False
            End If
            
        Next x
    Loop
    
    SortArray = True
hell:
    
End Function
