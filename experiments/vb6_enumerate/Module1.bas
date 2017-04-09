Attribute VB_Name = "Module1"
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSource As Any, ByVal ByteLen As Long)

Public Function keyForIndex(index, c) As String
    ' Get a key based on its index value.  Must be in range, or error.
    Dim i     As Long
    Dim Ptr   As Long
    Dim sKey  As String
    
    index = CLng(index)
    
    If index < 1 Or index > c.Count Then
        Err.Raise 9
        Exit Function
    End If
    '
    If index <= c.Count / 2 Then                                ' Start from front.
        CopyMemory Ptr, ByVal ObjPtr(c) + &H18, 4               ' First item pointer of collection header.
        For i = 2 To index
            CopyMemory Ptr, ByVal Ptr + &H18, 4                 ' Next item pointer of collection item.
        Next i
    Else                                                        ' Start from end and go back.
        CopyMemory Ptr, ByVal ObjPtr(c) + &H1C, 4               ' Last item pointer of collection header.
        For i = c.Count - 1 To index Step -1
            CopyMemory Ptr, ByVal Ptr + &H14, 4                 ' Previous item pointer of collection item.
        Next i
    End If
    '
    i = StrPtr(sKey)                                            ' Save string pointer because we're going to borrow the string.
    CopyMemory ByVal VarPtr(sKey), ByVal Ptr + &H10, 4          ' Key string of collection item.
    keyForIndex = sKey 'Base16Decode(sKey)                                ' Move key into property's return.
    CopyMemory ByVal VarPtr(sKey), i, 4                         ' Put string pointer back to keep memory straight.
End Function

Function clearVal(ByRef o As Variant)
    On Error Resume Next 'one or the other will work we dont care which..
    Set o = Nothing
    o = Empty
End Function

Function AryIsEmpty(ary) As Boolean
  On Error GoTo oops
    i = UBound(ary)  '<- throws error if not initalized
    AryIsEmpty = False
  Exit Function
oops: AryIsEmpty = True
End Function

