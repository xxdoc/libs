VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "StringEx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'VARIABLES

Private p_Value As String
Private p_ValueLen As Long
Private p_ValueLenReal As Long
Private p_ValuePtr As Long

'WINAPI

Private Declare Function CharLowerBuffW Lib "user32" (ByVal lStr As Long, ByVal lLength As Long) As Long
Private Declare Function CharUpperBuffW Lib "user32" (ByVal lStr As Long, ByVal lLength As Long) As Long
Private Declare Function IsTextUnicode Lib "advapi32" (ByRef uBuffer As Any, ByVal lBufferLen As Long, ByRef lResult As Long) As Long
Private Declare Function MultiByteToWideChar Lib "kernel32" (ByVal lCodePage As Long, ByVal lFlags As Long, uMultiByte As Any, ByVal lMultiByteLen As Long, uWideCharStr As Any, ByVal lWideCharStrLen As Long) As Long
Private Declare Sub RtlMoveMemory Lib "kernel32" (uTarget As Any, uSource As Any, ByVal lLen As Long)

'EVENTS

Private Sub Class_Initialize()

    p_ValueLenReal = 256&

End Sub

'PROPERTIES

Public Property Get Capacity() As Long

    Capacity = p_ValueLenReal

End Property

Public Property Get length() As Long

    length = p_ValueLen

End Property

Public Property Get Pointer() As Long

    Pointer = p_ValuePtr

End Property

Public Property Get Value() As String
Attribute Value.VB_UserMemId = 0
Attribute Value.VB_MemberFlags = "200"

    If p_ValueLen Then Value = VBA.left$(p_Value, p_ValueLen)

End Property

Public Property Let Value(ByRef NewValue As String)
        
    Dim l As Long

    l = Len(NewValue)

    If l Then
    
        BufferRedim l, True
    
        RtlMoveMemory ByVal p_ValuePtr, ByVal StrPtr(NewValue), p_ValueLen + p_ValueLen
    
    Else
        p_ValueLen = 0&
    End If
    
End Property

'ROUTINES

Public Function charCodeAt(ByVal index As Long) As Byte
    If index > 0& And index <= p_ValueLen Then charCodeAt = VBA.Asc(Mid$(p_Value, index, 1&))
End Function

Public Function charAt(index As Long) As String
    If index > 0& And index <= p_ValueLen Then charAt = Mid$(p_Value, index, 1&)
End Function

Public Function split(delimiter As String, Optional compare As VbCompareMethod = vbBinaryCompare) As Variant()
    split = VBA.split(p_Value, delimiter, , compare)
End Function

Public Function clone() As StringEx

    Set clone = New StringEx
    
    If p_ValueLen Then clone.Value = VBA.left$(p_Value, p_ValueLen)

End Function

Public Function concat(ByRef ConcatString As String) As StringEx
    
    Dim l As Long
    
    Set concat = Me
    
    l = Len(ConcatString)
    
    If l Then
    
        BufferRedim l, False
    
        RtlMoveMemory ByVal p_ValuePtr + ((p_ValueLen - l) * 2&), ByVal StrPtr(ConcatString), l + l
    
    End If

End Function

'supports negative offsets
Public Function subString(Optional start As Long = 1, Optional ByVal endAt As Long = 0) As StringEx
    Dim s As New StringEx
    Dim v As String
    Dim leng As Long
    
    Set subString = s
    
    If start > Me.length Then Exit Function
    
    If start > 0 Then
    
        If endAt = 0 Then
            endAt = Me.length - start + 1
        ElseIf endAt < 0 Then
            endAt = Me.length + 1 - endAt
        Else
            If endAt > Me.length - start Then
                endAt = Me.length - start + 1
            End If
        End If
        
        s.Value = Mid(p_Value, start, endAt)
    Else
        
        If endAt > 0 Then Exit Function
        If endAt < start Then Exit Function
        If Abs(start) > Me.length Then Exit Function
        
        v = VBA.right(Me.Value, Abs(start))
        If endAt < 0 Then
            v = VBA.left(v, Len(v) - Abs(endAt))
        End If
        
        s.Value = v
         
    End If
    
            
        
End Function

Public Function subStr(start As Long, Optional ByVal length As Long = -1) As StringEx
    Dim s As New StringEx
    
    Set subStr = s
    
    If start > Me.length Then Exit Function
    
    If length = -1 Or length > Me.length - start Then
        length = Me.length - start + 1
    End If
    
    s.Value = Mid(p_Value, start, length)
    
End Function

Public Function find(ByRef SearchString As String, Optional ByVal start As Long, Optional ByVal Reverse As Boolean, Optional ByVal CompareMethod As VbCompareMethod) As Long

    Dim x As Long
    
    If p_ValueLen > 0& And Len(SearchString) > 0& Then
    
        If Reverse Then
        
            If start > 0& And start < p_ValueLen Then x = start Else x = p_ValueLen
        
            find = InStrRev(p_Value, SearchString, x, CompareMethod)
        
        Else
        
            If start > 0& Then x = start Else x = 1&
        
            x = InStr(x, p_Value, SearchString, CompareMethod)
            
            If x <= p_ValueLen Then find = x
        
        End If
        
    End If

End Function

Public Function insert(ByVal index As Long, ByRef InsertString As String) As StringEx

    Dim l As Long

    Set insert = Me
    
    l = Len(InsertString)
        
    If l > 0& And index >= 0& And index <= p_ValueLen Then
    
        BufferRedim l, False
        
        RtlMoveMemory ByVal p_ValuePtr + (index * 2&) + l + l, ByVal p_ValuePtr + (index * 2&), (p_ValueLen - (index + l)) * 2&
        RtlMoveMemory ByVal p_ValuePtr + (index * 2&), ByVal StrPtr(InsertString), l + l
    
    End If

End Function

Public Function left(ByVal length As Long) As StringEx

    Set left = Me
    
    If p_ValueLen > 0& And length >= 0& And length < p_ValueLen Then p_ValueLen = length

End Function

Public Function lTrim() As StringEx

    Dim x As Long

    Set lTrim = Me
    
    If p_ValueLen Then
    
        If AscW(p_Value) = 32 Then
    
            Do
                x = x + 1&
            Loop While InStr(x + 1&, p_Value, ChrW$(32&), vbBinaryCompare) = (x + 1&)
        
            If x Then
        
                p_ValueLen = p_ValueLen - x
            
                RtlMoveMemory ByVal p_ValuePtr, ByVal p_ValuePtr + x + x, p_ValueLen + p_ValueLen
            
            End If
        
        End If
    
    End If

End Function



Public Function remove(ByVal index As Long, ByVal length As Long) As StringEx

    Dim x As Long

    Set remove = Me

    If p_ValueLen > 0& And index > 0& And index <= p_ValueLen And length > 0& Then
    
        x = length - (length - (p_ValueLen - index + 1&))
        
        If length <= x Then x = length
    
        If x < p_ValueLen Then
    
            RtlMoveMemory ByVal p_ValuePtr + ((index + -1) * 2&), ByVal p_ValuePtr + ((index + -1) * 2&) + (x * 2&), (p_ValueLen - (index + x) + 1&) * 2&
            
            p_ValueLen = p_ValueLen - x
            
        Else
            p_ValueLen = 0&
        End If
    
    End If

End Function

Public Function replace(ByRef SearchString As String, ByRef ReplaceString As String, Optional ByVal start As Long = 1&, Optional ByVal Count As Long = -1, Optional ByVal CompareMethod As VbCompareMethod) As StringEx

    Set replace = Me

    If p_ValueLen Then Value = VBA.replace(VBA.left$(p_Value, p_ValueLen), SearchString, ReplaceString, start, Count, CompareMethod)

End Function

Public Function strip(ByRef SearchString As String, Optional ByVal CompareMethod As VbCompareMethod) As StringEx

    Set strip = Me

    If p_ValueLen Then Value = VBA.replace(VBA.left$(p_Value, p_ValueLen), SearchString, Empty, , , CompareMethod)

End Function

Public Function strip2(ParamArray args()) As StringEx

    Dim tmp As String
    Dim v
    
    Set strip2 = Me
    
    tmp = Me.Value
    For Each v In args
       tmp = VBA.replace(tmp, v, Empty, , , vbTextCompare)
       If Len(tmp) = 0 Then Exit For
    Next
    
    Value = tmp

End Function

Public Function right(ByVal length As Long) As StringEx

    Set right = Me
    
    If p_ValueLen > 0& And length >= 0& And length < p_ValueLen Then
    
        RtlMoveMemory ByVal p_ValuePtr, ByVal p_ValuePtr + ((p_ValueLen - length) * 2&), length + length
    
        p_ValueLen = length
    
    End If

End Function

Public Function rTrim() As StringEx

    Set rTrim = Me
    
    If p_ValueLen Then
    
        If AscW(Mid$(p_Value, p_ValueLen, 1&)) = 32 Then
            
            Do
                p_ValueLen = p_ValueLen + -1
            Loop While InStrRev(p_Value, ChrW$(32&), p_ValueLen, vbBinaryCompare) = p_ValueLen
        
        End If
    
    End If

End Function

Public Function toArray(Optional ByVal AsANSI As Boolean = True) As Byte()

    Dim x As Long
    Dim b() As Byte
   
    If p_ValueLen Then

        x = p_ValueLen * (((Not AsANSI) * -1) + 1&)

        ReDim b(x + -1)

        If AsANSI Then
            RtlMoveMemory ByVal VarPtr(b(0)), ByVal p_Value, x
        Else
            RtlMoveMemory ByVal VarPtr(b(0)), ByVal p_ValuePtr, x
        End If

    End If
    
    toArray = b()

End Function

Public Function toLower() As StringEx
    
    Set toLower = Me
    
    If p_ValueLen Then CharLowerBuffW p_ValuePtr, p_ValueLen

End Function

Public Function toUpper() As StringEx
    
    Set toUpper = Me
    
    If p_ValueLen Then CharUpperBuffW p_ValuePtr, p_ValueLen

End Function

Public Function trim() As StringEx

    Set trim = Me
    
    lTrim
    rTrim

End Function

Private Sub BufferRedim(ByVal NewStringLen As Long, ByVal FromScratch As Boolean)

    Dim b As String
    Dim r As Boolean
    Dim x As Long
        
    p_ValueLen = (p_ValueLen * ((Not FromScratch) * -1)) + NewStringLen

    If p_ValueLen > (p_ValueLenReal \ 2&) Then p_ValueLenReal = p_ValueLenReal + (p_ValueLen * 2&) Else r = p_ValuePtr

    If Not r Then
    
        If FromScratch Then
        
            p_Value = Space$(p_ValueLenReal)
            p_ValuePtr = StrPtr(p_Value)
        
        Else
            
            x = p_ValueLen - NewStringLen
            
            If x Then b = VBA.left$(p_Value, x)
            
            p_Value = Space$(p_ValueLenReal)
            p_ValuePtr = StrPtr(p_Value)
            
            If x Then RtlMoveMemory ByVal p_ValuePtr, ByVal StrPtr(b), x + x
        
        End If

    End If
    
End Sub



'Public Function Duplicate(ByVal Count As Long) As StringEx
'
'    Dim i As Long
'    Dim l As Long
'
'    Set Duplicate = Me
'
'    If p_ValueLen > 0& And Count > 0& Then
'
'        l = p_ValueLen
'
'        BufferRedim l * Count, False
'
'        For i = 1& To Count
'            RtlMoveMemory ByVal p_ValuePtr + ((l * i) * 2&), ByVal p_ValuePtr, l + l
'        Next i
'
'    End If
'
'End Function


'Private Type TPOINT
'    x As Long
'    y As Long
'End Type
'
'Public Type TSAFEARRAY
'    iDims As Integer
'    iFeatures As Integer
'    lElementSize As Long
'    lLocks As Long
'    lData As Long
'    lVarType As Long
'    uBounds() As TPOINT
'End Type
'
''CONSTANTS
'
'Public Const B_MX As Byte = 255
'Public Const D_MN As Double = 1E+308
'Public Const D_MX As Double = -1E+308
'Public Const I_MN As Integer = -32768
'Public Const I_MX As Integer = 32767
'Public Const L_NG As Long = -1&
'Public Const L_MN As Long = &H80000000
'Public Const L_MX As Long = 2147483647
'
''VARIABLES
'
'Private gUnc As Boolean
'
''WINAPI
'
'Private Declare Function GetDesktopWindow Lib "user32" () As Long
'Private Declare Function GetLocaleInfoA Lib "kernel32" (ByVal lLocale As Long, ByVal lType As Long, ByVal sBuffer As String, ByVal lBufferLen As Long) As Long
'Private Declare Function GetLocaleInfoW Lib "kernel32" (ByVal lLocale As Long, ByVal lType As Long, ByVal lBuffer As Long, ByVal lBufferLen As Long) As Long
'Private Declare Function IsWindowUnicode Lib "user32" (ByVal lhWnd As Long) As Long
'
'Public Declare Sub RtlMoveMemory Lib "kernel32" (uTarget As Any, uSource As Any, ByVal lLen As Long)
'
''ROUTINES
'
'Private Sub Main()
'
'    gUnc = IsWindowUnicode(GetDesktopWindow)
'
'End Sub


'Public Function Parse(ByRef Value As Variant) As StringEx
'
'    Set Parse = Me
'
'    Dim a As TSAFEARRAY
'    Dim b As Integer
'    Dim i As Long
'    Dim j As Long
'    Dim x As Long
'
'    x = VarType(Value)
'
'    Select Case x
'
'        Case 2& To 8&, 10&, 11&, 14&, 17& 'INTEGER, LONG, SINGLE, DOUBLE, CURRENCY, DATE, STRING, ERROR, BOOLEAN, DECIMAL, BYTE
'
'            Me.Value = Value
'
'        Case 8209& 'BYTE array
'
'            If RArrPt(Value, a) = 1& Then 'get array info
'
'                x = a.uBounds(0).x 'get items count
'
'                If x Mod 2& Then 'odd bytes count (ANSI, UTF-8)
'
'                    If x >= 3& Then
'
'                        If Value(0) = 239 And Value(1) = 187 And Value(2) = 191 Then j = 3& 'check for UTF-8 byte order mark
'
'                        i = MultiByteToWideChar(65001, 8&, ByVal a.lData, x, ByVal 0&, 0&) 'estimate buffer size (with MB_ERR_INVALID_CHARS flag)
'
'                        If i Then '7-bit ANSI or UTF-8
'
'                            BufferRedim i - ((j > 0&) * L_NG), True
'                            MultiByteToWideChar 65001, 0&, ByVal a.lData + j, x, ByVal p_ValuePtr, p_ValueLen 'convert multi-bytes
'
'                            Exit Function
'
'                        End If
'
'                    End If
'
'                    BufferRedim x, True
'                    RtlMoveMemory ByVal p_ValuePtr, ByVal a.lData, x 'copy bytes as 8-bit ANSI
'
'                ElseIf x > 0& Then 'even bytes count (ANSI, UTF-8, UTF-16 BE/LE)
'
'                    Select Case Value(0) 'check if first byte BOM-like
'
'                        Case 255
'
'                            If Value(1) = 254 Then i = 2& 'UTF-16 Little Endian
'
'                        Case 254
'
'                            If Value(1) = 255 Then 'UTF-16 Big Endian
'
'                                For j = 2& To x + L_NG Step 2& 'convert to Little Endian order
'                                    RtlMoveMemory b, ByVal a.lData + j, 2&
'                                    b = (((b And &HFF00) \ &H100) And &HFF) Or (b And &H7F) * &H100 Or (b And &H80) * &HFF00 'swap integer low & high words
'                                    RtlMoveMemory ByVal a.lData + j, b, 2&
'                                Next j
'
'                                i = 2&
'
'                            End If
'
'                        Case 239
'
'                            If x >= 3& Then
'
'                                If Value(1) = 187 And Value(2) = 191 Then 'UTF-8
'
'                                    i = MultiByteToWideChar(65001, 8&, ByVal a.lData, x, ByVal 0&, 0&) 'estimate buffer size (with MB_ERR_INVALID_CHARS flag)
'
'                                    If i Then
'
'                                        BufferRedim i + L_NG, True
'                                        MultiByteToWideChar 65001, 0&, ByVal a.lData + 3&, x, ByVal p_ValuePtr, p_ValueLen 'convert multi-bytes
'
'                                        Exit Function
'
'                                    End If
'
'                                End If
'
'                            End If
'
'                    End Select
'
'                    If i = 0& Then 'no BOM
'
'                        If IsTextUnicode(ByVal a.lData, x, 7&) = 0& Then 'if IS_TEXT_UNICODE_ASCII16, IS_TEXT_UNICODE_STATISTICS and IS_TEXT_UNICODE_CONTROLS tests fail
'
'                            i = MultiByteToWideChar(65001, 8&, ByVal a.lData, x, ByVal 0&, 0&) 'estimate buffer size (with MB_ERR_INVALID_CHARS flag)
'
'                            If i Then '7-bit ANSI or UTF-8
'                                BufferRedim i, True
'                                MultiByteToWideChar 65001, 0&, ByVal a.lData, x, ByVal p_ValuePtr, p_ValueLen 'convert multi-bytes
'                            Else '8-bit ANSI or UTF-16 case (currently difference can not be determined correctly, f.e. Chinese double-byte encoded text is recognized as ANSI)
'                                BufferRedim x, True
'                                RtlMoveMemory ByVal p_Value, ByVal a.lData, x 'copy bytes as 8-bit ANSI
'                            End If
'
'                            Exit Function
'
'                        End If
'
'                    End If
'
'                    BufferRedim (x - i) \ 2&, True
'                    RtlMoveMemory ByVal p_ValuePtr, ByVal a.lData + i, x - i 'copy bytes as Unicode
'
'                End If
'
'            End If
'
'        Case Else
'
'            p_ValueLen = 0&
'
'    End Select

'End Function


'
'Public Function RArrPt(ByRef uArr As Variant, ByRef uPtr As TSAFEARRAY) As Long
'
'    '------------------------------------------------------------------------------------------------------------------------------------------'
'    '
'    ' PURPOSE   : Fill structure with array information.
'    '
'    ' RETURN    : Array dimensions count.
'    '
'    ' ARGUMENTS : uArr - source array
'    '             uPtr - returns array information structure
'    '
'    ' NOTES     : Expected array declare syntax - Dim Array(dim1 (cols), [dim2 (rows)], [dim 3], ... , [dim N]) As Type
'    '             Array item data adress = first item data start adress + ((col + row + ((cols - 1) * row)) * item size in bytes).
'    '
'    '------------------------------------------------------------------------------------------------------------------------------------------'
'
'    Dim i As Long
'    Dim j As Long
'    Dim x As Long
'
'    RtlMoveMemory x, ByVal VarPtr(uArr) + 8&, 4& 'get pointer to array variable
'
'    If x Then
'
'        RtlMoveMemory j, uArr, 2& 'get variable type
'
'        If j = 0& Then 'do not proceed empty variable
'            Exit Function
'        ElseIf j And 16384& Then 'if variable is passed by reference (pointer to pointer)
'            RtlMoveMemory x, ByVal x, 4& 'get real variable pointer
'            j = j - 16384& 'remove VT_BYREF flag
'        End If
'
'        If x Then 'if pointer data is array and is not empty array then
'
'            RtlMoveMemory uPtr, ByVal x, 16& 'fill first fixed 16 bytes of structure from pointer
'
'            With uPtr
'                .lVarType = j - 8192& 'remove VT_ARRAY flag for convenient use
'                i = .iDims 'get array dimensions count
'                ReDim .uBounds(i + L_NG) 'allocate structure member
'                RtlMoveMemory .uBounds(0), ByVal x + 16&, i * 8& 'fill structure member with array dimensions info (in descending order) bytes starting from member pointer adress + 16 bytes offset
'            End With
'
'            RArrPt = i 'return dimensions count
'
'        End If
'
'    End If
'
'End Function
