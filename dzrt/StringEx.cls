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
Attribute VB_Exposed = True
Option Explicit

'so this class is a mash up between a StringBuffer class which vb6 really needs for large strings
'and a string class which extends built in string primitives.
'
'for functions which modify the parent string value, they will return the same StringEx reference not a new one..
'this is so you can chain them like str.trim().tolower()
'technically its probably to always return a new string object but performance... so I am not sold on this yet...

'some functions do however return new string instances..
'these are generally data extraction routines and they all have their first letter capatilized as a subtle reminder.
'if you need to make sure you are working on a copy and not modifying the original string you can use clone.
'set freshStr = myStrEx.Clone().trim()

'not that get/let value is the default method for the class so it is compatiable with vb native
'routines such as
'
'dim x as string, myStrEx as new StringEx
'x = trim(myStrEx) also you can do
'myStrEx = x
'
'keep in mind that
'myStrEx2 = myStrEx              'this will make a copy of myStrEx.value into the old instance of myStrEx2 through the default .value props
'set myStrEx2 = myStrEx          'both objects now reference the same class instance..changes to one changes the other..
'set myStrEx2 = myStrEx.Clone()  'myStrEx2 is now an independant object and has a full copy of myStrEx.value
'
'so really using this class will be a bit more nuanced than i would like..but should have good performance
'and with some practice can do everything you need efficiently i think...still playing with it but this is where i am now..

'this will let me experiment with how I want it to behave...eventually maybe a private const..
'ok this wont work because some methods use private class members and I dont want to friend it..humm...
'Public alwaysNew As Boolean 'default is true for now..

Private Const LANG_US = &H409
Private Const MAX_LONG As Long = 2147483647
Private selIndex As Long 'used in setLowest

Private p_Value As String
Private p_ValueLen As Long
Private p_ValueLenReal As Long
Private p_ValuePtr As Long     'Strptr(p_Value)

'WINAPI

Private Declare Function CharLowerBuffW Lib "user32" (ByVal lStr As Long, ByVal lLength As Long) As Long
Private Declare Function CharUpperBuffW Lib "user32" (ByVal lStr As Long, ByVal lLength As Long) As Long
Private Declare Function IsTextUnicode Lib "advapi32" (ByRef uBuffer As Any, ByVal lBufferLen As Long, ByRef lResult As Long) As Long
Private Declare Function MultiByteToWideChar Lib "kernel32" (ByVal lCodePage As Long, ByVal lFlags As Long, uMultiByte As Any, ByVal lMultiByteLen As Long, uWideCharStr As Any, ByVal lWideCharStrLen As Long) As Long
Private Declare Sub RtlMoveMemory Lib "kernel32" (uTarget As Any, uSource As Any, ByVal lLen As Long)


Private Declare Function URLDownloadToFile Lib "urlmon" Alias _
    "URLDownloadToFileA" (ByVal pCaller As Long, _
    ByVal szURL As String, _
    ByVal szFileName As String, _
    ByVal dwReserved As Long, _
    ByVal lpfnCB As Long) As Long

Private Sub Class_Initialize()
    p_ValueLenReal = 256&
    'alwaysNew = True
End Sub

Public Property Get value() As String
Attribute value.VB_UserMemId = 0
    If p_ValueLen Then value = VBA.Left$(p_Value, p_ValueLen)
End Property

Public Property Let value(ByRef newvalue As String)
        
    Dim l As Long

    l = Len(newvalue)

    If l Then
    
        BufferRedim l, True
    
        RtlMoveMemory ByVal p_ValuePtr, ByVal StrPtr(newvalue), p_ValueLen + p_ValueLen
    
    Else
        p_ValueLen = 0&
    End If
    
End Property



'--------------------------------------------------------------------------------
'methods which modify the parent string value...
'--------------------------------------------------------------------------------

Public Function concat(ByRef ConcatString As String) As StringEx
    
    Dim l As Long
    
    Set concat = Me
    l = Len(ConcatString)
    
    If l Then
    
        BufferRedim l, False
    
        RtlMoveMemory ByVal p_ValuePtr + ((p_ValueLen - l) * 2&), ByVal StrPtr(ConcatString), l + l
    
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

Public Function insert(ByVal index As Long, ByRef InsertString As String) As StringEx
    
'    Dim ret As StringEx
'
'    If alwaysNew Then
'        Set insert = New StringEx
'        insert.value = Me.value
'    Else
'        Set insert = Me
'    End If
    
    Dim l As Long

    l = Len(InsertString)
        
    If l > 0& And index >= 0& And index <= p_ValueLen Then
    
        BufferRedim l, False
        
        RtlMoveMemory ByVal p_ValuePtr + (index * 2&) + l + l, ByVal p_ValuePtr + (index * 2&), (p_ValueLen - (index + l)) * 2&
        RtlMoveMemory ByVal p_ValuePtr + (index * 2&), ByVal StrPtr(InsertString), l + l
    
    End If

End Function

Function stripFromEnd(Optional char As String = vbCrLf, Optional method As VbCompareMethod = vbBinaryCompare) As StringEx
    
    Set stripFromEnd = Me
    While endsWith(char, method)
        'value = Mid(value, 1, Len(value) - Len(char))
        p_ValueLen = p_ValueLen - Len(char)
    Wend
        
End Function

Public Function stripAnyFromEnd2(ByVal csv_args As String, Optional seperator = ",") As StringEx
    Set stripAnyFromEnd2 = Me
    stripAnyFromEnd VBA.split(csv_args, seperator)
End Function

'Function stripAnyFromEnd(ParamArray chars()) As StringEx
'
'        Dim hadResult As Boolean
'        Dim cnt As Long
'        Dim index As Long
'        Dim b As String
'        Dim c
'
'        Set stripAnyFromEnd = Me
'
'        index = Len(value) - 1
'        Do
'            hadResult = False
'            For Each c In chars
'                b = Mid(value, index, 1)
'                If LCase(b) = LCase(c) Then
'                    index = index - 1
'                    If index < 1 Then Exit Do
'                    hadResult = True
'                    Exit For 'we found the char, move to next
'                End If
'            Next
'        Loop While hadResult
'
'        If index = Len(value) - 1 Then Exit Function 'nothing to do
'
'        If index < 1 Then
'            value = Empty
'        Else
'            value = Mid(value, 1, index)
'        End If
'
'
'End Function

Function stripAnyFromEnd(ParamArray chars()) As StringEx

        Dim hadResult As Boolean
        Dim cnt As Long
        Dim i As Long
        Dim b As String
        Dim c

        Set stripAnyFromEnd = Me

        i = p_ValueLen
        Do
            hadResult = False
            For Each c In chars
                stripFromEnd c
                If p_ValueLen <> i Then
                    If p_ValueLen = 0 Then Exit Do
                    hadResult = True
                End If
            Next
        Loop While hadResult

End Function

Public Function toLower() As StringEx
    Set toLower = Me
    If p_ValueLen Then CharLowerBuffW p_ValuePtr, p_ValueLen
End Function

Public Function toUpper() As StringEx
    Set toUpper = Me
    If p_ValueLen Then CharUpperBuffW p_ValuePtr, p_ValueLen
End Function

Public Function trim(Optional allWhiteSpace As Boolean = False) As StringEx
    Set trim = Me
    lTrim allWhiteSpace
    rTrim allWhiteSpace
End Function

Function isSpace(Optional pos As Long = -1, Optional allWhiteSpace As Boolean = False) As Boolean
    
    Dim b As Integer
    
    If p_ValueLen < 1 Then Exit Function
    
    If pos = -1 Then 'check last character
        b = Asc(Mid$(p_Value, p_ValueLen, 1&))
    Else
        b = Asc(Mid(p_Value, pos, 1))
    End If
    
    If allWhiteSpace = False Then
        isSpace = (b = 32)
    Else
        'return true for Space, CR, LF, Tab
        isSpace = (b = 32)
        If b = 10 Then isSpace = True
        If b = 13 Then isSpace = True
        If b = 9 Then isSpace = True
    End If
        
End Function

Public Function rTrim(Optional allWhiteSpace As Boolean = False) As StringEx
    
    Set rTrim = Me

    While isSpace(True, allWhiteSpace)
         p_ValueLen = p_ValueLen + -1
    Wend

End Function

Public Function lTrim(Optional allWhiteSpace As Boolean = False) As StringEx

    Dim x As Long

    Set lTrim = Me

    x = 1
    'how many white space characters are at the beginning of the string
    While isSpace(x, allWhiteSpace)
        x = x + 1
    Wend

    If x > 1 Then
        x = x - 1 'zero based...we had to be 1 based for instr...
        p_ValueLen = p_ValueLen - x
        'take the whole block, and move it back x bytes (*2 is for unicode char width)
        RtlMoveMemory ByVal p_ValuePtr, ByVal p_ValuePtr + (x * 2), p_ValueLen * 2
    End If

End Function

Function unixToDOS() As StringEx
    Dim tmp As String, isMixed As Boolean
    
    Set unixToDOS = Me
    
    tmp = value
    isMixed = (InStr(tmp, vbCrLf) > 0)
    If isMixed Then tmp = VBA.Replace(tmp, vbCrLf, Chr(5))
    tmp = VBA.Replace(tmp, vbLf, vbCrLf)
    If isMixed Then tmp = VBA.Replace(tmp, Chr(5), vbCrLf)
    value = tmp
    
End Function

Public Function strip(ByRef SearchString As String, Optional seperator = "", Optional ByVal CompareMethod As VbCompareMethod) As StringEx
    Set strip = Me
    If p_ValueLen Then
        If seperator = Empty Then
            value = VBA.Replace(VBA.Left$(p_Value, p_ValueLen), SearchString, Empty, , , CompareMethod)
        Else
            strip2 VBA.split(SearchString, seperator)
        End If
    End If
End Function

Public Function strip2(ParamArray args()) As StringEx
    
    Dim tmp As String
    Dim v
    
    Set strip2 = Me
    tmp = Me.value
    For Each v In args
       tmp = VBA.Replace(tmp, v, Empty, , , vbTextCompare)
       If Len(tmp) = 0 Then Exit For
    Next
    
    value = tmp

End Function

Function reverse() As StringEx
    Set reverse = Me
    value = StrReverse(value)
End Function

'Examples:
' ("number 0x%x", 16)
' ("number  %08x", &HCC)
' ("number '%8X'", &HCC)
' ("this is %s and can also be %S", "my string", "upper case")
' ("%s", "string with no prefix")
' ("0xCC in decimal = %d", &HCC)
' ("chr(&h41) = %c", &H41)
' ("chr(&h61) = %c it is my %s", &H61, "favorite letter")
' ("this is\n\ta test of \n\t\tnew lines tabs \t and %% escapes")

Function printf(ByVal msg As String, ParamArray vars()) As StringEx
'supports %x, %c, %s, %d, %10d \t \n %%

    Dim t
    Dim ret As String
    Dim i As Long
    Dim value As String
    Dim marker As String
    
    Set printf = Me
    
    msg = VBA.Replace(msg, Chr(0), Empty)
    msg = VBA.Replace(msg, "\t", vbTab)
    msg = VBA.Replace(msg, "\n", vbCrLf) 'simplified
    msg = VBA.Replace(msg, "%%", Chr(0))
    
    t = VBA.split(msg, "%")
    If UBound(t) <> UBound(vars) + 1 Then
        MsgBox "Format string mismatch.."
        Exit Function
    End If
    
    ret = t(0)
    For i = 1 To UBound(t)
        value = t(i)
        marker = ExtractSpecifier(value)
        If Len(marker) > 0 Then
            ret = ret & HandleMarker(value, marker, vars(i - 1))
        Else
            ret = ret & value
        End If
    Next
    
    ret = VBA.Replace(ret, Chr(0), "%")
    value = ret
    
End Function



'------------------------------------------------------------------------------------------------
'these ones return a new string....

Function Clone() As StringEx
    Set Clone = New StringEx
    Clone.value = Me.value
End Function

'sometimes you need the loop like when replacing two spaces with one..
Function Replace(find As String, Optional rep As String = Empty, Optional start As Long = 1, _
                Optional count As Long = -1, Optional method As VbCompareMethod = vbBinaryCompare _
) As StringEx

    Dim tmp As String
    Dim s As New StringEx
    
    Set Replace = s
    
    tmp = value
    While InStr(start, tmp, find, method) > 0
        tmp = VBA.Replace(tmp, find, rep, start, count, method)
    Wend
    s.value = tmp
End Function

'supports negative offsets
Public Function SubString(Optional start As Long = 1, Optional ByVal endAt As Long = 0) As StringEx
    Dim s As New StringEx
    Dim v As String
    Dim leng As Long
    
    Set SubString = s
    
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
        
        s.value = Mid(p_Value, start, endAt)
    Else
        
        If endAt > 0 Then Exit Function
        If endAt < start Then Exit Function
        If Abs(start) > Me.length Then Exit Function
        
        v = VBA.Right(Me.value, Abs(start))
        If endAt < 0 Then
            v = VBA.Left(v, Len(v) - Abs(endAt))
        End If
        
        s.value = v
         
    End If
    
            
        
End Function

Public Function SubStr(start As Long, Optional ByVal length As Long = -1) As StringEx
    Dim s As New StringEx
    
    Set SubStr = s
    
    If start > Me.length Then Exit Function
    
    If length = -1 Or length > Me.length - start Then
        length = Me.length - start + 1
    End If
    
    s.value = Mid(p_Value, start, length)
    
End Function

Function GetLine(Optional lNo As Long = 0) As StringEx
    
    Dim lineEnd As Long
    Dim lineStart As Long
    Dim cnt As Long
    Dim s As New StringEx
    
    Set GetLine = s
    
    lineStart = 1
    lineEnd = InStr(value, vbCrLf)
    While cnt < lNo
        If lineEnd < 1 Then Exit Function
        lineStart = lineEnd + 2
        lineEnd = InStr(lineEnd + 1, value, vbCrLf)
        cnt = cnt + 1
    Wend
    
    If lineEnd < 1 Then Exit Function
    If lineStart > lineEnd Then Exit Function
    
    s.value = Mid(value, lineStart, lineEnd - lineStart)
    
End Function

Function Extract(marker1, Optional marker2 = "", _
                Optional start As Long = 1, _
                Optional ByRef lastPos As Long, _
                Optional method As VbCompareMethod = vbBinaryCompare _
) As StringEx

    Dim a As Long, b As Long
    Dim s As New StringEx
    
    Set Extract = s
    
    a = InStr(start, value, marker1, method)
    If a < 1 Then Exit Function
    
    a = a + Len(marker1)
    If Len(marker2) = 0 Then
        lastPos = Len(value)
        s.value = Mid(value, a)
    Else
        b = InStr(a, value, marker2, method)
        If b < 1 Then Exit Function
        
        lastPos = b + Len(marker2)
        s.value = Mid(value, a, b - a)
    End If
    
End Function

'this should now be unicode safe on foreign systems..
Function Unescape(Optional newVal As String, Optional hadErr As Boolean) As StringEx   '%uxxxx and %xx
    
    'On Error GoTo hell
    
    Dim tmp() As String
    Dim b1 As String, b2 As String
    Dim i As Long, t
    Dim r() As Byte
    Dim elems As Long
    Dim s As New StringEx
    
    hadErr = False
    Set Unescape = s
    If Len(newVal) > 0 Then
        tmp = VBA.split(newVal, "%")
    Else
        tmp = VBA.split(value, "%")
    End If
    
    s_bpush r(), tmp(0) 'any prefix before encoded part..
    
    For i = 1 To UBound(tmp)
        t = tmp(i)
        
        If LCase(VBA.Left(t, 1)) = "u" Then
        
            If Len(t) < 5 Then '%u21 -> %u0021
                t = "u" & String(5 - Len(t), "0") & Mid(t, 2)
            End If

            b1 = Mid(t, 2, 2)
            b2 = Mid(t, 4, 2)
            
            If isHexChar(b1) And isHexChar(b2) Then
                hex_bpush r(), b2
                hex_bpush r(), b1
            Else
                s_bpush r(), "%u" & b1 & b2
            End If
            
            If Len(t) > 5 Then s_bpush r(), Mid(t, 6)
             
        Else
               b1 = Mid(t, 1, 2)
               If Not hex_bpush(r(), b1) Then s_bpush r(), "%" & b1
               If Len(t) > 2 Then s_bpush r(), Mid(t, 3)
        End If
        
    Next
            
hell:
     s.value = StrConv(r(), vbUnicode, LANG_US)
     hadErr = (Err.Number <> 0)
     
End Function

Function Pad(Optional l As Long = 8) As StringEx
    On Error GoTo hell
    Dim x As Long
    Dim s As New StringEx
    Set Pad = s
    x = Len(value)
    If x < l Then
        s.value = String(l - x, " ") & value
    Else
hell:
        s.value = value
    End If
End Function

Function Rpad(v, Optional l As Long = 10) As StringEx
    On Error GoTo hell
    Dim x As Long
    Dim s As New StringEx
    Set Rpad = s
    x = Len(value)
    If x < l Then
        s.value = value & String(l - x, " ")
    Else
hell:
        s.value = value
    End If
End Function

'=================================================================================================
' load from /save to functions
'-------------------------------------------------------------------------------------------------

Function loadFromBytes(bArray) As Boolean
    On Error GoTo hell
    value = StrConv(bArray, vbUnicode, LANG_US)
    loadFromBytes = True
hell:
End Function

Public Function loadFromHexString(str As String) As Boolean

    Dim ret As String
    Dim x As String
    Dim errCount As Long
    Dim r() As Byte
    Dim b As Byte
    Dim i As Long
    
    On Error Resume Next

    str = Replace(str, " ", Empty)
    str = Replace(str, vbCrLf, Empty)
    str = Replace(str, vbCr, Empty)
    str = Replace(str, vbLf, Empty)
    str = Replace(str, vbTab, Empty)
    str = Replace(str, Chr(0), Empty)
     
    For i = 1 To Len(str) Step 2
        x = Mid(str, i, 2)
        If isHexChar(x, b) Then
            bpush r(), b
        Else
            errCount = errCount + 1
            s_bpush r(), x
        End If
    Next

    value = StrConv(r(), vbUnicode, LANG_US)
    If Err.Number = 0 Then loadFromHexString = True
    
End Function

Function loadFromWeb(url) As Boolean
    On Error Resume Next
    Dim tmp As String
    tmp = Environ("temp") & "\StringEx_dl.tmp"
    If FileExists(tmp) Then Kill tmp
    If DownloadFile(url, tmp) Then
        value = ReadFile(tmp)
        loadFromWeb = True
        Kill tmp
    End If
End Function

Public Function loadFromFile(path As String) As Boolean
    On Error GoTo hell
    If Not FileExists(path) Then Exit Function
    value = ReadFile(path)
    loadFromFile = True
hell:
End Function

Function saveToFile(path As String) As Boolean
On Error GoTo hell
    If FileExists(path) Then Kill path
    WriteFile path, value
    saveToFile = True
hell:
End Function






'------------------------------------------------------------------------------------------------
'data retrieval or string tests (no mods to string unless a blatant loadfrom method which is expected to mod value..

Function split(Optional at As String, Optional method As VbCompareMethod = vbBinaryCompare)
    If Len(at) = 0 Then
        split = Me.toBytes()
    Else
        split = VBA.split(value, at, , method)
    End If
End Function

Function indexOf(find As String, Optional index As Long = 1, Optional method As VbCompareMethod = vbBinaryCompare) As Long
    indexOf = InStr(index, value, find, method)
End Function

Function lastIndexOf(find As String, Optional index As Long = -1, Optional method As VbCompareMethod = vbBinaryCompare) As Long
    lastIndexOf = InStrRev(value, find, index, method)
End Function

Function endsWith(ByVal char As String, Optional method As VbCompareMethod = vbBinaryCompare) As Boolean
    
    On Error GoTo hell
    Dim tmp As String
    
    If Len(value) < Len(char) Then Exit Function
    
    tmp = Right(value, Len(char))
    
    If method <> vbBinaryCompare Then
        tmp = LCase(tmp)
        char = LCase(char)
    End If
    
    If tmp = char Then endsWith = True
hell:
    
End Function

Function charAt(x As Long)
    On Error Resume Next
    charAt = Mid(value, x, 1)
End Function

Function charCodeAt(x As Long) As Byte
    On Error Resume Next
    charCodeAt = Asc(Mid(value, x, 1))
End Function

Public Property Get length() As Long
    length = p_ValueLen
End Property

Function startsWith(str As String) As Boolean
    On Error Resume Next
    startsWith = (LCase(str) = LCase(Left(value, Len(str))))
End Function

Function anyLikeThese(csvCmp) As Boolean
    Dim tmp() As String, i As Integer, sin As String
    tmp() = split(LCase(csvCmp), ",")
    sin = LCase(value)
    For i = 0 To UBound(tmp)
        tmp(i) = VBA.trim(tmp(i))
        If Len(tmp(i)) > 0 And sin Like "*" & tmp(i) & "*" Then
            anyLikeThese = True
            Exit Function
        End If
    Next
End Function

Function anyOfTheseInstr(csvCmp As String) As Boolean
    Dim tmp() As String, i As Integer
    tmp() = split(csvCmp, ",")
    For i = 0 To UBound(tmp)
        tmp(i) = VBA.trim(tmp(i))
        If Len(tmp(i)) > 0 And InStr(1, value, tmp(i), vbTextCompare) > 0 Then
            anyOfTheseInstr = True
            Exit Function
        End If
    Next
End Function

Function countOccurances(find, Optional method As VbCompareMethod = vbTextCompare) As Long
    Dim tmp() As String
    If InStr(1, value, find, method) < 1 Then
        countOccurances = 0
        Exit Function
    End If
    tmp = VBA.split(value, find, , method)
    countOccurances = UBound(tmp)
End Function

Public Function toBytes(Optional ByVal AsANSI As Boolean = True) As Byte()

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
    
    toBytes = b()

End Function

Function toHexString(Optional str) As String
    Dim b() As Byte
    Dim ret() As String
    Dim t As String
    Dim i As Long
    
    If IsMissing(str) Then
        If Len(value) = 0 Then Exit Function
        t = value
    Else
        t = str
    End If
    
    b() = StrConv(t, vbFromUnicode, LANG_US)
     
    For i = 0 To UBound(b)
        If b(i) < &H10 Then
            push ret, "0" & Hex(b(i))
        Else
            push ret, Hex(b(i))
        End If
    Next
    
    toHexString = VBA.Join(ret, "")
        
End Function

Public Function isHexChar(hexValue As String, Optional b As Byte) As Boolean
    On Error Resume Next
    Dim v As Long
    
    
    If Len(hexValue) = 0 Then GoTo nope
    If Len(hexValue) > 2 Then GoTo nope 'expecting hex char code like FF or 90
    
    v = CLng("&h" & hexValue)
    If Err.Number <> 0 Then GoTo nope 'invalid hex code
    
    b = CByte(v)
    If Err.Number <> 0 Then GoTo nope  'shouldnt happen.. > 255 cant be with len() <=2 ?

    isHexChar = True
    
    Exit Function
nope:
    Err.Clear
    isHexChar = False
End Function

Function HexDump(Optional start As Long = 1, Optional leng As Long = -1, Optional hexOnly As Boolean = False, Optional val) As String
    Dim tmp As String
    
    If IsMissing(val) Then
        HexDump = internal_HexDump(value, hexOnly, start, leng)
    Else
        If IsArray(val) Then
            If TypeName(val) = "Byte()" Then
                tmp = StrConv(val, vbUnicode, LANG_US)
                HexDump = internal_HexDump(tmp, hexOnly, start, leng)
            End If
        Else
            HexDump = internal_HexDump(CStr(val), hexOnly, start, leng)
        End If
    End If
    
End Function

Function findNextChar(charList, Optional ByRef whichFound As String, Optional ByRef startAt As Long = 1) As Long
    Dim tmp() As String, i As Long
    For i = 1 To Len(charList)
        push tmp, Mid(charList, i, 1)
    Next
    findNextChar = findNextOf(tmp, whichFound, startAt)
End Function

Function findNextOf(strArray() As String, Optional ByRef whichFound As String, Optional ByRef startAt As Long = 1) As Long
    
    Dim m() As Long
    Dim i As Long
    Dim low As Long
    Dim tmp() As String
    
    whichFound = Empty
    If startAt = 0 Then startAt = 1
    
    tmp = strArray()
    
    For i = 0 To UBound(tmp)
        If Len(tmp(i)) > 0 Then
            push m, InStr(startAt, value, tmp(i))
        End If
    Next
    
    low = lowest(m()) '-1 for none found..
    
    If low = -1 Then
        findNextOf = -1
        Exit Function
    End If
    
    whichFound = tmp(selIndex)
    findNextOf = InStr(startAt, value, whichFound)
    
End Function










'-----------------------------------------------------------------------------
' private functions
'-----------------------------------------------------------------------------

Private Function internal_HexDump(str As String, Optional hexOnly As Boolean = False, Optional start As Long = 0, Optional leng As Long = -1) As String
    
    On Error Resume Next
    
    Dim ary() As Byte
    Dim offset As Long
    Dim s() As String, chars As String, tmp As String
    Dim i As Long, tt, x, h
    
    offset = 0
    str = " " & str
    ary = StrConv(str, vbFromUnicode, LANG_US)
    
    If start < 0 Then start = 0
    If start > UBound(ary) Then Exit Function
    
    If leng = -1 Then
        leng = UBound(ary)
    Else
        leng = start + leng - 1
        If leng > UBound(ary) Then leng = UBound(ary)
    End If
    
    chars = "   "
    For i = start To leng
        tt = Hex(ary(i))
        If Len(tt) = 1 Then tt = "0" & tt
        tmp = tmp & tt & " "
        x = ary(i)
        'chars = chars & IIf((x > 32 And x < 127) Or x > 191, Chr(x), ".") 'x > 191 causes \x0 problems on non us systems... asc(chr(x)) = 0
        chars = chars & IIf((x > 32 And x < 127), Chr(x), ".")
        If i > 1 And i Mod 16 = 0 Then
            h = Hex(offset)
            While Len(h) < 6: h = "0" & h: Wend
            If Not hexOnly Then
                push s, h & "   " & tmp & chars
            Else
                push s, tmp
            End If
            offset = offset + 16
            tmp = Empty
            chars = "   "
        End If
    Next
    'if read length was not mod 16=0 then
    'we have part of line to account for
    If tmp <> Empty Then
        If Not hexOnly Then
            h = Hex(offset)
            While Len(h) < 6: h = "0" & h: Wend
            h = h & "   " & tmp
            While Len(h) <= 56: h = h & " ": Wend
            push s, h & chars
        Else
            push s, tmp
        End If
    End If
    
    internal_HexDump = VBA.Join(s, vbCrLf)
    
    If hexOnly Then
        internal_HexDump = VBA.Replace(internal_HexDump, " ", "")
        internal_HexDump = VBA.Replace(internal_HexDump, vbCrLf, "")
    End If
    
End Function

Private Sub bpush(bAry() As Byte, b As Byte) 'this modifies parent ary object
    On Error GoTo init
    Dim x As Long
    
    x = UBound(bAry) '<-throws Error If Not initalized
    ReDim Preserve bAry(UBound(bAry) + 1)
    bAry(UBound(bAry)) = b
    
    Exit Sub

init:
    ReDim bAry(0)
    bAry(0) = b
    
End Sub

Private Sub s_bpush(bAry() As Byte, sValue As String)
    Dim tmp() As Byte
    Dim i As Long
    tmp() = StrConv(sValue, vbFromUnicode, LANG_US)
    For i = 0 To UBound(tmp)
        bpush bAry, tmp(i)
    Next
End Sub


Private Sub push(ary, value) 'this modifies parent ary object
    On Error GoTo init
    Dim x As Long
    x = UBound(ary) '<-throws Error If Not initalized
    ReDim Preserve ary(UBound(ary) + 1)
    ary(UBound(ary)) = value
    Exit Sub
init:     ReDim ary(0): ary(0) = value
End Sub



Private Function FileExists(path As String) As Boolean
  On Error GoTo hell
    
  If Len(path) = 0 Then Exit Function
  If Right(path, 1) = "\" Then Exit Function
  If Dir(path, vbHidden Or vbNormal Or vbReadOnly Or vbSystem) <> "" Then FileExists = True
  
  Exit Function
hell: FileExists = False
End Function

Private Function ReadFile(filename) As String
  Dim f As Long
  Dim b() As Byte
  f = FreeFile
  Open filename For Binary As f
  ReDim b(LOF(f))
  Get f, , b()
  Close f
  ReadFile = StrConv(b(), vbUnicode, LANG_US)
End Function

Private Sub WriteFile(path, it)
    Dim b() As Byte, f As Long
    b() = StrConv(it, vbFromUnicode, LANG_US)
    f = FreeFile
    Open path For Binary As f
    Put f, , b()
    Close f
End Sub



Private Function hex_bpush(bAry() As Byte, hexValue As String) As Boolean   'this modifies parent ary object
    On Error Resume Next
    Dim b As Byte
    If Not isHexChar(hexValue, b) Then Exit Function
    bpush bAry, b
    hex_bpush = True
End Function

'note: if you pass in an array() dont utilize the paramarray() and submit multiple
'      you actually can, but dont expect selIndex to work..
'maxlong = error basically..
 Function lowest(ParamArray vals())
       
    Dim tmp As Long, i As Long, x, xx
    
    tmp = MAX_LONG
    
    selIndex = -1
      
    For Each x In vals
        If IsArray(x) Then
            For Each xx In x
                setIfLowest xx, tmp, i
                i = i + 1
            Next
            Exit For
        Else
            setIfLowest x, tmp, i
            i = i + 1
        End If
    Next
                    
    If tmp = MAX_LONG Then tmp = -1
    lowest = tmp
        
 End Function

 

 Private Sub setIfLowest(curVal, ByRef curLow, index As Long)

    Const allowNeg As Boolean = False
    Const allowZero As Boolean = False
    
        If curVal < 0 Then
            If allowNeg Then
                If curVal < curLow Then
                    curLow = curVal
                    selIndex = index
                End If
            End If
        ElseIf curVal = 0 Then
            If allowZero Then
                If curVal < curLow Then
                    curLow = curVal
                    selIndex = index
                End If
            End If
        Else
            If curVal < curLow Then
                curLow = curVal
                selIndex = index
            End If
        End If
        
 End Sub

Private Function AryIsEmpty(ary) As Boolean
  On Error GoTo oops
  Dim i As Long
  
    i = UBound(ary)  '<- throws error if not initalized
    AryIsEmpty = False
  Exit Function
oops: AryIsEmpty = True
End Function

Public Function DownloadFile(url, LocalFilename) As Boolean
    Dim lngRetVal As Long
    lngRetVal = URLDownloadToFile(0, CStr(url), CStr(LocalFilename), 0, 0)
    If lngRetVal = 0 Then DownloadFile = True
End Function



Private Function HandleMarker(value, ByVal marker, var) As String
    Dim newvalue As String
    Dim mType As Integer
    Dim nVal As String
    Dim spacer As String
    Dim prefix As String
    Dim count As Long
    
    If Len(value) > Len(marker) Then
        newvalue = Mid(value, Len(marker) + 1) 'remove the marker..
    End If
    
    mType = Asc(Mid(marker, Len(marker), 1))  'last character
    
    Select Case mType
        Case Asc("x"): nVal = Hex(var)
        Case Asc("X"): nVal = UCase(Hex(var))
        Case Asc("s"): nVal = var
        Case Asc("S"): nVal = UCase(var)
        Case Asc("c"): nVal = Chr(var)
        Case Asc("d"): nVal = var
        
        Case Else: nVal = var
    End Select
    
    If Len(marker) > 1 Then 'it has some more formatting involved..
        marker = Mid(marker, 1, Len(marker) - 1) 'trim off type
        If Left(marker, 1) = "0" Then
            spacer = "0"
            marker = Mid(marker, 2)
        Else
            spacer = " "
        End If
        count = CLng(marker) - Len(nVal)
        If count > 0 Then prefix = VBA.String(count, spacer)
    End If
    
    HandleMarker = prefix & nVal & newvalue
            
End Function

Private Function ExtractSpecifier(v)
    
    Dim ret As String
    Dim b() As Byte, i As Long
    
    Const LANG_US = &H409
    If Len(v) = 0 Then Exit Function
    
    b() = StrConv(v, vbFromUnicode, LANG_US)
    
    For i = 0 To UBound(b)
        ret = ret & Chr(b(i))
        If b(i) = Asc("x") Then Exit For
        If b(i) = Asc("X") Then Exit For
        If b(i) = Asc("c") Then Exit For
        If b(i) = Asc("s") Then Exit For
        If b(i) = Asc("S") Then Exit For
        If b(i) = Asc("d") Then Exit For
    Next
    
    ExtractSpecifier = ret
    
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
            
            If x Then b = VBA.Left$(p_Value, x)
            
            p_Value = Space$(p_ValueLenReal)
            p_ValuePtr = StrPtr(p_Value)
            
            If x Then RtlMoveMemory ByVal p_ValuePtr, ByVal StrPtr(b), x + x
        
        End If

    End If
    
End Sub

