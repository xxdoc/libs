VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CString"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private base As String
Private Const LANG_US = &H409
Private Const MAX_LONG As Long = 2147483647
Private selIndex As Long 'used in setLowest

Private Declare Function URLDownloadToFile Lib "urlmon" Alias _
    "URLDownloadToFileA" (ByVal pCaller As Long, _
    ByVal szURL As String, _
    ByVal szFileName As String, _
    ByVal dwReserved As Long, _
    ByVal lpfnCB As Long) As Long


'this property is the default method..get and let
Property Get text() As String
Attribute text.VB_UserMemId = 0
    text = base
End Property

Property Let text(x As String)
    base = x
End Property

'sometimes you need the loop like when replacing two spaces with one..
Function replace(find As String, Optional rep As String, Optional start As Long = 1, Optional count As Long = -1, Optional method As VbCompareMethod = vbBinaryCompare)
    tmp = base
    While InStr(start, tmp, find, method) > 0
        tmp = VBA.replace(tmp, find, rep, start, count, method)
    Wend
    replace = tmp
End Function

Function indexOf(find As String, Optional index As Long = 1, Optional method As VbCompareMethod = vbBinaryCompare) As Long
    indexOf = InStr(index, base, find, method)
End Function

Function lastIndexOf(find As String, Optional index As Long = -1, Optional method As VbCompareMethod = vbBinaryCompare) As Long
    lastIndexOf = InStrRev(base, find, index, method)
End Function

Function endsWith(ByVal char As String, Optional method As VbCompareMethod = vbBinaryCompare) As Boolean
    
    On Error GoTo hell
    Dim tmp As String
    
    If Len(base) < Len(char) Then Exit Function
    
    tmp = Right(base, Len(char))
    
    If method <> vbBinaryCompare Then
        tmp = LCase(tmp)
        char = LCase(char)
    End If
    
    If tmp = char Then endsWith = True
hell:
    
End Function

Sub stripFromEnd(Optional char As String = vbCrLf, Optional method As VbCompareMethod = vbBinaryCompare)
        
        While endsWith(char, method)
            base = Mid(base, 1, Len(base) - Len(char))
        Wend
        
End Sub

Sub stripAnyFromEnd(ParamArray chars())
        
        Dim hadResult As Boolean
        Dim cnt As Long
        Dim index As Long
        Dim b As String
        
        index = Len(base) - 1
        Do
            hadResult = False
            For Each c In chars
                b = Mid(base, index, 1)
                If LCase(b) = LCase(c) Then
                    index = index - 1
                    If index < 1 Then Exit Do
                    hadResult = True
                    Exit For 'we found the char, move to next
                End If
            Next
        Loop While hadResult
        
        If index = Len(base) - 1 Then Exit Sub 'nothing to do
        
        If index < 1 Then
            base = Empty
        Else
            base = Mid(base, 1, index)
        End If

    'much simplier but very inefficient for long strings..
    'If LCase(Right(base, 1)) = LCase(c) Then
    '    base = Mid(base, 1, Len(base) - Len(c))
    '    hadResult = True
    'End If
    
End Sub

Function readX(start As Long, length As Long) As String
    readX = Mid(base, start, length)
End Function

Function substr(start As Long, Optional endat As Long) As String
    If IsMissing(endat) Then
        substr = Mid(base, start)
    Else
        substr = Mid(base, start, endat - start)
    End If
End Function

Function charAt(x As Long)
    charAt = Mid(base, x, 1)
End Function

Function charCodeAt(x As Long) As Byte
    charCodeAt = Asc(Mid(base, x, 1))
End Function

Function split(at As String, Optional method As VbCompareMethod = vbBinaryCompare)
    split = VBA.split(base, at, , method)
End Function

Function join(wth As String) As String
    join = VBA.join(ary, wth)
End Function

Property Get length() As Long
    length = Len(base)
End Property

Function trim() As String
    trim = VBA.trim(base)
End Function

Function toLower() As String
    toLower = LCase(base)
End Function

Function toUpper() As String
    toUpper = UCase(base)
End Function

Function startsWith(str As String) As Boolean
    startsWith = (LCase(str) = LCase(Left(base, Len(str))))
End Function

Function concat(str As String)
    concat = base & str
End Function

Function getLine(Optional lNo As Long = 0) As String
    
    Dim lineEnd As Long
    Dim lineStart As Long
    Dim cnt As Long
    
    lineStart = 1
    lineEnd = InStr(base, vbCrLf)
    While cnt < lNo
        If lineEnd < 1 Then Exit Function
        lineStart = lineEnd + 2
        lineEnd = InStr(lineEnd + 1, base, vbCrLf)
        cnt = cnt + 1
    Wend
    
    If lineEnd < 1 Then Exit Function
    If lineStart > lineEnd Then Exit Function
    
    getLine = Mid(base, lineStart, lineEnd - lineStart)
    
End Function

'--------------- [ extras ] -------------------------------



Function AnyLikeThese(csvCmp) As Boolean
    Dim tmp() As String, i As Integer, sin As String
    tmp() = split(LCase(csvCmp), ",")
    sin = LCase(base)
    For i = 0 To UBound(tmp)
        tmp(i) = VBA.trim(tmp(i))
        If Len(tmp(i)) > 0 And sin Like "*" & tmp(i) & "*" Then
            AnyLikeThese = True
            Exit Function
        End If
    Next
End Function

Function AnyOfTheseInstr(csvCmp As String) As Boolean
    Dim tmp() As String, i As Integer
    tmp() = split(csvCmp, ",")
    For i = 0 To UBound(tmp)
        tmp(i) = VBA.trim(tmp(i))
        If Len(tmp(i)) > 0 And InStr(1, base, tmp(i), vbTextCompare) > 0 Then
            AnyOfTheseInstr = True
            Exit Function
        End If
    Next
End Function

Function CountOccurances(find, Optional method As VbCompareMethod = vbTextCompare) As Integer
    Dim tmp() As String
    If InStr(1, base, find, vbTextCompare) < 1 Then CountOccurances = 0: Exit Function
    tmp = VBA.split(base, find, , method)
    CountOccurances = UBound(tmp)
End Function

Function reverse() As String
    reverse = StrReverse(base)
End Function

Function toBytes() As Byte()
    toBytes = StrConv(base, vbFromUnicode, LANG_US)
End Function

Function extract(marker1, marker2, Optional start As Long = 1, Optional ByRef lastPos As Long, Optional method As VbCompareMethod = vbBinaryCompare) As String
    
    a = InStr(start, base, marker1, method)
    If a < 1 Then Exit Function
    
    a = a + Len(marker1)
    
    b = InStr(a, base, marker2, method)
    If b < 1 Then Exit Function
    
    lastPos = b + Len(marker2)
    extract = Mid(base, a, b - a)
    
End Function

Function toHexString(Optional str)
    Dim b() As Byte
    Dim ret() As String
    Dim t As String
    
    If IsMissing(str) Then
        If Len(base) = 0 Then Exit Function
        t = base
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
    
    toHexString = VBA.join(ret, "")
        
End Function

Function pad(v, Optional l As Long = 8)
    On Error GoTo hell
    Dim x As Long
    x = Len(v)
    If x < l Then
        pad = String(l - x, " ") & v
    Else
hell:
        pad = v
    End If
End Function

Function rpad(v, Optional l As Long = 10)
    On Error GoTo hell
    Dim x As Long
    x = Len(v)
    If x < l Then
        rpad = v & String(l - x, " ")
    Else
hell:
        rpad = v
    End If
End Function

Function LoadFromBytes(bArray) As Boolean
    On Error GoTo hell
    base = StrConv(bArray, vbUnicode, LANG_US)
    LoadFromBytes = True
hell:
End Function

Public Function LoadFromHexString(str As String) As Boolean

    Dim ret As String
    Dim x As String
    Dim errCount As Long
    Dim r() As Byte
    Dim b As Byte
    
    On Error Resume Next

    str = replace(str, " ", Empty)
    str = replace(str, vbCrLf, Empty)
    str = replace(str, vbCr, Empty)
    str = replace(str, vbLf, Empty)
    str = replace(str, vbTab, Empty)
    str = replace(str, Chr(0), Empty)
     
    For i = 1 To Len(str) Step 2
        x = Mid(str, i, 2)
        If isHexChar(x, b) Then
            bpush r(), b
        Else
            errCount = errCount + 1
            s_bpush r(), x
        End If
    Next

    base = StrConv(r(), vbUnicode, LANG_US)
    If Err.Number = 0 Then LoadFromHexString = True
    
End Function

Function LoadFromWeb(url) As Boolean
    On Error Resume Next
    Dim tmp As String
    tmp = Environ("temp") & "\cstring_dl.tmp"
    If FileExists(tmp) Then Kill tmp
    If DownloadFile(url, tmp) Then
        base = ReadFile(tmp)
        LoadFromWeb = True
        Kill tmp
    End If
End Function

Public Function LoadFromFile(path As String) As Boolean
    On Error GoTo hell
    If Not FileExists(path) Then Exit Function
    base = ReadFile(path)
    LoadFromFile = True
hell:
End Function

Function SaveToFile(path As String) As Boolean
On Error GoTo hell
    If FileExists(path) Then Kill path
    WriteFile path, base
    SaveToFile = True
hell:
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

Function HexDump(Optional val, Optional hexOnly As Boolean = False, Optional start As Long = 1, Optional leng As Long = -1) As String
    Dim tmp As String
    
    If IsMissing(val) Then
        HexDump = internal_HexDump(base, hexOnly, start, leng)
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


Private Function internal_HexDump(str As String, Optional hexOnly As Boolean = False, Optional start As Long = 0, Optional leng As Long = -1) As String
    
    On Error Resume Next
    
    Dim ary() As Byte
    Dim offset As Long
    Dim s() As String, chars As String, tmp As String
    
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
    
    internal_HexDump = VBA.join(s, vbCrLf)
    
    If hexOnly Then
        internal_HexDump = VBA.replace(internal_HexDump, " ", "")
        internal_HexDump = VBA.replace(internal_HexDump, vbCrLf, "")
    End If
    
End Function

'this should now be unicode safe on foreign systems..
Function unescape() As String '%uxxxx and %xx
    
    'On Error GoTo hell
    
    Dim tmp() As String
    Dim b1 As String, b2 As String
    Dim i As Long
    Dim r() As Byte
    Dim elems As Long
    
    tmp = VBA.split(base, "%")
    
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
     unescape = StrConv(r(), vbUnicode, LANG_US)
     
     If Err.Number <> 0 Then
        MsgBox "Error in unescape: " & Err.Description
     End If
     
End Function

Function findNextChar(charList, Optional ByRef whichFound As String, Optional ByRef startAt As Long = 1) As Long
    Dim tmp() As String
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
            push m, InStr(startAt, base, tmp(i))
        End If
    Next
    
    low = lowest(m()) '-1 for none found..
    
    If low = -1 Then
        findNextOf = -1
        Exit Function
    End If
    
    whichFound = tmp(selIndex)
    findNextOf = InStr(startAt, base, whichFound)
    
End Function



'--------------------- [ support functions ] -------------------------

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
    Dim b() As Byte
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
       
    Dim tmp As Long, i As Long
    
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
    Const allZero As Boolean = False
    
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

