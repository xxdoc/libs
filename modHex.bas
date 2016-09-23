Attribute VB_Name = "Module1"
Function HexDump(ByVal bAryOrStrData, Optional hexOnly = 0) As String
    Dim s() As String, chars As String, tmp As String
    On Error Resume Next
    Dim ary() As Byte
    Dim offset As Long
    Const LANG_US = &H409
    Dim i As Long, tt, h, X

    offset = 0
    
    If TypeName(bAryOrStrData) = "Byte()" Then
        ary() = bAryOrStrData
    Else
        ary = StrConv(bAryOrStrData, vbFromUnicode, &H409)
    End If
    
    chars = "   "
    For i = 1 To UBound(ary) + 1
        tt = Hex(ary(i - 1))
        If Len(tt) = 1 Then tt = "0" & tt
        tmp = tmp & tt & " "
        X = ary(i - 1)
        'chars = chars & IIf((x > 32 And x < 127) Or x > 191, Chr(x), ".") 'x > 191 causes \x0 problems on non us systems... asc(chr(x)) = 0
        chars = chars & IIf((X > 32 And X < 127), Chr(X), ".")
        If i > 1 And i Mod 16 = 0 Then
            h = Hex(offset)
            While Len(h) < 6: h = "0" & h: Wend
            If hexOnly = 0 Then
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
        If hexOnly = 0 Then
            h = Hex(offset)
            While Len(h) < 6: h = "0" & h: Wend
            h = h & "   " & tmp
            While Len(h) <= 56: h = h & " ": Wend
            push s, h & chars
        Else
            push s, tmp
        End If
    End If
    
    HexDump = Join(s, vbCrLf)
    
    If hexOnly <> 0 Then
        HexDump = Replace(HexDump, " ", "")
        HexDump = Replace(HexDump, vbCrLf, "")
    End If
    
End Function

Sub push(ary, value) 'this modifies parent ary object
    On Error GoTo init
    X = UBound(ary) '<-throws Error If Not initalized
    ReDim Preserve ary(UBound(ary) + 1)
    ary(UBound(ary)) = value
    Exit Sub
init:     ReDim ary(0): ary(0) = value
End Sub

Public Function toByteStr(hexstr, Optional isDecimal As Boolean = False) As String
    On Error Resume Next
    Dim b() As Byte
    b() = toBytes(hexstr, isDecimal)
    toByteStr = StrConv(b, vbUnicode, &H409)
End Function

Public Function toBytes(hexstr, Optional isDecimal As Boolean = False) As Byte()

'supports:
'11 22 33 44   spaced hex chars
'11223344      run together hex strings
'11,22,33,44   csv hex
'1,2,3,4       csv hex with no lead 0
'121,99,44,255 decimal csv or spaced values
'isDecimal flag requires csv or spaced values..
'ignores common C source prefixes and characters

    Dim ret As String, X As String, str As String
    Dim r() As Byte, b As Byte
    Dim foundDecimal As Boolean
    
    On Error GoTo hell
    
    If Len(hexstr) > 4 Then
        b = Asc((Mid(hexstr, 3, 1)))
        If b = Asc(" ") Or b = Asc(",") Then 'make sure all are double digit hex chars...
            tmp = Split(hexstr, Chr(b))
            
            If isDecimal Then
                For i = 0 To UBound(tmp)
                    tmp(i) = Hex(CLng(tmp(i)))
                Next
            End If
            
            For i = 0 To UBound(tmp)
                If Len(tmp(i)) = 1 Then tmp(i) = "0" & tmp(i)
            Next
        End If
    End If
        
    str = Replace(hexstr, " ", Empty)
    str = Replace(str, vbCrLf, Empty)
    str = Replace(str, vbCr, Empty)
    str = Replace(str, vbLf, Empty)
    str = Replace(str, vbTab, Empty)
    str = Replace(str, Chr(0), Empty)
    str = Replace(str, ",", Empty)
    str = Replace(str, "0x", Empty)
    str = Replace(str, "{", Empty)
    str = Replace(str, "}", Empty)
    str = Replace(str, ";", Empty)
    
    For i = 1 To Len(str) Step 2
        X = Mid(str, i, 2)
        If Not isHexChar(X, b) Then Exit Function
        bpush r(), b
    Next
    
    toBytes = r
    
hell:
End Function

Private Sub bpush(bAry() As Byte, b As Byte) 'this modifies parent ary object
    On Error GoTo init
    Dim X As Long
    
    X = UBound(bAry) '<-throws Error If Not initalized
    ReDim Preserve bAry(UBound(bAry) + 1)
    bAry(UBound(bAry)) = b
    
    Exit Sub

init:
    ReDim bAry(0)
    bAry(0) = b
    
End Sub

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

Public Function RC4(ByVal data As Variant, ByVal Password As Variant) As String
On Error Resume Next
    Dim RB(0 To 255) As Integer, X As Long, Y As Long, Z As Long, Key() As Byte, ByteArray() As Byte, Temp As Byte
    
    Dim plen As Long
    
    If TypeName(data) = "Byte()" Then
        ByteArray() = data
    Else
        If Len(data) = 0 Then Exit Function
        ByteArray() = StrConv(CStr(data), vbFromUnicode)
    End If
    
    If TypeName(Password) = "Byte()" Then
        Key() = Password
        If UBound(Key) > 255 Then ReDim Preserve Key(255)
    Else
        If Len(Password) = 0 Then
            Exit Function
        End If

        If Len(Password) > 256 Then
            Key() = StrConv(Left$(CStr(Password), 256), vbFromUnicode)
        Else
            Key() = StrConv(CStr(Password), vbFromUnicode)
        End If
    End If
    
    plen = UBound(Key) + 1
 
    'Debug.Print "key=" & HexDump(Key)
    'Debug.Print "data=" & HexDump(ByteArray)
    
    For X = 0 To 255
        RB(X) = X
    Next X
    
    X = 0
    Y = 0
    Z = 0
    For X = 0 To 255
        Y = (Y + RB(X) + Key(X Mod plen)) Mod 256
        Temp = RB(X)
        RB(X) = RB(Y)
        RB(Y) = Temp
    Next X
    
    X = 0
    Y = 0
    Z = 0
    For X = 0 To UBound(ByteArray)
        Y = (Y + 1) Mod 256
        Z = (Z + RB(Y)) Mod 256
        Temp = RB(Y)
        RB(Y) = RB(Z)
        RB(Z) = Temp
        ByteArray(X) = ByteArray(X) Xor (RB((RB(Y) + RB(Z)) Mod 256))
    Next X
    
    RC4 = StrConv(ByteArray, vbUnicode)
    
End Function



Function FileExists(path As String) As Boolean
  On Error GoTo hell
    
  If Len(path) = 0 Then Exit Function
  If Right(path, 1) = "\" Then Exit Function
  If Dir(path, vbHidden Or vbNormal Or vbReadOnly Or vbSystem) <> "" Then FileExists = True
  
  Exit Function
hell: FileExists = False
End Function

Function ReadFile(filename)
  f = FreeFile
  Temp = ""
   Open filename For Binary As #f        ' Open file.(can be text or image)
     Temp = Input(FileLen(filename), #f) ' Get entire Files data
   Close #f
   ReadFile = Temp
End Function

Function KeyExists(c As Collection, val As String) As Boolean
    On Error GoTo nope
    Dim t
    t = c(val)
    KeyExists = True
 Exit Function
nope:
End Function




Function GetAllElements(lv As ListView) As String
    Dim ret() As String, i As Integer, tmp As String
    Dim li As ListItem
    
    For i = 1 To lv.ColumnHeaders.Count
        tmp = tmp & lv.ColumnHeaders(i).Text & vbTab
    Next
    
    push ret, tmp
    push ret, String(50, "-")
        
    For Each li In lv.ListItems
        tmp = li.Text & vbTab
        For i = 1 To lv.ColumnHeaders.Count - 1
            tmp = tmp & li.SubItems(i) & vbTab
        Next
        push ret, tmp
    Next
    
    GetAllElements = Join(ret, vbCrLf)
    
End Function



