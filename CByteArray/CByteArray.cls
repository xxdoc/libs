VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CByteArray"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByVal pDest As Long, ByVal pSrc As Long, ByVal ByteLen As Long)

Private base() As Byte
Private Const LANG_US = &H409

Property Get Length() As Long
    If isEmpty Then
        Length = -1
    Else
        Length = UBound(base)
    End If
End Property

Function isEmpty() As Boolean
  On Error GoTo oops
    i = UBound(base)  '<- throws error if not initalized
    isEmpty = False
  Exit Function
oops: isEmpty = True
End Function

Sub Clear()
    Erase base
End Sub

Function LoadFromStr(x As String)
    base() = StrConv(x, vbFromUnicode, LANG_US)
End Function

'space seperated or run together hex string, no other white space..
Function LoadFromHexStr(ByVal x As String) As Boolean
    base() = internal_LoadFromHexStr(x)
    If Not isEmpty Then LoadFromHexStr = True
End Function

Function LoadFromFile(filename) As Boolean
  On Error GoTo hell
  f = FreeFile
  Open filename For Binary As #f
  ReDim base(LOF(f) - 1)
  Get f, , base()
  Close #f
  ReadFile = True
  Exit Function
hell:
End Function

Function SaveToFile(path) As Boolean
    On Error GoTo hell
    If FileExists(path) Then Kill path
    f = FreeFile
    Open path For Binary As #f
    Put f, , base()
    Close f
    SaveToFile = True
    Exit Function
hell:
End Function

Function asByte() As Variant
    asByte = base
End Function

Function asString() As String
    asString = StrConv(base(), vbUnicode, LANG_US)
End Function

Function asHexStr() As String
    Dim t() As String
    For i = 0 To UBound(b)
        push t, sHex(base(i))
    Next
    asHexStr = Join(t, "")
End Function

Function hexdump(Optional tmp As String) As String
    Dim my, i, c, s, a, b
    Dim lines() As String
    Dim it As String
    
    If Len(tmp) > 0 Then
        it = tmp
    Else
        it = asString()
    End If
    
    my = ""
    For i = 1 To Len(it)
        a = Asc(Mid(it, i, 1))
        c = Hex(a)
        c = IIf(Len(c) = 1, "0" & c, c)
        b = b & IIf(a > 33 And a < 123, Chr(a), ".")
        my = my & c & " "
        If i Mod 16 = 0 Then
            push lines(), my & "  [" & b & "]"
            my = Empty
            b = Empty
        End If
    Next
    
    If Len(b) > 0 Then
        If Len(my) < 48 Then
            my = my & String(48 - Len(my), " ")
        End If
        If Len(b) < 16 Then
             b = b & String(16 - Len(b), " ")
        End If
        push lines(), my & "  [" & b & "]"
    End If
        
    If Len(it) < 16 Then
        hexdump = my & "  [" & b & "]"
    Else
        hexdump = Join(lines, vbCrLf)
    End If
    
    
End Function




'----------------------------------[   reads and writes    ]----------------------------------

Function ReadLong(index As Long, Optional ByRef hadErr As Boolean) As Long
    hadErr = True
    If isEmpty Then Exit Function
    If index < LBound(base) Then Exit Function
    If index + 3 > UBound(base) Then Exit Function
    Dim retval As Long
    CopyMemory VarPtr(retval), VarPtr(base(index)), 4
    ReadLong = retval
    hadErr = False
End Function

Function ReadShort(index As Long, Optional ByRef hadErr As Long) As Integer
    hadErr = True
    If isEmpty Then Exit Function
    If index < LBound(base) Then Exit Function
    If index + 1 > UBound(base) Then Exit Function
    Dim retval As Integer
    CopyMemory VarPtr(retval), VarPtr(base(index)), 2
    ReadShort = retval
    hadErr = False
End Function

Function WriteLong(index As Long, value As Long) As Boolean
    If isEmpty Then Exit Function
    If index < LBound(base) Then Exit Function
    If index + 3 > UBound(base) Then Exit Function
    CopyMemory VarPtr(base(index)), VarPtr(value), 4
    WriteLong = True
End Function

Function WriteShort(index As Long, value As Integer) As Boolean
    If isEmpty Then Exit Function
    If index < LBound(base) Then Exit Function
    If index + 1 > UBound(base) Then Exit Function
    CopyMemory VarPtr(base(index)), VarPtr(value), 2
    WriteShort = True
End Function





'----------------------------------[   pushs and pops    ]----------------------------------
Function pushByte(bb As Byte)

    If isEmpty Then
        ReDim base(0)
    Else
        ReDim Preserve base(UBound(base) + 1)
    End If
    
    base(UBound(base)) = bb
    
End Function

Function pushShort(s As Integer)
    If isEmpty Then
        ReDim base(1)
    Else
        ReDim Preserve base(UBound(base) + 2)
    End If
    WriteShort UBound(base) - 1, s
End Function

Function pushLong(l As Long)
    If isEmpty Then
        ReDim base(3)
        WriteLong 0, l
    Else
        ReDim Preserve base(UBound(base) + 4)
        WriteLong UBound(base) - 3, l
    End If
End Function

Function pushStr(s As String)
    If isEmpty Then
        LoadFromStr s
    Else
        Dim tmp() As Byte
        tmp = StrConv(s, vbFromUnicode, LANG_US)
        pushBlock tmp()
    End If
End Function

Function pushHexStr(s As String) As Boolean
    Dim b() As Byte
    b() = internal_LoadFromHexStr(s)
    If AryIsEmpty(b) Then Exit Function
    pushBlock b()
End Function


'tested ok
Function pushBlock(b() As Byte)
    If isEmpty Then
        base() = b()
    Else
        Dim i As Long
        i = UBound(base)
        ReDim Preserve base(i + UBound(b) + 1)
        CopyMemory VarPtr(base(i + 1)), VarPtr(b(0)), UBound(b) + 1
    End If
End Function

Function readBlock(index As Long, size As Long) As Byte()
    Dim b() As Byte
    If index + size > UBound(base) + 1 Then Exit Function
    ReDim b(size - 1)
    CopyMemory VarPtr(b(0)), VarPtr(base(index)), size
    readBlock = b()
End Function

Function popBlock(size As Long) As Byte()
    On Error Resume Next
    Dim b() As Byte
    
    If isEmpty Then Exit Function
    
    b() = readBlock(UBound(base) + 1 - size, size)
    
    If UBound(base) = size Then
        Erase base
    Else
        ReDim Preserve base(UBound(base) - size)
    End If
    
    popBlock = b()
    
End Function

Function popByte() As Byte
    On Error Resume Next
    popByte = base(UBound(base))
    If UBound(base) = 0 Then
        Erase base
    Else
        ReDim Preserve base(UBound(base) - 1)
    End If
End Function

'these ones can throw errors..
Function popLong() As Long
    popLong = ReadLong(UBound(base) - 3)
    If UBound(base) = 3 Then
        Erase base
    Else
        ReDim Preserve base(UBound(base) - 4)
    End If
End Function

Function popShort()
    popShort = ReadShort(UBound(base) - 1)
    If UBound(base) = 1 Then
        Erase base
    Else
        ReDim Preserve base(UBound(base) - 2)
    End If
End Function









'--------------------------------[    private functions below    ]-----------------------------

Private Function IsHex(it)
    On Error GoTo out
      IsHex = Chr(Int("&H" & it))
    Exit Function
out:  IsHex = Empty
End Function

Private Function sHex(x) As String
    sHex = Right("0" & Hex(x), 2)
End Function

Private Sub push(ary, value) 'this modifies parent ary object
    On Error GoTo init
    x = UBound(ary) '<-throws Error If Not initalized
    ReDim Preserve ary(UBound(ary) + 1)
    ary(UBound(ary)) = value
    Exit Sub
init:     ReDim ary(0): ary(0) = value
End Sub

Private Function AryIsEmpty(ary) As Boolean
  On Error GoTo oops
    i = UBound(ary)  '<- throws error if not initalized
    AryIsEmpty = False
  Exit Function
oops: AryIsEmpty = True
End Function

 
Private Function FileExists(path) As Boolean
  On Error GoTo hell
    
  If Len(path) = 0 Then Exit Function
  If Right(path, 1) = "\" Then Exit Function
  If Dir(path, vbHidden Or vbNormal Or vbReadOnly Or vbSystem) <> "" Then FileExists = True
  
  Exit Function
hell: FileExists = False
End Function

'space seperated or run together hex string, no other white space..
Private Function internal_LoadFromHexStr(ByVal x As String) As Byte()
    Dim tmp() As String
    Dim i As Long
    Dim b() As Byte
    
    On Error GoTo hell

    If InStr(x, vbTab) > 0 Then x = Replace(x, vbTab, Empty)
    If InStr(x, vbCr) > 0 Then x = Replace(x, vbCr, Empty)
    If InStr(x, vbLf) > 0 Then x = Replace(x, vbLf, Empty)
    
    x = Trim(x)
    i = InStr(x, " ")
    
    '90 90 eb 15
    If i > 0 And i < 5 Then
        tmp = Split(x, " ")
    Else
        
        '9090eb15
        If InStr(x, " ") > 0 Then x = Replace(x, " ", Empty)
        If Len(x) Mod 2 <> 0 Then Exit Function
        
        For i = 1 To Len(x) Step 2
            push tmp, Mid(x, i, 2)
        Next
        
    End If
        
    ReDim b(UBound(tmp))
    
    For i = 0 To UBound(tmp)
        b(i) = CInt("&h" & tmp(i))
    Next

    internal_LoadFromHexStr = b()
        
hell:
    
End Function
