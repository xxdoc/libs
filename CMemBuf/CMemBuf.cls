VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CMemBuf"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Base 1

'the idea of this class is to provide an easy way to read/write various data types
'to a memory buffer using a file type pointer in the same way we can do it to disk with Get/Put
'the main functions are read/write and they will behave appropiratly based on the data type passed
'in to the data function.
'
'--> It is expected that the variables passed in will be strongly typed and not variants. <--
'
' We need to know their size to write them to a byte buffer, we accept a variant arg so we can accept
' any type..but variant itself is not accepted, except for the writeStr function which converts everything
' to a string as expected
'

'todo:
'     strong type primitives ReadByte/WriteByte ?
'     resize ?
'
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSrc As Any, ByVal ByteLen As Long)

Private buf() As Byte 'can not be public vb sayz
Private pointer As Long
Public lastErr As String
Public optBase0 As Boolean
Public optRaiseErr As Boolean

Enum mb_StrTypes
    mb_cstr
    mb_bstr
    mb_fixedSize
End Enum

Public Enum hexOutFormats
    hoDump
    hoSpaced
    hoHexOnly
End Enum

Private Const LANG_US = &H409

Sub clear()
    pointer = 1
    lastErr = Empty
    Erase buf
End Sub

Sub truncate(size As Long)
    If optBase0 Then size = size + 1
    If size < 1 Then
        Me.clear
    Else
        ReDim Preserve buf(size)
        If pointer > size Then pointer = size
    End If
End Sub

Property Get Buffer() As Byte()

    Dim b() As Byte
   
    If optBase0 Then
       ReDim b(0 To UBound(buf) - 1)
       CopyMemory ByVal VarPtr(b(0)), ByVal VarPtr(buf(1)), UBound(buf)
       Buffer = b
    Else
       Buffer = buf()
    End If
    
End Property

Property Let Buffer(b() As Byte)
    
    Dim ub As Long, lb As Long, ab As Long
    
    Me.clear
    If AryIsEmpty(b) Then Exit Property
        
    If LBound(b) = 1 Then
        buf() = b()
    Else
        ub = UBound(b)
        lb = LBound(b)
        ab = ub - lb + 1
        'Debug.Print "Array bounds: " & ub & "," & lb & "," & ab
        ReDim buf(1 To ab)
        CopyMemory ByVal VarPtr(buf(1)), ByVal VarPtr(b(lb)), ab
    End If
    
    
End Property

Property Get curOffset() As Long
    curOffset = pointer
    If optBase0 Then curOffset = curOffset - 1
End Property

Property Let curOffset(ByVal X As Long)
    
    If optBase0 Then X = X + 1
    
    If X < 1 Then
        lastErr = "Invalid offset"
        Err.Raise 122, "CMemBuf.curOffset", lastErr
    End If
    
    pointer = X
    
End Property

Private Sub Class_Initialize()
    pointer = 1
End Sub

Function ensureSize(ByVal size As Long, Optional offset As Long = -1) As Boolean
    
    Dim newSize As Long
    Dim start As Long
    
    'do not change pointer here, we are only testing and allocating
    start = pointer
    If offset <> -1 Then start = offset
    If start < 1 Then Exit Function
    
    If AryIsEmpty(buf) Then
        ReDim buf(size)
    Else
        newSize = start + size - 1
        If newSize > UBound(buf) Then
            ReDim Preserve buf(newSize)
        End If
    End If
    
    ensureSize = True
    
End Function

Function sizeInBounds(size As Long, Optional offset As Long = -1) As Boolean
    
    Dim start As Long
    
    'do not change pointer here, we are only testing
    start = pointer
    If offset <> -1 Then start = offset
    If start < 0 Then Exit Function
    If AryIsEmpty(buf) Then Exit Function
    
    If size + start <= UBound(buf) + 1 Then sizeInBounds = True

End Function

Function HexDump(Optional ByVal Length As Long = -1, Optional ByVal startAt As Long = 1, Optional hexFormat As hexOutFormats = hoDump) As String
    HexDump = HexDump_Internal(buf, Length, startAt, hexFormat)
End Function

Function readStr(Optional offset As Long = -1, Optional readType As mb_StrTypes = mb_cstr, Optional fixedWidth As Long = 0) As String
    
    Dim sz As Integer
    Dim b() As Byte, bb As Byte, tmp As String
    
    lastErr = Empty
    If offset <> -1 Then curOffset = offset 'raise error on bad value..
    
    If readType = mb_bstr Then
        If Not read(sz) Then Exit Function
        ReDim b(sz)
        If Not read(b) Then Exit Function
        readStr = StrConv(b, vbUnicode, &H409)
    ElseIf readType = mb_cstr Then
        Do While pointer <= UBound(buf)
            bb = buf(pointer)
            pointer = pointer + 1
            If bb <> 0 Then
                tmp = tmp & Chr(bb)
            Else
                Exit Do
            End If
        Loop
        readStr = tmp
    ElseIf readType = mb_fixedSize Then
        If fixedWidth = 0 Then
            lastErr = "Invalid Fixed width"
            Exit Function
        End If
        ReDim b(fixedWidth)
        If Not read(b) Then Exit Function
        readStr = StrConv(b, vbUnicode, &H409)
    End If
    
End Function


Function writeStr(data As Variant, Optional offset As Long = -1, Optional writeType As mb_StrTypes = mb_cstr) As Boolean
    
    If Not optRaiseErr Then
        On Error GoTo hell
    End If
    
    Dim s As String, n As String, bb() As Byte, X, sz As Integer
    
    lastErr = Empty
    If offset <> -1 Then curOffset = offset 'raise error on bad value..
   
    writeStr = True
    If IsArray(data) Then
        For Each X In data
            GoSub writeStr_internal
        Next
    Else
        X = data
        GoSub writeStr_internal
    End If
    
Exit Function
    
writeStr_internal: 'gosub is ugly but we dont want to have two copies of this code to keep synced on bugfix doing an internal fx is ok too but equally sloppy
    s = X
    If writeType = mb_cstr Then s = s & Chr(0)
    bb = StrConv(s, vbFromUnicode)
    sz = UBound(bb) + 1
    If writeType = mb_bstr Then write_ CInt(sz)
    ensureSize sz
    CopyMemory ByVal VarPtr(buf(pointer)), ByVal VarPtr(bb(LBound(bb))), sz
    pointer = pointer + sz
    Return
    
Exit Function
hell:
     writeStr = False
     lastErr = "WriteStr Err: " & Err.Description
End Function

Function writeFromMem(memAddress As Long, size As Long, Optional offset As Long = -1) As Boolean
    Dim b() As Byte
    ReDim b(1 To size)
    CopyMemory ByVal VarPtr(b(1)), ByVal memAddress, size
    writeFromMem = write_(b, offset)
End Function

'you better know what your doing or crashity crash in mysterious ways at some point in future...
Function saveToMem(memAddress As Long, Optional size As Long, Optional offset As Long = -1) As Boolean
    
    Dim o As Long
    
    o = 1
    If offset <> -1 Then    'its a partial copy of some latter portion of the buf
        curOffset = offset 'performs validation
        o = curOffset
        If size + o > UBound(buf) + 1 Then Exit Function
    Else
        If size > UBound(buf) + 1 Then Exit Function
    End If
    
    CopyMemory ByVal memAddress, ByVal VarPtr(buf(o)), size
    saveToMem = True
    
End Function

Function toFile(ByVal fPath As String, Optional fileOffset As Long = 1) As Boolean
    If Not optRaiseErr Then On Error Resume Next
    Dim f As Long
    f = FreeFile
    Open fPath For Binary As f
    If Err.Number <> 0 Then Exit Function
    Put f, fileOffset, buf()
    Close f
    toFile = (Err.Number = 0)
End Function

Function fromFile(ByVal fPath As String, Optional fileOffset As Long = 1, Optional ByVal dataSize As Long = -1) As Boolean
    If Not optRaiseErr Then On Error Resume Next
    Dim f As Long
    
    Me.clear
    If dataSize = -1 Then dataSize = FileLen(fPath)
    If fileOffset <> 1 Then dataSize = dataSize - fileOffset
    
    If Err.Number <> 0 Then
        lastErr = Err.Description
        Exit Function
    End If
    
    f = FreeFile
    Open fPath For Binary As f
    If Err.Number <> 0 Then Exit Function
    
    ReDim buf(1 To dataSize)
    Get f, fileOffset, buf
    Close f
    
    If Err.Number <> 0 Then lastErr = Err.Description
    fromFile = (Err.Number = 0)
    
End Function

Function write_(data As Variant, Optional offset As Long = -1) As Boolean
    
    Dim sz As Long, n As String, X
    Dim i As Integer, b As Byte, l As Long, s As String, c As Currency, d As Double
    Dim ii() As Integer, bb() As Byte, ll() As Long, ss() As String, cc() As Currency, dd() As Double
    Dim lb As Long, ub As Long, ab As Long
    
    If Not optRaiseErr Then On Error GoTo hell
    
    If offset <> -1 Then curOffset = offset 'raise error on bad value..
    
    
    lastErr = Empty
    write_ = True
    n = TypeName(data)
    'sz = LenB(data) 'this only gives the size of the variant..not contained data type...
    
    If InStr(n, "()") > 0 Then 'isArray
        ub = UBound(data)
        lb = LBound(data)
        ab = ub - lb + 1
        'Debug.Print "Array bounds: " & ub & "," & lb & "," & ab
    End If
    
    
    If n = "Integer" Then
        sz = 2
        i = data
        ensureSize sz
        CopyMemory ByVal VarPtr(buf(pointer)), ByVal VarPtr(i), sz
        pointer = pointer + sz
       
        
    ElseIf n = "Byte" Then
        sz = 1
        b = data
        ensureSize sz
        CopyMemory ByVal VarPtr(buf(pointer)), ByVal VarPtr(b), sz
        pointer = pointer + sz
        
    ElseIf n = "Long" Then
        sz = 4
        l = data
        ensureSize sz
        CopyMemory ByVal VarPtr(buf(pointer)), ByVal VarPtr(l), sz
        pointer = pointer + sz
        
    ElseIf n = "Currency" Then
        c = data
        sz = 8
        ensureSize sz
        CopyMemory ByVal VarPtr(buf(pointer)), ByVal VarPtr(c), sz
        pointer = pointer + sz
        
    ElseIf n = "Double" Then
        d = data
        sz = 8
        ensureSize sz
        CopyMemory ByVal VarPtr(buf(pointer)), ByVal VarPtr(d), sz
        pointer = pointer + sz
        
    ElseIf n = "Byte()" Then
        bb = data
        sz = ab * 1
        ensureSize sz
        CopyMemory ByVal VarPtr(buf(pointer)), ByVal VarPtr(bb(lb)), sz
        pointer = pointer + sz
        
    ElseIf n = "Integer()" Then
        ii = data
        sz = ab * 2
        ensureSize sz
        CopyMemory ByVal VarPtr(buf(pointer)), ByVal VarPtr(ii(lb)), sz
        pointer = pointer + sz
    
    ElseIf n = "Long()" Then
        ll = data
        sz = ab * 4
        ensureSize sz
        CopyMemory ByVal VarPtr(buf(pointer)), ByVal VarPtr(ll(lb)), sz
        pointer = pointer + sz
    
    ElseIf n = "Currency()" Then
        cc = data
        sz = ab * 8
        ensureSize sz
        CopyMemory ByVal VarPtr(buf(pointer)), ByVal VarPtr(cc(lb)), sz
        pointer = pointer + sz
        
    ElseIf n = "Double()" Then
        dd = data
        sz = ab * 8
        ensureSize sz
        CopyMemory ByVal VarPtr(buf(pointer)), ByVal VarPtr(dd(lb)), sz
        pointer = pointer + sz
        
    Else
        lastErr = "Unsupported type: " & n
        write_ = False
        If optRaiseErr Then Err.Raise 2323, "CMembuf.write", lastErr
    End If
    
Exit Function
hell:
    write_ = False
    lastErr = Err.Description
    Debug.Print "Write_ err w/ " & n & " - " & Err.Description
End Function



Function read(ByRef data As Variant, Optional offset As Long = -1) As Boolean
    
    Dim sz As Long, n As String, X
    Dim i As Integer, b As Byte, l As Long, s As String, c As Currency, d As Double
    Dim ii() As Integer, bb() As Byte, ll() As Long, ss() As String, cc() As Currency, dd() As Double
    Dim lb As Long, ub As Long, ab As Long
    
    If Not optRaiseErr Then On Error GoTo hell
    
    If offset <> -1 Then curOffset = offset 'raise error on bad value..
    
    lastErr = Empty
    read = True
    n = TypeName(data)
    'sz = LenB(data) 'this only gives the size of the variant..not contained data type...
    
    If InStr(n, "()") > 0 Then 'isArray
        ub = UBound(data)
        lb = LBound(data)
        ab = ub - lb + 1
        'Debug.Print "Array bounds: " & ub & "," & lb & "," & ab
    End If
    
    
    If n = "Integer" Then
        sz = 2
        If Not sizeInBounds(sz) Then GoTo outOfBounds
        CopyMemory ByVal VarPtr(i), ByVal VarPtr(buf(pointer)), sz
        pointer = pointer + sz
        data = i
        
    ElseIf n = "Byte" Then
        sz = 1
        If Not sizeInBounds(sz) Then GoTo outOfBounds
        CopyMemory ByVal VarPtr(b), ByVal VarPtr(buf(pointer)), sz
        pointer = pointer + sz
        data = b
        
    ElseIf n = "Long" Then
        sz = 4
        If Not sizeInBounds(sz) Then GoTo outOfBounds
        CopyMemory ByVal VarPtr(l), ByVal VarPtr(buf(pointer)), sz
        pointer = pointer + sz
        data = l
   
    ElseIf n = "Currency" Then
        sz = 8
        If Not sizeInBounds(sz) Then GoTo outOfBounds
        CopyMemory ByVal VarPtr(c), ByVal VarPtr(buf(pointer)), sz
        pointer = pointer + sz
        data = c
        
    ElseIf n = "Double" Then
        sz = 8
        If Not sizeInBounds(sz) Then GoTo outOfBounds
        CopyMemory ByVal VarPtr(d), ByVal VarPtr(buf(pointer)), sz
        pointer = pointer + sz
        data = d
        
    ElseIf n = "Byte()" Then
        sz = ab * 1
        If Not sizeInBounds(sz) Then GoTo outOfBounds
        ReDim bb(lb To ub)
        CopyMemory ByVal VarPtr(bb(lb)), ByVal VarPtr(buf(pointer)), sz
        pointer = pointer + sz
        data = bb
        
    ElseIf n = "Integer()" Then
        sz = ab * 2
        If Not sizeInBounds(sz) Then GoTo outOfBounds
        ReDim ii(lb To ub)
        CopyMemory ByVal VarPtr(ii(lb)), ByVal VarPtr(buf(pointer)), sz
        pointer = pointer + sz
        data = ii
        
    ElseIf n = "Long()" Then
        sz = ab * 4
        If Not sizeInBounds(sz) Then GoTo outOfBounds
        ReDim ll(lb To ub)
        CopyMemory ByVal VarPtr(ll(lb)), ByVal VarPtr(buf(pointer)), sz
        pointer = pointer + sz
        data = ll
        
    ElseIf n = "Currency()" Then
        sz = ab * 8
        If Not sizeInBounds(sz) Then GoTo outOfBounds
        ReDim cc(lb To ub)
        CopyMemory ByVal VarPtr(cc(lb)), ByVal VarPtr(buf(pointer)), sz
        pointer = pointer + sz
        data = cc
        
    ElseIf n = "Double()" Then
        sz = ab * 8
        If Not sizeInBounds(sz) Then GoTo outOfBounds
        ReDim dd(lb To ub)
        CopyMemory ByVal VarPtr(dd(lb)), ByVal VarPtr(buf(pointer)), sz
        pointer = pointer + sz
        data = dd
        
    Else
        lastErr = "Unsupported type: " & n
        read = False
        If optRaiseErr Then Err.Raise 2323, "CMemBuf.read", lastErr
    End If
    
Exit Function
hell:
    read = False
    lastErr = Err.Description
    Debug.Print "Read err w/ " & n & " - " & Err.Description
    
Exit Function

outOfBounds:
    read = False
    lastErr = "Read Out of bounds. Offset = " & pointer & " Size: " & sz & " Buffer Size: " & UBound(buf) + 1
    Debug.Print lastErr
    
End Function

'are these worth it?
Function ReadLong(ByRef X As Long, Optional offset As Long = -1) As Boolean
    ReadLong = read(X, offset)
End Function

Function ReadByte(ByRef X As Byte, Optional offset As Long = -1) As Boolean
    ReadByte = read(X, offset)
End Function

Function WriteLong(ByVal X As Long, Optional offset As Long = -1) As Boolean
    WriteLong = write_(X, offset)
End Function

Public Function rc4(ByVal password As Variant, Optional overWriteBuf As Boolean = True) As Byte()
    
    If Not optRaiseErr Then On Error GoTo hell
    
    Dim RB(0 To 255) As Integer, X As Long, Y As Long, Z As Long
    Dim plen As Long, key() As Byte, ByteArray() As Byte, temp As Byte
    
    lastErr = Empty
    ByteArray() = buf()
    
    If TypeName(password) = "Byte()" Then
        key() = password
        If UBound(key) > 255 Then ReDim Preserve key(255)
    Else
        If Len(password) = 0 Then password = "xxx"

        If Len(password) > 256 Then
            key() = StrConv(Left$(CStr(password), 256), vbFromUnicode, LANG_US)
        Else
            key() = StrConv(CStr(password), vbFromUnicode, LANG_US)
        End If
    End If
    
    plen = UBound(key) + 1
 
    'Debug.Print "key=" & HexDump(Key)
    'Debug.Print "data=" & HexDump(ByteArray)
    
    For X = 0 To 255
        RB(X) = X
    Next X
    
    X = 0: Y = 0: Z = 0
    For X = 0 To 255
        Y = (Y + RB(X) + key(X Mod plen)) Mod 256
        temp = RB(X)
        RB(X) = RB(Y)
        RB(Y) = temp
    Next X
    
    X = 0: Y = 0: Z = 0
    For X = LBound(ByteArray) To UBound(ByteArray)
        Y = (Y + 1) Mod 256
        Z = (Z + RB(Y)) Mod 256
        temp = RB(Y)
        RB(Y) = RB(Z)
        RB(Z) = temp
        ByteArray(X) = ByteArray(X) Xor (RB((RB(Y) + RB(Z)) Mod 256))
    Next X
    
    If overWriteBuf Then buf = ByteArray
    rc4 = ByteArray
    
Exit Function
hell:
    lastErr = "rc4 err: " & Err.Description

End Function


'library code below
Private Function HexDump_Internal(bAryOrStrData, Optional ByVal Length As Long = -1, Optional ByVal startAt As Long = 1, Optional hexFormat As hexOutFormats = hoDump) As String
    Dim s() As String, chars As String, tmp As String
    On Error Resume Next
    Dim ary() As Byte
    Dim offset As Long
    Const LANG_US = &H409
    Dim i As Long, tt, h, X
    Dim hexOnly As Long
    
    offset = 0
    If hexFormat <> hoDump Then hexOnly = 1
    
    If TypeName(bAryOrStrData) = "Byte()" Then
        ary() = bAryOrStrData
    Else
        ary = StrConv(CStr(bAryOrStrData), vbFromUnicode, LANG_US)
    End If
    
    If startAt < 1 Then startAt = 1
    If Length < 1 Then Length = -1
    
    While startAt Mod 16 <> 0
        startAt = startAt - 1
    Wend
    
    startAt = startAt + 1
    
    chars = "   "
    For i = startAt To UBound(ary) + 1
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
        If Length <> -1 Then
            Length = Length - 1
            If Length = 0 Then Exit For
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
    
    HexDump_Internal = Join(s, vbCrLf)
    
    If hexOnly <> 0 Then
        If hexFormat = hoHexOnly Then HexDump_Internal = Replace(HexDump_Internal, " ", "")
        HexDump_Internal = Replace(HexDump_Internal, vbCrLf, "")
    End If
    
End Function


Sub push(ary, Value) 'this modifies parent ary object
    On Error GoTo Init
    Dim X, lb As Long
       
    X = UBound(ary)
    ReDim Preserve ary(X + 1)
    
    If IsObject(Value) Then
        Set ary(X + 1) = Value
    Else
        ary(X + 1) = Value
    End If
    
    Exit Sub
Init:
    lb = optionBase 'in case the module uses option base 1 (as this one does)
    ReDim ary(lb)
    If IsObject(Value) Then
        Set ary(lb) = Value
    Else
        ary(lb) = Value
    End If
End Sub

Private Function optionBase() As Byte
  optionBase = LBound(Array())
End Function

Function AryIsEmpty(ary) As Boolean
  On Error GoTo oops
  Dim X
    If IsObject(UBound(ary)) Then AryIsEmpty = False
    'x = UBound(ary)
  Exit Function
oops: AryIsEmpty = True
End Function


