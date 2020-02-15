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
' while we can not support UDTs file vb get/put you could use my CMemStruct class with it

'todo:
'     support optionBase0 for read/write user offsets?
'     strong type primitives ReadByte/WriteByte ?
'     option raise err on fail?
'     toFile/fromFile
'     reset/resize
'     encrypt/decrypt
'
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSrc As Any, ByVal ByteLen As Long)

Private buf() As Byte 'can not be public vb sayz
Private pointer As Long
Public lastErr As String
'Public optBase0 As Boolean
'public optRaiseErr as boolean

Enum mb_StrTypes
    mb_cstr
    mb_bstr
    mb_fixedSize
End Enum

Property Get Buffer() As Byte()
   Buffer = buf()
End Property

Property Let Buffer(b() As Byte)
    buf() = b()
End Property

Property Get curOffset() As Long
    curOffset = pointer
End Property

Property Let curOffset(x As Long)
    If x < 1 Then
        lastErr = "Offset can not be < 1"
        Err.Raise 122, "CMemBuf", lastErr
    End If
    pointer = x
End Property

Private Sub Class_Initialize()
    pointer = 1
    'OptionBase1 = True
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
    HexDump = Module1.HexDump(buf, Length, startAt, hexFormat)
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
    
    On Error GoTo hell
    
    Dim s As String, n As String, bb() As Byte, x, sz As Integer
    
    lastErr = Empty
    If offset <> -1 Then curOffset = offset 'raise error on bad value..
   
    writeStr = True
    If IsArray(data) Then
        For Each x In data
            GoSub writeStr_internal
        Next
    Else
        x = data
        GoSub writeStr_internal
    End If
    
Exit Function
    
writeStr_internal: 'gosub is ugly but we dont want to have two copies of this code to keep synced on bugfix doing an internal fx is ok too but equally sloppy
    s = x
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


Function write_(data As Variant, Optional offset As Long = -1) As Boolean
    
    Dim sz As Long, n As String, x
    Dim i As Integer, b As Byte, l As Long, s As String, c As Currency, d As Double
    Dim ii() As Integer, bb() As Byte, ll() As Long, ss() As String, cc() As Currency, dd() As Double
    Dim lb As Long, ub As Long, ab As Long
    
    On Error GoTo hell
    
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
        
'    ElseIf n = "String" Then
'        s = data
'        bb = StrConv(s, vbFromUnicode)
'        sz = UBound(bb) + 1
'        ensureSize sz
'        CopyMemory ByVal VarPtr(buf(pointer)), ByVal VarPtr(bb(LBound(bb))), sz
'        pointer = pointer + sz
        
        
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
        
'    ElseIf n = "String()" Then
'        For Each x In data
'            s = x
'            bb = StrConv(s, vbFromUnicode)
'            sz = UBound(bb) + 1
'            ensureSize sz
'            CopyMemory ByVal VarPtr(buf(pointer)), ByVal VarPtr(bb(LBound(bb))), sz
'            pointer = pointer + sz
'        Next
        
    Else
        lastErr = "Unsupported type: " & n
        write_ = False
    End If
    
Exit Function
hell:
    write_ = False
    lastErr = Err.Description
    Debug.Print "Write_ err w/ " & n & " - " & Err.Description
End Function



Function read(ByRef data As Variant, Optional offset As Long = -1) As Boolean
    
    Dim sz As Long, n As String, x
    Dim i As Integer, b As Byte, l As Long, s As String, c As Currency, d As Double
    Dim ii() As Integer, bb() As Byte, ll() As Long, ss() As String, cc() As Currency, dd() As Double
    Dim lb As Long, ub As Long, ab As Long
    
    On Error GoTo hell
    
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
        
'    ElseIf n = "String" Then
'        s = data
'        sz = UBound(bb) + 1
'        ensureSize sz
'        CopyMemory ByVal VarPtr(buf(pointer)), ByVal VarPtr(bb(LBound(bb))), sz
'        pointer = pointer + sz
'        data = StrConv(bb, vbUnicode, &H409)
        
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
        
'    ElseIf n = "String()" Then
'        For Each x In data
'            s = x
'            bb = StrConv(s, vbFromUnicode)
'            sz = UBound(bb) + 1
'            ensureSize sz
'            CopyMemory ByVal VarPtr(buf(pointer)), ByVal VarPtr(bb(LBound(bb))), sz
'            pointer = pointer + sz
'        Next
'
    Else
        lastErr = "Unsupported type: " & n
        read = False
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
Function ReadLong(ByRef x As Long, Optional offset As Long = -1) As Boolean
    ReadLong = read(x, offset)
End Function

Function ReadByte(ByRef x As Byte, Optional offset As Long = -1) As Boolean
    ReadByte = read(x, offset)
End Function

Function WriteLong(ByVal x As Long, Optional offset As Long = -1) As Boolean
    WriteLong = write_(x, offset)
End Function
