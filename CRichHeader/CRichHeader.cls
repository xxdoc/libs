VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CRichHeader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public ErrorMsg As String
Public entries As Collection 'of CToolID
Public DanS As Long
Public Rich As Long
Public size As Long
Public clearData 'variant as byte()
Public checkSum As Long
Public calcChecksum As Long
Public exists As Boolean

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSrc As Any, ByVal ByteLen As Long)

Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" ( _
       ByVal lpPrevWndFunc As Long, _
       ByVal hWnd As Long, _
       ByVal Msg As Long, _
       ByVal wParam As Long, _
       ByVal lParam As Long _
) As Long

Const asm_rol32 = "\x55\x8B\xEC\x56\x8B\x4D\x10\x83\xE1\x1F\x8B\x45\x0C" & _
                  "\xD3\xE0\x8B\x4D\x10\x83\xE1\x1F\xBA\x20\x00\x00\x00" & _
                  "\x2B\xD1\x83\xE2\x1F\x8B\x75\x0C\x8B\xCA\xD3\xEE\x0B" & _
                  "\xC6\x5E\x5D\xC2\x10\x00"

Const asm_add32 = "\x55\x8B\xEC\x8B\x45\x0C\x03\x45\x10\x5d\xC2\x10\x00"

Dim asm() As Byte
Dim asm2() As Byte
Const LANG_US = &H409

Enum rhcdModes
    rhcd_Readable
    rhcd_Binary
    rhcd_Hex
End Enum
     
Private Sub initVars()
    ErrorMsg = Empty
    clearData = Empty
    Set entries = New Collection 'of CToolID
    DanS = 0
    Rich = 0
    size = 0
    checkSum = 0
    calcChecksum = 0
    exists = False
End Sub

'yara pe.rich_signature.version(verion,[toolId]) compatiable
Function version(compiler As Long, Optional toolID As rhToolID) As Boolean
    
    Dim t As CToolId
    
    For Each t In entries
        If compiler = t.compiler Then
            If toolID > 0 Then
                If toolID = t.toolID Then
                    version = True
                    Exit Function
                End If
            Else
                version = True
                Exit Function
            End If
        End If
    Next
                
End Function

'yara pe.rich_signature.toolid(toolid, [version]) compatiable
Function toolID(tid As rhToolID, Optional compiler As Long) As Boolean
    
    Dim t As CToolId
    
    For Each t In entries
        If tid = t.toolID Then
            If compiler > 0 Then
                If compiler = t.compiler Then
                    toolID = True
                    Exit Function
                End If
            Else
                toolID = True
                Exit Function
            End If
        End If
    Next
                
End Function



Function dump() As String

    Dim tmp() As String, t As CToolId
    
    If Len(ErrorMsg) > 0 Then
        dump = "ErrorMsg: " & ErrorMsg
        Exit Function
    End If
    
    push tmp, "offset DanS:" & DanS
    push tmp, "offset Rich:" & Rich
    push tmp, "Size: " & size
    
    If checkSumOk Then
        push tmp, "checkSum: 0x" & Hex(checkSum) & " Matches!"
    Else
        push tmp, "checkSum: 0x" & Hex(checkSum) & " Failed!"
        push tmp, "calculatedCheckSum: 0x" & Hex(calcChecksum)
    End If
    
    push tmp, ""
    push tmp, "clearData: " & strClearData(rhcd_Readable)
    push tmp, ""
    
    push tmp, entries.Count & " entries"
    push tmp, String(50, "-")
    
    For Each t In entries
        push tmp, t.dump(vbCrLf, True) & vbCrLf
    Next
     
    dump = Join(tmp, vbCrLf)
    
End Function

Property Get checkSumOk() As Boolean
    checkSumOk = (checkSum = calcChecksum)
End Property

Property Get strClearData(Optional mode As rhcdModes = rhcd_Binary) As String
    
    Dim b() As Byte, i, x
    
    If AryIsEmpty(clearData) Then Exit Property
    
    b() = clearData
    
    If mode = rhcd_Readable Then
        For i = 0 To UBound(b)
            If b(i) > 32 And b(i) < 127 Then
                x = x & Chr(b(i))
            Else
                x = x & "\x" & Right("00" & Hex(b(i)), 2)
            End If
        Next
    ElseIf mode = rhcd_Binary Then
        x = StrConv(b, vbUnicode, LANG_US)
    ElseIf mode = rhcd_Hex Then
        For i = 0 To UBound(b)
            x = x & "\x" & Right("00" & Hex(b(i)), 2)
        Next
    End If
        
    strClearData = x
        
End Property

' Text1 = getRich("D:\_code\libs\pe_lib2\_sppe2.dll")

'rich.py output
    ' Dans offset: 128
    ' 7299,0x000e,0x00000001,prodidMasm613,<unknown>,00.00
    ' 8041,0x0009,0x00000019,prodidUtc12_Basic,<unknown>,00.00
    ' 8169,0x000d,0x00000001,prodidVisualBasic60,<unknown>,00.00
    ' 8168,0x0004,0x00000001,prodidLinker600,<unknown>,00.00
    ' Checksums match! (0x5a8dff07)

'vb output
    'Checksum is 0x5A8DFF07
    'Dans file offset: 128
    '7299 E 1 prodidMasm613 <unknown> (00.00)
    '8041 9 19 prodidUtc12_Basic <unknown> (00.00)
    '8169 D 1 prodidVisualBasic60 <unknown> (00.00)
    '8168 4 1 prodidLinker600 <unknown> (00.00)
    'Checksums match!


Private Sub Class_Initialize()
    asm() = toBytes(asm_rol32)
    asm2() = toBytes(asm_add32)
End Sub


Function Load(fpath) As Boolean

   Dim pth As String, f As Long, e_lfanew As Long, i As Long
   Dim b() As Byte, tool As CToolId
   Dim sig As String * 2, sig4 As String * 4
   Dim tmp As Long, uPack() As Long
   
   Const SIZE_DOS_HEADER = &H40
   Const POS_E_LFANEW = &H3C
   
   On Error GoTo cleanup
   
   initVars

   If Not FileExists(fpath) Then
        ErrorMsg = "File not found"
        Exit Function
   End If
    
   f = FreeFile
   Open fpath For Binary Access Read As f
   
   Get f, , sig
   If sig <> "MZ" Then
        ErrorMsg = "MZ header not found"
        GoTo cleanup
   End If
   
   Get f, POS_E_LFANEW + 1, e_lfanew
   Get f, e_lfanew + 1, sig
   
   If sig <> "PE" Then
        ErrorMsg = "PE header not found"
        GoTo cleanup
   End If
   
   'IMPORTANT: Do not assume the data to start at 0x80, this is not always
   ' the case (modified DOS stub). Instead, start searching backwards for
   ' 'Rich', stopping at the end of the DOS header.
    For i = e_lfanew To SIZE_DOS_HEADER Step -1
        Get f, i, sig4
        If sig4 = "Rich" Then
            Rich = i
            Exit For
        End If
    Next
        
    If Rich = 0 Then
        ErrorMsg = "Rich signature not found. This file probably has no Rich header."
        GoTo cleanup
    End If
        
    'We found a valid 'Rich' signature in the header
    exists = True
    Get f, Rich + 4, checkSum
    'Debug.Print "Checksum is 0x" & Hex(checkSum)
    
    'xor backwards with csum until either 'DanS' or end of the DOS header,
    'inverse the list to get original order
    For i = (Rich - 4) To SIZE_DOS_HEADER Step -4
        Get f, i, tmp
        tmp = tmp Xor checkSum
        push uPack, tmp
        If tmp = &H536E6144 Then Exit For 'DanS marker
    Next
    
    If i <= SIZE_DOS_HEADER Then
        ErrorMsg = "DanS signature not found. Rich header corrupt."
        GoTo cleanup
    End If

    DanS = i
    'Debug.Print "Dans file offset: " & DanS - 1
    
    uPack = AryReverse(uPack)
    'Open "C:\Users\home\Desktop\New folder\x.dat" For Binary As 11
    'For i = 0 To UBound(upack)
    '    Put 11, , upack(i)
    'Next
    'Close 11
    
    'DanS is _always_ followed by three zero dwords
    For i = 1 To 3
        If uPack(i) <> 0 Then
            ErrorMsg = "DanS not followed by 0 @ " & i & " = " & Hex(uPack(i))
            GoTo cleanup
        End If
    Next
 
    '------------- start calculate checksum -------------------------------
    'get a copy of the entire MZ header + rich header
    '(+1 for file offset, +1 for 0 based array) now at end of Rich
    ReDim b(Rich + 2)
    Get f, 1, b() 'get the entire DOS stub + encrypted rich header to end of Rich (we only use up to DanS)
    
    'Mask out the e_lfanew field as it's not initialized at checksum calculation time
    i = 0 '&HAAAAAAAA
    CopyMemory ByVal VarPtr(b(POS_E_LFANEW)), i, 4
    'Debug.Print HexDump(b)

    'Calculate correct checksum, first Dos stub with elfanew 0'ed - end at DanS start
    calcChecksum = DanS - 1 '-1 because of 1 based vb file offsets
    For i = 0 To DanS - 2
        calcChecksum = add32(calcChecksum, rol32(CLng(b(i)), CInt(i)))
        'Debug.Print Join(Array(i, Hex(b(i)), Hex(calcChecksum)), " ")
    Next
    
    'now add checksum for the clear_data fields ignoring the DanS marker and the alignment 0 0 0 fields..
    For i = 4 To UBound(uPack) Step 2
    
        calcChecksum = add32(calcChecksum, rol32(uPack(i), uPack(i + 1)))
        'Debug.Print Join(Array(i + dans - 1, Hex(calcChecksum), Hex(upack(i)), Hex(upack(i + 1))), " ")
        
        'parse the entries and add to global collection
        Set tool = New CToolId
        tool.LoadSelf uPack(i), uPack(i + 1)
        entries.Add tool
        'Debug.Print tool.dump()
        
    Next
    '------------- end calculate checksum -------------------------------
    
    size = (UBound(uPack) + 1) * 4
    ReDim b(size - 1)
    CopyMemory ByVal VarPtr(b(0)), ByVal VarPtr(uPack(0)), size
    'Debug.Print HexDump(b)
    clearData = b()
    
    'to 0 based offsets
    Rich = Rich - 1
    DanS = DanS - 1
    
    Load = True

cleanup:
    If Err.Number <> 0 Then ErrorMsg = ErrorMsg & " - err: " & Err.Description
    On Error Resume Next
    Close f
    
End Function



Private Function toBytes(s) As Byte()
    Dim b() As Byte, tmp() As String, i As Long
    tmp = Split(s, "\x")
    ReDim b(UBound(tmp))
    For i = 1 To UBound(tmp)
        If Len(tmp(i)) > 0 Then
            b(i) = CByte(CInt("&h" & tmp(i)))
        End If
    Next
    toBytes = b()
End Function

Private Function rol32(base As Long, bits As Long)
    rol32 = CallWindowProc(VarPtr(asm(1)), 0, base, bits, 0)
End Function

Private Function add32(v1 As Long, v2 As Long) 'no overflow, allows wrap
    add32 = CallWindowProc(VarPtr(asm2(1)), 0, v1, v2, 0)
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

Private Function AryIsEmpty(ary) As Boolean
  On Error GoTo oops
    Dim i As Long
    i = UBound(ary)  '<- throws error if not initalized
    AryIsEmpty = False
  Exit Function
oops: AryIsEmpty = True
End Function

Private Function AryReverse(ary)
    Dim tmp, i, al() As Long, ai() As Long, ass() As String, av() As Variant, ab() As Byte
    
    If Not IsArray(ary) Then Exit Function
    
    If TypeName(ary) = "Long()" Then
        tmp = al
    ElseIf TypeName(ary) = "Integer()" Then
        tmp = ai
    ElseIf TypeName(ary) = "String()" Then
        tmp = ass
    ElseIf TypeName(ary) = "Variant()" Then
        tmp = av
    ElseIf TypeName(ary) = "Byte()" Then
        tmp = ab
    Else
        MsgBox "Add support for: " & TypeName(ary)
    End If
        
    If AryIsEmpty(ary) Then
        AryReverse = tmp
        Exit Function
    End If
    
    For i = UBound(ary) To LBound(ary) Step -1
        'Debug.Print i & " " & Hex(ary(i))
        push tmp, ary(i)
    Next
    
    AryReverse = tmp
End Function

Private Function FileExists(path) As Boolean
  On Error GoTo hell
    
  If Len(path) = 0 Then Exit Function
  If Right(path, 1) = "\" Then Exit Function
  If Dir(path, vbHidden Or vbNormal Or vbReadOnly Or vbSystem) <> "" Then FileExists = True
  
  Exit Function
hell: FileExists = False
End Function


'Function test_longs_to_bytes()
'
'    Dim uPack() As Long, size As Long, b() As Byte
'
'    ReDim uPack(1)
'    uPack(0) = &H11223344
'    uPack(1) = &H55667788
'
'    size = (UBound(uPack) + 1) * 4
'    ReDim b(size - 1)
'
'    CopyMemory ByVal VarPtr(b(0)), ByVal VarPtr(uPack(0)), size
'    Debug.Print HexDump(b) 'dzrt reference
'
'
'End Function


