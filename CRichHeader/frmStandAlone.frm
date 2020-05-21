VERSION 5.00
Begin VB.Form frmStandAlone 
   Caption         =   "Form1"
   ClientHeight    =   3135
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3135
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "frmStandAlone"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'this standalone form was used for debugging and testing first...before CRichHeader

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

Function toBytes(s) As Byte()
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

Function rol32(base As Long, bits As Long)
    rol32 = CallWindowProc(VarPtr(asm(1)), 0, base, bits, 0)
End Function

Function add32(v1 As Long, v2 As Long) 'no overflow, allows wrap
    add32 = CallWindowProc(VarPtr(asm2(1)), 0, v1, v2, 0)
End Function


Sub push(ary, value) 'this modifies parent ary object
    On Error GoTo init
    Dim x
    x = UBound(ary) '<-throws Error If Not initalized
    ReDim Preserve ary(UBound(ary) + 1)
    ary(UBound(ary)) = value
    Exit Sub
init:     ReDim ary(0): ary(0) = value
End Sub

Function AryIsEmpty(ary) As Boolean
  On Error GoTo oops
    Dim i As Long
    i = UBound(ary)  '<- throws error if not initalized
    AryIsEmpty = False
  Exit Function
oops: AryIsEmpty = True
End Function

Function AryReverse(ary)
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


Private Sub Form_Load()
    
    asm() = toBytes(asm_rol32)
    asm2() = toBytes(asm_add32)
    
    'MsgBox Hex(rol32(1, 1))
    'End
    
   Dim pth As String, f As Long, sig As String * 2, e_lfanew As Long, sig4 As String * 4
   Dim i As Long, Rich As Long, checkSum As Long
   Dim b() As Byte
   
   Const SIZE_DOS_HEADER = &H40
   Const POS_E_LFANEW = &H3C

   pth = "D:\_code\libs\pe_lib2\_sppe2.dll"
   
   f = FreeFile
   Open pth For Binary Access Read As f
   
   Get f, , sig
   If sig <> "MZ" Then
        Debug.Print "MZ not found"
        GoTo cleanup
   End If
   
   Get f, POS_E_LFANEW + 1, e_lfanew
   Get f, e_lfanew + 1, sig
   If sig <> "PE" Then
        Debug.Print "PE not found"
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
        Debug.Print "Rich signature not found. This file probably has no Rich header."
        GoTo cleanup
    End If
    
    'get a copy of the entire MZ header + rich header
    '(+1 for file offset, +1 for 0 based array) now at end of Rich
    ReDim b(Rich + 2)
    Get f, 1, b() 'get the entire DOS stub + encrypted rich header to end of Rich
    
    '## Mask out the e_lfanew field as it's not initialized at checksum calculation time
    i = 0 '&HAAAAAAAA
    CopyMemory ByVal VarPtr(b(POS_E_LFANEW)), i, 4
    'Debug.Print HexDump(b)
    
    '## We found a valid 'Rich' signature in the header from here on
    Get f, Rich + 4, checkSum
    Debug.Print "Checksum is 0x" & Hex(checkSum)
    
    '## xor backwards with csum until either 'DanS' or end of the DOS header,
    '## inverse the list to get original order
    'upack = [ u32(dat[i:][:4]) ^ csum for i in range(rich - 4, SIZE_DOS_HEADER, -4) ][::-1]
    'if u32(b'DanS') not in upack:
    '    return {'err': -7}
    Dim tmp As Long, upack() As Long, DanS As Long
    For i = (Rich - 4) To SIZE_DOS_HEADER Step -4
        Get f, i, tmp
        tmp = tmp Xor checkSum
        push upack, tmp
        If tmp = &H536E6144 Then 'DanS
            Exit For
        End If
    Next
    
    If i = SIZE_DOS_HEADER Then
        Debug.Print "DanS signature not found. Rich header corrupt."
        GoTo cleanup
    End If

    DanS = i
    Debug.Print "Dans file offset: " & DanS - 1
    
    upack = AryReverse(upack)
    'Open "C:\Users\home\Desktop\New folder\x.dat" For Binary As 11
    'For i = 0 To UBound(upack)
    '    Put 11, , upack(i)
    'Next
    'Close 11
    
    '## DanS is _always_ followed by three zero dwords
    For i = 1 To 3
        If upack(i) <> 0 Then
            Debug.Print "DanS not followed by 0 @ " & i & " = " & Hex(upack(i))
            GoTo cleanup
        End If
    Next
    
    'copy the rich.clear_data back over buffer
    'CopyMemory ByVal VarPtr(b(dans + 1)), ByVal VarPtr(upack(0)), ((UBound(upack) + 1) * 4)
    'Debug.Print HexDump(b)

    '## Bonus feature: Calculate and check the checksum csum
    Dim calcChecksum As Long 'New ULong
    calcChecksum = DanS - 1
    For i = 0 To DanS - 2
        'calcChecksum = calcChecksum.Add(rol32(CLng(b(i)), CInt(i)))
        calcChecksum = add32(calcChecksum, rol32(CLng(b(i)), CInt(i)))
        'Debug.Print Join(Array(i, Hex(b(i)), Hex(calcChecksum)), " ")
    Next
    
    Dim tools As New Collection, tool As CToolId
    
    For i = 4 To UBound(upack) Step 2 'ignore the DanS marker and the 0 0 0 fields..
    
        'calcChecksum = calcChecksum.Add(rol32(upack(i), upack(i + 1)))
        calcChecksum = add32(calcChecksum, rol32(upack(i), upack(i + 1)))
        'Debug.Print Join(Array(i + dans - 1, Hex(calcChecksum), Hex(upack(i)), Hex(upack(i + 1))), " ")
        
        Set tool = New CToolId
        tool.LoadSelf upack(i), upack(i + 1)
        Debug.Print tool.dump()
        tools.Add tool
        
    Next
    
    If calcChecksum = checkSum Then
        Debug.Print "Checksums match!"
    Else
        Debug.Print "Checksum Corrupted!"
    End If
    

    
    DoEvents
    
        
        
        
        
        
        

cleanup:
    Close f
End Sub
