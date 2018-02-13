Attribute VB_Name = "modMath"
'dep safe way to call asm byte buffers in VB for some math operations that do not have support native

Private Declare Function CallAsm Lib "user32" Alias "CallWindowProcA" (ByRef lpBytes As Any, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function VirtualAlloc Lib "kernel32" (ByVal lpAddress As Long, ByVal dwSize As Long, ByVal flAllocationType As Long, ByVal flProtect As Long) As Long
Private Declare Function VirtualFree Lib "kernel32" (ByVal lpAddress As Long, ByVal dwSize As Long, ByVal dwFreeType As Long) As Long
Private Declare Function VirtualLock Lib "kernel32" (ByVal lpAddress As Long, ByVal dwSize As Long) As Long
Private Declare Function VirtualUnlock Lib "kernel32" (ByVal lpAddress As Long, ByVal dwSize As Long) As Long
Private Declare Sub RtlMoveMemory Lib "kernel32" (ByVal Destination As Long, ByVal Source As Long, ByVal length As Long)

Private Const PAGE_RWX      As Long = &H40
Private Const MEM_COMMIT    As Long = &H1000

Private base As Long
Private offset As Long
Private initilized As Boolean

Private Function buildThunk(asm As String) As Long
    
    Dim b() As Byte
    Dim nextOffset As Long
    
    If base = 0 Then base = VirtualAlloc(ByVal 0&, &H1000, MEM_COMMIT, PAGE_RWX)
    If base = 0 Then Exit Function
    
    b() = toBytes(asm)
    RtlMoveMemory base + offset, VarPtr(b(0)), UBound(b) + 1
    buildThunk = base + offset
    
    nextOffset = offset + UBound(b) + 1
    nextOffset = nextOffset + (nextOffset Mod 16)
    offset = nextOffset + 1
    
End Function

'unsigned math without overflow..cant do in vb6 naturally..
Function RelocCalc(CurValue As Long, CurBase As Long, newBase As Long) As Long
    '8B45 0C          MOV EAX,DWORD PTR SS:[EBP+C]    arg1
    '2B45 10          SUB EAX,DWORD PTR SS:[EBP+10]   arg2
    '8B4D 14          MOV ECX,DWORD PTR SS:[EBP+14]   arg3
    '2BC8             SUB ECX,EAX
    '8BC1             MOV EAX,ECX
    'C2 1000          RETN 10
    Static lpfnRelocCalc As Long
    Const asm As String = "8B 45 0C 2B 45 10 8B 4D 14 2b C8 8b C1 C2 10 00"
    
    If lpfnRelocCalc = 0 Then lpfnRelocCalc = buildThunk(asm)
    If lpfnRelocCalc = 0 Then Exit Function
    
    RelocCalc = CallAsm(ByVal oRelocCalc, CurBase, newBase, CurValue, 0)
    
End Function

Function ShlX(x As Long, shift As Byte) As Long
    '8B45 0C        MOV EAX,DWORD PTR SS:[EBP+12]
    'C1E0 12        SHL EAX,12
    'C2 10 00       RETN 10h
    Static lpfnShlX As Long
    Const asm As String = "8B 45 0C C1 E0 __ C2 10 00"
    
    If lpfnShlX = 0 Then lpfnShlX = buildThunk(Replace(asm, "__", Hex(shift)))
    If lpfnShlX = 0 Then Exit Function
    
    ShlX = CallAsm(ByVal lpfnShlX, x, 0, 0, 0)
    
End Function

Function ShrX(x As Long, shift As Byte) As Long
    '8B45 0C        MOV EAX,DWORD PTR SS:[EBP+12]
    'C1E8 12        SHR EAX,12
    'C2 10 00       RETN 10h
    
    Static pShrX As Long
    Const asm As String = "8B 45 0C C1 E8 __ C2 10 00"
    
    If pShrX = 0 Then pShrX = buildThunk(Replace(asm, "__", Hex(shift)))
    If pShrX = 0 Then Exit Function
    
    ShrX = CallAsm(ByVal pShrX, x, 0, 0, 0)
End Function

Function Shl(x As Long) As Long
    '8B45 0C        MOV EAX,DWORD PTR SS:[EBP+12]
    'D1E0           SHL EAX,1
    'C2 10 00       RETN 10h
    Static pShl As Long
    Const asm As String = "8B 45 0C D1 E0 C2 10 00"

    If pShl = 0 Then pShl = buildThunk(asm)
    If pShl = 0 Then Exit Function
    
    Shl = CallAsm(ByVal pShl, x, 0, 0, 0)
End Function

Function Shr(x As Long) As Long
    '8B45 0C        MOV EAX,DWORD PTR SS:[EBP+12]
    'D1E8           SHR EAX,1
    'C2 10 00       RETN 10h
    
    Static pShr As Long
    Const asm As String = "8B 45 0C D1 E8 C2 10 00"
   
    If pShr = 0 Then pShr = buildThunk(asm)
    If pShr = 0 Then Exit Function
    
    Shr = CallAsm(ByVal pShr, x, 0, 0, 0)
End Function

Private Function toBytes(x As String) As Byte()
    Dim tmp() As String
    Dim fx() As Byte
    Dim i As Long
    
    tmp = Split(x, " ")
    ReDim fx(UBound(tmp))
    
    For i = 0 To UBound(tmp)
        If Len(tmp(i)) = 1 Then tmp(i) = "0" & tmp(i)
        fx(i) = CInt("&h" & tmp(i))
    Next
    
    toBytes = fx()

End Function



