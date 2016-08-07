Attribute VB_Name = "modAsm"
'calls raw byte buffers in VB for some math operations vb does not support native
'todo: make dep safe or replace...

Private Declare Function CallAsm Lib "user32" Alias "CallWindowProcA" (ByRef lpBytes As Any, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Function RelocCalc(CurValue As Long, CurBase As Long, newBase As Long) As Long
    '8B45 0C          MOV EAX,DWORD PTR SS:[EBP+C]    arg1
    '2B45 10          SUB EAX,DWORD PTR SS:[EBP+10]   arg2
    '8B4D 14          MOV ECX,DWORD PTR SS:[EBP+14]   arg3
    '2BC8             SUB ECX,EAX
    '8BC1             MOV EAX,ECX
    'C2 1000          RETN 10
    Dim o() As Byte
    Const sl As String = "8B 45 0C 2B 45 10 8B 4D 14 2b C8 8b C1 C2 10 00"
    o() = toBytes(sl)
    RelocCalc = CallAsm(o(0), CurBase, newBase, CurValue, 0)
End Function

Function ShlX(x As Long, shift As Byte) As Long
    '8B45 0C        MOV EAX,DWORD PTR SS:[EBP+12]
    'C1E0 12        SHL EAX,12
    'C2 10 00       RETN 10h
    Dim o() As Byte
    Const sl As String = "8B 45 0C C1 E0 __ C2 10 00"
    o() = toBytes(Replace(sl, "__", Hex(shift)))
    ShlX = CallAsm(o(0), x, 0, 0, 0)
End Function

Function ShrX(x As Long, shift As Byte) As Long
    '8B45 0C        MOV EAX,DWORD PTR SS:[EBP+12]
    'C1E8 12        SHR EAX,12
    'C2 10 00       RETN 10h
    Dim o() As Byte
    Const sr As String = "8B 45 0C C1 E8 __ C2 10 00"
    o() = toBytes(Replace(sr, "__", Hex(shift)))
    ShrX = CallAsm(o(0), x, 0, 0, 0)
End Function

Function Shl(x As Long) As Long
    '8B45 0C        MOV EAX,DWORD PTR SS:[EBP+12]
    'D1E0           SHL EAX,1
    'C2 10 00       RETN 10h
    Dim o() As Byte
    Const sl As String = "8B 45 0C D1 E0 C2 10 00"
    o() = toBytes(sl)
    Shl = CallAsm(o(0), x, 0, 0, 0)
End Function

Function Shr(x As Long) As Long
    '8B45 0C        MOV EAX,DWORD PTR SS:[EBP+12]
    'D1E8           SHR EAX,1
    'C2 10 00       RETN 10h
    Dim o() As Byte
    Const sr As String = "8B 45 0C D1 E8 C2 10 00"
    o() = toBytes(sr)
    Shr = CallAsm(o(0), x, 0, 0, 0)
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



