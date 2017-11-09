Attribute VB_Name = "Module1"
Private Declare Function lstrlen Lib "kernel32.dll" Alias "lstrlenA" (ByVal lpString As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

Public Function StrFromPtr(ByVal lpStr As Long) As String
    Dim b() As Byte, x As Long
    On Error Resume Next
    x = lstrlen(lpStr)
    'form1.dbg "size: " & x
    If x Then
        ReDim b(0 To x - 1)
        Call CopyMemory(b(0), ByVal lpStr, x)
        StrFromPtr = StrConv(b, vbUnicode)
    End If
End Function


Function myCallBack(ByVal lpString As Long, ByVal arg2 As Long) As Long
    
    Form1.dbg "In VBCallback! Received args: " & Hex(lpString) & " arg2=" & Hex(arg2)
    Form1.dbg "Extracted string: " & StrFromPtr(lpString)
    
    myCallBack = 1234
    
End Function




