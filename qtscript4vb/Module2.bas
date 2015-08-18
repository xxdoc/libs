Attribute VB_Name = "modCOM"

'this is used for script to host app object integration..
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef Destination As Any, Source As Any, ByVal length As Long)
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Declare Function lstrlen Lib "kernel32.dll" Alias "lstrlenA" (ByVal lpString As Long) As Long
Public Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Public Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long


Enum op
    op_reset = 0
    op_setdbg = 1
    op_setTimeout = 2
    op_setResolverHandler = 3
    op_setRetValInt = 4
    op_setRetValStr = 5
    op_getVarFromCtx = 6
    op_qtShutdown = 7
End Enum
 
Public Declare Function QtOp Lib "dbg.dll" (ByVal operation As op, Optional ByVal v1 As Long, Optional ByVal v2 As Long, Optional ByVal v3 As Long) As Long
Public Declare Function AddFile Lib "dbg.dll" (ByVal fPath As String) As Boolean
Public Declare Function Eval Lib "dbg.dll" (ByVal code As String, ByRef outVal As Variant) As Boolean


Public Function HostResolver(ByVal buf As Long, ByVal ctx As Long, ByVal argCnt As Long, ByVal hInst As Long) As Long
    Dim key As String
    Dim v1 As Variant
    
    On Error Resume Next
    'we could switch to numeric ids..but it would be harder to manage/debug when more complex..
    key = StringFromPointer(buf)
    
    'this is just a quick demo not the full setup see duk4vb project for a full COM relay using same structure
    If key = "list1.additem" Then
        If argCnt > 1 Then
            If CBool(QtOp(op_getVarFromCtx, ctx, 1, VarPtr(v1))) Then
                Form1.List1.AddItem CStr(v1)
            Else
                Debug.Print "list1.additem.op_getVarFromCtx(" & ctx & ", 1) failed?"
            End If
        End If
    End If
    
    If key = "txtFile.text" Then
         SetRetStr Form1.txtFile.Text
         HostResolver = 1
    End If
            
End Function

Private Function SetRetStr(s As String)
    Dim b() As Byte
    b() = StrConv(s & Chr(0), vbFromUnicode)
    QtOp op_setRetValStr, VarPtr(b(0))
End Function

Private Function StringFromPointer(buf As Long) As String
    Dim sz As Long
    Dim tmp As String
    Dim b() As Byte
    
    If buf = 0 Then Exit Function
       
    sz = lstrlen(buf)
    If sz = 0 Then Exit Function
    
    ReDim b(sz)
    CopyMemory b(0), ByVal buf, sz
    tmp = StrConv(b, vbUnicode)
    If Right(tmp, 1) = Chr(0) Then tmp = Left(tmp, Len(tmp) - 1)
    
    StringFromPointer = tmp
 
End Function

