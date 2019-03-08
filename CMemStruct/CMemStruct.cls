VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CMemStruct"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public name As String      'in case instances get held in a collection or want to be named
Public filePath As String

Private m_size As Long
Private fHandle As Long
Private fields As New Collection

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSrc As Any, ByVal ByteLen As Long)

'todo: fromMemory/toMemory (just a copymemory call away with toBytes() fromBytes()
'      test more: multiple reads/writes to same open file using offset and only opening once with filepath arg in first call..


'this way consumers cant mess with our collection and do unexpected
'stuff like set new or add/remove manually and mess up cached m_size
'also this gives us intellisense when using..
Function field(fieldIndexOrName) As CStructField
    Set field = fields(fieldIndexOrName)
End Function

Property Get fieldCount() As Long
    fieldCount = fields.Count
End Property

Function CurOffset() As Long
    On Error Resume Next
    CurOffset = Seek(fHandle) - 1
    If Err.Number <> 0 Then CurOffset = -1
End Function

Property Get size() As Long
    size = m_size                'should we cache or always recalculate?, user must use call .reset rather than manually setting fields = new collection
    'Dim sf As CStructField      'I could add a recalculate size method..but not needed i dont think..
    'For Each sf In fields       'also we dont have to make fields public we could wrap in .count and .getField(nameOrIndex) as CStructField props..
    '    sizeOf = sizeOf + sf.size
    'Next
End Property

Sub reset()
    Set fields = New Collection
    m_size = 0
End Sub

Private Sub Class_Terminate()
    CloseFile
End Sub

Sub CloseFile()
    On Error Resume Next
    If fHandle <> 0 Then Close fHandle
    fHandle = 0
End Sub

Function offsetOf(fieldIndexOrName) As Long
    On Error Resume Next
    Dim sf As CStructField, ff As CStructField

    Set sf = fields(fieldIndexOrName)
    If Err.Number <> 0 Then
        offsetOf = -1
    Else
        For Each ff In fields
            If ObjPtr(sf) = ObjPtr(ff) Then Exit For
            offsetOf = offsetOf + ff.size
        Next
    End If
    
End Function

Function LoadFromFile(Optional ByVal offset As Long, Optional ByVal filePath As String)
        
    If fields.Count = 0 Then Exit Function
    
    If Len(filePath) > 0 Then
        If fHandle <> 0 Then Close fHandle
        fHandle = FreeFile
        Me.filePath = filePath
        Open filePath For Binary As fHandle
    End If
    
    If offset <> 0 Then Seek fHandle, offset + 1
    
    Dim b() As Byte
    ReDim b(m_size - 1)
    Get fHandle, , b()
    fromBytes b
    
End Function

Function SaveToFile(Optional ByVal offset As Long, Optional ByVal filePath As String)
        
    If fields.Count = 0 Then Exit Function
    
    If Len(filePath) > 0 Then
        If fHandle <> 0 Then Close fHandle
        fHandle = FreeFile
        Me.filePath = filePath
        Open filePath For Binary As fHandle
    End If
    
    If offset <> 0 Then Seek fHandle, offset + 1
    
    Dim b() As Byte
    b() = toBytes()
    Put fHandle, , b()
    
End Function

Function toHexString() As String
    Dim b() As Byte
    Dim i As Long
    Dim x As String
    
    b() = Me.toBytes()
    For i = 0 To UBound(b)
        x = x & Right("00" & Hex(b(i)), 2)
    Next
    
    toHexString = x
    
End Function

Function fromHexString(str, Optional errMsg) As Boolean
    
    Dim b() As Byte
    Dim i As Long, j As Long
    Dim x As String
    
    On Error GoTo hell
        
    errMsg = Empty
    If Len(str) Mod 2 <> 0 Then
        errMsg = "String length must be divisible by 2"
        Exit Function
    End If
    
    ReDim b((Len(str) / 2) - 1)
    
    For i = 1 To Len(str) Step 2
        x = Mid(str, i, 2)
        b(j) = CByte(CInt("&h" & x))
        j = j + 1
    Next
    
    fromBytes b
    fromHexString = True
    
    Exit Function
hell: errMsg = "CharPos: " & i & " " & Err.Description
End Function

Function toBytes() As Byte()

    Dim buf() As Byte
    Dim b As Byte
    Dim i As Integer
    Dim l As Long
    Dim c As Currency
    Dim bb() As Byte
    Dim offset As Long
    Dim sf As CStructField
    
    ReDim buf(Me.size() - 1)
    
    For Each sf In fields
        Select Case sf.fieldType
            Case msf_byte
                b = sf.Value
                CopyMemory ByVal VarPtr(buf(offset)), ByVal VarPtr(b), 1
                offset = offset + 1
            Case msf_int:
                i = sf.Value
                CopyMemory ByVal VarPtr(buf(offset)), ByVal VarPtr(i), 2
                offset = offset + 2
            Case msf_long:
                l = sf.Value
                CopyMemory ByVal VarPtr(buf(offset)), ByVal VarPtr(l), 4
                offset = offset + 4
            Case msf_cur:
                c = sf.Value
                CopyMemory ByVal VarPtr(buf(offset)), ByVal VarPtr(c), 8
                offset = offset + 8
            Case msf_blob:
                bb() = sf.Value
                CopyMemory ByVal VarPtr(buf(offset)), ByVal VarPtr(bb(0)), sf.size
                offset = offset + sf.size
        End Select
    Next
    
    toBytes = buf()
    
End Function

Function fromBytes(buf() As Byte) As Boolean

    Dim b As Byte
    Dim i As Integer
    Dim l As Long
    Dim c As Currency
    Dim bb() As Byte
    Dim offset As Long
    Dim sf As CStructField
    
    If UBound(buf) < (m_size - 1) Then
        Exit Function
    End If
    
    For Each sf In fields
        Select Case sf.fieldType
            Case msf_byte
                CopyMemory ByVal VarPtr(b), ByVal VarPtr(buf(offset)), 1
                offset = offset + 1
                sf.SetValue b
            Case msf_int:
                CopyMemory ByVal VarPtr(i), ByVal VarPtr(buf(offset)), 2
                sf.SetValue i
                offset = offset + 2
            Case msf_long:
                CopyMemory ByVal VarPtr(l), ByVal VarPtr(buf(offset)), 4
                sf.SetValue l
                offset = offset + 4
            Case msf_cur:
                CopyMemory ByVal VarPtr(c), ByVal VarPtr(buf(offset)), 8
                sf.SetValue c
                offset = offset + 8
            Case msf_blob:
                ReDim bb(sf.size - 1)
                CopyMemory ByVal VarPtr(bb(0)), ByVal VarPtr(buf(offset)), sf.size
                sf.SetValue bb()
                offset = offset + sf.size
        End Select
    Next

    fromBytes = True
    
End Function




Function dump(Optional asHex As Boolean = True) As String
    Dim t() As String, n As String, hi As Long, lo As Long
    Dim sf As CStructField
    
    For Each sf In fields
       n = sf.name & " = "
       If sf.fieldType = msf_blob Then
            n = n & "Byte[0x" & sf.size & "]"
            push t, n & vbCrLf & sf.asString(asHex)
       Else
            push t, n & sf.asString(asHex)
       End If
    Next
    
    dump = Join(t, vbCrLf)
    
End Function


Sub AddField(ByVal name As String, ft As ms_fieldType, Optional size As Long = 0)
    Dim sf As New CStructField
    sf.init ft, name, size
    fields.Add sf, name
    m_size = m_size + sf.size
End Sub

'format: name1*type_identifier,name2*type_identifier
'format specifiers (case insensitive)
'   b = byte, l = long, i = int, c = currency, numeric val = blob size

Function AddFields(ByVal formatString As String, Optional errMsg, Optional reset As Boolean = True, Optional ByVal name As String) As Boolean
    Dim f() As String, ff, t() As String, size As Long
    Dim ft As ms_fieldType, sf As CStructField
    
    On Error Resume Next
    
    Me.name = name
    errMsg = Empty
    f = Split(formatString, ",")
    
    If reset Then
        m_size = 0
        Set fields = New Collection 'if an addfields failed before we need to do a clean reset..
    End If
    
    For Each ff In f
        'If Len(ff) = 0 Then Exit For
        t() = Split(ff, "*")
        
        If Len(t(0)) = 0 Then
            errMsg = "No field name specified" 'for index?
            Exit Function
        End If
        
        If UBound(t) = 0 Then
            errMsg = "field " & ff & " no type"
            Exit Function
        End If
        
        If IsNumeric(t(1)) Then
            ft = msf_blob
            size = CLng(t(1))
            If size = 0 Then
                errMsg = "Blob type can not have size 0"
                Exit Function
            End If
        Else
            Select Case LCase(t(1))
                Case "l": ft = msf_long
                Case "i": ft = msf_int
                Case "b": ft = msf_byte
                Case "c": ft = msf_cur
            End Select
        End If
        
        If ft = 0 Then
            errMsg = "Unknown type '" & t(1) & "' for field " & t(0)
            Exit Function
        End If
        
        Set sf = New CStructField
        sf.init ft, t(0), size
        m_size = m_size + sf.size
        fields.Add sf, t(0)
        
        If Err.Number <> 0 Then 'todo get exact err number..
            errMsg = "Duplicate field name? '" & t(0) & "'"
            Exit Function
        End If
    
    Next
        
        AddFields = True
        
End Function


