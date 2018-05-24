VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CStructField"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public fieldType As ms_fieldType
Private m_Value                      'note: m_Value is a variant, its vartype is explicitly set when loaded,  bit tests will fail if you reset it to the wrong type..
Public name As String
Private m_size As Long

Enum ms_fieldType 'const m_Value = size dont change..
    msf_byte = 1
    msf_int = 2
    msf_long = 4
    msf_cur = 8
    msf_blob = 10
End Enum

Property Get Value()
    Value = m_Value
End Property

'we need be careful end user does not inadvertantly change the variant value type must match field type so we control it..
Function SetValue(v, Optional errMsg) As Boolean
    On Error GoTo hell
    
    errMsg = Empty
    If fieldType = msf_byte Then v = CByte(v)
    If fieldType = msf_int Then v = CInt(v)
    If fieldType = msf_long Then v = CLng(v)
    If fieldType = msf_cur Then v = CCur(v)
    
    If fieldType = msf_blob Then
        SetValue = SetBlobValue(v, errMsg)
        If Len(errMsg) > 0 Then Exit Function
    Else
        m_Value = v
    End If
    
    SetValue = True
    
    Exit Function
hell:
    errMsg = Err.Description

End Function


Private Function SetBlobValue(bAryOrStr, Optional errMsg) As Boolean
    Dim tmp() As Byte
    Dim s As String
    
    On Error GoTo hell
    
    If TypeName(bAryOrStr) = "Byte()" Then
        If UBound(bAryOrStr) > m_size - 1 Then
            errMsg = "Byte array to large can not fit into blob"
            Exit Function
        Else
            tmp() = bAryOrStr
            If UBound(bAryOrStr) < m_size - 1 Then ReDim Preserve tmp(m_size - 1)
            m_Value = tmp
        End If
    Else
        s = CStr(bAryOrStr)
        tmp() = StrConv(s, vbFromUnicode, &H409)
        If UBound(tmp) > m_size - 1 Then
            errMsg = "String to large can not fit into blob"
            Exit Function
        Else
            If UBound(tmp) < m_size - 1 Then ReDim Preserve tmp(m_size - 1)
            m_Value = tmp
        End If
    End If
    
    Exit Function
hell: errMsg = Err.Description
    
End Function

Property Get size() As Long
    size = m_size
End Property

Friend Sub init(ft As ms_fieldType, name As String, Optional size As Long = 0)
    
    Me.fieldType = ft
    Me.name = name
    
    If ft = msf_blob Then
        If size = 0 Then Err.Raise 1, "CStructField.init", "blob type must specify size"
        m_size = size
    Else
        m_size = ft
    End If
    
End Sub

Function asString(Optional asHex As Boolean) As String
    
    Dim tmp As String
    
    If fieldType = msf_blob Then
         tmp = StrConv(m_Value, vbUnicode, &H409)
         If asHex Then
            asString = HexDump(tmp)
         Else
            asString = tmp
         End If
    ElseIf fieldType = msf_cur Then
        If asHex Then
            asString = CurToHex(m_Value)
        Else
            asString = CStr(m_Value)
        End If
    Else
        If asHex Then
            asString = Hex(m_Value)
        Else
            asString = CStr(m_Value)
        End If
     End If
     
End Function

Private Function validateBitSizeForVar(bit, source)
     If fieldType = msf_byte And bit >= 8 Then Err.Raise 1, source, "Byte m_Values can test bits 0=7"
     If fieldType = msf_int And bit >= 16 Then Err.Raise 1, source, "Int m_Values can test bits 0=15"
     If fieldType = msf_long And bit >= 32 Then Err.Raise 1, source, "Long m_Values can test bits 0=31"
     If fieldType = msf_cur Or fieldType = msf_blob Then Err.Raise 1, source, "Can not bit test currency or blob types"
End Function

Function isBitSet(bit) As Boolean
     validateBitSizeForVar bit, "isBitSet"
     isBitSet = (m_Value And bitFlags(bit)) <> 0
End Function

Function setBit(bit) As CStructField
    validateBitSizeForVar bit, "setBit"
    'm_Value = m_Value Or bitFlags(bit) 'will change m_Value to long
    If fieldType = msf_byte Then m_Value = CByte(m_Value Or bitFlags(bit))
    If fieldType = msf_int Then m_Value = CInt(m_Value Or bitFlags(bit))
    If fieldType = msf_long Then m_Value = CLng(m_Value Or bitFlags(bit))
    Set setBit = Me
End Function

Function clearBit(bit) As CStructField
    validateBitSizeForVar bit, "clearBit"
    'm_Value = m_Value And (Not bitFlags(bit)) 'will change m_Value to long
    If fieldType = msf_byte Then m_Value = CByte(m_Value And (Not bitFlags(bit)))
    If fieldType = msf_int Then m_Value = CInt(m_Value And (Not bitFlags(bit)))
    If fieldType = msf_long Then m_Value = CLng(m_Value And (Not bitFlags(bit)))
    Set clearBit = Me
End Function

Function toBinary() As String
    
    Dim x
    Dim bits As Long
    Dim bytes As Long
    Dim l As Long
    Dim b As Byte
    Dim i As Integer
    Dim v As Long
    
    If TypeName(m_Value) = "Long" Then bytes = 4
    If TypeName(m_Value) = "Integer" Then bytes = 2
    If TypeName(m_Value) = "Byte" Then bytes = 1
    
    If bytes = 0 Then
        toBinary = TypeName(m_Value) & " unsupported variable type"
        Exit Function
    End If
    
    v = CLng(m_Value)
    bits = (bytes * 8) - 1
    
    For i = bits To 0 Step -1
        If i > 0 And i Mod 8 = 0 Then x = x & " "
        x = x & IIf((v And bitFlags(i)) <> 0, "1", "0")
    Next
    
    toBinary = x
    
End Function



Private Sub Class_Initialize()
    Dim i As Long
    
    'there will be lots of instances of this class..we could hold this ary internally...
    'its not that sloppy and would be self contained..i will decide latter.
    
    If bitFlags(1) = 0 Then
        For i = 0 To UBound(bitFlags) - 1
            bitFlags(i) = 2 ^ i
        Next
        bitFlags(31) = &H80000000 'vb would consider this an overflow if calc as above
    End If
    
End Sub
