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
Public value
Public name As String
Private m_size As Long

Enum ms_fieldType 'const value = size dont change..
    msf_byte = 1
    msf_int = 2
    msf_long = 4
    msf_cur = 8
    msf_blob = 10
End Enum

Function SetBlobValue(bAryOrStr, Optional errMsg) As Boolean
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
            value = tmp
        End If
    Else
        s = CStr(bAryOrStr)
        tmp() = StrConv(s, vbFromUnicode, &H409)
        If UBound(tmp) > m_size - 1 Then
            errMsg = "String to large can not fit into blob"
            Exit Function
        Else
            If UBound(tmp) < m_size - 1 Then ReDim Preserve tmp(m_size - 1)
            value = tmp
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
         tmp = StrConv(value, vbUnicode, &H409)
         If asHex Then
            asString = HexDump(tmp)
         Else
            asString = tmp
         End If
    ElseIf fieldType = msf_cur Then
        If asHex Then
            asString = CurToHex(value)
        Else
            asString = CStr(value)
        End If
    Else
        If asHex Then
            asString = Hex(value)
        Else
            asString = CStr(value)
        End If
     End If
     
End Function



