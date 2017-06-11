VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CResEntry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'private to dll clients dont use it..

Public Name As String
Public id As Long
Public DataOffset As Long
Public SubDirOffset As Long
Public isSubDir As Boolean
Public subDir As CResDir
Public path As String

Friend Sub setName(hFile As Long, base As Long, NameOffset As Long)
    
    Dim no As Long, i As Long
    Dim sz As Integer
    Dim dbg
    Dim b() As Byte
    
    no = NameOffset And &H7FFFFFFF
    no = base + no
    dbg = Hex(no)
    
    Get hFile, no, sz 'two byte size of unicode string..
    ReDim b(1 To (sz * 2))
    Get hFile, , b()
    
    For i = 1 To UBound(b)
        If b(i) <> 0 Then Name = Name & Chr(b(i))
    Next
    
End Sub

Friend Sub setData(data As Long)
    
    If (data And &H80000000) = &H80000000 Then
        isSubDir = True
        SubDirOffset = data And &H7FFFFFFF
        Set subDir = New CResDir
    Else
        DataOffset = data
    End If
    
End Sub



